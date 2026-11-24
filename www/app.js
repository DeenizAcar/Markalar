/* Markalar - tek dosyalık uygulama mantığı */
(function () {
  'use strict';

  var K_MARKA = 'mk_marka', K_SILINEN = 'mk_silinen';

  var $ = function (id) { return document.getElementById(id); };
  var Cap = window.Capacitor;
  var isNative = !!(Cap && Cap.isNativePlatform && Cap.isNativePlatform());
  var plug = function (n) { return Cap && Cap.Plugins ? Cap.Plugins[n] : null; };

  /* ---------- Depolama ---------- */
  function oku(k, d) { try { var v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
  function yaz(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { toast('Kayıt yapılamadı!'); } }

  var markalar = oku(K_MARKA, []);    // [{id, ad, pasif}]
  var silinen = oku(K_SILINEN, []);   // [{id, zaman, ad, pasif}]
  function kaydetMarka() { yaz(K_MARKA, markalar); }
  function kaydetSilinen() { yaz(K_SILINEN, silinen); }

  /* ---------- Yardımcılar ---------- */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function gunAnahtar(d) { d = d || new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function tarihYaz(d) { return pad(d.getDate()) + '.' + pad(d.getMonth() + 1) + '.' + d.getFullYear(); }
  function saatYaz(d) { return pad(d.getHours()) + ':' + pad(d.getMinutes()); }
  function trKucuk(s) { return (s || '').toLocaleLowerCase('tr-TR'); }
  function adSirala(a, b) { return a.ad.localeCompare(b.ad, 'tr', { sensitivity: 'base' }); }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  var toastT;
  function toast(m) {
    var t = $('toast'); t.textContent = m; t.classList.remove('hidden');
    clearTimeout(toastT); toastT = setTimeout(function () { t.classList.add('hidden'); }, 2500);
  }

  /* ---------- Ekranlar ---------- */
  var aktifEkran = 'scrTanim';
  function ekranGoster(id) {
    document.querySelectorAll('.screen').forEach(function (s) { s.classList.toggle('active', s.id === id); });
    aktifEkran = id;
    menuKapat();
    if (id === 'scrTanim') renderTanim();
    if (id === 'scrSilinen') renderSilinen();
  }

  /* İsimleri sahaya sığdır: taşarsa yazı boyutunu küçült */
  function sigdir(kok) {
    requestAnimationFrame(function () {
      (kok || document).querySelectorAll('.fit').forEach(function (el) {
        el.style.fontSize = '';
        var fs = parseFloat(getComputedStyle(el).fontSize) || 16;
        while (fs > 9 && el.scrollWidth > el.clientWidth + 0.5) { fs -= 0.5; el.style.fontSize = fs + 'px'; }
      });
    });
  }
  var sigT;
  window.addEventListener('resize', function () { clearTimeout(sigT); sigT = setTimeout(function () { sigdir(); }, 150); });

  /* ---------- Paylaşım / indirme ---------- */
  function paylas(dataUrl, dosyaAdi, baslik, indirMesaji) {
    var Fs = plug('Filesystem'), Sh = plug('Share');
    if (isNative && Fs && Sh) {
      return Fs.writeFile({ path: dosyaAdi, data: dataUrl.split(',')[1], directory: 'CACHE' })
        .then(function (r) { return Sh.share({ title: baslik, files: [r.uri], dialogTitle: baslik }); })
        .catch(function (e) { if (!/cancel/i.test(String(e && e.message))) throw e; });
    }
    return fetch(dataUrl).then(function (r) { return r.blob(); }).then(function (blob) {
      var file = new File([blob], dosyaAdi, { type: blob.type });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        return navigator.share({ files: [file], title: baslik })
          .catch(function (e) { if (e && e.name !== 'AbortError') indir(dataUrl, dosyaAdi); });
      }
      indir(dataUrl, dosyaAdi);
      toast(indirMesaji);
    });
  }
  function indir(dataUrl, ad) { var a = document.createElement('a'); a.href = dataUrl; a.download = ad; document.body.appendChild(a); a.click(); a.remove(); }

  /* ---------- Excel Dışa / İçe Aktar ---------- */
  var xlsxSoz = null;
  function xlsxYukle() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxSoz) xlsxSoz = new Promise(function (ok, hata) {
      var s = document.createElement('script'); s.src = 'vendor/xlsx.full.min.js';
      s.onload = function () { ok(window.XLSX); };
      s.onerror = function () { xlsxSoz = null; hata(new Error('xlsx yüklenemedi')); };
      document.head.appendChild(s);
    });
    return xlsxSoz;
  }

  function disaAktar() {
    xlsxYukle().then(function (X) {
      var rows = markalar.slice().sort(adSirala).map(function (m) {
        return { 'Marka İsmi': m.ad, 'Durumu': m.pasif ? 'Pasif' : 'Aktif' };
      });
      var wb = X.utils.book_new();
      var ws = X.utils.json_to_sheet(rows, { header: ['Marka İsmi', 'Durumu'] });
      ws['!cols'] = [{ wch: 40 }, { wch: 10 }];
      X.utils.book_append_sheet(wb, ws, 'Markalar');
      var b64 = X.write(wb, { bookType: 'xlsx', type: 'base64' });
      var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      return paylas('data:' + mime + ';base64,' + b64, 'markalar_' + gunAnahtar() + '.xlsx', 'Markalar Excel', 'Excel dosyası indirildi.');
    }).catch(function (e) { console.error(e); toast('Dışa aktarılamadı.'); });
  }

  function hucre(row, adlar) {
    for (var i = 0; i < adlar.length; i++) {
      for (var k in row) if (trKucuk(k).replace(/\s+/g, ' ').trim() === adlar[i]) return row[k];
    }
    return undefined;
  }

  function iceAktar(dosya) {
    Promise.all([xlsxYukle(), dosya.arrayBuffer()]).then(function (r) {
      var X = r[0], wb = X.read(r[1], { type: 'array' });
      var sayfa = wb.Sheets['Markalar'] || wb.Sheets[wb.SheetNames[0]];
      var yeni = [];
      var ayniAd = function (liste, ad) { return liste.some(function (q) { return trKucuk(q.ad) === trKucuk(ad); }); };
      if (sayfa) X.utils.sheet_to_json(sayfa, { defval: '' }).forEach(function (row) {
        var ad = String(hucre(row, ['marka i̇smi', 'marka ismi', 'marka adı', 'marka adi', 'marka', 'isim', 'ad']) || row[Object.keys(row)[0]] || '')
          .replace(/\s+/g, ' ').trim();
        if (!ad || ayniAd(markalar, ad) || ayniAd(yeni, ad)) return;
        var pasif = trKucuk(String(hucre(row, ['durumu', 'durum']) || '')).trim() === 'pasif';
        yeni.push({ id: uid(), ad: ad, pasif: pasif });
      });
      if (!yeni.length) { toast('Dosyada eklenecek yeni marka bulunamadı.'); return; }
      onay('İçe Aktar', yeni.length + ' yeni marka eklensin mi? (Var olan kayıtlar silinmez)', function () {
        yeni.forEach(function (m) { markalar.push(m); });
        kaydetMarka(); renderTanim();
        toast(yeni.length + ' marka eklendi.');
      });
    }).catch(function (e) { console.error(e); toast('Dosya okunamadı. Excel (.xlsx) dosyası seçiniz.'); });
  }

  /* ---------- Satır etkileşimi: uzun basma + çift dokunma ---------- */
  function satirEtkilesim(tr, onLong, onDouble) {
    var timer = null, sx = 0, sy = 0, uzun = false, sonTik = 0;
    tr.addEventListener('pointerdown', function (e) {
      if (e.button && e.button !== 0) return;
      uzun = false; sx = e.clientX; sy = e.clientY;
      clearTimeout(timer);
      timer = setTimeout(function () { uzun = true; satirSec(tr); onLong(tr, sx, sy); if (navigator.vibrate) navigator.vibrate(30); }, 550);
    });
    tr.addEventListener('pointermove', function (e) { if (Math.abs(e.clientX - sx) > 10 || Math.abs(e.clientY - sy) > 10) clearTimeout(timer); });
    ['pointerup', 'pointercancel', 'pointerleave'].forEach(function (ev) { tr.addEventListener(ev, function () { clearTimeout(timer); }); });
    tr.addEventListener('contextmenu', function (e) { e.preventDefault(); });
    tr.addEventListener('click', function () {
      if (uzun) { uzun = false; return; }
      satirSec(tr);
      var t = Date.now();
      if (t - sonTik < 400) { sonTik = 0; onDouble(tr); } else sonTik = t;
    });
  }
  function satirSec(tr) {
    tr.parentNode.querySelectorAll('.selected').forEach(function (r) { r.classList.remove('selected'); });
    tr.classList.add('selected');
  }

  var menuSil = null, menuEk = null;
  // ek: { yazi, fn } verilirse menüde 2. sırada gösterilir
  function menuAc(x, y, silFn, ek) {
    var m = $('popMenu'); menuSil = silFn; menuEk = ek ? ek.fn : null;
    $('popEk').textContent = ek ? ek.yazi : '';
    $('popEk').classList.toggle('hidden', !ek);
    m.classList.remove('hidden');
    var w = m.offsetWidth, h = m.offsetHeight;
    m.style.left = Math.max(8, Math.min(x, window.innerWidth - w - 8)) + 'px';
    m.style.top = Math.max(8, Math.min(y, window.innerHeight - h - 8)) + 'px';
    $('overlay').classList.remove('hidden'); $('overlay').style.background = 'transparent';
  }
  function menuKapat() {
    $('popMenu').classList.add('hidden'); menuSil = null; menuEk = null;
    if (!acikDialog()) { $('overlay').classList.add('hidden'); }
    $('overlay').style.background = '';
  }
  $('popVazgec').addEventListener('click', menuKapat);
  $('popSil').addEventListener('click', function () { var f = menuSil; menuKapat(); if (f) f(); });
  $('popEk').addEventListener('click', function () { var f = menuEk; menuKapat(); if (f) f(); });

  /* ---------- Pencereler ---------- */
  function acikDialog() { return document.querySelector('.dialog:not(.hidden)'); }
  function dialogAc(id) { $('overlay').classList.remove('hidden'); $('overlay').style.background = ''; $(id).classList.remove('hidden'); }
  function dialogKapat(id) {
    $(id).classList.add('hidden');
    if (!acikDialog() && $('popMenu').classList.contains('hidden')) $('overlay').classList.add('hidden');
  }
  document.querySelectorAll('[data-close]').forEach(function (b) {
    b.addEventListener('click', function () { dialogKapat(b.closest('.dialog').id); });
  });
  $('overlay').addEventListener('click', function () {
    if (!$('popMenu').classList.contains('hidden')) { menuKapat(); return; }
  });

  var onayCb = null;
  function onay(baslik, mesaj, cb) {
    $('onayTitle').textContent = baslik; $('onayMsg').textContent = mesaj; onayCb = cb;
    dialogAc('dlgOnay');
  }
  $('onayEvet').addEventListener('click', function () { var f = onayCb; onayCb = null; dialogKapat('dlgOnay'); if (f) f(); });
  $('onayHayir').addEventListener('click', function () { onayCb = null; dialogKapat('dlgOnay'); });

  /* ---------- MARKA TANIMLARI ---------- */
  var araKelimeler = [];
  function aramaKelimeleri() {
    return araKelimeler.concat($('txtAra').value.split(/\s+/).filter(Boolean)).map(trKucuk);
  }
  function renderChips() {
    var box = $('araEtiketler'); box.innerHTML = '';
    araKelimeler.forEach(function (k, i) {
      var c = document.createElement('span'); c.className = 'chip';
      c.innerHTML = '<span>' + esc(k) + '</span><button class="xbtn" title="Kelimeyi sil">✕</button>';
      c.querySelector('button').addEventListener('click', function (e) {
        e.stopPropagation(); araKelimeler.splice(i, 1); renderTanim();
      });
      box.appendChild(c);
    });
    $('txtAra').placeholder = araKelimeler.length ? '' : 'Marka ara...';
  }
  function araGirdi() {
    var v = $('txtAra').value;
    if (/\s/.test(v)) {
      var parca = v.split(/\s+/), son = parca.pop();
      parca.filter(Boolean).forEach(function (k) { araKelimeler.push(k); });
      $('txtAra').value = son;
    }
    renderTanim();
  }

  function renderTanim() {
    renderChips();
    var kel = aramaKelimeleri();
    var liste = markalar.slice().sort(adSirala).filter(function (m) {
      if (!kel.length) return true;
      var ad = trKucuk(m.ad);
      return kel.some(function (k) { return ad.indexOf(k) >= 0; });
    });
    var tb = $('tanimBody'); tb.innerHTML = '';
    var bos = $('tanimEmpty');
    bos.classList.toggle('hidden', liste.length > 0);
    bos.innerHTML = markalar.length ? 'Aramaya uyan marka bulunamadı.' : 'Henüz marka yok.<br>“Marka Ekle” ile ekleyebilirsiniz.';
    liste.forEach(function (m) {
      var tr = document.createElement('tr');
      tr.innerHTML = '<td><span class="fit"><span class="pname' + (m.pasif ? ' pasif' : '') + '">' + esc(m.ad) + '</span>' +
        (m.pasif ? '<span class="pasiftag">(PASİF)</span>' : '') + '</span></td>';
      satirEtkilesim(tr, function (r, x, y) {
        menuAc(x, y, function () {
          onay('Marka Sil', '“' + m.ad + '” silinsin mi?', function () {
            markalar = markalar.filter(function (q) { return q.id !== m.id; });
            silinen.push({ id: uid(), zaman: new Date().toISOString(), ad: m.ad, pasif: !!m.pasif });
            kaydetMarka(); kaydetSilinen(); renderTanim(); toast('Marka silindi ve arşive alındı.');
          });
        }, {
          yazi: m.pasif ? 'Aktif Yap' : 'Pasif Yap',
          fn: function () {
            m.pasif = !m.pasif;
            kaydetMarka(); renderTanim();
            toast('“' + m.ad + '” ' + (m.pasif ? 'pasif yapıldı.' : 'aktif yapıldı.'));
          }
        });
      }, function () { markaDialog(m); });
      tb.appendChild(tr);
    });
    sigdir(tb);
  }

  var duzenlenen = null;
  function markaDialog(m) {
    duzenlenen = m || null;
    $('dlgMarkaTitle').textContent = m ? 'Marka Düzenle' : 'Marka Ekle';
    $('inpAd').value = m ? m.ad : '';
    $('inpPasif').checked = !!(m && m.pasif);
    $('dlgMarkaErr').textContent = '';
    dialogAc('dlgMarka');
    setTimeout(function () { $('inpAd').focus(); }, 50);
  }
  function markaKaydet() {
    var ad = $('inpAd').value.replace(/\s+/g, ' ').trim();
    if (!ad) { $('dlgMarkaErr').textContent = 'Marka ismi boş olamaz.'; return; }
    var ayni = markalar.some(function (q) { return q !== duzenlenen && trKucuk(q.ad) === trKucuk(ad); });
    if (ayni) { $('dlgMarkaErr').textContent = 'Bu isimde bir marka zaten var.'; return; }
    var pasif = $('inpPasif').checked;
    if (duzenlenen) { duzenlenen.ad = ad; duzenlenen.pasif = pasif; }
    else markalar.push({ id: uid(), ad: ad, pasif: pasif });
    kaydetMarka(); dialogKapat('dlgMarka'); renderTanim();
  }
  $('btnMarkaKaydet').addEventListener('click', markaKaydet);
  $('inpAd').addEventListener('keydown', function (e) { if (e.key === 'Enter') markaKaydet(); });
  $('btnMarkaEkle').addEventListener('click', function () { markaDialog(null); });
  $('txtAra').addEventListener('input', araGirdi);
  $('txtAra').addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && $('txtAra').value.trim()) { araKelimeler.push($('txtAra').value.trim()); $('txtAra').value = ''; renderTanim(); }
    else if (e.key === 'Backspace' && !$('txtAra').value && araKelimeler.length) { araKelimeler.pop(); renderTanim(); }
  });
  $('araKutu').addEventListener('click', function () { $('txtAra').focus(); });
  $('btnAraTemizle').addEventListener('click', function (e) { e.stopPropagation(); araKelimeler = []; $('txtAra').value = ''; renderTanim(); });
  $('btnSilinenArsiv').addEventListener('click', function () { ekranGoster('scrSilinen'); });
  $('btnDisaAktar').addEventListener('click', disaAktar);
  $('btnIceAktar').addEventListener('click', function () { $('inpIceAktar').value = ''; $('inpIceAktar').click(); });
  $('inpIceAktar').addEventListener('change', function () { if (this.files && this.files[0]) iceAktar(this.files[0]); });

  /* ---------- SİLİNEN MARKA ARŞİVİ ---------- */
  function renderSilinen() {
    var liste = silinen.slice().sort(function (a, b) { return a.zaman < b.zaman ? 1 : a.zaman > b.zaman ? -1 : 0; });
    var tb = $('silinenBody'); tb.innerHTML = '';
    $('silinenEmpty').classList.toggle('hidden', liste.length > 0);
    liste.forEach(function (k) {
      var d = new Date(k.zaman), z = tarihYaz(d) + ' ' + saatYaz(d);
      var tr = document.createElement('tr');
      tr.innerHTML = '<td><b>' + z + '</b></td><td><span class="fit"><span class="pname">' + esc(k.ad) + '</span>' +
        '</span>' + (k.pasif ? '<span class="pasifken">(PASİFKEN)</span>' : '') + '</td>';
      satirEtkilesim(tr, function (r, x, y) {
        menuAc(x, y, function () {
          onay('Kayıt Sil', '“' + k.ad + '” silme kaydı arşivden silinsin mi?', function () {
            silinen = silinen.filter(function (q) { return q.id !== k.id; }); kaydetSilinen(); renderSilinen(); toast('Kayıt silindi.');
          });
        });
      }, function () {});
      tb.appendChild(tr);
    });
    var ilk = tb.querySelector('tr'); if (ilk) satirSec(ilk);
    sigdir(tb);
  }
  $('btnSilinenTumSil').addEventListener('click', function () {
    if (!silinen.length) { toast('Silinecek kayıt yok.'); return; }
    onay('Tümünü Sil', 'Tüm arşiv kayıtları (' + silinen.length + ' adet) silinsin mi?', function () {
      silinen = []; kaydetSilinen(); renderSilinen(); toast('Tüm arşiv silindi.');
    });
  });
  $('btnSilinenKapat').addEventListener('click', function () { ekranGoster('scrTanim'); });

  /* ---------- Kapat / Android geri tuşu ---------- */
  var App = plug('App');
  function cikisSor() { onay('Çıkış', 'Uygulamadan çıkılsın mı?', function () { if (App) App.exitApp(); }); }
  // Web sürümünde uygulama kapatılamaz; Kapat butonu sadece Android'de görünür
  $('btnTanimKapat').classList.toggle('hidden', !(isNative && App));
  $('btnTanimKapat').addEventListener('click', cikisSor);
  if (isNative && App) {
    App.addListener('backButton', function () {
      if (!$('popMenu').classList.contains('hidden')) { menuKapat(); return; }
      var d = acikDialog();
      if (d) { if (d.id === 'dlgOnay') onayCb = null; dialogKapat(d.id); return; }
      if (aktifEkran !== 'scrTanim') { ekranGoster('scrTanim'); return; }
      cikisSor();
    });
  }

  // diğer sekmede/pencerede yapılan değişiklikleri al
  window.addEventListener('storage', function () {
    markalar = oku(K_MARKA, []); silinen = oku(K_SILINEN, []);
    ekranGoster(aktifEkran);
  });

  if ('serviceWorker' in navigator && !isNative && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(function () {});
  }

  ekranGoster('scrTanim');
})();
