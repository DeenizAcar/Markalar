# Markalar

Marka tanımlama uygulaması (Mini Puantaj'ın Tanımlar ekranı mantığıyla).

- Web: https://deenizacar.github.io/Markalar/
- Android APK: https://deenizacar.github.io/Markalar/apk/

## Özellikler
- **Arama:** Yazılan her kelime boşluk/Enter ile etikete dönüşür; kelimelerden herhangi birine uyan markalar listelenir.
- **Marka Ekle** / satıra çift dokunma: Marka İsmi + **Pasife Al**. Aynı isim ikinci kez eklenemez.
- **Basılı tutma:** `Kaydı Sil` / `Pasif Yap`–`Aktif Yap` / `Vazgeç`. Silinen marka **Silinen Marka Arşivi**'ne gider.
- **Dışa / İçe Aktar:** `.xlsx` (sayfa `Markalar`: Marka İsmi, Durumu). İçe aktarma sadece yeni markaları ekler; ilk kolonda isim olan elle hazırlanmış liste de okunur.
- **Kapat** (sadece Android) ve geri tuşu çıkış onayı sorar.

## Teknik
Saf HTML/CSS/JS (`www/`), Capacitor 7 ile Android'e paketlenir. Veriler cihazda `localStorage`'da tutulur:
`mk_marka` = `[{id, ad, pasif}]`, `mk_silinen` = `[{id, zaman, ad, pasif}]`.

`main` dalına her push'ta `.github/workflows/deploy.yml` APK'yı derler ve siteyle birlikte GitHub Pages'e yayınlar.
İmza için repo'da `DEBUG_KEYSTORE_B64` secret'ı gerekir (paket kimliği: `com.deenizacar.markalar`).

Yerelde: `npm run serve` → http://localhost:5173
