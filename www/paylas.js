/* Android uygulamasında verileri Sipariş Hazırlama ile paylaşmak için Preferences'a da yazar.
   (Uygulama localStorage kullanır; Sipariş Hazırlama onu doğrudan okuyamaz.) */
(function () {
  'use strict';
  var ANAHTAR = 'mk_marka';
  var C = window.Capacitor;
  if (!C || !C.isNativePlatform || !C.isNativePlatform()) return;
  var P = C.Plugins && C.Plugins.Preferences;
  if (!P) return;
  function aktar(v) {
    try {
      var p = v == null ? P.remove({ key: ANAHTAR }) : P.set({ key: ANAHTAR, value: String(v) });
      if (p && p.catch) p.catch(function () {});
    } catch (e) { /* yok say */ }
  }
  var setItem = Storage.prototype.setItem, removeItem = Storage.prototype.removeItem;
  Storage.prototype.setItem = function (k, v) {
    setItem.call(this, k, v);
    if (this === window.localStorage && k === ANAHTAR) aktar(v);
  };
  Storage.prototype.removeItem = function (k) {
    removeItem.call(this, k);
    if (this === window.localStorage && k === ANAHTAR) aktar(null);
  };
  aktar(localStorage.getItem(ANAHTAR));
})();
