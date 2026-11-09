// PatiPay site kayıtları: Google E-Tablo'ya yazar.
// Kurulum: E-Tablo > Uzantılar > Apps Script > bu kodu yapıştır > Dağıt > Yeni dağıtım >
// Tür: Web uygulaması, Erişim: Herkes > Dağıt. Çıkan /exec adresini siteye ENDPOINT olarak koy.

var HAREKET_BASLIK = ['Zaman', 'Ziyaretçi', 'Oturum', 'Olay', 'Detay', 'Bölüm', 'Kanal', 'Kampanya', 'Cihaz', 'Ekran genişliği'];
var KAYIT_BASLIK = ['Zaman', 'Tür', 'E-posta', 'Dostlar', 'Dost sayısı', 'Aylık tutar', 'Asistan önerisi', 'Başlangıç', 'Sigorta', 'İstekler', 'Kanal', 'Kampanya', 'Cihaz', 'Ziyaretçi'];

function sayfa(ad, baslik) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ad) || ss.insertSheet(ad);
  if (sh.getLastRow() === 0) {
    sh.appendRow(baslik);
    sh.getRange(1, 1, 1, baslik.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function detay(p) {
  if (!p) return '';
  return Object.keys(p).map(function (k) {
    var v = p[k];
    return k + ': ' + (typeof v === 'object' ? JSON.stringify(v) : v);
  }).join(' | ');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var simdi = new Date();
    if (d.type === 'events') {
      var sh = sayfa('Hareketler', HAREKET_BASLIK);
      var rows = (d.events || []).map(function (ev) {
        return [new Date(ev.t), d.visitor, d.session, ev.e, detay(ev.p), ev.v, d.kanal, d.utm_campaign, d.cihaz, d.ekran];
      });
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HAREKET_BASLIK.length).setValues(rows);
    } else {
      var sk = sayfa('Kayıtlar', KAYIT_BASLIK);
      sk.appendRow([simdi, d.type === 'vet' ? 'Veteriner' : 'Kullanıcı', d.email, d.dostlar || '', d.dost_sayisi || '',
        d.aylik_tutar || '', d.asistan_onerisi || '', d.baslangic || '', d.sigorta || '', d.istekler || '',
        d.kanal, d.utm_campaign, d.cihaz, d.visitor]);
      if (d.email && d.type === 'signup') {
        MailApp.sendEmail(d.email, 'PatiPay listesine katıldın',
          'Merhaba,\n\nPatiPay listesine katıldığın için teşekkürler. Uygulama hazır olduğunda ilk haberi sana vereceğiz.\n\nInstagram: @pati_pay_\n\nPatiPay');
      }
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
