// TEKNOFEST KULÜBÜ ÜYELİK SİSTEMİ - Google E-Tablolar bağlantısı
// 1) TABLO_ID kısmına E-Tablonun ID'sini yazın (adresteki /d/ ile /edit arası)
// 2) Bir kez "kurulum" fonksiyonunu çalıştırın (başlıkları oluşturur, izin ister)
// 3) Dağıt > Yeni dağıtım > Web uygulaması (Çalıştıran: Ben, Erişim: Herkes)

const TABLO_ID = 'TABLO_ID';
const SEKME = 'Kayitlar';

function kurulum() {
  const ss = SpreadsheetApp.openById(TABLO_ID);
  const sh = ss.getSheetByName(SEKME) || ss.insertSheet(SEKME);
  sh.getRange(1, 1, 1, 7).setValues([['Zaman', 'Ad', 'Soyad', 'Bölüm', 'Öğrenci No', 'Telefon', 'Ödeme Kontrolü']])
    .setFontWeight('bold').setBackground('#0B57D0').setFontColor('#ffffff');
  sh.getRange('E:F').setNumberFormat('@');
  sh.setFrozenRows(1);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const d = JSON.parse(e.postData.contents);
    if (d.web) return cevap(true, 'ok'); // bot tuzağı

    const ad = temiz(d.ad), soyad = temiz(d.soyad), bolum = temiz(d.bolum);
    const no = temiz(d.no).replace(/\s/g, ''), tel = temiz(d.tel).replace(/\s/g, '');
    if (!ad || !soyad || !bolum || !/^\d{5,15}$/.test(no) || !/^\d{10,11}$/.test(tel))
      return cevap(false, 'Bilgiler eksik veya hatalı.');

    const sh = SpreadsheetApp.openById(TABLO_ID).getSheetByName(SEKME);
    const son = sh.getLastRow();

    // Aynı öğrenci numarası tekrar gelirse yeni satır açma, öğrenciyi 3. aşamaya geçir
    if (son > 1) {
      const nolar = sh.getRange(2, 5, son - 1, 1).getDisplayValues().flat();
      if (nolar.includes(no)) return cevap(true, 'Bu öğrenci numarası zaten kayıtlı.');
    }

    const r = son + 1;
    sh.getRange(r, 5, 1, 2).setNumberFormat('@'); // baştaki 0 silinmesin
    const zaman = Utilities.formatDate(new Date(), 'Europe/Istanbul', 'dd.MM.yyyy HH:mm:ss');
    sh.getRange(r, 1, 1, 6).setValues([[zaman, ad, soyad, bolum, no, tel]]);
    return cevap(true, 'Kayıt alındı.');
  } catch (err) {
    return cevap(false, 'Sunucu hatası, tekrar dene.');
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet() { return cevap(true, 'Sistem çalışıyor.'); }

function temiz(s) {
  return String(s || '').trim().replace(/^[=+\-@\s]+/, '').slice(0, 100);
}

function cevap(ok, msg) {
  return ContentService.createTextOutput(JSON.stringify({ ok: ok, msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
