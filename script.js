/**
 * BAHAR BEAUTY STUDIO - MERKEZİ BULUT VERİTABANI API MOTORU
 * Sürüm: 16.1 (Eksiksiz Master Sürüm - ID Senkronizasyonu ve Kategori Update Eklendi)
 */

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var sonuc = { status: "error", message: "Geçersiz Eylem veya Parametre" };

  try {
    if (action === "login") sonuc = kullaniciGirisKontroluVeLog(e.parameter.username, e.parameter.password);
    else if (action === "getkullanicilar") sonuc = tumVerileriListele("salonKullanicilari");
    else if (action === "kaydetkullanici") sonuc = personelEkleVeyaGuncelle(e.parameter);
    else if (action === "silkullanici") sonuc = satirSil("salonKullanicilari", e.parameter.id);
    else if (action === "getdanisanlar") sonuc = tumVerileriListele("salonDanisanlari");
    else if (action === "ekledanisan" || action === "kaydetdanisan") sonuc = yeniDanisanEkleMetodu(e.parameter);
    else if (action === "guncelledanisan") sonuc = danisanGuncelleMetodu(e.parameter);
    else if (action === "sildanisan") sonuc = satirSil("salonDanisanlari", e.parameter.id);
    
    // 🎯 STOK KATEGORİ
    else if (action === "getkategoriler") sonuc = tumVerileriListele("stokKategori");
    else if (action === "addkategori") sonuc = yeniKategoriEkleMetodu(e.parameter.kategoriAdi);
    else if (action === "editkategori") sonuc = kategoriGuncelleMetodu(e.parameter.eskiKategoriAdi, e.parameter.yeniKategoriAdi); // YENİ EKLENDİ
    else if (action === "silkategori") sonuc = satirSil("stokKategori", e.parameter.id);
    
    // 🎯 DEPO STOK
    else if (action === "getstok") sonuc = tumVerileriListele("stokDepoHafizasi");
    else if (action === "kaydeturun") sonuc = urunEkleVeyaGuncelleMetodu(e.parameter);
    else if (action === "silurun") sonuc = satirSil("stokDepoHafizasi", e.parameter.id);
    
    else if (action === "getlogs") sonuc = tumVerileriListele("girisLoglari");
    else if (action === "clearlogs") sonuc = loglariTemizleFonksiyonu();
    
    // 🎯 HİZMET KÜTÜPHANESİ
    else if (action === "gethizmetler" || action === "gethizmetkategori") sonuc = tumVerileriListele("hizmetKatogri"); 
    else if (action === "kaydethizmetkutuphane") sonuc = hizmetKutuphanesineEkle(e.parameter.hizmetAdiid, e.parameter.hizmetislemSuresi);
    else if (action === "silhizmetkutuphane") sonuc = satirSil("hizmetKatogri", e.parameter.id);
    else if (action === "guncellehizmetkutuphane") sonuc = hizmetKutuphaneGuncelle(e.parameter.id, e.parameter.ad, e.parameter.sure);
    
    // 🎯 PAKET ARŞİVİ
    else if (action === "getarsiv" || action === "getpaketler") sonuc = tumVerileriListele("hizmetPaketGiris");
    else if (action === "savearsiv" || action === "kaydethizmetpaketgiris") sonuc = arsivKaydiOlustur(e.parameter);
    else if (action === "updatearsiv") sonuc = arsivCiroGuncelle(e.parameter.id, e.parameter.ciro);
    else if (action === "updatefullseans") sonuc = arsivSeansVeRandevuGuncelle(e.parameter);
    else if (action === "deletearsiv" || action === "silhizmetpaket") sonuc = satirSil("hizmetPaketGiris", e.parameter.id || e.parameter.hizmetNo);
    
    // 🎯 RANDEVU
    else if (action === "getrandevular") sonuc = tumVerileriListele("randevuTakvimi");
    else if (action === "kaydetrandevu") sonuc = randevuKaydetMetodu(e.parameter);
    else if (action === "silrandevu") sonuc = satirSil("randevuTakvimi", e.parameter.id);

    // 🚨 KABİN SEANS GEÇMİŞİ (ÇİFT KOLON YAPISI)
    else if (action === "getseanslar") sonuc = tumVerileriListele("kabinSeansGecmisi");
    else if (action === "addseans") sonuc = seansEkleVeStokDus(e.parameter);
    else if (action === "deleteseans" || action === "silseans") {
      var seansId = e.parameter.id || e.parameter.seansid || e.parameter.islemid || e.parameter.no;
      sonuc = seansSilVeStokIade(seansId);
    }

  } catch (error) {
    sonuc = { status: "error", message: error.toString() };
  }

  if (callback) {
    var jsOutput = callback + "(" + JSON.stringify(sonuc) + ");";
    return ContentService.createTextOutput(jsOutput).setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService.createTextOutput(JSON.stringify(sonuc)).setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================
// 🚨 KABİN SEANS: EKSİKSİZ "YENİ KAYIT VE GÜNCELLEME" MOTORU
// =========================================================================
function seansEkleVeStokDus(p) {
  var seansSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("kabinSeansGecmisi");
  var stokSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("stokDepoHafizasi");
  
  var gelenId = (p.id && String(p.id).trim() !== "undefined" && String(p.id).trim() !== "null" && String(p.id).trim() !== "") ? String(p.id).trim() : null;
  var isUpdate = gelenId ? true : false; 
  
  var sId = isUpdate ? gelenId : "D-" + String(1000 + seansSheet.getLastRow()).substring(1);
  var sTarih = p.islemTarihi || p.tarih || Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
  var uName = p.uzmanAdSoyad || p.uzmanAdi || p.uzman || "";
  var dName = p.danisanAdSoyad || p.danisanAdi || p.danisan || "";
  var kUrun = p.urunAdi || p.kullanilanUrun || "";
  var kMiktar = parseFloat(p.tuketimMiktari || p.kullanilanMiktar || 0);
  
  var satisFiyati = parseFloat(p.satisFiyati || p.gelir || 0);
  var odenenMiktar = parseFloat(p.odenenMiktar || p.odenen || 0);
  var kategori = p.islemTipi || p.kategori || "";
  var masraf = parseFloat(p.masraf || p.maliyet || 0);
  
  if (kUrun && masraf === 0) {
    var stokData = stokSheet.getDataRange().getValues();
    for (var i = 1; i < stokData.length; i++) {
      if (String(stokData[i][2]).trim() === String(kUrun).trim()) {
        kategori = stokData[i][1]; 
        var birimMaliyet = parseFloat(stokData[i][7] || 0); 
        masraf = kMiktar * birimMaliyet; 
        
        if (!isUpdate) {
          var mevcutKalan = parseFloat(stokData[i][4] || 0); 
          stokSheet.getRange(i + 1, 5).setValue(mevcutKalan - kMiktar);
        }
        break;
      }
    }
  }
  
  var kar = odenenMiktar - masraf; 
  
  if (isUpdate) {
    var data = seansSheet.getDataRange().getValues();
    for (var row = 1; row < data.length; row++) {
      if (String(data[row][0]).trim() === String(sId).trim()) {
        seansSheet.getRange(row + 1, 2).setValue(sTarih);
        seansSheet.getRange(row + 1, 3).setValue(uName);
        seansSheet.getRange(row + 1, 4).setValue(dName);
        seansSheet.getRange(row + 1, 5).setValue(kategori);
        seansSheet.getRange(row + 1, 6).setValue(kUrun);
        seansSheet.getRange(row + 1, 7).setValue(satisFiyati);
        seansSheet.getRange(row + 1, 8).setValue(odenenMiktar);
        seansSheet.getRange(row + 1, 9).setValue(masraf);
        seansSheet.getRange(row + 1, 10).setValue(kar);
        
        SpreadsheetApp.flush(); 
        return { status: "success", message: "Seans başarıyla güncellendi." };
      }
    }
  }
  
  seansSheet.appendRow([sId, sTarih, uName, dName, kategori, kUrun, satisFiyati, odenenMiktar, masraf, kar]);
  SpreadsheetApp.flush();
  return { status: "success", message: "Seans başarıyla işlendi ve yeni kayıt eklendi." };
}

// =========================================================================
// 🚨 İPTAL EDİLEN SEANSIN STOĞUNU DİNAMİK SÜTUNDAN İADE ETME
// =========================================================================
function seansSilVeStokIade(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var seansSheet = ss.getSheetByName("kabinSeansGecmisi");
  var stokSheet = ss.getSheetByName("stokDepoHafizasi");
  
  var data = seansSheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "error", message: "Kayıt yok." };

  var headers = data[0].map(function(k){ return String(k).trim().toLowerCase().replace(/\s/g, ""); });
  var urunIdx = headers.indexOf("urunadi"); if(urunIdx === -1) urunIdx = 5; 
  var masrafIdx = headers.indexOf("masraflar"); if(masrafIdx === -1) masrafIdx = headers.indexOf("masraf"); if(masrafIdx === -1) masrafIdx = 8; 
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      var urunAdi = data[i][urunIdx]; 
      var masrafTutari = parseFloat(data[i][masrafIdx] || 0); 
      
      if (urunAdi && urunAdi !== "-" && urunAdi !== "") {
        var stokData = stokSheet.getDataRange().getValues();
        for (var j = 1; j < stokData.length; j++) {
          if (String(stokData[j][2]).trim() === String(urunAdi).trim()) { 
            var birimMaliyet = parseFloat(stokData[j][7] || 0); 
            var iadeMiktari = birimMaliyet > 0 ? (masrafTutari / birimMaliyet) : 0;
            
            var mevcutKalan = parseFloat(stokData[j][4] || 0); 
            stokSheet.getRange(j + 1, 5).setValue(mevcutKalan + iadeMiktari); 
            break;
          }
        }
      }
      seansSheet.deleteRow(i + 1);
      return { status: "success", message: "Seans iptal edildi ve stok iade edildi." };
    }
  }
  return { status: "error", message: "Seans kaydı bulunamadı." };
}

function geceTarihiFavoriGuncelle() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var danisanSheet = ss.getSheetByName("salonDanisanlari");
  var seansSheet = ss.getSheetByName("kabinSeansGecmisi");
  var paketSheet = ss.getSheetByName("hizmetPaketGiris");
  if (!danisanSheet) return;

  var danisanData = danisanSheet.getDataRange().getValues();
  var seansData = seansSheet ? seansSheet.getDataRange().getValues() : [];
  var paketData = paketSheet ? paketSheet.getDataRange().getValues() : [];

  for (var i = 1; i < danisanData.length; i++) {
    var adSoyad = String(danisanData[i][1]).trim().toLowerCase();
    if (adSoyad === "") continue;

    var maxCount = 0; var favori = "-"; var seansCount = {}; var paketCount = {};

    for (var j = 1; j < seansData.length; j++) {
      if (String(seansData[j][3]).trim().toLowerCase() === adSoyad) {
        var kategori = String(seansData[j][4]).trim().toLowerCase();
        if (kategori.indexOf("sarf") === -1) {
          var urun = String(seansData[j][5]).trim();
          if (urun !== "" && urun !== "-") seansCount[urun] = (seansCount[urun] || 0) + 1;
        }
      }
    }

    for (var k = 1; k < paketData.length; k++) {
      if (String(paketData[k][2]).trim().toLowerCase() === adSoyad) {
        var paket = String(paketData[k][3]).trim();
        if (paket !== "" && paket !== "-") paketCount[paket] = (paketCount[paket] || 0) + 1;
      }
    }

    for (var sUrun in seansCount) {
      if (seansCount[sUrun] > maxCount) { maxCount = seansCount[sUrun]; favori = sUrun; }
    }
    for (var pPaket in paketCount) {
      if (paketCount[pPaket] > maxCount) { maxCount = paketCount[pPaket]; favori = pPaket; }
    }

    var eskiFavori = String(danisanData[i][7]).trim();
    if (favori !== eskiFavori) danisanSheet.getRange(i + 1, 8).setValue(favori);
  }
}

function otomatikTetikleyiciKur() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "geceTarihiFavoriGuncelle") ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("geceTarihiFavoriGuncelle").timeBased().atHour(6).everyDays(1).create();
}

// -------------------------------------------------------------------------
// 🚀 DİĞER TÜM ORTAK MODÜLLER 
// -------------------------------------------------------------------------
function satirSil(sheetName, id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { status: "error", message: "Sayfa bulunamadı." };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush(); 
      return { status: "success", message: "Kayıt kalıcı olarak silindi." };
    }
  }
  return { status: "error", message: "Silinecek kayıt bulunamadı." };
}

function tumVerileriListele(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var range = sheet.getDataRange();
  var data = range.getValues();
  var displayData = range.getDisplayValues(); 
  if (data.length <= 1) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var liste = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var hucreVerisi = data[i][j];
      if (hucreVerisi instanceof Date) hucreVerisi = displayData[i][j];
      obj[headers[j]] = hucreVerisi;
    }
    liste.push(obj);
  }
  return liste;
}

function randevuKaydetMetodu(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("randevuTakvimi");
  if (!sheet) {
    sheet = ss.insertSheet("randevuTakvimi");
    sheet.appendRow(["id", "tarih", "saat", "uzman", "danisan", "kategori", "sure", "notlar"]);
  }

  var rId = "R-" + new Date().getTime(); 
  var safeDate = "'" + String(p.tarih || "").trim();
  var safeSaat = "'" + String(p.saat || "").trim();

  sheet.appendRow([ rId, safeDate, safeSaat, String(p.uzman || "").trim(), String(p.danisan || "").trim(), String(p.kategori || "").trim(), String(p.sure || "").trim(), String(p.notlar || "").trim() ]);
  SpreadsheetApp.flush();
  return { status: "success", message: "Randevu başarıyla buluta kaydedildi.", id: rId };
}

function arsivKaydiOlustur(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("hizmetPaketGiris");
  if (!sheet) return { status: "error", message: "hizmetPaketGiris sayfası bulunamadı." };
  
  var pId = p.id || p.hizmetNo || "";
  var uzman = p.uzman || p.hizmetUzman || "";
  var danisan = p.danisan || p.hizmetDanisan || "";
  var hizmetTercihi = p.hizmetTercihi || p.hizmetPaketTercihi || "";
  var sure = p.sure || p.hizmetPaketSuresi || "0";
  var seansSayisi = p.seansSayisi || p.hizmetSeansSayisi || "1";
  var toplamTutar = parseFloat(p.toplamTutar || p.hizmetToplamOdeme || "0");
  var odenenMiktar = parseFloat(p.odenenMiktar || p.hizmetAraOdeme || "0");
  var kalanBorc = toplamTutar - odenenMiktar;
  var seansDurumu = p.seansDurumu || p.hizmeteGelmeDurumu || "{}";
  var kalanSeans = p.kalanSeans || p.hizmetSeansSayisi || "1";
  var paketTarihi = p.paketTarihi || new Date().toISOString().slice(0,10);
  var randevuTarihi = p.randevuTarihi || p.sonrakiSeansTarihi || "-";
  var randevuSaati = p.randevuSaati || p.sonrakiSeansSaati || "-";

  sheet.appendRow([pId, uzman, danisan, hizmetTercihi, sure, seansSayisi, toplamTutar, odenenMiktar, kalanBorc, seansDurumu, kalanSeans, paketTarihi, randevuTarihi, randevuSaati]);
  return { status: "success", message: "Paket kartı oluşturuldu." };
}

function arsivCiroGuncelle(id, yeniCiro) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("hizmetPaketGiris");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      var toplam = parseFloat(data[i][6] || 0); 
      sheet.getRange(i + 1, 8).setValue(yeniCiro); 
      sheet.getRange(i + 1, 9).setValue(toplam - yeniCiro); 
      return { status: "success", message: "Ödeme güncellendi." };
    }
  }
  return { status: "error", message: "Kayıt bulunamadı." };
}

function arsivSeansVeRandevuGuncelle(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("hizmetPaketGiris");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(p.id || p.hizmetNo).trim()) {
      sheet.getRange(i + 1, 10).setValue(p.seansDurumu || p.hizmeteGelmeDurumu);
      sheet.getRange(i + 1, 11).setValue(p.kalanSeans);
      sheet.getRange(i + 1, 13).setValue(p.randevuTarihi || p.sonrakiSeansTarihi);
      sheet.getRange(i + 1, 14).setValue(p.randevuSaati || p.sonrakiSeansSaati);
      return { status: "success", message: "Seans matrisi güncellendi." };
    }
  }
  return { status: "error", message: "Kayıt bulunamadı." };
}

function kullaniciGirisKontroluVeLog(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("salonKullanicilari");
  var logSheet = ss.getSheetByName("girisLoglari");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === String(username).trim() && String(data[i][3]).trim() === String(password).trim()) {
      if (logSheet) logSheet.appendRow(["L-00" + logSheet.getLastRow(), Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss"), username, "BAŞARILI", "Giriş Yapıldı"]);
      return { status: "success", adSoyad: data[i][1], rol: data[i][4] };
    }
  }
  return { status: "error", message: "Kullanıcı adı veya şifre hatalı!" };
}

function personelEkleVeyaGuncelle(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("salonKullanicilari");
  var data = sheet.getDataRange().getValues();
  var keys = data[0].map(function(k) { return String(k).trim().toLowerCase(); });
  var idIdx = keys.indexOf("id");
  if (p.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === String(p.id).trim()) {
        sheet.getRange(i + 1, keys.indexOf("adsoyad") + 1).setValue(p.adSoyad);
        sheet.getRange(i + 1, keys.indexOf("eposta") + 1).setValue(p.eposta);
        sheet.getRange(i + 1, keys.indexOf("sifre") + 1).setValue(p.sifre);
        sheet.getRange(i + 1, keys.indexOf("rol") + 1).setValue(p.rol);
        return { status: "success", message: "Personel güncellendi." };
      }
    }
  }
  sheet.appendRow(["P-" + (100 + sheet.getLastRow()), p.adSoyad, p.eposta, p.sifre, p.rol]);
  return { status: "success", message: "Personel başarıyla eklendi." };
}

function urunEkleVeyaGuncelleMetodu(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("stokDepoHafizasi");
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(k) { return String(k).trim().toLowerCase(); });
  var idIdx = headers.indexOf("urunid") > -1 ? headers.indexOf("urunid") : headers.indexOf("id");
  var hesapHacim = parseFloat(p.toplamHacim || p.tophamHacim || 0); // Hata düzeltildi
  var maliyet = parseFloat(p.toplamMaliyet || 0);
  var birimMaliyet = hesapHacim > 0 ? (maliyet / hesapHacim) : 0;

  // 🚨 İŞTE BÜYÜK HATA BURADAYDI. ŞİMDİ GELEN ID VARSA ONU KULLAN, YOKSA YENİ ÜRET.
  var targetId = (p.id && String(p.id).trim() !== "undefined" && String(p.id).trim() !== "") ? String(p.id).trim() : ("U-" + (Date.now()).toString().slice(-6));

  if (p.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === String(p.id).trim()) {
        sheet.getRange(i + 1, headers.indexOf("urunadi") + 1).setValue(p.urunAdi);
        sheet.getRange(i + 1, headers.indexOf("kategori") + 1).setValue(p.kategori);
        sheet.getRange(i + 1, headers.indexOf("toplamhacim") + 1).setValue(hesapHacim);
        sheet.getRange(i + 1, headers.indexOf("kalanhacim") + 1).setValue(hesapHacim);
        sheet.getRange(i + 1, headers.indexOf("toplammaliyet") + 1).setValue(maliyet);
        sheet.getRange(i + 1, headers.indexOf("birimmaliyet") + 1).setValue(birimMaliyet);
        sheet.getRange(i + 1, headers.indexOf("birim") + 1).setValue(p.birim);
        return { status: "success", message: "Ürün güncellendi." };
      }
    }
  }
  
  // EĞER BULAMAZSA VEYA YENİ İSE TARGET ID (FRONTEND'DEN GELEN ID) İLE YENİ SATIR EKLE
  sheet.appendRow([targetId, p.kategori, p.urunAdi, hesapHacim, hesapHacim, p.birim, maliyet, birimMaliyet]);
  SpreadsheetApp.flush();
  return { status: "success", message: "Yeni Ürün eklendi." };
}

function yeniKategoriEkleMetodu(kategoriAdi) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("stokKategori");
  sheet.appendRow(["K-" + (100 + sheet.getLastRow()), kategoriAdi.trim()]);
  return { status: "success", message: "Kategori başarıyla eklendi." };
}

// 🚨 YENİ EKLENEN KATEGORİ GÜNCELLEME FONKSİYONU
function kategoriGuncelleMetodu(eskiAd, yeniAd) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("stokKategori");
  if(!sheet) return {status: "error", message: "Sayfa yok."};
  var data = sheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
     if(String(data[i][1]).trim() === String(eskiAd).trim()) {
        sheet.getRange(i+1, 2).setValue(yeniAd.trim());
        return {status:"success", message:"Kategori güncellendi."};
     }
  }
  return {status:"error", message:"Eski kategori bulunamadı."};
}

function hizmetKutuphanesineEkle(ad, sure) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("hizmetKatogri"); 
  sheet.appendRow(["H-" + (100 + sheet.getLastRow()), ad ? ad.trim() : "", sure ? sure.toString().trim() : ""]);
  SpreadsheetApp.flush(); 
  return { status: "success", message: "Hizmet başarıyla eklendi." };
}

function hizmetKutuphaneGuncelle(id, ad, sure) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("hizmetKatogri");
  var data = sheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
     if(String(data[i][0]).trim() === String(id).trim()) {
        sheet.getRange(i+1, 2).setValue(ad);
        sheet.getRange(i+1, 3).setValue(sure);
        return {status: "success", message: "Hizmet başarıyla güncellendi."};
     }
  }
  return {status: "error", message: "Hizmet bulunamadı."};
}

function danisanGuncelleMetodu(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("salonDanisanlari");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(p.id).trim()) {
      if (p.adSoyad) sheet.getRange(i + 1, 2).setValue(p.adSoyad.trim());
      if (p.telefon) sheet.getRange(i + 1, 4).setValue(String(p.telefon).trim());
      if (p.dogumTarihi) sheet.getRange(i + 1, 5).setValue(String(p.dogumTarihi).trim());
      if (p.not) sheet.getRange(i + 1, 7).setValue(String(p.not).trim());
      return { status: "success", message: "Danışan güncellendi." };
    }
  }
  return { status: "error", message: "Danışan bulunamadı." };
}

function yeniDanisanEkleMetodu(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("salonDanisanlari");
  var adSoyad = typeof p === 'object' ? (p.adSoyad || "") : (p || "");
  sheet.appendRow(["D-" + (100 + sheet.getLastRow()), String(adSoyad).trim(), Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm"), String(p.telefon||"").trim(), String(p.dogumTarihi||"").trim(), "", String(p.not||"").trim(), "-"]);
  return { status: "success", message: "Danışan başarıyla eklendi." };
}

function loglariTemizleFonksiyonu() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("girisLoglari");
  var basliklar = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, basliklar[0].length).setValues(basliklar);
  return { status: "success", message: "Loglar sıfırlandı." };
}