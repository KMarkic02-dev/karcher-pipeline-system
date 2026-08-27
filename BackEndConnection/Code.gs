function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Kärcher InfoBoard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function preveriStatusInPosljiEmail() {
  var doc = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID");
  var sheet = doc.getSheetByName("Status") || doc.getSheets()[0];
  
  var vrednostA1 = sheet.getRange("A1").getDisplayValue().trim();
  Logger.log("Trenutna vrednost v A1: '" + vrednostA1 + "'");
  
  // Preverimo, da vsebuje ERROR: in se NE konča na E
  if (vrednostA1.indexOf("ERROR:") !== -1 && !vrednostA1.endsWith("E")) {
    Logger.log("-> POGOJ JE IZPOLNJEN! Pošiljam email...");
    
    var prejemnik = "user@example.com";
    var zadeva = "ERROR: Kärcher InfoBoard Osveževanje";
    var vsebina = "Pri samodejnem osveževanju SAP podatkov je prišlo do napake.\n\n" +
                  "\n" + vrednostA1 + "\n\n" +
                  "Prosimo, preverite lokalne loge na strežniku.";
                  
    MailApp.sendEmail(prejemnik, zadeva, vsebina);
    
    // Na konec dodamo E brez presledka (da trim() ne pobriše oznake)
    sheet.getRange("A1").setValue(vrednostA1 + "E");
    Logger.log("-> E-pošta poslana in A1 posodobljen!");
  } else {
    Logger.log("-> POGOJ NI BIL IZPOLNJEN. Razlog: ali ni 'ERROR:' ali pa se že konča z 'E'.");
  }
}

function pridobiZadnjiStatus() {
  try {
    var doc = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID");
    var sheet = doc.getSheetByName("Status");
    if (!sheet) {
      sheet = doc.getSheets()[0];
    }
    var vrednost = sheet.getRange("A1").getValue();
    if (vrednost instanceof Date) {
      vrednost = Utilities.formatDate(vrednost, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
    }
    return String(vrednost || "Neznano");
  } catch (e) {
    return "Ni podatka";
  }
}

function pridobiVseCeneMap() {
  try {
    var doc = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID");
    var sheet = doc.getSheetByName("Cenik") || doc.getSheets()[6];
    if (!sheet) return {};
    
    var data = sheet.getDataRange().getValues();
    var cenikMap = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rawArtikel = String(row[0]).trim();
      var cistiArtikel = rawArtikel.toLowerCase().replace(/[\.\-\s]/g, '');

      if (cistiArtikel !== '') {
        cenikMap[cistiArtikel] = {
          mpc_ddv: row[3] !== '' && row[3] !== null ? row[3] : null,
          mpc: row[4] !== '' && row[4] !== null ? row[4] : null,
          vpc: row[5] !== '' && row[5] !== null ? row[5] : null
        };
      }
    }
    return cenikMap;
  } catch (e) {
    Logger.log("Napaka v pridobiVseCeneMap: " + e.toString());
    return {};
  }
}

function pridobiPodatkeZaCenik() {
  var doc = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID");
  var sheet = doc.getSheetByName("Cenik") || doc.getSheets()[6];
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var rows = data.slice(1);
  
  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

function pridobiPodatkeZaDashboard() {
  var doc = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID");
  var sheet = doc.getSheetByName("Zaloga");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var rows = data.slice(1);
  
  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

function pridobiPodatkeZaKupce() {
  var doc = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID");
  var sheet = doc.getSheetByName("Kupci");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var rows = data.slice(1);
  
  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

function pridobiNarocilaZaKupca(sifraKupca) {
  var doc = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID");
  var sheet = doc.getSheetByName("Delivery") || doc.getSheetByName("Deilvery");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var rows = data.slice(1);
  
  function ocistiNiz(val) {
    if (val === null || val === undefined) return "";
    var s = String(val)
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/\s+/g, '')
      .trim();
    if (s.endsWith(".0")) {
      s = s.substring(0, s.length - 2);
    }
    return s;
  }
  
  var cistaTargetSifra = ocistiNiz(sifraKupca);
  var customerIndex = -1;
  
  headers.forEach(function(header, idx) {
    var h = String(header).trim().toUpperCase();
    if (h === "WW CUSTOMER" || h === "CUSTOMER" || h === "ŠIFRA PODJETJA" || h === "ŠIFRA KUPCA") {
      customerIndex = idx;
    }
  });
  
  if (customerIndex === -1) customerIndex = 1;
  
  var filtriranaNarocila = [];
  var timeZone = Session.getScriptTimeZone();
  
  rows.forEach(function(row) {
    var vrediSifra = ocistiNiz(row[customerIndex]);
    if (vrediSifra !== "" && vrediSifra === cistaTargetSifra) {
      var obj = {};
      headers.forEach(function(header, index) {
        var vrednost = row[index];
        
        // --- PRAVILNO FORMATIRANJE DATUMA ---
        if (vrednost instanceof Date) {
          vrednost = Utilities.formatDate(vrednost, timeZone, "dd.MM.yyyy");
        } else if (typeof vrednost === 'string' && /^\d{4}-\d{2}-\d{2}/.test(vrednost.trim())) {
          var deli = vrednost.trim().split('T')[0].split('-');
          if (deli.length === 3) {
            vrednost = deli[2] + '.' + deli[1] + '.' + deli[0];
          }
        }
        
        obj[header] = vrednost;
      });
      filtriranaNarocila.push(obj);
    }
  });
  
  return filtriranaNarocila;
}