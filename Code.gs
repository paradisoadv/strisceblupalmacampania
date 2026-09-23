const SPREADSHEET_ID = "1x8Yx2sU4M_GD4OM3w1vlcqz-cW-tx8XZ";
const CACHE_SECONDS = 300; // 5 minuti

function doGet(e) {
  try {
    const targa = normalizza_(e && e.parameter ? e.parameter.targa : "");
    if (!targa) return json_({ok:false, error:"Targa mancante"});

    const cache = CacheService.getScriptCache();
    const cacheKey = "plate_" + targa;
    const cached = cache.get(cacheKey);
    if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let risultato = {ok:true, found:false, plate:targa};

    for (const sh of sheets) {
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow < 3 || lastCol < 5) continue;

      // Legge solo la colonna E (targa), dalla riga 3 in poi.
      const plates = sh.getRange(3, 5, lastRow - 2, 1).getDisplayValues();
      let idx = -1;
      for (let i = 0; i < plates.length; i++) {
        if (normalizza_(plates[i][0]) === targa) {
          idx = i + 3;
          break;
        }
      }
      if (idx < 0) continue;

      // Legge una sola riga, soltanto quando trova la targa.
      const row = sh.getRange(idx, 1, 1, lastCol).getDisplayValues()[0];

      // Struttura del foglio: J:M Residenti/Docenti/Studenti,
      // N:Q Commercianti/Lavoratori, R ISEE, S Spendo a Palma,
      // T Non residenti, U Residenti nella zona, W/X validità.
      const categorie = [
        ["RESIDENTI/DOCENTI/STUDENTI - MENSILE", 9],
        ["RESIDENTI/DOCENTI/STUDENTI - TRIMESTRALE", 10],
        ["RESIDENTI/DOCENTI/STUDENTI - SEMESTRALE", 11],
        ["RESIDENTI/DOCENTI/STUDENTI - ANNUALE", 12],
        ["COMMERCIANTI/LAVORATORI - MENSILE", 13],
        ["COMMERCIANTI/LAVORATORI - TRIMESTRALE", 14],
        ["COMMERCIANTI/LAVORATORI - SEMESTRALE", 15],
        ["COMMERCIANTI/LAVORATORI - ANNUALE", 16],
        ["AGEVOLAZIONE ISEE - ANNUALE", 17],
        ["SPENDO A PALMA - ANNUALE", 18],
        ["NON RESIDENTI - ANNUALE", 19],
        ["RESIDENTI NELLA ZONA - ANNUALE", 20]
      ];

      let tipo = "ABBONAMENTO";
      for (const [nome, col] of categorie) {
        if (row[col] !== undefined && String(row[col]).trim() !== "") {
          tipo = nome;
          break;
        }
      }

      const expiryRaw = row[23] || ""; // X = AL
      const expiryDate = parseDate_(expiryRaw);
      const today = new Date();
      today.setHours(0,0,0,0);

      risultato = {
        ok:true,
        found:true,
        active: expiryDate ? expiryDate >= today : true,
        plate:targa,
        type:tipo,
        expiry:formatDate_(expiryDate, expiryRaw)
      };
      break;
    }

    const payload = JSON.stringify(risultato);
    cache.put(cacheKey, payload, CACHE_SECONDS);
    return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ok:false, found:false, error:String(err && err.message ? err.message : err)});
  }
}

function normalizza_(v) {
  return String(v || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function parseDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) return v;
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function formatDate_(d, fallback) {
  if (!d) return fallback || "—";
  return Utilities.formatDate(d, Session.getScriptTimeZone() || "Europe/Rome", "dd/MM/yyyy");
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
