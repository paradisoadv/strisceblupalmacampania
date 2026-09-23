// Google Apps Script - da collegare al Google Sheet degli abbonamenti.
// Pubblicare come: Distribuisci > Nuova distribuzione > App web.
const SPREADSHEET_ID = '1x8Yx2sU4M_GD4OM3w1vlcqz-cW-tx8XZ';

function doGet(e) {
  const targa = normalizza(e.parameter.targa || '');
  const out = cercaTarga(targa);
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizza(v){ return String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,''); }
function fmt(d){ return Utilities.formatDate(new Date(d), Session.getScriptTimeZone() || 'Europe/Rome', 'dd/MM/yyyy'); }

function cercaTarga(targa){
  if(!targa) return {found:false};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  for (const sh of ss.getSheets()) {
    const values = sh.getDataRange().getValues();
    if(values.length < 3) continue;
    // Struttura osservata: riga 1 categorie, riga 2 sottocategorie; targa colonna F; DAL/AL ultime colonne.
    for(let r=2;r<values.length;r++){
      const row=values[r];
      if(normalizza(row[5]) !== targa) continue;
      const expiry=row[23] || row[row.length-1];
      const start=row[22] || row[row.length-2];
      const type=tipoAbbonamento(values[0], values[1], row);
      const expDate=expiry ? new Date(expiry) : null;
      const now=new Date(); now.setHours(0,0,0,0);
      const active=!!expDate && expDate >= now;
      return {found:true, active, plate:targa, type:type, start:start?fmt(start):'', expiry:expiry?fmt(expiry):''};
    }
  }
  return {found:false, plate:targa};
}

function tipoAbbonamento(top, sub, row){
  // Colonne tariffarie J:U (indici 9..20). Cerca la cella valorizzata e combina categoria + durata.
  let lastTop='';
  for(let c=0;c<top.length;c++){
    if(String(top[c]).trim()) lastTop=String(top[c]).trim();
    if(c>=9 && c<=20 && row[c]!=='' && row[c]!=null){
      const s=String(sub[c]||'').trim();
      return [lastTop,s].filter(Boolean).join(' - ');
    }
  }
  return 'Abbonamento';
}
