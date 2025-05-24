## Összegyűjti az UGYELETI_BEOSZTAS_2025 megosztott lapról és email-t küld.
```
/*
 * Szkript: Aktuális hónap szombat és vasárnapi napjain szereplő "2"-es beosztások és az "Ü" betűk összesítése,
 * valamint találatok dátumainak kiírása RichText formátumban, havi kifizetések, szabadnap megváltás és összesített kifizetés számítása.
 * A dátumokat e-mailben HTML formátumban küldi, inline színekkel: '2'-es dátumok zölddel, 'Ü'-k pirossal.
 * Az e-mail fejlécében: zölddel "Hétvégi készenlét", pirossal "Hétvégi ügyelet".
 * A név és az összes kifizetés fehér színnel jelenik meg az e-mailben.
 */

// TODO: Állítsd be a forrás táblázat ID-jét és a címzett e-mailt!
var SOURCE_SS_ID = '1oycDmFTobQqmoblDTzlvfYNgNjPntUtPy0SVefiQeA8';
var RECIPIENT    = 'kesmarkizoltan@gmail.com';

function countWeekendTwosAndU() {
  var ui = SpreadsheetApp.getUi();
  if (SOURCE_SS_ID.indexOf('1oyc') === -1 || RECIPIENT.indexOf('@') === -1) {
    ui.alert('Kérlek, állítsd be a SOURCE_SS_ID és RECIPIENT változókat a valós értékekre!');
    return;
  }

  // Aktuális hónap
  var today = new Date(),
      year  = today.getFullYear(),
      monthIdx = today.getMonth(),
      month = ('0' + (monthIdx + 1)).slice(-2);

  // Forrás megnyitása
  var source;
  try { source = SpreadsheetApp.openById(SOURCE_SS_ID); }
  catch(e){ ui.alert('Hiba forrás megnyitásakor: ' + e.message); return; }
  var sheet = source.getSheetByName(month);
  if (!sheet) { ui.alert('Nincs lap: ' + month); return; }

  // AH oszlop kizárása
  var ignoreCol = sheet.getRange('AH1').getColumn();
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) { ui.alert('Nincs adat a hónap lapján.'); return; }

  // Nevek és cellák
  var names = sheet.getRange(4,1,lastRow-3).getValues().map(r=>r[0]);
  var data  = sheet.getRange(4,3,lastRow-3,lastCol-2).getDisplayValues();

  // Dátum formázó
  function fmt(day){
    var d = new Date(year, monthIdx, day);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy.MM.dd');
  }

  // Hétvégi napok indexei
  var weekendCols = [];
  for (var i=0;i<lastCol-2;i++){
    var col = i+3;
    if (col===ignoreCol) continue;
    var day = new Date(year, monthIdx, i+1).getDay();
    if (day===0||day===6) weekendCols.push(col);
  }

  // Eredmények gyűjtése
  var counts = [], dateRuns = [];
  names.forEach(function(name,idx){
    var c2=0, cU=0, runs=[];
    weekendCols.forEach(function(col){
      var v = String(data[idx][col-3]).trim().toUpperCase();
      var dayNum = col-2;
      if (v==='2') { c2++; runs.push({text:fmt(dayNum), type:'2'}); }
      if (v==='Ü') { cU++; runs.push({text:fmt(dayNum), type:'Ü'}); }
    });
    counts.push([c2, cU]);
    dateRuns.push(runs);
  });

  // Kimeneti lap létrehozása/frissítése
  var dest = SpreadsheetApp.getActiveSpreadsheet();
  var outName = 'Weekend_2_&_Ü_Count';
  var out = dest.getSheetByName(outName) || dest.insertSheet(outName);
  out.clear();

  // Fejléc
  out.getRange(1,1,1,7).setValues([['Név','Dátum','Hétvégi készenlét','Hétvégi ügyelet','Készenlét/Ügyelet havi','Szabadnap megváltás','Összes kifizetés']]);

  // Nevek és dátumok írása
  out.getRange(2,1,names.length,1).setValues(names.map(n=>[n]));

  // RichText dátumok
  var rich = dateRuns.map(function(runs){
    var text = runs.map(r=>r.text).join(', ');
    var builder = SpreadsheetApp.newRichTextValue().setText(text);
    var pos=0;
    runs.forEach(function(run,i){
      var len = run.text.length;
      var style = SpreadsheetApp.newTextStyle()
        .setForegroundColor(run.type==='2'?'green':'red')
        .build();
      builder.setTextStyle(pos, pos+len, style);
      pos += len + (i<runs.length-1?2:0);
    });
    return [builder.build()];
  });
  out.getRange(2,2,rich.length,1).setRichTextValues(rich);

  // Számlálások
  out.getRange(2,3,counts.length,2).setValues(counts);

  // Képletek
  names.forEach(function(_,i){
    var r = i+2;
    if (r>=2&&r<=5) out.getRange(r,5).setFormula(`=C${r}*10000`);
    if (r>=8&&r<=10) out.getRange(r,5).setFormula(`=D${r}*7500`);
    if ((r>=2&&r<=7)||(r>=11&&r<=25)) out.getRange(r,6).setFormula(`=D${r}*5000`);
    out.getRange(r,7).setFormula(`=E${r}+F${r}`);
  });

  // Formázás
  [2,3,4,5,6,7].forEach(c=>out.autoResizeColumn(c));
  out.setColumnWidth(1,out.getColumnWidth(1)+30);
  out.setColumnWidth(2,out.getColumnWidth(2)+40);
  out.getRange('B9').setHorizontalAlignment('left');
  // C-G oszlopok szélességének növelése +20 ponttal
  [3,4,5,6,7].forEach(function(col){
  var w = out.getColumnWidth(col);
  out.setColumnWidth(col, w + 5);
});


  // E-mail összeállítása HTML-ben
  var lines = [];
  dateRuns.forEach(function(runs,i){
    var total = out.getRange(i+2,7).getValue();
    if (total!==0) {
      var name = names[i];
      var htmlDates = runs.map(function(run){
        var color = run.type==='2'?'green':'red';
        return `<span style="color:${color}">${run.text}</span>`;
      }).join(', ');
      lines.push(`<p>` +
                 `<strong style="color:white">${name}</strong> ` +
                 `(${htmlDates}): ` +
                 `<span style="color:white">${total}</span>` +
                 `</p>`);
    }
  });

  if (lines.length) {
    var headerHtml = '<p>' +
      '<span style="color:green">Hétvégi készenlét</span> ' +
      '<span style="color:red">Hétvégi ügyelet/Szabadnap megváltás</span>' +
      '</p>';
    MailApp.sendEmail({
      to: RECIPIENT,
      subject: `Készenlét/Szabadnap megváltás kifizetés – ${month}`,
      htmlBody: headerHtml + lines.join('')
    });
  }

  ui.alert('Futás befejezve – email elküldve, ha volt nem-nullás sor.');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ügyeleti Összesítés')
    .addItem('Számolás','countWeekendTwosAndU')
    .addItem('Parancs gomb','showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutput(
    '<button onclick="google.script.run.countWeekendTwosAndU()" ' +
    'style="font-size:14px;padding:10px;">Futtatás</button>'
  ).setTitle('Ügyeleti Össz');
  SpreadsheetApp.getUi().showSidebar(html);
}
```
