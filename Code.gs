/**
 * Live Skor LCC — backend Google Sheets
 * Pasangan dari index.html (papan skor live 3 regu).
 *
 * CARA PAKAI:
 * 1. Buat spreadsheet baru di Google Sheets (atau pakai yang sudah ada).
 * 2. Salin ID spreadsheet dari alamatnya:
 *    https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit
 *    lalu tempel ke SPREADSHEET_ID di bawah (ganti tulisan GANTI_DENGAN_ID_SPREADSHEET).
 * 3. Di Apps Script: tempel seluruh berkas ini, Simpan, lalu
 *    Deploy > New deployment > Web app:
 *      Execute as : Me
 *      Who has access : Anyone
 * 4. Salin URL Web App hasil deploy, tempel ke kolom
 *    "URL Deployment Web App" di halaman papan skor, lalu Tes Koneksi.
 *
 * API yang dilayani:
 *  GET  ?action=ping       -> {status:"pong", ok:true}
 *  GET  ?action=getScores  -> {status:"success", teams:[{name, score}, ...]}
 *  POST {action:"updateScores", teams:[{name, score}, ...]}
 *       -> {status:"success", ok:true, updated:3}
 *
 * Isi Sheet "Skor": kolom A = Nama Regu, kolom B = Skor, kolom C = Terakhir Diperbarui.
 */

var SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET';
var SHEET_NAME = 'Skor';

function jawab_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetSkor_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['Nama Regu', 'Skor', 'Terakhir Diperbarui']]);
  }
  return sh;
}

function bacaSkor_() {
  var sh = sheetSkor_();
  var last = sh.getLastRow();
  var teams = [];
  if (last >= 2) {
    var data = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var nama = String(data[i][0] || '').trim();
      if (!nama) continue;
      teams.push({ name: nama, score: Number(data[i][1]) || 0 });
    }
  }
  // Bila Sheet masih kosong, kembalikan 3 regu awal supaya papan tidak kosong.
  if (teams.length === 0) {
    teams = [
      { name: 'REGU A', score: 0 },
      { name: 'REGU B', score: 0 },
      { name: 'REGU C', score: 0 }
    ];
  }
  return teams;
}

function tulisSkor_(teams) {
  var sh = sheetSkor_();
  // Kosongkan isi lama (kecuali header), lalu tulis yang baru.
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  }
  var waktu = new Date();
  var rows = [];
  for (var i = 0; i < teams.length; i++) {
    var nama = String(teams[i].name || teams[i].nama || '').trim();
    if (!nama) continue;
    rows.push([nama, Number(teams[i].score) || 0, waktu]);
  }
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  return rows.length;
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'ping') {
      return jawab_({ status: 'pong', ok: true });
    }
    if (action === 'getScores') {
      return jawab_({ status: 'success', ok: true, teams: bacaSkor_() });
    }
    return jawab_({ status: 'success', ok: true, info: 'Live Skor LCC aktif. Gunakan action=ping atau action=getScores.' });
  } catch (err) {
    return jawab_({ status: 'error', ok: false, pesan: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (body.action === 'updateScores' && Array.isArray(body.teams)) {
      var n = tulisSkor_(body.teams);
      return jawab_({ status: 'success', ok: true, updated: n });
    }
    return jawab_({ status: 'error', ok: false, pesan: 'Aksi tidak dikenal. Kirim {action:"updateScores", teams:[{name, score}]}.' });
  } catch (err) {
    return jawab_({ status: 'error', ok: false, pesan: String(err) });
  }
}
