function doPost(e) {
  try {
    var action = e.parameter.action;
    if (action === 'report') {
      return createReport(e);
    }
    return saveRecord(e);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, err: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createReport(e) {
  var parsed = JSON.parse(e.postData.contents);
  var rec = parsed.record;
  var photos = parsed.photos || [];

  // 保存先フォルダをIDで直接指定
  var folder = DriveApp.getFolderById('1zBJ37OWogXNedv3s9hk9Ka8x6WaPumCL');

  // ファイル名：日付の「/」を「-」に変換してから 日付_現場名_機種名
  var dateStr = rec.date.replace(/\//g, '-');
  var baseName = dateStr + '_' + rec.site + '_' + rec.model;

  // 写真を保存（ファイル名: 日付_現場名_機種名_写真1.jpg）
  for (var i = 0; i < photos.length; i++) {
    var b64 = photos[i].replace(/^data:image\/\w+;base64,/, '');
    var raw = Utilities.base64Decode(b64);
    var blob = Utilities.newBlob(raw, 'image/jpeg', baseName + '_写真' + (i + 1) + '.jpg');
    folder.createFile(blob);
  }

  // 点検記録テキストを保存
  // ★修正: createFile(ファイル名, 本文内容, MIMEタイプ) の順
  var kinds = {ice: '製氷機', cooler: '冷蔵庫', freezer: '冷凍庫', ac: 'エアコン', other: 'その他'};
  var lines = [
    '【冷凍設備点検記録】',
    '',
    '現場名　: ' + rec.site,
    '日付　　: ' + rec.date,
    '機種名　: ' + rec.model,
    '機器種別: ' + (kinds[rec.type] || rec.type),
    '冷媒　　: ' + rec.refrigerant,
    '',
    '--- 計測値 ---',
    '外気温　: ' + rec.amb + ' ℃',
    '庫内温度: ' + rec.indoor + ' ℃',
    '高圧　　: ' + rec.hp + ' MPa',
    '低圧　　: ' + rec.lp + ' MPa',
    '吐出温度: ' + rec.td + ' ℃',
    '吸入温度: ' + rec.ts + ' ℃',
    '電流　　: ' + rec.amp + ' A',
    'メモ　　: ' + (rec.note || ''),
  ];
  var textFile = folder.createFile(baseName + '.txt', lines.join('\n'), 'text/plain');

  // スプレッドシートに記録
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  sh.appendRow([
    rec.id, rec.date, rec.site, rec.model, rec.type, rec.refrigerant,
    rec.hp, rec.lp, rec.td, rec.ts, rec.amb, rec.indoor, rec.amp,
    rec.note, textFile.getUrl()
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ok: true, url: textFile.getUrl()}))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveRecord(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  sh.appendRow([
    data.id, data.date, data.site, data.model, data.type, data.refrigerant,
    data.hp, data.lp, data.td, data.ts, data.amb, data.indoor, data.amp,
    data.note, ''
  ]);
  return ContentService.createTextOutput('ok');
}

function testAuth() {
  DriveApp.getFolderById('1zBJ37OWogXNedv3s9hk9Ka8x6WaPumCL');
  Logger.log('OK');
}
