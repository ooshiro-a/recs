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

  // ルートフォルダを取得または作成
  var rootName = 'Rec Reports';
  var rootIter = DriveApp.getFoldersByName(rootName);
  var rootFolder = rootIter.hasNext() ? rootIter.next() : DriveApp.createFolder(rootName);

  // 記録ごとのフォルダを作成
  var folderName = rec.site + '_' + rec.date;
  var recFolder = rootFolder.createFolder(folderName);
  recFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 写真を保存
  for (var i = 0; i < photos.length; i++) {
    var b64 = photos[i].replace(/^data:image\/\w+;base64,/, '');
    var raw = Utilities.base64Decode(b64);
    var blob = Utilities.newBlob(raw, 'image/jpeg', '写真' + (i + 1) + '.jpg');
    recFolder.createFile(blob);
  }

  // 点検記録テキストを保存
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
  recFolder.createFile(lines.join('\n'), '点検記録.txt', 'text/plain');

  // スプレッドシートに記録
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  sh.appendRow([
    rec.id, rec.date, rec.site, rec.model, rec.type, rec.refrigerant,
    rec.hp, rec.lp, rec.td, rec.ts, rec.amb, rec.indoor, rec.amp,
    rec.note, recFolder.getUrl()
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ok: true, url: recFolder.getUrl()}))
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
