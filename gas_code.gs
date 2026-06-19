function doPost(e) {
  var action = e.parameter.action;
  if (action === 'report') {
    return createReport(e);
  }
  return saveRecord(e);
}

function createReport(e) {
  var body = JSON.parse(e.postData.contents);
  var rec = body.record;
  var photos = body.photos || [];

  var folderName = 'Rec Reports';
  var iter = DriveApp.getFoldersByName(folderName);
  var folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(folderName);

  var docTitle = rec.site + '_' + rec.date;
  var doc = DocumentApp.create(docTitle);
  var b = doc.getBody();

  var kinds = {ice:'製氷機', cooler:'冷蔵庫', freezer:'冷凍庫', ac:'エアコン', other:'その他'};

  b.appendParagraph('冷凍設備点検記録').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  b.appendParagraph(rec.site + '　' + rec.date);
  b.appendParagraph('');

  b.appendTable([
    ['現場名',       rec.site],
    ['日付',         rec.date],
    ['機種名',       rec.model],
    ['機器種別',     kinds[rec.type] || rec.type],
    ['冷媒',         rec.refrigerant],
    ['外気温',       rec.amb + ' ℃'],
    ['庫内温度',     rec.indoor + ' ℃'],
    ['高圧（吐出）', rec.hp + ' MPa'],
    ['低圧（吸入）', rec.lp + ' MPa'],
    ['吐出温度',     rec.td + ' ℃'],
    ['吸入温度',     rec.ts + ' ℃'],
    ['運転電流',     rec.amp + ' A'],
    ['メモ',         rec.note || '']
  ]);

  if (photos.length > 0) {
    b.appendParagraph('');
    b.appendParagraph('現場写真').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    for (var i = 0; i < photos.length; i++) {
      var b64 = photos[i].replace(/^data:image\/\w+;base64,/, '');
      var raw = Utilities.base64Decode(b64);
      var blob = Utilities.newBlob(raw, 'image/jpeg', 'photo' + i + '.jpg');
      var img = b.appendImage(blob);
      var origW = img.getWidth();
      var origH = img.getHeight();
      if (origW > 400) {
        img.setWidth(400);
        img.setHeight(Math.round(origH * 400 / origW));
      }
      b.appendParagraph('');
    }
  }

  doc.saveAndClose();

  var pdf = doc.getAs('application/pdf');
  pdf.setName(docTitle + '.pdf');
  var pdfFile = folder.createFile(pdf);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  DriveApp.getFileById(doc.getId()).setTrashed(true);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  sh.appendRow([
    rec.id, rec.date, rec.site, rec.model, rec.type, rec.refrigerant,
    rec.hp, rec.lp, rec.td, rec.ts, rec.amb, rec.indoor, rec.amp,
    rec.note, pdfFile.getUrl()
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ok: true, url: pdfFile.getUrl()}))
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
