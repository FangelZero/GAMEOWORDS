function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  const rows = sheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.user && rows[i][1] === data.categoria) {
      sheet.getRange(i + 1, 3, 1, 5).setValues([data.valores]);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([data.user, data.categoria, ...data.valores]);
  }

  return ContentService.createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const user = e.parameter.user;
  const categoria = e.parameter.categoria;

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === user && rows[i][1] === categoria) {
      return ContentService.createTextOutput(
        JSON.stringify(rows[i].slice(2, 7))
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(
    JSON.stringify(["", "", "", "", ""])
  ).setMimeType(ContentService.MimeType.JSON);
}