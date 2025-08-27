function generateQRCodesOnSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  var name = sheet.getRange(lastRow, 2).getValue();     // Column B = Name
  var email = sheet.getRange(lastRow, 3).getValue();    // Column C = Email
  var phone = sheet.getRange(lastRow, 4).getValue();    // Column D = WhatsApp
  var people = sheet.getRange(lastRow, 5).getValue();   // Column E = No. of People

  // Create formatted text for QR
  var qrData = "Name: " + name +
               "\nEmail: " + email +
               "\nPhone: " + phone +
               "\nNo. of People: " + people;

  var url = "https://api.qrserver.com/v1/create-qr-code/?data=" 
            + encodeURIComponent(qrData) + "&size=200x200";

  // Put QR in Column H
  sheet.getRange(lastRow, 8).setFormula('=IMAGE("' + url + '")');
  
  // Generate sequential ID for Column I
  var idColumn = 9;
  var prevId = sheet.getRange(lastRow - 1, idColumn).getValue();
  var newId;

  if (prevId && /^VMID\d+$/.test(prevId)) {
    var num = parseInt(prevId.replace("VMID", ""), 10) + 1;
    newId = "VMID" + num;
  } else {
    newId = "VMID1"; // First ID
  }

  sheet.getRange(lastRow, idColumn).setValue(newId);
}

