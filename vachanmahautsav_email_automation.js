function sendConfirmationEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2;  // Assuming row 1 is header
  var numRows = sheet.getLastRow() - 1;
  var dataRange = sheet.getRange(startRow, 2, numRows, 9); // From column B to I
  var data = dataRange.getValues();
  
  var subject = "Vachan Mahautsav 2025 Registration Confirmation ✅ ";
  var imageUrl = "https://drive.google.com/file/d/1axjDcnmStoxcNHUF5HZc9asTIOKh8KWT/view?usp=sharing"; 

  // File ID of the image stored in Google Drive
  var fileId = "1yZ-SYNPGlvwK15Fp1b9D3fRCiRqeYTAz"; 
  var file = DriveApp.getFileById(fileId);

  for (var i = 0; i < data.length; i++) {
    var name = data[i][0];// Column B
    var emailAddress = data[i][1]; // Column C
    var phone = data[i][2];    // Column D = WhatsApp
    var numPeople = data[i][3];   // Column E = No. of People
    var vmidseq = data[i][7]; // Column I 
    var confirmStatus = data[i][8]; // Column I (8th col in this range)
    var sentStatus = sheet.getRange(startRow + i, 11).getValue(); // Column J (10th col in sheet)
    var qrData = `ID: ${vmidseq}\nName: ${name}\nEmail: ${emailAddress}\nPhone: ${phone}\nNo. of People: ${numPeople}`;
    var qrUrl = "https://api.qrserver.com/v1/create-qr-code/?data=" + encodeURIComponent(qrData) + "&size=200x200";

    if (confirmStatus === "CONFIRM" && sentStatus !== "SENT" && emailAddress) {
      var message = `
        <p>Hello ${name},</p>
        <p>Praise the Lord!</p>
        <p>We’re excited to confirm your registration for <strong>VACHAN MAHAUTSAV 2025</strong></p>
        <p>Attached to this email, you’ll find your <strong>Entry Pass</strong> for the event.</p>
        <p>Additionally, we’ve sent your personal <strong>QR code</strong> below. You can either:</p>
        <ul>
          <li>Take a screenshot and save it on your phone, or</li>
          <li>Show this email at the entrance</li>
        </ul>

        <p><img src="${qrUrl}" width="200"></p>

        <p>We are delighted to have you join us for Vahan Mahautsav. May God Bless you and your family through this event.</p>

        <p>Best regards,<br>
        NLF CBD Zone Media Team</p>
      `;

      GmailApp.sendEmail(emailAddress, subject, "", {
        htmlBody: message,
        attachments: [file.getAs(MimeType.PNG)] // or JPG/PDF
      });

      // Mark column J as "SENT"
      sheet.getRange(startRow + i, 11).setValue("SENT");
    }
  }
}
