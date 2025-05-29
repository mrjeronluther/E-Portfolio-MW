
function processFormData(data) {
  const folderId = '1A_YUkMwWikOLq0kxQ3WY-f4M9IoJkWoV';
  const folder = DriveApp.getFolderById(folderId);

  // Upload primary file
  const blob = Utilities.newBlob(Utilities.base64Decode(data.fileData), MimeType.PDF, data.fileName);
  const file = folder.createFile(blob);
  const fileUrl = file.getUrl();

  let revisedFileName = '';
  let revisedFileUrl = '';

  // If resubmission, upload revised file
  if (data.isResubmission && data.revisedFileData && data.revisedFileName) {
    const revisedBlob = Utilities.newBlob(
      Utilities.base64Decode(data.revisedFileData),
      MimeType.PDF,
      data.revisedFileName
    );
    const revisedFile = folder.createFile(revisedBlob);
    revisedFileName = data.revisedFileName;
    revisedFileUrl = revisedFile.getUrl();
  }

  const now = new Date();
  const refNumber = generateReferenceNumber();

  const timestamp = now;
  const name = data.name;
  const email = data.email;
  const email1 = data.email1;
  const memoType = data.memoType;
  const dateType = data.dateType;
  const subjectmemo = data.additionalMemoInfo;
  const addprop = data.otherExtraOption;

  const formattedDate = Utilities.formatDate(new Date(dateType), Session.getScriptTimeZone(), "MMM d, yyyy");

  let rowData = [
    refNumber,
    timestamp,
    name,
    email,
    memoType,
    email1,
    subjectmemo,
    dateType,
    data.fileName,
    fileUrl
  ];

  rowData.push(...data.additionalOptions);
  rowData.push(addprop);
  rowData.push(...data.departments);
  rowData.push(...data.approvers);

  // Push revised file name and URL if available
  if (revisedFileName && revisedFileUrl) {
    rowData.push(revisedFileName);
    rowData.push(revisedFileUrl);
  }

  // Limit total columns to 58 (adjust if needed)
  rowData = rowData.slice(0, 60);

  const spreadsheet = SpreadsheetApp.openById('1GLdze5owg9I3QdaaHfzd89it0ZvkKkufMcKlO-tzvnY');
  const sheet = spreadsheet.getSheetByName('Conso');

  if (!sheet) {
    throw new Error('Sheet named "Conso" does not exist.');
  }

  // ==== Check for spreadsheet size limit ====
  if (isSheetNearLimit(sheet, rowData.length)) {
    throw new Error('Cannot save data: Spreadsheet has reached or is near the 10 million cell limit.');
  }

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(1, 5, lastRow + 1000, 60); // Start from column E
  const values = range.getValues();

  let targetRow = 0;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const isEmpty = row.every(cell => cell === '' || cell === null);
    if (isEmpty) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === 0) {
    targetRow = lastRow + 1;
  }

  try {
    sheet.getRange(targetRow, 5, 1, rowData.length).setValues([rowData]); // Start from column E
  } catch (err) {
    if (err.message.includes('above the limit of 10000000')) {
      throw new Error('Cannot save data: Spreadsheet has exceeded the 10 million cell limit.');
    } else {
      throw err;
    }
  }

  // === EMAIL CONFIRMATION ===
  const subject = `📨 Memo Submitted - Reference: ${refNumber}`;
  const htmlMessage = `
    <div style="font-family: Arial, sans-serif; padding: 10px;">
      <table width="100%" style="margin-bottom: 20px;">
        <tr>
          <td align="center">
            <img src="https://images.contentstack.io/v3/assets/blt827157d7af7bc6d4/blt539b6710c5d23513/63306748763d011cd3e925bd/megaworld-logo.png" 
                 alt="Header Image" 
                 style="max-width: 600px; width: 100%; height: auto; display: block;">
          </td>
        </tr>
      </table>

      <p>Hi <strong>${name}</strong>,</p>
      <p>Your memo has been <span style="color:green;"><strong>successfully submitted</strong></span>.</p>

      <table style="border-collapse: collapse; margin-top: 10px;">
        <tr><td style="padding: 4px;"><strong>Subject:</strong></td><td style="padding: 4px;">${subjectmemo}</td></tr>
        <tr><td style="padding: 4px;"><strong>Date of Memo:</strong></td><td style="padding: 4px;">${formattedDate}</td></tr>
        <tr><td style="padding: 4px;"><strong>Preparer of Memo:</strong></td><td style="padding: 4px;">${memoType}</td></tr>
        <tr><td style="padding: 4px;"><strong>Reference No.:</strong></td><td style="padding: 4px; font-weight: bold;">${refNumber}</td></tr>
      </table>

      <p><i>Please save this reference number for your records.</i></p>
      <p style="margin-top: 16px;">You may access the uploaded file <a href="${fileUrl}" target="_blank">here</a>.</p>
      
      <p style="margin-top: 30px; text-align: center;">
        <hr style="margin-top: 30px; border: none; border-top: 1px solid #eee;">
        <small style="color: #888;">This is an automated message. Please do not reply to this email.</small>
      </p>
    </div>
  `;

  const recipients = [email, email1].filter(e => e).join(',');
  if (recipients) {
    MailApp.sendEmail({
      to: email,
      cc: email1,
      subject: subject,
      htmlBody: htmlMessage,
      name: "MCD Memo Submission Notification"
    });
  }

  return refNumber;
}

// ==== Helper function to check cell usage ====
function isSheetNearLimit(sheet, rowDataLength) {
  const spreadsheet = sheet.getParent();
  const totalCells = spreadsheet.getSheets()
    .map(s => s.getMaxRows() * s.getMaxColumns())
    .reduce((a, b) => a + b, 0);

  const newCells = 1 * rowDataLength; // Writing 1 row
  return (totalCells + newCells >= 10000000);
}

// Global Set to store already used reference numbers (in-memory for this example)
const usedReferenceNumbers = new Set();

// Generate a unique reference number like REF#******, where "******" is a 6 alphanumeric string
function generateReferenceNumber() {
  let randomStr;

  // Ensure uniqueness by checking if the generated string already exists
  do {
    randomStr = Math.random().toString(36).substring(2, 8).toUpperCase(); // 6 alphanumeric characters
  } while (usedReferenceNumbers.has(randomStr)); // Keep generating until unique

  usedReferenceNumbers.add(randomStr);

  return `REF#${randomStr}`;
}
