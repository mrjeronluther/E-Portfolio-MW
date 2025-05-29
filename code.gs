
function getSheetData(sheetName) {
  const spreadsheet = SpreadsheetApp.openById("1GLdze5owg9I3QdaaHfzd89it0ZvkKkufMcKlO-tzvnY");
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found.`);
    return null;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return null;

  const range = sheet.getRange(1, 1, lastRow, 62); // A to BJ
  const data = range.getValues();

  const cleanedData = data.map(row => row.map(cell => {
    if (cell instanceof Date) {
      return Utilities.formatDate(cell, Session.getScriptTimeZone(), "MMM d, yyyy hh:mm a");
    }
    return String(cell).trim();
  }));

  const filteredData = cleanedData.filter(row => {
    const statusColA = row[0];
    const statusColF = row[5];
    const colJ = row[9]; // Column J is index 9

    // Exclude if column A is APPROVED or DISAPPROVED
    if (statusColA === "APPROVED" || statusColA === "DISAPPROVED") return false;
    if (statusColF === "APPROVED" || statusColF === "DISAPPROVED") return false;
    console.log("statusColF:", statusColF);


    


    // Exclude if column J is empty
    if (!colJ) return false;

    // Include row only if it has any non-empty cell
    return row.some(cell => cell !== "");
  }).map(row => row.slice(4)); // Show data from column E onward

  Logger.log('Filtered Data: ' + JSON.stringify(filteredData));
  return filteredData.length > 0 ? filteredData : null;
}




function updateRowStatusWithValidation(rowData, status, reason, approvalReason, sheetName) {
  const spreadsheet = SpreadsheetApp.openById("1GLdze5owg9I3QdaaHfzd89it0ZvkKkufMcKlO-tzvnY");
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const data = sheet.getDataRange().getValues();

  //Match using reference number (column J = index 9)
  const referenceNumberToMatch = String(rowData[0]).trim();
  Logger.log(`Looking for reference: "${referenceNumberToMatch}" in sheet "${sheetName}"`);

  const timestamp = new Date();
  const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MMM d, yyyy HH:mm");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const referenceNumber = String(row[9]).trim(); // Column J in target sheet

    Logger.log(`Row ${i + 1} - Checking reference: "${referenceNumber}"`);

    if (referenceNumber === referenceNumberToMatch) {
      const rowToUpdate = i + 1;

      // Update columns
      const statusCol = 1;
      const reasonCol = 2;
      const timeCol = 3;
      const intervalCol = 4;
      const remarksCol = 2;

      // Decide which to save based on status
      const remarks = status === "DISAPPROVED" ? reason : approvalReason;

      // Then save only one
      sheet.getRange(rowToUpdate, remarksCol).setValue(remarks || '');

      sheet.getRange(rowToUpdate, statusCol).setValue(status);

      sheet.getRange(rowToUpdate, timeCol).setValue(formattedTimestamp);

      const colBValue = row[1];
      if (colBValue) {
        const startDate = new Date(colBValue);
        const diffTime = timestamp - startDate;
        const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        sheet.getRange(rowToUpdate, intervalCol).setValue(diffDays);
      } else {
        sheet.getRange(rowToUpdate, intervalCol).setValue('');
      }

      Logger.log('Match found and row updated.');
      return true;
    }
  }

  Logger.log('No match found for reference number: "' + referenceNumberToMatch + '"');
  throw new Error('No matching row found to update.');
}

// NEW Multi-user Authentication
function authenticateUser(username, password) {
  const users = {
    'GMC': { password: 'pass1', displayName: 'Graham Coates', sheet: 'GMC' },
    'RCS': { password: 'pass2', displayName: 'Rosalyn Segura', sheet: 'RCS' },
    'MGL': { password: 'pass3', displayName: 'Michael Lao', sheet: 'MGL' },
    'MLP': { password: 'pass3a', displayName: 'Louise Piamonte', sheet: 'MGL' },
    'KMV': { password: 'pass4', displayName: 'Tinay Villanueva', sheet: 'KMV' },
    'LDA': { password: 'pass5', displayName: 'Lorence Aurelio', sheet: 'LDA' },
    'JAYE': { password: 'pass6', displayName: 'Jaye Pizarro', sheet: 'JAYE' },
    'JOY': { password: 'pass7', displayName: 'Jocelyn Melitante', sheet: 'JOY' },
    'ALEX': { password: 'pass8', displayName: 'Alex Flores', sheet: 'ALEX' },
    'AU': { password: 'pass9', displayName: 'Aurora Palostero', sheet: 'AU' },
    'MAR': { password: 'pass10', displayName: 'Mariano Caleja', sheet: 'MAR' },
    'EBA': { password: 'pass11', displayName: 'Ernesto Andrade', sheet: 'EBA' },
    'DSM': { password: 'pass12', displayName: 'Dustin Sta. Maria', sheet: 'DSM' },
    'JAC': { password: 'pass13', displayName: 'Jenel Ann Cruzgarcia', sheet: 'JAC' },
    'DP': { password: 'pass14', displayName: 'Doreen Penilla', sheet: 'DP' },
    'FA': { password: 'pass15', displayName: 'Frances Armian', sheet: 'FA' },
    'DM': { password: 'pass16', displayName: 'Denisse Malong', sheet: 'DM' },
    'KL': { password: 'pass17', displayName: 'Kevin Lin', sheet: 'KL' },
    'JM': { password: 'pass18', displayName: 'Juvi Masilungan', sheet: 'JM' },
    'JG': { password: 'pass19', displayName: 'Jhoanalyn Gatdula', sheet: 'JG' },
    'MA': { password: 'pass20', displayName: 'Mary Arceo', sheet: 'MA' },
    'JLC': { password: 'luther2024', displayName: 'Jeron Luther Castro', sheet: 'GMC' },
  };

  const user = users[username];
  if (user && user.password === password) {
    return { displayName: user.displayName, sheet: user.sheet };
  }
  return null;
}
function submitApproverValues(data) {
  const spreadsheet = SpreadsheetApp.openById("1GLdze5owg9I3QdaaHfzd89it0ZvkKkufMcKlO-tzvnY");
  const sheet = spreadsheet.getSheetByName("Conso");

  if (!sheet) throw new Error("Sheet 'Conso' not found.");

  const referenceNumber = data.referenceNumber;
  const rawApproverValues = data.approverValues;

  Logger.log("Incoming data: %s", JSON.stringify(data));

  // Normalize: convert string values to objects { label, value }
  const normalizedApprovers = (rawApproverValues || []).map(val => {
    if (typeof val === 'string') {
      return { label: val, value: val };
    } else if (val && typeof val === 'object' && val.value) {
      return val;
    } else {
      return null;
    }
  }).filter(val => val && typeof val.value === 'string' && val.value.trim() !== "");

  if (!referenceNumber || normalizedApprovers.length === 0) {
    throw new Error("Submission aborted: No valid approvers selected or reference number is missing.");
  }

  const lastRow = sheet.getLastRow();
  const refColumn = sheet.getRange("E2:E" + lastRow).getValues();
  const startRow = 2;

  const matchRowIndex = refColumn.findIndex(row => row[0] === referenceNumber);
  Logger.log("Searching for reference number: " + referenceNumber);

  if (matchRowIndex === -1) {
    throw new Error(`Reference number '${referenceNumber}' not found in column E.`);
  }

const actualRow = matchRowIndex + startRow;
Logger.log("Matched reference number at row: " + actualRow);

const targetRange = sheet.getRange(actualRow, 16, 1, 46); // P (16) to BI (61) = 61 - 16 + 1 = 46 columns
const rowData = targetRange.getValues()[0];
Logger.log("Existing row data (P:BI): %s", JSON.stringify(rowData));

// Check for duplicates
const duplicates = normalizedApprovers.filter(a => rowData.includes(a.value));
if (duplicates.length > 0) {
  const duplicateLabels = duplicates.map(a => a.label || a.value);
  throw new Error(`The following approver(s) are already in the list: ${duplicateLabels.join(", ")}`);
}

// Check for space
const emptySlots = rowData.filter(val => !val || val.toString().trim() === "").length;
if (emptySlots === 0) {
  throw new Error("All approvers are already included — no empty slots in P:BI.");
}


  // Insert values
  let insertIndex = 0;
  for (let approver of normalizedApprovers) {
    while (insertIndex < rowData.length && rowData[insertIndex]) {
      insertIndex++;
    }

    if (insertIndex >= rowData.length) {
      throw new Error("No available columns (P:BI) left to insert new approvers.");
    }

    Logger.log("Inserting approver '%s' at column offset: %s", approver.value, insertIndex);
    rowData[insertIndex] = approver.value;
    insertIndex++;
  }

  sheet.getRange(actualRow, 16, 1, 46).setValues([rowData]);
  Logger.log("✅ Successfully wrote approvers to sheet.");
  
}






