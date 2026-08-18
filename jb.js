// ========================================
// GOOGLE SPREADSHEET CONFIGURATION
// ========================================

// If Code.gs is attached to the Google Sheet,
// you can leave this empty.
const SPREADSHEET_ID = "1ePUAoJLAkNi_91akGOdw2mpnqcFSk46_e5vZ_-WaU2g";

// Name of your spreadsheet sheet/tab
const SHEET_NAME = "Messages";


// ========================================
// OPEN SPREADSHEET
// ========================================

function getSpreadsheet() {

  // If this Apps Script is bound to the spreadsheet
  if (SPREADSHEET_ID === "") {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  // If using a standalone Apps Script project
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}


// ========================================
// OPEN MESSAGE SHEET
// ========================================

function getMessageSheet() {

  const spreadsheet = getSpreadsheet();

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  // Create sheet automatically if it does not exist
  if (!sheet) {

    sheet = spreadsheet.insertSheet(SHEET_NAME);

    sheet.appendRow([
      "ID",
      "Encrypted Message",
      "Cipher",
      "Date/Time"
    ]);
  }

  return sheet;
}


// ========================================
// WEB APP
// ========================================

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Cipher Message Board")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


// ========================================
// INCLUDE HTML FILES
// ========================================

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}


// ========================================
// SAVE MESSAGE TO GOOGLE SHEETS
// ========================================

function saveMessage(encryptedMessage) {

  try {

    if (
      !encryptedMessage ||
      encryptedMessage.trim() === ""
    ) {

      return {
        success: false,
        message: "Encrypted message cannot be empty."
      };

    }


    const sheet = getMessageSheet();


    // Generate unique ID
    const id =
      Utilities.getUuid();


    // Current date and time
    const date =
      new Date();


    // Save data to spreadsheet
    sheet.appendRow([
      id,
      encryptedMessage,
      "Vigenere",
      date
    ]);


    return {

      success: true,

      message:
        "Message successfully saved to Google Sheets.",

      id: id

    };

  }

  catch (error) {

    console.error(error);

    return {

      success: false,

      message:
        "Unable to save message: " +
        error.message

    };

  }

}


// ========================================
// GET MESSAGES FROM GOOGLE SHEETS
// ========================================

function getMessages() {

  try {

    const sheet =
      getMessageSheet();


    const lastRow =
      sheet.getLastRow();


    // Only header exists
    if (lastRow <= 1) {
      return [];
    }


    const data =
      sheet.getRange(
        2,
        1,
        lastRow - 1,
        4
      ).getValues();


    return data.map(function(row) {

      return {

        id: row[0],

        encryptedMessage:
          row[1],

        cipher:
          row[2],

        date:
          row[3]

      };

    });

  }

  catch (error) {

    console.error(error);

    throw new Error(
      "Unable to retrieve messages from Google Sheets."
    );

  }

}
