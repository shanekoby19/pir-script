/**
 * Returns the FIRST attachment of the message with the same subject being passed in.
 * @param messageSubject - A string representing the subject of the message you want to retrieve the attachment from.
 * @param daysOffset (optional) - The number of days you want to look in the past. Defaults to yesterday if not specified.
 * @param startPos (optional) - The starting position to begin searching your email threads. Defaults to the first email in your inbox.
 * @param endPos (optional) - The position at which you want to stop searching your email threads. Defaults to the last email thread on the first page of your inbox.
 */
 function getAttachment(messageSubject, daysOffset=1, startPos=0, endPos=50) {
    // determine the timestamp for today.
    const cutOffTimeStamp = Date.now() - (daysOffset * 24 * 60 * 60 * 1000);
  
    // Get only the first 50 threads from your inbox.
    const thread = GmailApp.getInboxThreads(startPos, endPos);
    let messages = [];
    thread.forEach(gmailThread => messages.push(...gmailThread.getMessages()));
  
    // Find a message with the given subject that was received within the last 24 hours.
    const message = messages.find(message => message.getSubject() === messageSubject && message.getDate().getTime() >= cutOffTimeStamp);
    //message.moveToTrash();
    return message.getAttachments()[0];
  }
  
  /**
   * This function stores a file in your email on the drive.
   * @param msgSubject - The subject of the message containing the .xlsx attachment
   * @param fileName - The name you want the file to have on the drive. No extension is necessary.
   */
  function storeFileOnDrive(attachment, fileName) {
    // Location the Automated - MBI folder on the users drive or create it if it doesn't exists.
    const convertedSheetFolders = DriveApp.getFoldersByName(`Avance Health Dashboard Reports`);
    const folderFound = convertedSheetFolders.hasNext();
    let convertedSheetFolder = folderFound ? convertedSheetFolders.next() : DriveApp.createFolder(`Avance Health Dashboard Reports`);
  
    // Create a new .xlsx file on your drive.
    const xlsxFile = DriveApp.createFile(attachment.copyBlob());
    xlsxFile.setName(fileName);
    xlsxFile.moveTo(convertedSheetFolder);
  }
  
  /**
   * Written by Amit Agarwal
   * A function to convert an .xlsx file on your drive to a google spreadsheet.
   * @param fileName - The name you want the file to have on the drive. No extension is necessary.
   */
  function convertExceltoGoogleSpreadsheet(fileName) {
    try {
  
      // Written by Amit Agarwal
      // www.ctrlq.org;
  
      // Find and store the excel file from your drive with the given name.
      const excelFile = DriveApp.getFilesByName(fileName).next();
      const fileId = excelFile.getId();
      const folderId = Drive.Files.get(fileId).parents[0].id;
      const blob = excelFile.getBlob();
      // Build a resource object with the parent folder to pass to Drive.Files.insert();
      const resource = {
        title: excelFile.getName(),
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [{id: folderId}],
      };
  
      Drive.Files.insert(resource, blob);
  
    } catch (f) {
      Logger.log(`Error: f.toString()`); // Logs an error message if it occurs.
    }
  }
  
  /**
   * Writes a file from your drive to the spreadsheet.
   * @param fileName - The name of the file you want to find on the drive. No extension is necessary.
   */
  
  function writeFileToSpreadsheet(sheet, fileName) {
    const googleSheet = DriveApp.getFilesByName(fileName).next();
    const fileId = googleSheet.getId();
  
    sheet.clear();
  
    const sourceSheet = SpreadsheetApp.openById(fileId).getSheets()[0];
    const sourceValues = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()).getValues();
  
    sheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()).setValues(sourceValues);
  
    copyPasteFormatting(sourceSheet, sheet);
  } 
  
  
  
  /**
   * Cleans a data sheet by deleting meta data from top and bottom.
   * @param sheetName - A string with the same name as the google sheet to clean.
   * @param top (optional) - The number of rows to delete from the top of the sheet, if not specified no rows will be deleted from the top.
   * @param bottom (optional) - The number of rows to delete from the bottom of the sheet, if not specified no rows will be deleted from the bottom.
   */
  function removeMetaData(sheet, top=0, bottom=0) {
    if (sheet.getLastColumn() >= 3) {
      top !== 0 ? sheet.deleteRows(1, top) : null;
      bottom !== 0 ? sheet.deleteRows(sheet.getLastRow() - bottom + 1, bottom) : null;
    }
    else {
      sheet.clear();
    }
  }
  
  /**
   * Copy and paste formatting of data from of sheet to another. Takes all rows and columns of source sheet and copies there formats to all rows and columns of dest sheet.
   * @param sourceSheet - The sheet you want to copy the formatting from.
   * @param destSheet - The sheet you want to paste the formatting to.
   */
  function copyPasteFormatting(sourceSheet, destSheet) {
     // Get data and formatting from the source sheet
    const range = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn());
    const linkRange = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn());
    // Set the link range to plain text so that all numeric value links will pull over.
    linkRange.setNumberFormat(`@`);
  
    const background = range.getBackgrounds();
  
    // Put data and formatting in the destination sheet
    const destRange = destSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn());
    const destLinkRange = destSheet.getRange(1, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn());
  
    // Copies and all the link values on the sheet to a two dimensional array. 
    const richTextValues = linkRange.getRichTextValues();
    linkRange.getValues().forEach((r, i) => {
      r.forEach((c, j) => {
        if (typeof c == "number") {
          richTextValues[i][j] = SpreadsheetApp.newRichTextValue().setText(c).setLinkUrl(richTextValues[i][j].getLinkUrl()).build();
        }
      });
    });
    // Sets the link values that were copied to selected range. In our case the entire sheet.
    destLinkRange.setRichTextValues(richTextValues);
    destRange.setBackgrounds(background);
  }
  