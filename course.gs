/*
Course

0. To create a project, go to: https://script.google.com
Creating a project from Google Drive:
https://developers.google.com/apps-script/guides/projects
1. Attachments need to be uploaded in Google Drive root folder manually.  
2. The Administrator of the script has to build the application (setup).
3. Corresponding Google Services need to be authorized.

A maximum number of attendees of the course is set up (safeNumberOFAttendees).
The script creates a folder structure and necessary files
(createFoldersFiles).
Attachments (getAttachmentIds) are copied in the created folder and removed
from the root folder.
It also generates a basic website (doGet) with all files in the folder.
It creates a trigger (createOnFormSubmitTrigger) and other triggers
(createEveryDayTrigger) based on the number of attendees.
Redundant triggers are deleted automatically (deleteTrigger).
The script sends emails (sendFirstFile, sendNextFiles) with attachments
based on the form data provided by the attendees of the online course.
The data are stored in the created spreadsheet permanently or temporarily.
For more info go to the corresponding function.

4. The application can be deleted (uninstall).
This removes all triggers (deleteAllTriggers), the folder structure
and files (removeFiles), and script properties (deleteScriptProperties).

*/

// Build the application
function setup() {
  safeNumberOfAttendees(80);
  createFoldersFiles();
  getAttachmentIds('letter');
  doGet();
  createOnFormSubmitTrigger();
  Logger.log('Application successfully installed');
}


// Remove the application
function uninstall() {
  deleteAllTriggers();
  removeFoldersFiles();
  deleteScriptProperties();
  Logger.log('Application successfully uninstalled');
}

/* Set up safeNumberOfAttendees.  Current Gmail daily guota
(permitted number of sent emails): 100.
@param {number} safeNumber Recommended safeNumber value: 80. */
function safeNumberOfAttendees(safeNumber) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var safeNumberOfAttendees = scriptProperties
    .getProperty('safeNumberOfAttendees');
  if (safeNumberOfAttendees === null ||
      Number(safeNumberOfAttendees) !== safeNumber) {
    safeNumberOfAttendees = safeNumber;
    scriptProperties.setProperty('safeNumberOfAttendees',
                                 safeNumberOfAttendees);
    Logger.log('safeNumberOfAttendees property created/set up to ' +
      safeNumberOfAttendees);
  } else {
    Logger.log('safeNumberOfAttendees property is set up correctly');
  }
}


// Create a folder structure, a spreadsheet and a form in Google Drive
function createFoldersFiles() {
  // Create a folder structure and store properties
  var scriptProperties = PropertiesService.getScriptProperties();
  var folderId = scriptProperties.getProperty('folderId');
  Logger.log('folderId = ' + folderId);
  if (folderId === null) {
    var folder = DriveApp.createFolder('course');
    var folderUrl = folder.getUrl();
    scriptProperties.setProperty('folderUrl', folderUrl);
    Logger.log('folderUrl = ' + folderUrl);
    folderId = folder.getId();
    scriptProperties.setProperty('folderId', folderId);
    Logger.log('folderId = ' + folderId);
  } else {
    Logger.log('Folder has already been created');
  }

  // Create a spreadsheet and store properties
  var spreadsheetId = scriptProperties.getProperty('spreadsheetId');
  Logger.log('spreadsheetId = ' + spreadsheetId);
  if (spreadsheetId === null) {
    var spreadsheet = SpreadsheetApp.create('course');
    var spreadsheetUrl = spreadsheet.getUrl();
    scriptProperties.setProperty('spreadsheetUrl', spreadsheetUrl);
    Logger.log('spreadsheetUrl = ' + spreadsheetUrl);
    spreadsheetId = spreadsheet.getId();
    scriptProperties.setProperty('spreadsheetId', spreadsheetId);
    Logger.log('spreadsheetId = ' + spreadsheetId);
    var spreadsheetTemp = DriveApp.getFileById(spreadsheetId);
    DriveApp.getFolderById(folderId).addFile(spreadsheetTemp);
    DriveApp.getRootFolder().removeFile(spreadsheetTemp);
    SpreadsheetApp.flush();
  } else {
    Logger.log('Spreadsheet has already been created');
  }

  // Create a form and store properties
  var formId = scriptProperties.getProperty('formId');
  Logger.log('formId = ' + formId);
  if (formId === null) {
    var form = FormApp.create('course');
    var formUrl = form.getEditUrl();
    // Link the form to the spreadsheet
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
    scriptProperties.setProperty('formUrl', formUrl);
    Logger.log('formUrl = ' + formUrl);
    formId = form.getId();
    scriptProperties.setProperty('formId', formId);
    Logger.log('formId = ' + formId);
    var formTemp = DriveApp.getFileById(formId);
    DriveApp.getFolderById(folderId).addFile(formTemp);
    DriveApp.getRootFolder().removeFile(formTemp);
    // Create fields in the form
    form.setDescription('You can order the online course. ' +
                        'Please fill in and submit the form.\n' +
                        'If you do not get an email, check spam.');
    var firstName = form.addTextItem().setTitle('First Name')
      .setRequired(true);
    var surname = form.addTextItem().setTitle('Surname').setRequired(true);
    var email = form.addTextItem().setTitle('E-mail').setRequired(true);
    var city = form.addTextItem().setTitle('City').setRequired(true);
    var information = form.addCheckboxItem();
    information.setTitle('Information')
      .setChoices([information.createChoice('I would like to get the info ' +
                                            'about future activities ' +
                                            '(if unchecked, you will  be ' +
                                            'removed from the database)')]);
    var agreement = form.addCheckboxItem();
    agreement.setTitle('Collecting personal data and GDPR terms approval')
      .setChoices([agreement.createChoice('I agree with collecting ' +
                                          'my personal data and GDPR terms')])
      .setRequired(true);
    // Writing in the spreadsheet takes more time, it needs some time to wait
    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    // Store mainUrlKeys for later html links building
    var mainUrlKeys = 'spreadsheetUrl, formUrl, folderUrl';
    scriptProperties.setProperty('mainUrlKeys', mainUrlKeys);
    Logger.log('mainUrlKeys = ' + mainUrlKeys);
  } else {
    Logger.log('Form has already been created');
  }

  // Add a new title and format the spreadsheet
  var sheet = spreadsheet.getSheets()[0];
  sheet.setName('processed');
  spreadsheet.deleteSheet(spreadsheet.getSheets()[1]);
  SpreadsheetApp.flush();
  var lastColumn = sheet.getLastColumn();
  var newTitle = 'Letter No. sent';
  Logger.log('lastColumn = ' + lastColumn);
  if (sheet.getRange(1, lastColumn).getValue() != newTitle) {
    var range = sheet.getRange(1, lastColumn + 1);
    var newTitleCell = sheet.setActiveRange(range);
    newTitleCell.setValue(newTitle);
    Logger.log('Title of a new column set');
    SpreadsheetApp.flush();
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 250);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 100);
    sheet.setColumnWidth(lastColumn + 1, 100);
    Logger.log('Columns width set');
    SpreadsheetApp.flush();
    sheet.getRange(1, 1, 1, lastColumn + 1).setFontWeight('bold');
    Logger.log('Title font set to bold');
    SpreadsheetApp.flush();
    var sheet2 = sheet.copyTo(spreadsheet);
    SpreadsheetApp.flush();
    spreadsheet.getSheets()[1].setName('finished');
    SpreadsheetApp.flush();
  } else {
    Logger.log('Spreadsheet formatting has already been set');
  }
}


/* Store all attachment URLs and IDs.  Set up attachmentNames.
@param {string} attachmentNames Used for manipulation and in emails. */
function getAttachmentIds(attachmentNames) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('attachmentNames', attachmentNames);
  var folderId = scriptProperties.getProperty('folderId');
  var files = DriveApp.getRootFolder()
    .searchFiles('title contains "' + attachmentNames + '"');
  var filesArray = [];
  while (files.hasNext()) {
    var file = files.next();
    filesArray.push(file);
    DriveApp.getFolderById(folderId).addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }
  var numberOfAttachments = filesArray.length;
  scriptProperties.setProperty('numberOfAttachments', numberOfAttachments);
    Logger.log('numberOfAttachments = ' + numberOfAttachments);
  if (numberOfAttachments === 0) {
    Logger.log('No attachment containing name "' + attachmentNames +
      '" found');
  } else {
    filesArray.sort();
    Logger.log(filesArray);
    var urlKeys = '';
    var idKeys = '';
    var nameKeys = '';
    for (var i = 0; i < filesArray.length; i++) {
      var fileUrl = filesArray[i].getUrl();
      var fileId = filesArray[i].getId();
      scriptProperties.setProperty(filesArray[i] + 'Url', fileUrl);
      Logger.log(filesArray[i] + 'Url = ' + fileUrl);
      scriptProperties.setProperty(filesArray[i] + 'Id', fileId);
      Logger.log(filesArray[i] + 'Id = ' + fileId);
      if (i == 0) {
        urlKeys += filesArray[i] + 'Url';
        idKeys += filesArray[i] + 'Id';
        nameKeys += filesArray[i];
      } else {
        urlKeys += ', ' + filesArray[i] + 'Url';
        idKeys += ', ' + filesArray[i] + 'Id';
        nameKeys += ', ' + filesArray[i];
      }
    }
    scriptProperties.setProperty('urlKeys', urlKeys);
    Logger.log('urlKeys = ' + urlKeys);
    scriptProperties.setProperty('idKeys', idKeys);
    Logger.log('idKeys = ' + idKeys);
    scriptProperties.setProperty('nameKeys', nameKeys);
    Logger.log('nameKeys = ' + nameKeys);
  }
}


/* To activate the web page, go to File/Manage versions/Save new version.
Deploy version as a web app and open links navigating Publish/Deploy as web
app. */
function doGet() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var mainUrlKeys = scriptProperties.getProperty('mainUrlKeys').split(', ');
  Logger.log(mainUrlKeys);
  var urlKeys = scriptProperties.getProperty('urlKeys').split(', ');
  Logger.log(urlKeys);
  var html = '<!doctype html>\n' +
    '<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en-US" ' +
    'lang="en-US">\n' +
    '<head>\n' +
    '<title>course</title>\n' +
    '<meta charset="utf-8" />\n' +
    '<body>\n' +
    '<h1>createdFoldersFiles</h1>\n';
  for (var i = 0; i < mainUrlKeys.length; i++) {
    var mainUrlKey = scriptProperties.getProperty(mainUrlKeys[i]);
    html += '<p>' + mainUrlKeys[i] + ': <a href="' + mainUrlKey +
      '" target="_blank">' + mainUrlKey + '</a></p>\n';
  }
  html += '<h1>attachments</h1>\n';
  for (var i = 0; i < urlKeys.length; i++) {
    var urlKey = scriptProperties.getProperty(urlKeys[i]);
    html += '<p>' + urlKeys[i] + ': <a href="' + urlKey +
      '" target="_blank">' + urlKey + '</a></p>\n';
  }
  html += '</body>\n' +
    '</html>';
  Logger.log('html created');
  return HtmlService.createHtmlOutput(html);
}


/* Send the first email with an attachment if the user provided necessary
data.
@param {Event} e Triggered in case the form was submitted. */
function sendFirstFile(e) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var safeNumberOfAttendees = scriptProperties
      .getProperty('safeNumberOfAttendees');
  safeNumberOfAttendees = Number(safeNumberOfAttendees);
  var attachmentNames = scriptProperties
      .getProperty('attachmentNames');
  var numberOfAttachments = scriptProperties
      .getProperty('numberOfAttachments');
  var idKeys = scriptProperties.getProperty('idKeys').split(', ');
  Logger.log(idKeys);
  // Get the first attachment
  var number = 0;
  var fileId = scriptProperties.getProperty(idKeys[number]);
  var file = DriveApp.getFileById(fileId);
  Logger.log(file);

  var activeUserEmail = Session.getEffectiveUser().getEmail();
  var aliases = GmailApp.getAliases();
  Logger.log('aliases = ' + aliases);
  var firstName = e.values[1];
  var surname = e.values[2];
  var email = e.values[3];
  var city = e.values[4];
  var information = e.values[5];
  var agreement = e.values[6];

  var subject = attachmentNames + ' ' + (number + 1);
  var message = 'Hello ' + firstName + ',\n';
  message += 'You will receive letters approximately once a week.\n';
  message += 'Total number of letters: ' +
    Math.round(numberOfAttachments - 1) + '\n';
  message += 'You can find ' + file + ' in attachment.\n';
  message += 'If you did not receive any of the letters, ' +
    'please let us know on email: ' + aliases + '\n';
  message += 'If you no more want us to send you any email, ' +
    'please let us know on the email mentioned above.\n\n';
  message += 'Best regards\n';
  message += 'signature\n\n';
  message += '-----------------------------------\n';
  message += 'YOUR DATA:\n';
  message += 'First name: ' + firstName + '\n';
  message += 'Surname: ' + surname + '\n';
  message += 'E-mail: ' + email + '\n';
  message += 'City: ' + city + '\n';
  message += 'Information: ' + information + '\n';
  message += 'Agreement: ' + agreement;

  if (idKeys.length > 0) {
    if (aliases.length > 0) {
      GmailApp.sendEmail(email, subject, message, {
        name: 'sender',
        from: aliases[0],
        attachments: [file.getAs(MimeType.PDF)]
      });
      Logger.log('Email sent as alias');
    } else {
      GmailApp.sendEmail(email, subject, message, {
        name: 'sender',
        attachments: [file.getAs(MimeType.PDF)]
      });
      Logger.log('Email sent without alias');
    }

    // Write number of sent files into the last cell of the active sheet
    number += 1;
    var spreadsheetId = scriptProperties.getProperty('spreadsheetId');
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0];
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var lastCell = sheet.getRange(lastRow, lastColumn);
    Logger.log('lastRow:lastColumn = ' + lastRow + ':' + lastColumn);
    var range = sheet.getRange(lastRow, lastColumn);
    var currentCell = sheet.setActiveRange(range);
    currentCell.setValue(number);
    SpreadsheetApp.flush();
    // Move the last row at the position of row 2
    if (lastRow > 2) {
      sheet.insertRowBefore(2);
      SpreadsheetApp.flush();
      sheet.getRange(sheet.getLastRow(), 1, 1, lastColumn)
        .copyTo(sheet.getRange(2, 1, 1, lastColumn));
      SpreadsheetApp.flush();
      sheet.deleteRow(sheet.getLastRow());
      SpreadsheetApp.flush();
      Logger.log('Last row ' + sheet.getLastRow() + ' moved');
    }

    /* Build a trigger in case the last row in the spreadsheet exceeded
    safeNumberOfAttendees multiples.  Max number of daily triggers is 7.
    Send a message to your email if attendees limit reached (originally
    80 * 7 = 560). */
    switch (lastRow) {
      case 2:
        createEveryDayTrigger(0);
        break;
      case 2 + safeNumberOfAttendees:
        createEveryDayTrigger(1);
        break;
      case 2 + safeNumberOfAttendees * 2:
        createEveryDayTrigger(2);
        break;
      case 2 + safeNumberOfAttendees * 3:
        createEveryDayTrigger(3);
        break;
      case 2 + safeNumberOfAttendees * 4:
        createEveryDayTrigger(4);
        break;
      case 2 + safeNumberOfAttendees * 5:
        createEveryDayTrigger(5);
        break;
      case 2 + safeNumberOfAttendees * 6:
        createEveryDayTrigger(6);
        break;
      case 1 + safeNumberOfAttendees * 7:
        GmailApp.sendEmail(activeUserEmail, (7 * safeNumberOfAttendees) +
          ' attendees reached', 'Use another account. Meanwhile, increase ' +
          'safeNumber (be careful - you can reach the daily quota limit)');
        Logger.log(7 * safeNumberOfAttendees + ' attendees reached, ' +
          'Use another account. Meanwhile, increase safeNumber ' +
          '(be careful - you can reach the daily quota limit)');
        break;
      default:
        break;
    }
  } else {
    /* Send an error message to your email (no attachment found,
    email not sent) */
    GmailApp.sendEmail(activeUserEmail, 'Error - ' +
      subject, 'File ' + file + ' not found and not sent to ' + email);
    Logger.log('File ' + file + ' not found and not sent to ' + email);
  }
}


/* Send next files with attachments.
@param {Event} e Triggered every week in a specific day. */
function sendNextFiles(e) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var safeNumberOfAttendees = scriptProperties
      .getProperty('safeNumberOfAttendees');
  safeNumberOfAttendees = Number(safeNumberOfAttendees);
  var idKeys = scriptProperties.getProperty('idKeys').split(', ');
  Logger.log(idKeys);
  var spreadsheetId = scriptProperties.getProperty('spreadsheetId');
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0];
  var sheet2 = SpreadsheetApp.openById(spreadsheetId).getSheets()[1];
  var lastColumn = sheet.getLastColumn();
  // range according to a specific day trigger and the current day
  var dayIndex = Number(new Date().getDay());
  var range = sheet.getRange(2 + safeNumberOfAttendees * dayIndex, 1,
                             safeNumberOfAttendees, lastColumn);
  var values = range.getValues();
  var aliases = GmailApp.getAliases();
  Logger.log('aliases = ' + aliases);

  // Iterate over all rows in the range
  for (var row = values.length - 1; row >= 0; row--) {
    var cellValue = values[row][lastColumn - 1];
    var firstName = values[row][1];
    // Iterate over numbers of sent files
    for (var j = 0; j <= idKeys.length - 3; j++) {
      var number = j + 1;
      if (cellValue === number) {
        var email = values[row][3];
        var fileId = scriptProperties.getProperty(idKeys[number]);
        var file = DriveApp.getFileById(fileId);
        Logger.log('File to be sent: ' + file);
        var subject = 'letter ' + (number + 1);
        var message = 'Hello ' + firstName + ',\n';
        message += 'You can find ' + file + ' in attachment.\n';
        if (number === idKeys.length - 2) {
          var lastFileId = scriptProperties
            .getProperty(idKeys[idKeys.length - 1]);
          var lastFile = DriveApp.getFileById(lastFileId);
          message += 'This is the last letter. You can also find ' +
            lastFile + ' attached.\n';
        }
        message += 'If you did not receive any of the letters, ' +
          'please let us know on email: ' + aliases + '\n';
        message += 'If you no more want us to send you any email, ' +
          'please let us know on the email mentioned above.\n\n';
        message += 'Best regards\n';
        message += 'signature';
        if (number < idKeys.length - 2) {
          var attach = [file.getAs(MimeType.PDF)];
        } else {
          // 2 attachments in the last email
          var attach = [file.getAs(MimeType.PDF),
                        lastFile.getAs(MimeType.PDF)];
        }

        // Send email if attachments exist
        if (idKeys.length > 0) {
          if (aliases.length > 0) {
            GmailApp.sendEmail (email, subject, message, {
              name: 'sender',
              from: aliases[0],
              attachments: attach
            });
            Logger.log('Email sent as alias');
          } else {
            GmailApp.sendEmail (email, subject, message, {
              name: 'sender',
              attachments: attach
            });
            Logger.log('Email sent without alias');
          }

          // Increment the number of sent letters in the sheet
          number += 1
          var currentNumberRange = sheet.getRange(row + 2 +
            safeNumberOfAttendees * dayIndex, lastColumn);
          var sentNumber = currentNumberRange.setValue(number);
          Logger.log('Value changed to: ' + number);
          SpreadsheetApp.flush();

          /* Remove the attendee from the database if unchecked information
          else move the finished attendee at the second sheet */
          if (number === idKeys.length - 1) {
            var currentRange = sheet.getRange(row + 2 +
              safeNumberOfAttendees * dayIndex, 1, 1, lastColumn);
            var infoValue = currentRange.getValues()[0][5];
            var targetRange = sheet2
                .getRange(sheet2.getLastRow() + 1, 1, 1, lastColumn);
            if (infoValue == '') {
              sheet.deleteRow(row + 2 + safeNumberOfAttendees * dayIndex);
              Logger.log('The finished attendee removed from the database, ' +
              'row ' + (row + 2 + safeNumberOfAttendees * dayIndex));
            } else {
              currentRange.copyTo(targetRange);
              Logger.log('The finished attendee moved to sheet: finished, ' +
                'row ' + (row + 2 + safeNumberOfAttendees * dayIndex) +
                ': [' + currentRange.getValues() + '] ');
              SpreadsheetApp.flush();
              sheet.deleteRow(row + 2 + safeNumberOfAttendees * dayIndex);
            }
            SpreadsheetApp.flush();
          }
        } else {
          /* Send an error message to your email (no attachment found,
          email not sent) */
          var activeUserEmail = Session.getEffectiveUser().getEmail();
          GmailApp.sendEmail(activeUserEmail, 'Error - ' +
            subject, 'File ' + file + ' not found and not sent to ' + email);
          Logger.log('File ' + file + ' not found and not sent to ' + email);
        }
      }
    }
  }

  // Delete the trigger and the property if no more used
  var sendNextFilesSundayTriggerId = scriptProperties
    .getProperty('sendNextFilesSundayTriggerId');
  var sendNextFilesMondayTriggerId = scriptProperties
    .getProperty('sendNextFilesMondayTriggerId');
  var sendNextFilesTuesdayTriggerId = scriptProperties
    .getProperty('sendNextFilesTuesdayTriggerId');
  var sendNextFilesWednesdayTriggerId = scriptProperties
    .getProperty('sendNextFilesWednesdayTriggerId');
  var sendNextFilesThursdayTriggerId = scriptProperties
    .getProperty('sendNextFilesThursdayTriggerId');
  var sendNextFilesFridayTriggerId = scriptProperties
    .getProperty('sendNextFilesFridayTriggerId');
  var sendNextFilesSaturdayTriggerId = scriptProperties
    .getProperty('sendNextFilesSaturdayTriggerId');
  var sendNextFilesDayTriggerIdKeys = ['sendNextFilesSundayTriggerId',
                                       'sendNextFilesMondayTriggerId',
                                       'sendNextFilesTuesdayTriggerId',
                                       'sendNextFilesWednesdayTriggerId',
                                       'sendNextFilesThursdayTriggerId',
                                       'sendNextFilesFridayTriggerId',
                                       'sendNextFilesSaturdayTriggerId'];
  var sendNextFilesDayTriggerIds = [sendNextFilesSundayTriggerId,
                                    sendNextFilesMondayTriggerId,
                                    sendNextFilesTuesdayTriggerId,
                                    sendNextFilesWednesdayTriggerId,
                                    sendNextFilesThursdayTriggerId,
                                    sendNextFilesFridayTriggerId,
                                    sendNextFilesSaturdayTriggerId];
  var lastRow = sheet.getLastRow();
  for (var k = sendNextFilesDayTriggerIds.length - 1; k >= 0; k--) {
    if (lastRow < 2 + safeNumberOfAttendees * k &&
        sendNextFilesDayTriggerIds[k] !== null) {
      deleteTrigger(sendNextFilesDayTriggerIds[k]);
      scriptProperties.deleteProperty(sendNextFilesDayTriggerIdKeys[k]);
      Logger.log('The trigger and the property ' +
      sendNextFilesDayTriggerIdKeys[k] + ' have been deleted');
    }
  }
}


// Create a trigger - on form submit
function createOnFormSubmitTrigger() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sendFirstFileTriggerId = scriptProperties
    .getProperty('sendFirstFileTriggerId');
  Logger.log('sendFirstFileTriggerId = ' + sendFirstFileTriggerId);
  if (sendFirstFileTriggerId === null) {
    var spreadsheetId = scriptProperties.getProperty('spreadsheetId');
    var sendFirstFileTrigger = ScriptApp.newTrigger('sendFirstFile')
      .forSpreadsheet(spreadsheetId).onFormSubmit().create();
    var sendFirstFileTriggerId = sendFirstFileTrigger.getUniqueId();
    scriptProperties.setProperty('sendFirstFileTriggerId',
                                 sendFirstFileTriggerId);
    Logger.log('sendFirstFileTriggerId = ' + sendFirstFileTriggerId);
  } else {
    Logger.log('sendFirstFileTriggerId has already been created');
  }
}


/* Create a trigger - every defined day at 9am (+ dayIndex).
@param {number} dayIndex Index of the day (0-6). */
function createEveryDayTrigger(dayIndex) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sendNextFilesSundayTriggerId = scriptProperties
    .getProperty('sendNextFilesSundayTriggerId');
  var sendNextFilesMondayTriggerId = scriptProperties
    .getProperty('sendNextFilesMondayTriggerId');
  var sendNextFilesTuesdayTriggerId = scriptProperties
    .getProperty('sendNextFilesTuesdayTriggerId');
  var sendNextFilesWednesdayTriggerId = scriptProperties
    .getProperty('sendNextFilesWednesdayTriggerId');
  var sendNextFilesThursdayTriggerId = scriptProperties
    .getProperty('sendNextFilesThursdayTriggerId');
  var sendNextFilesFridayTriggerId = scriptProperties
    .getProperty('sendNextFilesFridayTriggerId');
  var sendNextFilesSaturdayTriggerId = scriptProperties
    .getProperty('sendNextFilesSaturdayTriggerId');
  var sendNextFilesDayTriggerIdKeys = ['sendNextFilesSundayTriggerId',
                                       'sendNextFilesMondayTriggerId',
                                       'sendNextFilesTuesdayTriggerId',
                                       'sendNextFilesWednesdayTriggerId',
                                       'sendNextFilesThursdayTriggerId',
                                       'sendNextFilesFridayTriggerId',
                                       'sendNextFilesSaturdayTriggerId'];
  var sendNextFilesDayTriggerIds = [sendNextFilesSundayTriggerId,
                                    sendNextFilesMondayTriggerId,
                                    sendNextFilesTuesdayTriggerId,
                                    sendNextFilesWednesdayTriggerId,
                                    sendNextFilesThursdayTriggerId,
                                    sendNextFilesFridayTriggerId,
                                    sendNextFilesSaturdayTriggerId];
  Logger.log(sendNextFilesDayTriggerIdKeys[dayIndex] + ' = ' +
             sendNextFilesDayTriggerIds[dayIndex]);
  if (sendNextFilesDayTriggerIds[dayIndex] === null) {
    switch (dayIndex) {
      case 0:
        var sendNextFilesDayTrigger = ScriptApp.newTrigger('sendNextFiles')
          .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY)
          .atHour(9 + dayIndex).create();
        break;
      case 1:
        var sendNextFilesDayTrigger = ScriptApp.newTrigger('sendNextFiles')
          .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY)
          .atHour(9 + dayIndex).create();
        break;
      case 2:
        var sendNextFilesDayTrigger = ScriptApp.newTrigger('sendNextFiles')
          .timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY)
          .atHour(9 + dayIndex).create();
        break;
      case 3:
        var sendNextFilesDayTrigger = ScriptApp.newTrigger('sendNextFiles')
          .timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
          .atHour(9 + dayIndex).create();
        break;
      case 4:
        var sendNextFilesDayTrigger = ScriptApp.newTrigger('sendNextFiles')
          .timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY)
          .atHour(9 + dayIndex).create();
        break;
      case 5:
        var sendNextFilesDayTrigger = ScriptApp.newTrigger('sendNextFiles')
          .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY)
          .atHour(9 + dayIndex).create();
        break;
      case 6:
        var sendNextFilesDayTrigger = ScriptApp.newTrigger('sendNextFiles')
          .timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY)
          .atHour(9 + dayIndex).create();
        break;
      default:
        break;
    }
    var sendNextFilesDayTriggerId = sendNextFilesDayTrigger.getUniqueId();
    scriptProperties.setProperty(sendNextFilesDayTriggerIdKeys[dayIndex],
                                 sendNextFilesDayTriggerId);
    Logger.log(sendNextFilesDayTriggerIdKeys[dayIndex] + ' = ' +
      sendNextFilesDayTriggerId);
  } else {
    Logger.log(sendNextFilesDayTriggerIdKeys[dayIndex] +
      ' has already been created');
  }
}


/* Delete a trigger.
@param {string} triggerId Trigger ID stored in script properties. */
function deleteTrigger(triggerId) {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      break;
    }
  }
}


// Delete all triggers in the project
function deleteAllTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {    
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log('All triggers deleted');
}


// Remove the folder (and the spreadsheet and the form)
function removeFoldersFiles() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var folderId = scriptProperties.getProperty('folderId');
  /*
  // Remove files.  The same could be done with attachments.
  var files = [];
  var spreadsheetId = scriptProperties.getProperty('spreadsheetId');
  var formId = scriptProperties.getProperty('formId');
  if (spreadsheetId !== null) {
    files.push(spreadsheetId);
  }
  if (formId !== null) {
    files.push(formId);
  }
  if (files.length > 0) {
    for (var i = 0; i < files.length; i++) {
      folder.removeFile(DriveApp.getFileById(files[i]));
      Logger.log('file removed (Id) = ' + files[i]);
    }
  }
  */
  if (folderId !== null) {
    var folder = DriveApp.getFolderById(folderId);
    // Remove the folder (and everything in the folder)
    DriveApp.removeFolder(folder);
    Logger.log('folder removed (Id) = ' + folderId);
  }
}


/* Delete script properties - script stores properties permanently
so they have to be deleted if needed empty properties */
function deleteScriptProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
  Logger.log('All script properties deleted');
}
