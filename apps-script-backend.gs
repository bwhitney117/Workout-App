// Google Apps Script backend for Training Tracker
// Deployed as a Web App bound to Google Sheet "Training Log"
// Sheet ID: 1sX0fRECWa6ctf-zE3vjU7b6JE-L0JcgEYT7V1uWIntk
//
// This file is a reference copy. The live version is deployed via
// Apps Script editor at:
// https://script.google.com/macros/s/AKfycbzQ9XdyksH1npXk3JsxitUBseQbJlUPQ_dETVxVV7gy8hOajn2CWh8_7wTeGzmvF-onxA/exec

var SHEET_NAME = 'Workouts';
var HEADERS = ['id', 'date', 'type', 'category', 'exercise', 'sets', 'cardioType', 'totalTime', 'totalMiles', 'splits', 'createdAt'];

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var workouts = [];

    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      workouts.push(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, workouts: workouts }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'add') {
      var sheet = getOrCreateSheet();
      var row = HEADERS.map(function(h) {
        return body[h] !== undefined ? (typeof body[h] === 'object' ? JSON.stringify(body[h]) : body[h]) : '';
      });
      sheet.appendRow(row);

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, action: 'added' }))
        .setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'delete') {
      var sheet = getOrCreateSheet();
      var data = sheet.getDataRange().getValues();
      var idCol = data[0].indexOf('id');

      for (var i = 1; i < data.length; i++) {
        if (data[i][idCol] === body.id) {
          sheet.deleteRow(i + 1);
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, action: 'deleted' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'Workout not found' }))
        .setMimeType(ContentService.MimeType.JSON);

    } else {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action: ' + action }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
