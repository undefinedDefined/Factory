var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1hmt0DZP1XTgYuNcHCGrLFZYqG6pi8veL0C3IkHiFIWo/edit#gid=0');
var ws = ss.getSheetByName('Exercice') ?? ss.getActiveSheet();
var range = ws.getRange(1, 1); // ou getRange('A1');

/**
 * WebApp
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Form');
}

/**
 * Exercice 1
 */
function setRange() {
  range.setValue('100');
}

/**
 * Exercice 2
 * add headers on the first row of the sheet 'Exercice'
 */
function setHeaders() {
  var headers = ['userID', 'userName', 'birthDate', 'isActive', 'favMusicType'];
  for (i = 1; i <= headers.length; i++) {
    range = ws.getRange(1, i);
    range.setValue(headers[i - 1]);
  }
}

/**
 * Trigger to add menu
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Data')
    .addItem('Show form', 'showForm')
    .addToUi();
}

/**
 * Trigger to display the form
 */
function showForm() {
  var form = HtmlService.createHtmlOutputFromFile('Form').setTitle('Formulaire');
  // SpreadsheetApp.getUi().showSidebar(form);
  SpreadsheetApp.getUi().showModalDialog(form, 'Formulaire d\'ajout de données');
}

/**
 * Script to process form data
 */
function processForm(form) {
  var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  var row = ws.getLastRow() + 1;

  for (i = 1; i <= headers.length; i++) {
    range = ws.getRange(row, i);

    switch (headers[i - 1]) {
      case 'userID':
        range.setValue(row - 1);
        break;
      default:
        range.setValue(form[headers[i - 1]]);
        break;
    }
  }
}

/**
 * Add sheet 'Data'
 */
function addSheet() {
  ss.insertSheet('Data'); // devient la nouvelle page active
  let data = ['Classique', 'RnB', 'Rap', 'Pop-Rock', 'Rai'];
  for (i = 1; i <= data.length; i++) {
    ss.getActiveSheet().getRange(i, 1).setValue(data[i - 1]);
  }
}

/**
 * Script to get items from the 'Data' sheet
 */
function getOptions() {
  ws = ss.getSheetByName('Data');

  if (!ws) throw new Error('La page Data n\'existe pas');

  var data = ws.getRange(1, 1, ws.getLastRow(), 1).getValues();

  return data.map(function (row) {
    return row[0];
  })
}
