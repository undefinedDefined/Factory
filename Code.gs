// var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1hmt0DZP1XTgYuNcHCGrLFZYqG6pi8veL0C3IkHiFIWo/edit#gid=0');
var ss = SpreadsheetApp.getActive();
// on active l'onglet 'Exercice' ou celui qui est actif si 'Exercice' n'existe pas
var ws = ss.getSheetByName('Exercice') ?? ss.getActiveSheet();
// On récupère la ligne (1,1) = A1
var range = ws.getRange(1, 1); // ou getRange('A1');

/**
 * WebApp
 * pour définir la page à afficher lors du déploiement
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Formulaire').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * Exercice 1
 * mettre la valeur 100 dans A1
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
 * Ajoute le menu 'Data' à l'ouverture du fichier et le sous-onglet 'Show form' qui déclenche la fonction showForm (en dessous)
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Data')
    .addItem('Show form', 'showForm')
    .addToUi();
}

/**
 * Trigger to display the form
 * fonction executée lorsqu'on clique sur le sous-menu 'Show Form'
 */
function showForm() {
  var form = HtmlService.createTemplateFromFile('Formulaire').evaluate();
  form.setTitle('Formulaire');
  // SpreadsheetApp.getUi().showSidebar(form);
  SpreadsheetApp.getUi().showModalDialog(form, 'Formulaire d\'ajout de données');
}

/**
 * Script to process form data
 * met les elements du formulaire dans le tableur
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
 * crée l'onglet Data
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

/**
 * Get last username
 */
function getLastUser(){
  var lastRow = ws.getLastRow();

  return ws.getRange(lastRow, 2).getValue();
}
