function generaID() {
  // Imposta l'ID del foglio e il nome del foglio
  var sheetID = '1WtxISvCYKJyX8c9blp8ROJcd0v-UrFDeUFUpfL9h7Wg';
  var sheetName = 'cronologia';
  var sheet = SpreadsheetApp.openById(sheetID).getSheetByName(sheetName);

  // Trova la colonna "id"
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idColumn = headers.indexOf('id') + 1;

  if (idColumn === 0) {
    throw new Error('Header "id" non trovato.');
  }

  // Ottieni tutti gli ID esistenti
  var data = sheet.getRange(2, idColumn, sheet.getLastRow() - 1).getValues();
  var existingIDs = data.flat();

  // Funzione per generare un ID a 4 cifre
  function generateID() {
    var id = '';
    for (var i = 0; i < 4; i++) {
      id += Math.floor(Math.random() * 10);
    }
    return id;
  }

  // Genera un ID univoco
  var newID;
  do {
    newID = generateID();
  } while (existingIDs.includes(newID));

  return newID;
}

function impostaID() {
  // Genera un nuovo ID univoco
  var newID = generateUniqueID();

  // Restituisci il nuovo ID
  return newID;
}
