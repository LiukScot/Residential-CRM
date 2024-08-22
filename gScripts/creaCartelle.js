function creaCartelle(tipoOpportunita, id, yy, nome, cognome) {
  Logger.log('Inizio della funzione creaCartelle');
  
  const parentFolderId = "1kpBsmlPAaeCFWvgCEIw38tEk5Q-xQpH_"; // ID della cartella madre
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  Logger.log('Cartella madre ID: ' + parentFolderId);
  
  const folderName = `${tipoOpportunita}-${id}-${yy} ${nome} ${cognome}`;
  Logger.log('Nome della cartella: ' + folderName);
  
  const sheetId = "1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ"; // ID del foglio di Google
  const sheetName = "cronologia"; // Nome della pagina nel foglio di Google
  Logger.log('ID del foglio: ' + sheetId);
  Logger.log('Nome della pagina del foglio: ' + sheetName);

  let mainFolderUrl;

  // Controlla se la cartella esiste già
  const existingFolders = parentFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    Logger.log('La cartella esiste già');
    
    // La cartella esiste già, ottieni l'URL esistente
    const existingFolder = existingFolders.next();
    mainFolderUrl = existingFolder.getUrl();
  } else {
    Logger.log('La cartella non esiste, creazione di una nuova cartella');
    
    // Crea la cartella principale e ottieni il link
    const mainFolder = parentFolder.createFolder(folderName);
    mainFolderUrl = mainFolder.getUrl();
    
    // Crea le sottocartelle
    const subfolders = ['progetto', 'documenti', 'allegati'];
    subfolders.forEach((folder) => {
      const createdFolder = mainFolder.createFolder(folder);
      Logger.log('Sottocartella creata: ' + folder);
    });
  }
  
  // Aggiorna il foglio di calcolo con l'URL della cartella
  aggiornaFoglioConURL(sheetId, sheetName, id, mainFolderUrl);
  
  Logger.log('Cartella principale URL: ' + mainFolderUrl);
  
  // Opzionale: restituisce l'URL della cartella principale se necessario
  return { mainFolderUrl: mainFolderUrl };
}

function aggiornaFoglioConURL(sheetId, sheetName, id, mainFolderUrl) {
  Logger.log('Inizio della funzione aggiornaFoglioConURL');
  
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const idColumnIndex = values[0].indexOf("id") + 1;
  const folderColumnIndex = values[0].indexOf("cartella") + 1;
  
  Logger.log('Indice colonna ID: ' + idColumnIndex);
  Logger.log('Indice colonna Cartella: ' + folderColumnIndex);
  
  if (idColumnIndex <= 0 || folderColumnIndex <= 0) {
    Logger.log('Errore: Non è possibile trovare le colonne "id" o "cartella".');
    throw new Error("Non è possibile trovare le colonne 'id' o 'cartella'.");
  }
  
  let targetRow;
  for (let i = 1; i < values.length; i++) {
    if (values[i][idColumnIndex - 1].toString() === id.toString()) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow) {
    Logger.log('Riga target trovata: ' + targetRow);
    sheet.getRange(targetRow, folderColumnIndex).setValue(mainFolderUrl);
  } else {
    Logger.log('Errore: Non è stato possibile trovare una riga con l\'ID specificato.');
    throw new Error("Non è stato possibile trovare una riga con l'ID specificato.");
  }
}