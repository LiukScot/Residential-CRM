function creaCartelle(tipoOpportunita, id, yy, nome, cognome) {
  const parentFolderId = "1kpBsmlPAaeCFWvgCEIw38tEk5Q-xQpH_"; // ID della cartella madre
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const folderName = `${tipoOpportunita}-${id}-${yy} ${nome} ${cognome}`;
  const sheetId = "1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ"; // ID del foglio di Google
  const sheetName = "cronologia"; // Nome della pagina nel foglio di Google

  let mainFolderUrl;

  // Controlla se la cartella esiste già
  const existingFolders = parentFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    // La cartella esiste già, ottieni l'URL esistente
    const existingFolder = existingFolders.next();
    mainFolderUrl = existingFolder.getUrl();
  } else {
    // Crea la cartella principale e ottieni il link
    const mainFolder = parentFolder.createFolder(folderName);
    mainFolderUrl = mainFolder.getUrl();
    
    // Crea le sottocartelle
    const subfolders = ['progetto', 'documenti', 'allegati'];
    subfolders.forEach((folder) => {
      const createdFolder = mainFolder.createFolder(folder);
     });
  }
  
  // Aggiorna il foglio di calcolo con l'URL della cartella
  aggiornaFoglioConURL(sheetId, sheetName, id, mainFolderUrl);

  // Opzionale: restituisce l'URL della cartella principale se necessario
  return { mainFolderUrl: mainFolderUrl };
}

function aggiornaFoglioConURL(sheetId, sheetName, id, mainFolderUrl) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const idColumnIndex = values[0].indexOf("id") + 1;
  const folderColumnIndex = values[0].indexOf("cartella") + 1;
  
  if (idColumnIndex <= 0 || folderColumnIndex <= 0) {
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
    sheet.getRange(targetRow, folderColumnIndex).setValue(mainFolderUrl);
  } else {
    throw new Error("Non è stato possibile trovare una riga con l'ID specificato.");
  }
}