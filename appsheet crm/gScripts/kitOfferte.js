function kitOfferte(id, codiceKit) {
    // ID del primo file e nome del foglio
    const sourceSpreadsheetId = '11Pg7GdIKjzAemVdvtXTnYy9xZpBaQRt1dvRwMLBq6GI';
    const sourceSheetName = 'kit offerte per appsheet';
    
    // ID del secondo file e nome del foglio
    const targetSpreadsheetId = '1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ';
    const targetSheetName = 'cronologia';
  
    // Aprire il foglio di origine
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    
    // Aprire il foglio di destinazione
    const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  
    // Trovare la colonna con header "Codice kit offerta" nel foglio di origine (riga 7)
    const sourceData = sourceSheet.getDataRange().getValues();
    Logger.log('Source Data: ' + JSON.stringify(sourceData));
    
    let codiceKitColIdx = null;
    const targetHeader = "Codice kit offerta".toLowerCase().trim();
    const headerRow = 6; // Indice della riga 7 (0-based index)
    for (let j = 0; j < sourceData[headerRow].length; j++) {
      if (sourceData[headerRow][j].toLowerCase().trim() === targetHeader) {
        codiceKitColIdx = j;
        break;
      }
    }
    
    Logger.log('Codice Kit Column Index: ' + codiceKitColIdx);
    
    if (codiceKitColIdx === null) {
      throw new Error('Header "Codice kit offerta" non trovato nel foglio di origine');
    }
  
    // Trovare la riga nel foglio di origine con il codiceKit (dalla riga 8 in poi)
    let sourceRow = null;
    for (let i = headerRow + 1; i < sourceData.length; i++) { // Partendo da headerRow + 1 per escludere l'header
      if (String(sourceData[i][codiceKitColIdx]).toLowerCase().trim() === codiceKit.toLowerCase().trim()) {
        sourceRow = sourceData[i].slice(2); // Escludere le prime due colonne
        Logger.log('Source Row Found: ' + JSON.stringify(sourceRow));
        break;
      }
    }
    
    if (!sourceRow) {
      throw new Error('Codice Kit non trovato nel foglio di origine');
    }
  
    // Trovare la colonna con header "id" nel foglio di destinazione
    const targetData = targetSheet.getDataRange().getValues();
    Logger.log('Target Data: ' + JSON.stringify(targetData));
    
    let idColIdx = null;
    for (let j = 0; j < targetData[0].length; j++) {
      if (String(targetData[0][j]).toLowerCase().trim() === 'id'.toLowerCase().trim()) {
        idColIdx = j;
        break;
      }
    }
    
    Logger.log('ID Column Index: ' + idColIdx);
    
    if (idColIdx === null) {
      throw new Error('Header "id" non trovato nel foglio di destinazione');
    }
  
    // Trovare la riga nel foglio di destinazione con l'id
    let targetRowIdx = null;
    let targetColIdx = null;
    
    for (let i = 1; i < targetData.length; i++) { // Partendo da 1 per escludere l'header
      Logger.log('Checking ID: ' + String(targetData[i][idColIdx]).toLowerCase().trim());
      if (String(targetData[i][idColIdx]).toLowerCase().trim() === String(id).toLowerCase().trim()) {
        targetRowIdx = i;
        Logger.log('Target Row Index Found: ' + targetRowIdx);
        break;
      }
    }
  
    if (targetRowIdx !== null) {
      for (let j = 0; j < targetData[0].length; j++) {
        Logger.log('Checking Header: ' + String(targetData[0][j]).toLowerCase().trim());
        if (String(targetData[0][j]).toLowerCase().trim() === 'codice kit offerta'.toLowerCase().trim()) {
          targetColIdx = j;
          Logger.log('Target Column Index Found: ' + targetColIdx);
          break;
        }
      }
    }
    
    Logger.log('Final Target Row Index: ' + targetRowIdx);
    Logger.log('Final Target Column Index: ' + targetColIdx);
    
    if (targetRowIdx === null || targetColIdx === null) {
      throw new Error('ID o colonna "codice kit offerta" non trovati nel foglio di destinazione');
    }
  
    // Incollare i valori copiati nella riga corrispondente
    for (let i = 0; i < sourceRow.length; i++) {
      targetSheet.getRange(targetRowIdx + 1, targetColIdx + 1 + i).setValue(sourceRow[i]);
    }
  }
  