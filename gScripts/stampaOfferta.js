function stampaOffertaV2(appID, tipo_opportunita, id, yy, nome, cognome, indirizzo, telefono, email, numero_moduli, numero_inverter, marca_moduli, 
  marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto, 
  potenza_impianto, produzione_impianto, alberi, testo_aggiuntivo, tipo_pagamento, 
  condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta,
  iva_offerta, prezzo_offerta, cartella, anni_finanziamento, conLayout, esposizione, area_m2_impianto, 
  numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita, 
  scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, 
  detrazione, anni_ritorno_investimento, utile_25_anni, consumi_annui, 
  profilo_di_consumo, provincia, prezzo_energia, rata_mensile, numero_rate_mensili, anni_finanziamento) {

// Log per debug
Logger.log('Tipo opportunità: ' + tipo_opportunita);

//CREA SOTTOCARTELLA "contratto"
var cartellaDestinazioneId = cartella.split('/folders/')[1];
var oggi = new Date();
var dataOggi = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(oggi);

// Creazione della sottocartella "contratto"
var cartellaContratto = DriveApp.getFolderById(cartellaDestinazioneId).getFoldersByName('contratto');
var cartellaContrattoId;
if (cartellaContratto.hasNext()) {
cartellaContrattoId = cartellaContratto.next().getId();
} else {
cartellaContrattoId = DriveApp.getFolderById(cartellaDestinazioneId).createFolder('contratto').getId();
}

// Creazione della sottocartella con la data
var nomeCartellaData = dataOggi;
var cartellaData = DriveApp.getFolderById(cartellaContrattoId).getFoldersByName(nomeCartellaData);
var cartellaDataId;
if (cartellaData.hasNext()) {
cartellaDataId = cartellaData.next().getId();
} else {
cartellaDataId = DriveApp.getFolderById(cartellaContrattoId).createFolder(nomeCartellaData).getId();
}

var datiDocumento = [];

//SCELTA DEI MODELLI DI OFFERTA
var presentazioneFinanz = '1zMIjekT-K_JWssZidSBjSuog_LfcHZjMLcEbePnP_t8' //v2
var offertaMateriale = '1gMJGZZA7LwdugXKEFTK5LbJU2iiIIs6Ee5zBnlW81es'; 
var presentazione = '1XYDLbJymoNqU8B1nYqJm0k52-SU5O19G1Xzph_rjShg'; //v2
var contratto = '1_PNr5Y6svOADvgKZIjFjKsoDFpNV6TkOxivLIVqcZdA';
var contrattoREDEN = '1mFtXfWCxKv2y4-kbkRLugrDx_Hnih_ZkMmFto0RtSVU';
var contrattoGSE = '1t5S9CYogDPAtKhy2ejMVELKjAkkieqfu31eIFF06GYg';
var contrattoFinanz = '1RCr8lgM98ryQwMiGFqecMHiWgIHsPN0tfV5HN82eYr4'; //v2

if (tipo_opportunita === "MAT") {
datiDocumento.push({
templateId: offertaMateriale,
nomeFile: "Offerta Myenergy " + nome + " " + cognome + " " + dataOggi
});
} else {
if (tipo_pagamento === "Finanziamento") {
datiDocumento.push({
templateId: presentazioneFinanz,
nomeFile: "Presentazione offerta Myenergy " + nome + " " + cognome
});
} else {
datiDocumento.push({
templateId: presentazione,
nomeFile: "Presentazione offerta Myenergy " + nome + " " + cognome
});
}

if (tipo_opportunita === "REDEN") {
datiDocumento.push({
templateId: contrattoREDEN,
nomeFile: "Offerta " + tipo_opportunita + "-" + id + "-" + yy+ " " + nome + " " + cognome + " " + dataOggi
});
datiDocumento.push({
templateId: contrattoGSE,
nomeFile: "Contratto GSE " + tipo_opportunita + "-" + id + "-" + yy + " " + nome + " " + cognome + " " + dataOggi
});
} else if (tipo_pagamento === 'Finanziamento') {
datiDocumento.push({
templateId: contrattoFinanz,
nomeFile: "Offerta " + tipo_opportunita + "-" + id + "-" + yy + " " + nome + " " + cognome + " " + dataOggi
});
} else {
datiDocumento.push({
templateId: contratto,
nomeFile: "Offerta " + tipo_opportunita + "-" + id + "-" + yy + " " + nome + " " + cognome + " " + dataOggi
});
}
}

// Se non è MAT, esegui le operazioni specifiche per altri tipi di opportunità
if (tipo_opportunita !== "MAT") {

// CREA FILE SHEET "dati tecnici"

// Estrai l'ultima offerta generata da sheet "CRM database", sheet "cronologia"
var CRMdatabase = SpreadsheetApp.openById('1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ');
var sheetOfferte = CRMdatabase.getSheetByName('offerte');
var data = sheetOfferte.getDataRange().getValues();




// Trova la colonna "appID" nell'header
var header = data[0];
var appIDColIndex = header.indexOf('appID');

if (appIDColIndex === -1) {
throw new Error('Colonna "appID" non trovata');
}

// Trova la riga con il valore corrispondente di appID
var selectedRow = null;
for (var i = 1; i < data.length; i++) {
if (data[i][appIDColIndex] === appID) {
selectedRow = data[i];
break;
}
}

if (selectedRow === null) {
throw new Error('Nessuna riga trovata con appID: ' + appID);
}



// crea cartella "progetto"
var nomeFileDatiTecnici = 'dati tecnici v4';
var cartellaProgetto = DriveApp.getFolderById(cartellaDestinazioneId).getFoldersByName('progetto');
var cartellaProgettoId;
if (cartellaProgetto.hasNext()) {
cartellaProgettoId = cartellaProgetto.next().getId();
} else {
var nuovaCartellaProgetto = DriveApp.getFolderById(cartellaDestinazioneId).createFolder('progetto');
cartellaProgettoId = nuovaCartellaProgetto.getId();
}

// crea o aggiorna "dati tecnici"
var fileDatiTecnici = DriveApp.getFolderById(cartellaProgettoId).getFilesByName(nomeFileDatiTecnici);
var nuovoFileDatiTecnici;
if (fileDatiTecnici.hasNext()) {
nuovoFileDatiTecnici = SpreadsheetApp.openById(fileDatiTecnici.next().getId());
} else {
var modelloDatiTecnici = DriveApp.getFileById('1cPaLSSNlz5snyD4q3vBCLlpIGsKtiOaKrvsmoi_8SCk').makeCopy(nomeFileDatiTecnici, DriveApp.getFolderById(cartellaProgettoId));
nuovoFileDatiTecnici = SpreadsheetApp.openById(modelloDatiTecnici.getId());
}

// aggiorna i valori nel foglio "log" con l'ultima offerta generata
var nuovoSheet = nuovoFileDatiTecnici.getActiveSheet();
var ultimaRigaVuota = nuovoSheet.getLastRow() + 1;
nuovoSheet.getRange(ultimaRigaVuota, 1, 1, selectedRow.length).setValues([selectedRow]);

/* Trova l'indice della colonna con header "id" nel foglio "log"
var headers = nuovoSheet.getRange(1, 1, 1, nuovoSheet.getLastColumn()).getValues()[0];
var colonnaIdIndex = headers.indexOf("appID") + 1; // Aggiungiamo 1 perché gli indici delle colonne partono da 1 in Apps Script

// Aggiungi la lettera corrispondente alla riga alla fine del valore "id" e riporta il valore come variabile di testo
var offerID = "";
if (colonnaIdIndex > 0) { // Controlla se la colonna "id" esiste
var idValue = nuovoSheet.getRange(ultimaRigaVuota, colonnaIdIndex).getValue();
var lettera = String.fromCharCode(64 + (ultimaRigaVuota - 2)); // Ajusta per iniziare da 'A' per la riga 3
offerID = idValue + lettera; // Aggiorna il valore nell'ID
nuovoSheet.getRange(ultimaRigaVuota, colonnaIdIndex).setValue(offerID);
}
*/

// INCOLLA E PRENDI DATI DA "ANALISI ENERGETICA"

// Apri il foglio "analisi energetica"
var sheetAnalisiEnergetica = nuovoFileDatiTecnici.getSheetByName('analisi energetica');

// Sostituisci le celle denominate con le variabili
sheetAnalisiEnergetica.getRange('consumi_annui').setValue(consumi_annui);
sheetAnalisiEnergetica.getRange('profilo_di_consumo').setValue(profilo_di_consumo);
sheetAnalisiEnergetica.getRange('provincia').setValue(provincia);
sheetAnalisiEnergetica.getRange('esposizione').setValue(esposizione);
sheetAnalisiEnergetica.getRange('offerta_analizzata').setValue(appID);

// Leggi i valori dalle celle denominate e assegnali alle variabili
var percentuale_autoconsumo = sheetAnalisiEnergetica.getRange('percentuale_autoconsumo').getValue();
var media_vendita = sheetAnalisiEnergetica.getRange('media_vendita').getValue();
var anni_ritorno_investimento = sheetAnalisiEnergetica.getRange('anni_ritorno_investimento').getValue();
var percentuale_risparmio_energetico = sheetAnalisiEnergetica.getRange('percentuale_risparmio_energetico').getValue();
var utile_25_anni = sheetAnalisiEnergetica.getRange('utile_25_anni').getValue();

// MODIFICA DEL GRAFICO DI RITORNO ANNUO

// Ottenere il foglio per il grafico del risparmio
const grafico = SpreadsheetApp.openById('1cfLNo1WU-poleX1i4hg9fynUP53i06d_6DpIlbFZcTc');
const foglio1GRAFICO = grafico.getSheetByName('foglio1');

// Apri il foglio "calcoli" sempre in dati tecnici
const sheetCalcoliDATITECNICI = nuovoFileDatiTecnici.getSheetByName('calcoli');
if (!sheetCalcoliDATITECNICI) {
throw new Error('Foglio "calcoli" non trovato.');
}

// Copia il valore delle celle denominate "utile_25_anni_grafico"
var valoriUtile25AnniGrafico = getNamedRangeValues(sheetCalcoliDATITECNICI, 'utile_25_anni_grafico');

function getNamedRangeValues(sheet, rangeName) {
var range = sheet.getRange(rangeName);
if (!range) {
throw new Error('Intervallo denominato "' + rangeName + '" non trovato nel foglio "' + sheet.getName() + '".');
}
return range.getValues();
}

// Incolla i valori nel foglio grafico in colonna B
foglio1GRAFICO.getRange(1, 2, valoriUtile25AnniGrafico.length, valoriUtile25AnniGrafico[0].length).setValues(valoriUtile25AnniGrafico);
}

//COMPILA DOC OFFERTA CON LE VARIABILI

datiDocumento.forEach(function(dato) {
var doc = createDocumentFromTemplate(dato.templateId, cartellaDataId, dato.nomeFile);
var corpo = doc.getBody();
var mappaturaSegnapostov2 = {
'{{tipo_opportunità}}': tipo_opportunita,
'{{id}}': id,
'{{yy}}': yy,
'{{nome}}': nome,
'{{cognome}}': cognome,
'{{indirizzo}}': indirizzo,
'{{telefono}}': telefono,
'{{email}}': email,
'{{data ultima modifica}}': dataOggi,
'{{numero_moduli}}': numero_moduli,
'{{marca_moduli}}': marca_moduli,
'{{numero_inverter}}': numero_inverter,
'{{marca_inverter}}': marca_inverter,
'{{numero_batteria}}': numero_batteria,
'{{capacità batteria}}': capacita_batteria,
'{{totale_capacità_batterie}}': totale_capacita_batterie,
'{{marca_batteria}}': marca_batteria,
'{{tetto}}': tetto,
'{{potenza_impianto}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(potenza_impianto),
'{{produzione_impianto}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(produzione_impianto),
'{{alberi}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(alberi),
'{{testo_aggiuntivo}}': testo_aggiuntivo,
'{{tipo_pagamento}}': tipo_pagamento,
'{{condizione_pagamento_1}}': condizione_pagamento_1,
'{{condizione_pagamento_2}}': condizione_pagamento_2,
'{{condizione_pagamento_3}}': condizione_pagamento_3,
'{{condizione_pagamento_4}}': condizione_pagamento_4,
'{{imponibile_offerta}}': new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(imponibile_offerta),
'{{iva_offerta}}': new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(iva_offerta),
'{{prezzo_offerta}}': new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(prezzo_offerta),
'{{anni_finanziamento}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(anni_finanziamento),
'{{rata_mensile}}': new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(rata_mensile),
'{{numero_rate_mensili}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(numero_rate_mensili),
'{{esposizione}}': esposizione,
'{{area_m2_impianto}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(area_m2_impianto),
'{{scheda_tecnica_moduli}}': 'Link scheda tecnica moduli',
'{{scheda_tecnica_inverter}}': 'Link scheda tecnica inverter',
'{{scheda_tecnica_batterie}}': 'Link scheda tecnica batterie',
'{{scheda_tecnica_ottimizzatori}}': 'Link scheda tecnica ottimizzatori',
'{{numero_colonnina_74kw}}': numero_colonnina_74kw,
'{{numero_colonnina_22kw}}': numero_colonnina_22kw,
'{{numero_ottimizzatori}}': numero_ottimizzatori,
'{{marca_ottimizzatori}}': marca_ottimizzatori,
'{{numero_linea_vita}}': numero_linea_vita,
'{{detrazione}}': new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(detrazione),
'{{anni_ritorno_investimento}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(anni_ritorno_investimento),
'{{utile_25_anni}}': new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(utile_25_anni),
'{{percentuale_autoconsumo}}': new Intl.NumberFormat('it-IT', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(percentuale_autoconsumo),
'{{media_vendita}}': new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(media_vendita),
'{{prezzo_energia}}': new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(prezzo_energia),
'{{percentuale_risparmio_energetico}}': new Intl.NumberFormat('it-IT', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(percentuale_risparmio_energetico),
};

// sostiuisci i placeholder con i valori corretti
replacePlaceholders(corpo, mappaturaSegnapostov2);
addHyperlink(corpo, 'Link scheda tecnica moduli', scheda_tecnica_moduli);
addHyperlink(corpo, 'Link scheda tecnica inverter', scheda_tecnica_inverter);
addHyperlink(corpo, 'Link scheda tecnica batterie', scheda_tecnica_batterie);
addHyperlink(corpo, 'Link scheda tecnica ottimizzatori', scheda_tecnica_ottimizzatori);

doc.saveAndClose();

// Log each variable for debugging
for (var key in mappaturaSegnapostov2) {
if (mappaturaSegnapostov2.hasOwnProperty(key)) {
Logger.log(key + ': ' + mappaturaSegnapostov2[key]);
}
}

});
}




// FUNZIONI CHIAMATE DALLO SCRIPT stampaOffertaV2

// Funzione per creare un documento da un template
function createDocumentFromTemplate(templateId, destinationFolderId, fileName) {
var documentCopy = DriveApp.getFileById(templateId).makeCopy(fileName, DriveApp.getFolderById(destinationFolderId));
return DocumentApp.openById(documentCopy.getId());
}

// Funzione per sostituire i segnaposto con i valori nel documento
function replacePlaceholders(body, placeholders) {
Object.keys(placeholders).forEach(function(placeholder) {
var value = placeholders[placeholder];
if (value !== null && value !== undefined && value !== "") { // Verifica che il valore sia valido
body.replaceText(placeholder, value.toString()); // Converte il valore in stringa per evitare errori
}
});
}

// Funzione per aggiungere un hyperlink
function addHyperlink(corpo, searchText, url) {
var foundElement = corpo.findText(searchText);
while (foundElement) {
var foundText = foundElement.getElement().asText();
var startOffset = foundElement.getStartOffset();
var endOffset = foundElement.getEndOffsetInclusive();
foundText.setLinkUrl(startOffset, endOffset, url);
foundElement = corpo.findText(searchText, foundElement);
}
}