function updateLatLong() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange(); // Ottiene il range di dati
    var values = range.getValues(); // Ottiene tutti i dati in un array di array
    
    for (var i = 1; i < values.length; i++) { // Inizia da 1 per saltare l'intestazione
      var address = values[i][0]; // Assumendo che gli indirizzi siano nella prima colonna
      if (values[i][1] === "" && address !== "") { // Controlla se la latlong è vuota ma l'indirizzo no
        var latlong = getLatLongFromAddress(address);
        if (latlong) { // Se latlong non è null
          sheet.getRange(i + 1, 2).setValue(latlong); // Imposta latlong nella seconda colonna
        }
      }
    }
  }
  
  function getLatLongFromAddress(address) {
    var encodedAddress = encodeURIComponent(address);
    var url = "https://nominatim.openstreetmap.org/search?format=json&q=" + encodedAddress;
    var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    if (response.getResponseCode() == 200) {
      var results = JSON.parse(response.getContentText());
      if (results.length > 0) {
        var lat = results[0].lat;
        var lon = results[0].lon;
        return lat + ", " + lon;
      }
    }
    return null; // Ritorna null se non ci sono risultati o se c'è un errore
  }