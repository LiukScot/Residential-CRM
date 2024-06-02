function stampaOffertaV2(tipo_opportunita, id, nome, cognome, indirizzo, telefono, email, numero_moduli, marca_moduli, marca_inverter, numero_batteria, capacita_batteria, totale_capacita_batterie, marca_batteria, tetto, potenza_impianto, produzione_impianto, risparmio_25_anni, alberi, testo_aggiuntivo, tipo_pagamento, condizione_pagamento_1, condizione_pagamento_2, condizione_pagamento_3, condizione_pagamento_4, imponibile_offerta, iva_offerta, prezzo_offerta, cartella, anni_finanziamento, conLayout, esposizione, area_m2_impianto, numero_colonnina_74kw, numero_colonnina_22kw, numero_ottimizzatori, marca_ottimizzatori, numero_linea_vita, scheda_tecnica_moduli, scheda_tecnica_inverter, scheda_tecnica_batterie, scheda_tecnica_ottimizzatori, detrazione) {

    var oggi = new Date();
    var data = new Intl.DateTimeFormat('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(oggi);
  
    var cartellaDestinazioneId = cartella.split('/folders/')[1];
    
    // Creazione della sottocartella "contratto"
    var cartellaContratto = DriveApp.getFolderById(cartellaDestinazioneId).getFoldersByName('contratto');
    var cartellaContrattoId;
    if (cartellaContratto.hasNext()) {
      cartellaContrattoId = cartellaContratto.next().getId();
    } else {
      cartellaContrattoId = DriveApp.getFolderById(cartellaDestinazioneId).createFolder('contratto').getId();
    }
  
    // Creazione della sottocartella con la data
    var nomeCartellaData = LibrerieMyenergySolutions.validaValore(data);
    var cartellaData = DriveApp.getFolderById(cartellaContrattoId).getFoldersByName(nomeCartellaData);
    var cartellaDataId;
    if (cartellaData.hasNext()) {
      cartellaDataId = cartellaData.next().getId();
    } else {
      cartellaDataId = DriveApp.getFolderById(cartellaContrattoId).createFolder(nomeCartellaData).getId();
    }
  
    var offertaStandard = '10HvhPMSPYG-vG-4r4C6Bd6cizT0YuZLMfQ9_smgTOds';
    var offertaMateriale = '1gMJGZZA7LwdugXKEFTK5LbJU2iiIIs6Ee5zBnlW81es';
    var offertaConLayout = '1dOsSzOQa30u7I2Z4FnU0guxiilO_VmsEjZ23UIrB1d0';
    var contratto = '1_PNr5Y6svOADvgKZIjFjKsoDFpNV6TkOxivLIVqcZdA';
  
    var datiDocumento = [];
  
    if (tipo_opportunita === "MAT") {
      datiDocumento.push({
        templateId: offertaMateriale,
        nomeFile: "Offerta Myenergy " + LibrerieMyenergySolutions.validaValore(nome) + " " + LibrerieMyenergySolutions.validaValore(cognome) + " " + LibrerieMyenergySolutions.validaValore(data)
      });
    } else {
      if (conLayout) {
        datiDocumento.push({
          templateId: offertaConLayout,
          nomeFile: "Presentazione offerta Myenergy " + LibrerieMyenergySolutions.validaValore(nome) + " " + LibrerieMyenergySolutions.validaValore(cognome)
        });
      } else {
        datiDocumento.push({
          templateId: offertaStandard,
          nomeFile: "Presentazione offerta Myenergy " + LibrerieMyenergySolutions.validaValore(nome) + " " + LibrerieMyenergySolutions.validaValore(cognome)
        });
      }
  
      datiDocumento.push({
        templateId: contratto,
        nomeFile: "Offerta " + LibrerieMyenergySolutions.validaValore(tipo_opportunita) + "-" + LibrerieMyenergySolutions.validaValore(id) + " " + LibrerieMyenergySolutions.validaValore(nome) + " " + LibrerieMyenergySolutions.validaValore(cognome) + " " + LibrerieMyenergySolutions.validaValore(data)
      });
    }
  
    datiDocumento.forEach(function(dato) {
      var doc = LibrerieMyenergySolutions.createDocumentFromTemplate(dato.templateId, cartellaDataId, dato.nomeFile);
      var corpo = doc.getBody();
  
      var mappaturaSegnapostov2 = {
        '{{tipo_opportunità}}': tipo_opportunita,
        '{{id}}': id,
        '{{nome}}': nome,
        '{{cognome}}': cognome,
        '{{indirizzo}}': indirizzo,
        '{{telefono}}': telefono,
        '{{email}}': email,
        '{{data ultima modifica}}': data,
        '{{numero_moduli}}': numero_moduli,
        '{{marca_moduli}}': marca_moduli,
        '{{marca_inverter}}': marca_inverter,
        '{{numero_batteria}}': numero_batteria,
        '{{capacità batteria}}': capacita_batteria,
        '{{totale_capacità_batterie}}': totale_capacita_batterie,
        '{{marca_batteria}}': marca_batteria,
        '{{tetto}}': tetto,
        '{{potenza_impianto}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(potenza_impianto),
        '{{produzione_impianto}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(produzione_impianto),
        '{{risparmio_25_anni}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(risparmio_25_anni),
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
        '{{detrazione}}': detrazione,
        '{{anni_finanziamento}}': new Intl.NumberFormat('it-IT', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(anni_finanziamento),
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
      };
  
      LibrerieMyenergySolutions.replacePlaceholders(corpo, mappaturaSegnapostov2);
      LibrerieMyenergySolutions.addHyperlink(corpo, 'Link scheda tecnica moduli', scheda_tecnica_moduli);
      LibrerieMyenergySolutions.addHyperlink(corpo, 'Link scheda tecnica inverter', scheda_tecnica_inverter);
      LibrerieMyenergySolutions.addHyperlink(corpo, 'Link scheda tecnica batterie', scheda_tecnica_batterie);
      LibrerieMyenergySolutions.addHyperlink(corpo, 'Link scheda tecnica ottimizzatori', scheda_tecnica_ottimizzatori);
  
      doc.saveAndClose();
    });
  
    var datiTecnici = LibrerieMyenergySolutions.getTechnicalDataById('1_QEo5ynx_29j3I3uJJff5g7ZzGZJnPcIarIXfr5O2gQ', id);
    var cartellaProgetto = DriveApp.getFolderById(cartellaDestinazioneId).getFoldersByName('progetto');
    var cartellaProgettoId = cartellaProgetto.hasNext() ? cartellaProgetto.next().getId() : DriveApp.getFolderById(cartellaDestinazioneId).createFolder('progetto').getId();
  
    var fileDatiTecnici = DriveApp.getFolderById(cartellaProgettoId).getFilesByName('dati tecnici');
    var nuovoFileDatiTecnici = fileDatiTecnici.hasNext() ? SpreadsheetApp.openById(fileDatiTecnici.next().getId()) : SpreadsheetApp.openById(DriveApp.getFileById('1U2kbWlfl9-LO1j-pNPbLNdVreJb8Gej4W13kjMyxq0k').makeCopy('dati tecnici', DriveApp.getFolderById(cartellaProgettoId)).getId());
  
    LibrerieMyenergySolutions.addTechnicalData(nuovoFileDatiTecnici.getActiveSheet(), datiTecnici);
  }