/**
 * Erstellt einen benutzerdefinierten Menüpunkt beim Öffnen der Tabelle.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Quiz-Generator')
      .addItem('Neues Quiz aus Text erstellen', 'showInputDialog')
      .addToUi();
}

/**
 * Zeigt ein Dialogfenster zur Eingabe des Quiz-Textes an.
 */
function showInputDialog() {
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Dialog')
      .setWidth(700) // Etwas breiter für die Vorschau
      .setHeight(550); 
  ui.showModalDialog(htmlOutput, 'Quiz-Generator');
}

/**
 * Mischt die Elemente eines Arrays zufällig durch.
 * @param {Array} array Das zu mischende Array.
 * @returns {Array} Das gemischte Array.
 */
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

/**
 * Bereinigt einen reinen Textstring für die Verwendung in LaTeX.
 * @param {string} text Der zu bereinigende Text.
 * @returns {string} Der bereinigte Text.
 */
function sanitizeOnlyText(text) {
    return text.replace(/"/g, "''")
               .replace(/ß/g, '{\\ss}')
               .replace(/ä/g, '{\\"a}')
               .replace(/ö/g, '{\\"o}')
               .replace(/ü/g, '{\\"u}')
               .replace(/Ä/g, '{\\"A}')
               .replace(/Ö/g, '{\\"O}')
               .replace(/Ü/g, '{\\"U}');
}

/**
 * Wandelt eine Zeile (Frage oder Antwort) intelligent in sauberen LaTeX-Code um,
 * indem es li: und lo: als Schalter verwendet.
 * @param {string} text Die Eingabezeile.
 * @returns {string} Der formatierte LaTeX-String.
 */
function parseLineToLatex(text) {
  const parts = text.split(/(li:|lo:)/i);
  let result = '';
  let isMathMode = false;

  parts.forEach(part => {
    if (!part) return;
    const lowerPart = part.toLowerCase();

    if (lowerPart === 'li:') {
      isMathMode = true;
    } else if (lowerPart === 'lo:') {
      isMathMode = false;
    } else {
      if (isMathMode) {
        result += ` ${part} `;
      } else {
        // Behandelt auch Zeilenumbrüche (\\) innerhalb von reinen Textteilen
        const subParts = part.split(/\\\\/g);
        const latexSubParts = subParts.map(p => `\\text{${sanitizeOnlyText(p)}}`);
        result += latexSubParts.join(' \\\\ ');
      }
    }
  });
  return result.trim();
}

/**
 * Dieser Trigger wird bei jeder neuen Formular-Antwort ausgeführt.
 * Er fügt den Quiz-Titel und den Bearbeitungs-Link in die neue Zeile ein.
 * @param {Object} e Das Event-Objekt der Formular-Antwort.
 */
function addExtraDataOnSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    const formUrl = sheet.getFormUrl();
    if (!formUrl) return; // Nicht von einem Formular

    const form = FormApp.openByUrl(formUrl);
    const formId = form.getId();
    
    // Gespeicherte Daten für dieses Formular abrufen
    const properties = PropertiesService.getScriptProperties();
    const formTitle = properties.getProperty(`${formId}_title`);
    const editUrl = properties.getProperty(`${formId}_editUrl`);

    if (!formTitle || !editUrl) return; // Keine Daten für dieses Formular gespeichert

    // Prüfen, ob die Spalten "Quiz-Titel" und "Link zum Bearbeiten" bereits existieren
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let titleColumnIndex = headers.indexOf("Quiz-Titel") + 1;
    let linkColumnIndex = headers.indexOf("Link zum Bearbeiten") + 1;

    // Wenn die Spalten nicht existieren, erstelle sie
    if (linkColumnIndex === 0) {
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue("Link zum Bearbeiten").setFontWeight("bold");
      linkColumnIndex = 1;
    }
    if (titleColumnIndex === 0) {
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue("Quiz-Titel").setFontWeight("bold");
      titleColumnIndex = 1;
      linkColumnIndex++; // Index der Link-Spalte hat sich verschoben
    }

    // Füge den Titel und den Link in die Zeile der neuen Antwort ein
    const row = e.range.getRow();
    sheet.getRange(row, titleColumnIndex).setValue(formTitle);
    sheet.getRange(row, linkColumnIndex).setFormula(`=HYPERLINK("${editUrl}"; "Bearbeiten")`);
    
  } catch (err) {
    Logger.log(`Fehler im onFormSubmit Trigger: ${err.toString()}`);
  }
}


/**
 * Generiert eine HTML-Vorschau der Quiz-Bilder.
 * @param {string} text Der vom Benutzer eingegebene Text.
 * @returns {string} Ein HTML-String mit den Vorschau-Bildern.
 */
function generatePreviewHtml(text) {
  if (!text || text.trim() === '') {
    return '<p style="color: red;">Bitte geben Sie zuerst einen Quiz-Text ein.</p>';
  }
  
  let previewHtml = '';
  const questionBlocks = text.trim().split(/\n\s*\n/);

  questionBlocks.forEach((block, index) => {
    if (block.trim() === '') return;
    
    previewHtml += `<h4 style="margin-top: 20px; font-weight: bold; border-bottom: 1px solid #ccc;">Vorschau Frage ${index + 1}</h4>`;

    const lines = block.trim().split('\n');
    const questionText = lines.shift().trim();
    
    const allChoicesData = [];
    lines.forEach(line => {
      const trimmedLine = line.trim();
      const isCorrect = trimmedLine.startsWith('*');
      const content = isCorrect ? trimmedLine.substring(1).trim() : trimmedLine;
      allChoicesData.push({ content: content, isCorrect: isCorrect });
    });

    // 1. Bild für die Frage erstellen
    const questionContent = parseLineToLatex(questionText);
    const questionLatexString = `\\normalsize \\begin{array}{l} ${questionContent} \\end{array}`;
    const questionImageUrl = `https://latex.codecogs.com/png.latex?${encodeURIComponent(questionLatexString)}`;
    previewHtml += `<p><b>Frage:</b></p><img src="${questionImageUrl}" style="border: 1px solid #ddd; padding: 5px; max-width: 100%;">`;

    // 2. Bild für die Antworten erstellen
    const shuffledChoicesData = shuffleArray(allChoicesData);
    let answersLatexString = '\\normalsize \\begin{array}{ll} \n';

    shuffledChoicesData.forEach((choice, i) => {
      const letter = String.fromCharCode(65 + i);
      const answerContent = parseLineToLatex(choice.content);
      answersLatexString += `\\text{${letter}) } & ${answerContent} \\\\ \\\\ \n`;
    });
    answersLatexString += '\\end{array}';
    const answersImageUrl = `https://latex.codecogs.com/png.latex?${encodeURIComponent(answersLatexString)}`;
    previewHtml += `<p style="margin-top: 15px;"><b>Antworten:</b></p><img src="${answersImageUrl}" style="border: 1px solid #ddd; padding: 5px; max-width: 100%;">`;
  });

  return previewHtml;
}


/**
 * Verarbeitet den eingegebenen Text und erstellt das Google Formular.
 * Diese Funktion wird vom HTML-Dialog aufgerufen.
 * @param {string} text Der vom Benutzer eingegebene Text.
 * @param {string} formTitle Der vom Benutzer eingegebene Titel.
 */
function createFormFromText(text, formTitle) {
  try {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Neues Google Formular erstellen und als Quiz konfigurieren
    const form = FormApp.create(formTitle)
        .setIsQuiz(true)
        .setCollectEmail(false)
        .setAllowResponseEdits(false)
        .setLimitOneResponsePerUser(false);

    // Benutzerdefinierte Bestätigungsnachricht ohne "Wiederholen"-Link
    form.setConfirmationMessage('Vielen Dank für deine Teilnahme! Deine Antworten wurden übermittelt.');

    // Ein Feld für den Namen hinzufügen, da E-Mails nicht mehr gesammelt werden
    form.addTextItem().setTitle('Bitte gib deinen Namen ein').setRequired(true);


    // Text parsen und Fragen erstellen
    const questionBlocks = text.trim().split(/\n\s*\n/);

    questionBlocks.forEach((block, index) => {
      if (block.trim() === '') return;

      const lines = block.trim().split('\n');
      const questionText = lines.shift().trim();
      
      // Alle Antworten parsen
      const allChoicesData = [];
      lines.forEach(line => {
        const trimmedLine = line.trim();
        const isCorrect = trimmedLine.startsWith('*');
        const content = isCorrect ? trimmedLine.substring(1).trim() : trimmedLine;
        allChoicesData.push({ content: content, isCorrect: isCorrect });
      });

      // --- ELEMENTE IN KORREKTER REIHENFOLGE ERSTELLEN ---
      
      form.addSectionHeaderItem().setTitle(`Aufgabe ${index + 1}`);

      // 1. Bild für die Frage erstellen
      const questionContent = parseLineToLatex(questionText);
      const questionLatexString = `\\normalsize \\begin{array}{l} ${questionContent} \\end{array}`;
      const questionImageUrl = `https://latex.codecogs.com/png.latex?${encodeURIComponent(questionLatexString)}`;
      
      try {
        const response = UrlFetchApp.fetch(questionImageUrl, { 'muteHttpExceptions': true });
        if (response.getResponseCode() == 200) {
          form.addImageItem().setImage(response.getBlob()).setAlignment(FormApp.Alignment.CENTER);
        } else { throw new Error(`Server-Fehler (Frage): Code ${response.getResponseCode()}`); }
      } catch (e) {
        form.addSectionHeaderItem().setTitle(`Fehler beim Erstellen des Fragen-Bildes: ${e.message}`);
      }

      // 2. Bild für die Antworten erstellen
      const shuffledChoicesData = shuffleArray(allChoicesData);
      let answersLatexString = '\\normalsize \\begin{array}{ll} \n';
      
      shuffledChoicesData.forEach((choice, i) => {
        const letter = String.fromCharCode(65 + i); // A, B, C...
        const answerContent = parseLineToLatex(choice.content);
        answersLatexString += `\\text{${letter}) } & ${answerContent} \\\\ \\\\ \n`;
      });
      answersLatexString += '\\end{array}';
      const answersImageUrl = `https://latex.codecogs.com/png.latex?${encodeURIComponent(answersLatexString)}`;

      try {
        const response = UrlFetchApp.fetch(answersImageUrl, { 'muteHttpExceptions': true });
        if (response.getResponseCode() == 200) {
          form.addImageItem().setImage(response.getBlob()).setAlignment(FormApp.Alignment.CENTER);
        } else { throw new Error(`Server-Fehler (Antworten): Code ${response.getResponseCode()}`); }
      } catch (e) {
        form.addSectionHeaderItem().setTitle(`Fehler beim Erstellen des Antwort-Bildes: ${e.message}`);
      }
      
      // 3. Standardisierte Multiple-Choice-Frage HINTERHER hinzufügen
      const questionItem = form.addMultipleChoiceItem();
      questionItem.setTitle(`${index + 1}. Bitte richtige Antwort ankreuzen`)
                  .setRequired(true)
                  .setPoints(1);
      
      const finalChoices = shuffledChoicesData.map((choice, i) => {
        const letter = String.fromCharCode(65 + i);
        const choiceText = `Antwort ${letter}`;
        return questionItem.createChoice(choiceText, choice.isCorrect);
      });
      questionItem.setChoices(finalChoices);
    });

    // Formular mit Tabelle verknüpfen und Sheet umbenennen
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    SpreadsheetApp.flush(); // Warten, bis die Verknüpfung aktiv ist

    // KORRIGIERTE LOGIK: Das neue Antwort-Blatt wird zuverlässig gefunden und umbenannt.
    const formUrl = form.getPublishedUrl();
    const allSheets = spreadsheet.getSheets();
    let newSheet = null;
    for (let i = 0; i < allSheets.length; i++) {
        const currentSheet = allSheets[i];
        try {
            if (currentSheet.getFormUrl() === formUrl) {
                newSheet = currentSheet;
                break;
            }
        } catch(e) { /* Ignoriere Blätter ohne Formular-Link */ }
    }
    
    if (newSheet) {
        try {
            let sheetName = formTitle.substring(0, 100);
            if (spreadsheet.getSheetByName(sheetName)) {
                sheetName = `${sheetName} (Antworten)`;
            }
            newSheet.setName(sheetName);
        } catch (e) {
            Logger.log(`Fehler beim Umbenennen des Antwort-Sheets: ${e.toString()}`);
        }
    }
    
    // Quiz-Titel und Edit-URL für den Trigger speichern
    const editUrl = form.getEditUrl();
    const formId = form.getId();
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty(`${formId}_title`, formTitle);
    properties.setProperty(`${formId}_editUrl`, editUrl);

    // Trigger erstellen, der bei jeder neuen Antwort feuert
    const triggers = ScriptApp.getProjectTriggers();
    let triggerExists = false;
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'addExtraDataOnSubmit' && triggers[i].getTriggerSourceId() === spreadsheet.getId()) {
            triggerExists = true;
            break;
        }
    }
    if (!triggerExists) {
        ScriptApp.newTrigger('addExtraDataOnSubmit')
            .forSpreadsheet(spreadsheet)
            .onFormSubmit()
            .create();
    }


    // Erfolgsmeldung mit klickbaren Links anzeigen
    const publishUrl = form.getPublishedUrl();
    
    const htmlMessage = `
      <style>
        body { font-family: Arial, sans-serif; padding: 10px; font-size: 14px; }
        a { color: #1a73e8; text-decoration: none; }
        a:hover { text-decoration: underline; }
        p { margin-bottom: 15px; }
        .link-block { margin-bottom: 10px; word-wrap: break-word; }
        button { background-color: #4285f4; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; }
        button:hover { background-color: #3367d6; }
      </style>
      <p><b>Quiz erfolgreich erstellt!</b></p>
      <div class="link-block">
        <b>Link zum Bearbeiten (für Sie):</b><br>
         <a href="${editUrl}" target="_blank">${editUrl}</a>
      </div>
      <div class="link-block">
        <b>Link zum Versenden (an Schüler):</b><br>
         <a href="${publishUrl}" target="_blank">${publishUrl}</a>
      </div>
      <br>
      <button onclick="google.script.host.close()">Schließen</button>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(htmlMessage)
        .setWidth(600)
        .setHeight(250);
    ui.showModalDialog(htmlOutput, 'Erfolg!');

  } catch (e) {
    SpreadsheetApp.getUi().alert('Ein Fehler ist aufgetreten:', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
