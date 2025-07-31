
Move setPoints method in Code.gs
31. Juli
MichaelGlaubitz/aeg
main
+14
-2

Archivieren

Weitergeben

PR erstellen


1. In `Code.gs` around lines 220‑232, move `questionItem.setPoints(1)` **after** calling `questionItem.setChoices(finalChoices)`.
2. Optionally validate that each question block contains exactly one answer marked with `*` and alert the user if not.
Kopieren
Summary

Added a validation loop ensuring each question block contains exactly one answer marked with “*” and alerting the user otherwise

Adjusted the multiple-choice question setup so that points are assigned after choices are set

Testing

❌ npm test (failed to run because package.json was not found)


Datei (1)

Code.gs
+14
-2


Diffs
Protokolle
Code.gs
+14
-2

@@ -141,50 +141,62 @@ function generatePreviewHtml(text) {
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

    // Sicherstellen, dass jede Frage genau eine mit * markierte Antwort enthält
    for (let i = 0; i < questionBlocks.length; i++) {
      const block = questionBlocks[i];
      if (block.trim() === '') continue;
      const answerLines = block.trim().split('\n').slice(1);
      const correctCount = answerLines.filter(line => line.trim().startsWith('*')).length;
      if (correctCount !== 1) {
        ui.alert(`Frage ${i + 1} muss genau eine Antwort mit * enthalten.`);
        throw new Error(`Ungültige Anzahl korrekter Antworten in Frage ${i + 1}`);
      }
    }

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
@@ -198,60 +210,60 @@ function createFormFromText(text, formTitle) {

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
                  .setRequired(true);
      
      // KORRIGIERTE LOGIK: Erstellt die Ankreuz-Optionen zuverlässig
      const finalChoices = shuffledChoicesData.map((choice, i) => {
        const letter = String.fromCharCode(65 + i);
        const choiceText = `Antwort ${letter}`;
        return questionItem.createChoice(choiceText, choice.isCorrect);
      });
      questionItem.setChoices(finalChoices);
      questionItem.setPoints(1);
    });

    // Sheet umbenennen und Titel-Spalte hinzufügen
    const sheetsBefore = spreadsheet.getSheets();
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    SpreadsheetApp.flush();

    const sheetsAfter = spreadsheet.getSheets();
    let newSheet = null;
    if (sheetsAfter.length > sheetsBefore.length) {
        const sheetsBeforeIds = sheetsBefore.map(s => s.getSheetId());
        for (let i = 0; i < sheetsAfter.length; i++) {
            if (sheetsBeforeIds.indexOf(sheetsAfter[i].getSheetId()) == -1) {
                newSheet = sheetsAfter[i];
                break;
            }
        }
    }

    if (newSheet) {
        try {
            let sheetName = formTitle.substring(0, 100);
            if (spreadsheet.getSheetByName(sheetName)) {
                sheetName = `${sheetName} (Antworten)`;
            }
