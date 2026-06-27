// À déployer dans Google Apps Script (script.google.com)
// → Nouveau projet → coller ce code → Déployer → Application web
//   Accès : Tout le monde (anonyme)
//   Copier l'URL de déploiement → la coller dans index.html (SHEET_URL)

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Ajoute les en-têtes si la feuille est vide
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Nom', 'Email', 'Samedi', 'Dimanche', 'Accompagné·e']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  sheet.appendRow([
    new Date(),
    e.parameter.nom      || '',
    e.parameter.email    || '',
    e.parameter.samedi   || 'Non précisé',
    e.parameter.dimanche || 'Non précisé',
    e.parameter.accompagne || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Test rapide : Exécuter → doTest dans l'éditeur Apps Script
function doTest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([new Date(), 'Test', 'test@test.com', 'Présent·e', 'Présent·e', 'Seul·e']);
}
