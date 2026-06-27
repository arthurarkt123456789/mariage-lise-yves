// À déployer dans Google Apps Script (script.google.com)
// → Déployer → Gérer les déploiements → modifier → Nouvelle version → Déployer

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // En-têtes au premier envoi
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Nom', 'Email', 'Samedi', 'Dimanche', 'Accompagné·e']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  sheet.appendRow([
    new Date(),
    e.parameter.nom       || '',
    e.parameter.email     || '',
    e.parameter.samedi    || 'Non précisé',
    e.parameter.dimanche  || 'Non précisé',
    e.parameter.accompagne || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Test rapide : Exécuter → doTest
function doTest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([new Date(), 'Test', 'test@test.com', 'Présent·e', 'Présent·e', 'Seul·e']);
}
