
function getSheetName() {
  //Renvoie le nom de l'onglet dans laquel la formule est appelé
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function SommaireMiseAJour() {
  /*
    Creer une tableau Sommaire dans l'onglet qui porte le nom A2-SOMMAIRE (voir ligne 8)
    Ce tableau liste tous les onglets du classeurs avec leurs noms et nombres de colonnes et de lignes afin d'avoir leur dimmenssion. 
  */
  Logger.log("Début du script");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = "A2-SOMMAIRE"; // Nom de l'onglet cible
  let targetSheet = ss.getSheetByName(targetSheetName);

  // Vérifie si l'onglet existe
  if (!targetSheet) {
    throw new Error(`La feuille "${targetSheetName}" n'existe pas. Veuillez la créer avant d'exécuter la fonction.`);
  }

  //const sheets = ss.getSheets();
  const sheets = ss.getSheets().filter(sheet => {
  try {
    return sheet.getLastRow() > 0 || sheet.getLastColumn() > 0;
  } catch (e) {
    console.log(`Feuille ignorée (probablement un graphique) : ${sheet.getName()}`);
    Logger.log(`Erreur lors du traitement de ${sheet.getName()} : ${e.message}`);
    return false;
  }
});
  // Ligne et colonne de départ pour effacer et écrire
  const startRow = 9; // Ligne de départ pour les données
  const startColumn = 1; // Colonne de départ (ex. C = 3)

  // Taille approximative d'une page imprimable (en nombre de lignes)
  const rowsPerPage = 50;

  // Ajoute les en-têtes
  const headerRange = targetSheet.getRange(startRow, startColumn, 1, 8);
  headerRange.setValues([[
    "NOM DE L'ONGLET", 
    "IMPRESSION PDF", 
    "NOMBRE DE PAGES (APPROX)", 
    "PAGE DE _ A _", 
    "DE", 
    "A",
    "Ligne",  // Nouvelle colonne pour le nombre de lignes
    "Colonne" // Nouvelle colonne pour le nombre de colonnes
  ]]);
  headerRange.setFontWeight("bold"); // Met en gras
  headerRange.setFontSize(12); // Taille de police
  headerRange.setBackground("#f4f4f4"); // Couleur d'arrière-plan
  headerRange.setHorizontalAlignment("center"); // Centre les en-têtes

  // Parcourt chaque onglet et calcule les données
  sheets.forEach((sheet, index) => {
    const currentRow = startRow + index + 1;

    // Vérifie si des valeurs existent déjà dans les colonnes "IMPRESSION PDF", "DE", et "A"
    const existingPdfValue = targetSheet.getRange(currentRow, startColumn + 1).getValue();
    const existingDeValue = targetSheet.getRange(currentRow, startColumn + 4).getValue();
    const existingAValue = targetSheet.getRange(currentRow, startColumn + 5).getValue();

    // Ajoute les données dans les colonnes
    targetSheet.getRange(currentRow, startColumn).setValue(sheet.getName()); // Nom de l'onglet
    Utilities.sleep(200);
    targetSheet.getRange(currentRow, startColumn + 2).setValue(
      Math.ceil(sheet.getLastRow() / rowsPerPage)
    ); // Nombre de pages
    Utilities.sleep(200);
    targetSheet.getRange(currentRow, startColumn + 1).setValue(
      existingPdfValue || "non"
    ); // Conserve "IMPRESSION PDF" ou met "oui" par défaut
    Utilities.sleep(200);
    targetSheet.getRange(currentRow, startColumn + 4).setValue(existingDeValue || ""); // Conserve "DE" ou laisse vide
    Utilities.sleep(200);
    targetSheet.getRange(currentRow, startColumn + 5).setValue(existingAValue || ""); // Conserve "A" ou laisse vide

    targetSheet.getRange(currentRow, startColumn + 6).setValue(sheet.getLastRow()); // Nombre de lignes
    targetSheet.getRange(currentRow, startColumn + 7).setValue(sheet.getLastColumn()); // Nombre de colonnes

    // Ajoute la formule CONCATENATE dans la colonne "PAGE DE _ A _"
    const formula = `=CONCATENATE("Page "; ${targetSheet.getRange(currentRow, startColumn + 4).getA1Notation()}; " à "; ${targetSheet.getRange(currentRow, startColumn + 5).getA1Notation()})`;
    Utilities.sleep(200);
    targetSheet.getRange(currentRow, startColumn + 3).setFormula(formula); // Colonne "PAGE DE _ A _"
  });

  // Crée un menu déroulant avec "oui" ou "non" dans la colonne IMPRESSION PDF
  const dropdownRange = targetSheet.getRange(startRow + 1, startColumn + 1, sheets.length, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["oui", "non"], true)
    .setAllowInvalid(false)
    .build();
  dropdownRange.setDataValidation(rule);

  // Appliquer des bordures autour des cellules utilisées
  const dataRange = targetSheet.getRange(
    startRow,
    startColumn,
    sheets.length + 1, // Nombre total de lignes (en-tête + données)
    8 // Nombre total de colonnes
  );
  dataRange.setBorder(true, true, true, true, true, true); // Ajoute des bordures
  dataRange.setHorizontalAlignment("center"); // Centre toutes les données

  SpreadsheetApp.flush(); // Applique les modifications
  Logger.log("Script terminé avec succès !");
}

