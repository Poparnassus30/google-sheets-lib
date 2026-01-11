function getSheetName() {
  // Renvoie le nom de l'onglet dans laquelle la formule est appelée
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * Met à jour (ou crée) le tableau Sommaire dans un onglet cible.
 *
 * @param {Object} opts Options de configuration (facultatif)
 * @param {string} opts.targetSheetName Nom de l'onglet sommaire (ex: "A2-SOMMAIRE" ou "A2_SOMMAIRE")
 * @param {number} opts.startRow Ligne de départ pour écrire les données (défaut: 9)
 * @param {number} opts.startColumn Colonne de départ (défaut: 1)
 * @param {number} opts.rowsPerPage Nombre de lignes par page (approx) (défaut: 50)
 * @param {number} opts.sleepMs Pause entre écritures (défaut: 200) (peut être 0 si tu veux accélérer)
 */
function SommaireMiseAJour(opts) {
  /*
    Crée un tableau Sommaire dans l'onglet cible.
    Ce tableau liste tous les onglets du classeur avec leurs noms et dimensions (lignes/colonnes),
    ainsi que des colonnes "IMPRESSION PDF", "DE", "A" conservées si déjà existantes.
  */
  Logger.log("Début du script");

  opts = opts || {};
  const cfg = {
    targetSheetName: "A2-SOMMAIRE",
    startRow: 9,
    startColumn: 1,
    rowsPerPage: 50,
    sleepMs: 200,
    ...opts
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = cfg.targetSheetName;
  const targetSheet = ss.getSheetByName(targetSheetName);

  // Vérifie si l'onglet existe
  if (!targetSheet) {
    throw new Error(`La feuille "${targetSheetName}" n'existe pas. Veuillez la créer avant d'exécuter la fonction.`);
  }

  // Filtre les feuilles "réelles" (évite certains objets/feuilles bizarres)
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
  const startRow = cfg.startRow;
  const startColumn = cfg.startColumn;

  // Taille approximative d'une page imprimable (en nombre de lignes)
  const rowsPerPage = cfg.rowsPerPage;

  const sleepMs = Number(cfg.sleepMs) || 0;
  const sleep = () => { if (sleepMs > 0) Utilities.sleep(sleepMs); };

  // Ajoute les en-têtes
  const headerRange = targetSheet.getRange(startRow, startColumn, 1, 8);
  headerRange.setValues([[
    "NOM DE L'ONGLET",
    "IMPRESSION PDF",
    "NOMBRE DE PAGES (APPROX)",
    "PAGE DE _ A _",
    "DE",
    "A",
    "Ligne",
    "Colonne"
  ]]);
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(12);
  headerRange.setBackground("#f4f4f4");
  headerRange.setHorizontalAlignment("center");

  // Parcourt chaque onglet et calcule les données
  sheets.forEach((sheet, index) => {
    const currentRow = startRow + index + 1;

    // Valeurs existantes (conserve si déjà remplies)
    const existingPdfValue = targetSheet.getRange(currentRow, startColumn + 1).getValue();
    const existingDeValue = targetSheet.getRange(currentRow, startColumn + 4).getValue();
    const existingAValue = targetSheet.getRange(currentRow, startColumn + 5).getValue();

    // Nom onglet
    targetSheet.getRange(currentRow, startColumn).setValue(sheet.getName());
    sleep();

    // Nombre de pages approx
    targetSheet.getRange(currentRow, startColumn + 2).setValue(
      Math.ceil(sheet.getLastRow() / rowsPerPage)
    );
    sleep();

    // Impression PDF (oui/non)
    targetSheet.getRange(currentRow, startColumn + 1).setValue(existingPdfValue || "non");
    sleep();

    // DE / A (conserve si présent)
    targetSheet.getRange(currentRow, startColumn + 4).setValue(existingDeValue || "");
    sleep();
    targetSheet.getRange(currentRow, startColumn + 5).setValue(existingAValue || "");
    sleep();

    // Dimensions
    targetSheet.getRange(currentRow, startColumn + 6).setValue(sheet.getLastRow());
    targetSheet.getRange(currentRow, startColumn + 7).setValue(sheet.getLastColumn());

    // Formule "PAGE DE _ A _"
    const formula = `=CONCATENATE("Page "; ${targetSheet.getRange(currentRow, startColumn + 4).getA1Notation()}; " à "; ${targetSheet.getRange(currentRow, startColumn + 5).getA1Notation()})`;
    sleep();
    targetSheet.getRange(currentRow, startColumn + 3).setFormula(formula);
  });

  // Menu déroulant oui/non dans la colonne IMPRESSION PDF
  const dropdownRange = targetSheet.getRange(startRow + 1, startColumn + 1, sheets.length, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["oui", "non"], true)
    .setAllowInvalid(false)
    .build();
  dropdownRange.setDataValidation(rule);

  // Bordures + alignement
  const dataRange = targetSheet.getRange(
    startRow,
    startColumn,
    sheets.length + 1,
    8
  );
  dataRange.setBorder(true, true, true, true, true, true);
  dataRange.setHorizontalAlignment("center");

  SpreadsheetApp.flush();
  Logger.log("Script terminé avec succès !");
}
