function getSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * Met à jour (ou crée) le tableau Sommaire dans un onglet cible.
 *
 * - Crée l'onglet si absent
 * - Met à jour les lignes selon les onglets existants
 * - Conserve les colonnes: IMPRESSION PDF / DE / A si déjà remplies
 *
 * @param {Object} opts Options de configuration (facultatif)
 * @param {string} opts.targetSheetName Nom de l'onglet sommaire (défaut: "A2-SOMMAIRE")
 * @param {number} opts.startRow Ligne de départ (défaut: 9)
 * @param {number} opts.startColumn Colonne de départ (défaut: 1)
 * @param {number} opts.rowsPerPage Taille approx d'une page (défaut: 50)
 * @param {number} opts.sleepMs Pause entre écritures (défaut: 0)
 * @param {boolean} opts.clearOldRows Effacer les anciennes lignes du sommaire (défaut: true)
 */
function SommaireMiseAJour(opts) {
  Logger.log("Début du script SommaireMiseAJour");

  opts = opts || {};
  const cfg = {
    targetSheetName: "A2-SOMMAIRE",
    startRow: 9,
    startColumn: 1,
    rowsPerPage: 50,
    sleepMs: 0,
    clearOldRows: true,
    ...opts
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Crée l'onglet sommaire si absent
  let targetSheet = ss.getSheetByName(cfg.targetSheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(cfg.targetSheetName);
    Logger.log(`Onglet sommaire créé: ${cfg.targetSheetName}`);
  }

  const startRow = cfg.startRow;
  const startColumn = cfg.startColumn;
  const rowsPerPage = cfg.rowsPerPage;
  const sleepMs = Number(cfg.sleepMs) || 0;
  const sleep = () => { if (sleepMs > 0) Utilities.sleep(sleepMs); };

  // 2) Liste des onglets "réels" (exclut le sommaire lui-même)
  const sheets = ss.getSheets().filter(sh => {
    if (sh.getName() === cfg.targetSheetName) return false;
    try {
      return sh.getLastRow() > 0 || sh.getLastColumn() > 0;
    } catch (e) {
      Logger.log(`Feuille ignorée: ${sh.getName()} (erreur: ${e.message})`);
      return false;
    }
  });

  // 3) Conserver les anciennes valeurs (IMPRESSION PDF / DE / A) via un index par nom d'onglet
  // On lit l'ancien tableau si existant
  const oldMap = {}; // { sheetName: { pdf, de, a } }

  const lastRow = targetSheet.getLastRow();
  const headerRow = startRow;

  if (lastRow > headerRow) {
    const dataStartRow = headerRow + 1;
    const numRows = lastRow - headerRow;

    // Colonnes: A..H (8 colonnes)
    const oldValues = targetSheet
      .getRange(dataStartRow, startColumn, numRows, 8)
      .getValues();

    // oldValues[i] = [NOM, PDF, NBPAGES, PAGE, DE, A, LIGNE, COL]
    oldValues.forEach(row => {
      const name = row[0];
      if (!name) return;
      oldMap[name] = {
        pdf: row[1],
        de: row[4],
        a: row[5]
      };
    });
  }

  // 4) Optionnel : nettoyage de la zone existante (sauf si tu veux garder des lignes fantômes)
  if (cfg.clearOldRows) {
    // On efface tout sous l'entête (A..H)
    // Largeur 8 colonnes
    const maxRowsToClear = Math.max(0, targetSheet.getMaxRows() - startRow);
    if (maxRowsToClear > 0) {
      targetSheet
        .getRange(startRow, startColumn, maxRowsToClear, 8)
        .clear({ contentsOnly: true, formatsOnly: false });
    }
  }

  // 5) Écrit l'entête
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

  // 6) Construit les lignes à écrire en mémoire (plus rapide que setValue ligne par ligne)
  const rows = sheets.map(sh => {
    const name = sh.getName();
    const keep = oldMap[name] || {};

    const nbPages = Math.ceil(sh.getLastRow() / rowsPerPage);

    const pdf = (keep.pdf && String(keep.pdf).trim() !== "") ? keep.pdf : "non";
    const de = (keep.de && String(keep.de).trim() !== "") ? keep.de : "";
    const a = (keep.a && String(keep.a).trim() !== "") ? keep.a : "";

    // colonne "PAGE DE _ A _" = formule, on la mettra après (range.setFormulas)
    return [
      name,
      pdf,
      nbPages,
      "",  // formule ensuite
      de,
      a,
      sh.getLastRow(),
      sh.getLastColumn()
    ];
  });

  // 7) Écrit toutes les lignes d’un coup
  if (rows.length > 0) {
    const dataRange = targetSheet.getRange(startRow + 1, startColumn, rows.length, 8);
    dataRange.setValues(rows);
    sleep();

    // 8) Ajoute les formules "PAGE DE _ A _" (colonne D = index 4ème)
    const formulas = rows.map((_, i) => {
      const r = startRow + 1 + i;
      const deCell = targetSheet.getRange(r, startColumn + 4).getA1Notation(); // col E
      const aCell  = targetSheet.getRange(r, startColumn + 5).getA1Notation(); // col F
      return [`=CONCATENATE("Page "; ${deCell}; " à "; ${aCell})`];
    });

    targetSheet.getRange(startRow + 1, startColumn + 3, rows.length, 1).setFormulas(formulas);
    sleep();

    // 9) Dropdown oui/non sur la colonne IMPRESSION PDF (col B)
    const dropdownRange = targetSheet.getRange(startRow + 1, startColumn + 1, rows.length, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["oui", "non"], true)
      .setAllowInvalid(false)
      .build();
    dropdownRange.setDataValidation(rule);

    // 10) Bordures + alignement
    const fullRange = targetSheet.getRange(startRow, startColumn, rows.length + 1, 8);
    fullRange.setBorder(true, true, true, true, true, true);
    fullRange.setHorizontalAlignment("center");
  } else {
    Logger.log("Aucun onglet à lister (hors sommaire).");
  }

  SpreadsheetApp.flush();
  Logger.log("SommaireMiseAJour terminé avec succès !");
}
