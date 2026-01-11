function onOpen() {
  // Appelle les fonctions suivantes  à l'ouverture du document :

  //Fichier A1-Sommaire.gs
  SommaireMiseAJour();
  
}

function onClick() {
// Appels les fonctions à chaque modification
  getSheetName()
}

function getSheetName() {
  //Renvoie le nom de l'onglet dans laquel la formule est appelé
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}