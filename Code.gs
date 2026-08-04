// ============================================================
// SUIVI COMPÉTENCES SPCL — Code.gs
// Géré via clasp + GitHub. Déploiement : clasp push, puis
// Déployer > Gérer les déploiements > modifier > nouvelle version.
// ============================================================

const SHEET_ELEVES = "Élèves";
const SHEET_REFERENTIEL = "Référentiel";
const SHEET_REPONSES = "Réponses";

// Liste des pages disponibles. Pour ajouter un nouveau module :
// 1. Créer le fichier NomDuModule.html
// 2. Ajouter "nomdumodule": "NomDuModule" ci-dessous
// L'URL du module devient alors : <url du déploiement>?page=nomdumodule
const PAGES = {
  "competences": "Index",
  "pt100-etalonnage": "PT100Etalonnage",
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Suivi SPCL")
    .addItem("Générer les codes élèves manquants", "genererCodesEleves")
    .addItem("Réinitialiser tous les codes (nouvelle année)", "reinitialiserCodes")
    .addItem("Initialiser le référentiel (S&P1 + S&P2)", "seedReferentiel")
    .addToUi();
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || "competences";
  const template = PAGES[page] || "Index";
  const tpl = HtmlService.createTemplateFromFile(template);
  tpl.scriptUrl = ScriptApp.getService().getUrl();
  return tpl.evaluate()
    .setTitle("Suivi SPCL")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ---------- Gestion des codes élèves ----------

function genererCodesEleves() {
  const sh = getSheet(SHEET_ELEVES);
  const data = sh.getDataRange().getValues();
  const existingCodes = new Set(data.slice(1).map(r => String(r[2])).filter(Boolean));
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const nom = data[i][0], code = data[i][2];
    if (nom && !code) {
      let newCode;
      do {
        newCode = String(Math.floor(1000 + Math.random() * 9000));
      } while (existingCodes.has(newCode));
      existingCodes.add(newCode);
      sh.getRange(i + 1, 3).setValue(newCode);
      updated++;
    }
  }
  SpreadsheetApp.getUi().alert(updated + " code(s) généré(s).");
}

function reinitialiserCodes() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert("Réinitialiser tous les codes ?", "Cela efface tous les codes existants (à faire en début d'année). Continuer ?", ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const sh = getSheet(SHEET_ELEVES);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 3, lastRow - 1, 1).clearContent();
  genererCodesEleves();
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === SHEET_ELEVES) sh.appendRow(["Nom", "Prénom", "Code"]);
    if (name === SHEET_REFERENTIEL) sh.appendRow(["Chapitre", "Activité", "Code compétence", "Intitulé"]);
    if (name === SHEET_REPONSES) sh.appendRow(["Horodatage", "Code", "Nom complet", "Chapitre", "Activité", "Code compétence", "Intitulé", "Niveau"]);
  }
  return sh;
}

// ---------- Référentiel (pré-remplissage) ----------

function seedReferentiel() {
  const sh = getSheet(SHEET_REFERENTIEL);
  const rows = [
    ["S&P1", "AE01", "S&P1-01", "Programmer une entrée/sortie numérique (digitalWrite, digitalRead)."],
    ["S&P1", "AE01", "S&P1-02", "Utiliser la communication série (Serial.begin, Serial.print)."],
    ["S&P1", "AE01", "S&P1-03", "Lire une entrée analogique (analogRead)."],
    ["S&P1", "AE02", "S&P1-04", "Numérisation d'une tension : Convertisseur Analogique-Numérique (CAN)."],
    ["S&P1", "AE02", "S&P1-05", "Citer les caractéristiques utiles d'un CAN : nombre de bits, quantum, fréquence d'échantillonnage."],
    ["S&P2", "AE01", "S&P2-01", "Capteur et conditionneur."],
    ["S&P2", "AE01", "S&P2-02", "Identifier les grandeurs d'entrée et de sortie du capteur et/ou du conditionneur."],
    ["S&P2", "AE01", "S&P2-03", "Utiliser un microcontrôleur pour afficher la valeur de la mesure."],
    ["S&P2", "AE01", "S&P2-04", "Pont diviseur de tension (rappel de 1ère SPCL)."],
    ["S&P2", "AE01", "S&P2-05", "Concevoir et réaliser un conditionneur à partir d'un pont diviseur de tension pour un capteur résistif."],
    ["S&P2", "AE02", "S&P2-06", "Capacité thermique massique d'un corps."],
    ["S&P2", "AE02", "S&P2-07", "Réaliser des calculs simples de calorimétrie."],
    ["S&P2", "AE02", "S&P2-08", "Mener une démarche d'investigation en autonomie (construire un protocole)."],
  ];
  const existing = sh.getDataRange().getValues().slice(1).map(r => r[2]);
  rows.forEach(r => {
    if (!existing.includes(r[2])) sh.appendRow(r);
  });
  SpreadsheetApp.getUi().alert("Référentiel initialisé (" + rows.length + " compétences).");
}

// ---------- Fonctions appelées depuis les pages web ----------

function verifierCode(code) {
  const data = getSheet(SHEET_ELEVES).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(code)) {
      return { ok: true, nom: data[i][0] + " " + data[i][1] };
    }
  }
  return { ok: false };
}

function getChapitres() {
  const data = getSheet(SHEET_REFERENTIEL).getDataRange().getValues();
  return [...new Set(data.slice(1).map(r => r[0]))];
}

function getActivites(chapitre) {
  const data = getSheet(SHEET_REFERENTIEL).getDataRange().getValues();
  return [...new Set(data.slice(1).filter(r => r[0] === chapitre).map(r => r[1]))];
}

function getCompetences(chapitre, activite) {
  const data = getSheet(SHEET_REFERENTIEL).getDataRange().getValues();
  return data.slice(1)
    .filter(r => r[0] === chapitre && r[1] === activite)
    .map(r => ({ code: r[2], intitule: r[3] }));
}

function soumettreEvaluation(payload) {
  const verif = verifierCode(payload.code);
  if (!verif.ok) return { ok: false, message: "Code invalide." };
  const sh = getSheet(SHEET_REPONSES);
  const now = new Date();
  payload.niveaux.forEach(item => {
    sh.appendRow([now, payload.code, verif.nom, payload.chapitre, payload.activite, item.code, item.intitule, item.niveau]);
  });
  return { ok: true };
}

// ---------- Enregistrement générique de résultats d'outils (ex. étalonnage Pt100) ----------
// Permet à n'importe quel module HTML d'enregistrer un résultat, sans créer une nouvelle
// feuille à chaque fois : tout part dans "Réponses_Outils" avec le nom de l'outil en colonne.

function enregistrerResultatOutil(payload) {
  const verif = verifierCode(payload.code);
  if (!verif.ok) return { ok: false, message: "Code invalide." };
  const sh = getSheet("Réponses_Outils");
  if (sh.getLastRow() === 0) sh.appendRow(["Horodatage", "Code", "Nom complet", "Outil", "Données (JSON)"]);
  sh.appendRow([new Date(), payload.code, verif.nom, payload.outil, JSON.stringify(payload.donnees)]);
  return { ok: true };
}
