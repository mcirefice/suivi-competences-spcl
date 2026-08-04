// ============================================================
// SUIVI COMPÉTENCES SPCL — Code.gs
// Géré via clasp + GitHub. Déploiement : clasp push, puis
// Déployer > Gérer les déploiements > modifier > nouvelle version.
// ============================================================

const SHEET_ELEVES = "Élèves";
const SHEET_REFERENTIEL = "Référentiel";
const SHEET_REPONSES = "Réponses";
const SHEET_CODES = "Codes";

// Liste des pages disponibles, organisées par chapitre pour la navigation.
// Pour ajouter un nouveau module :
// 1. Créer le fichier NomDuModule.html
// 2. Ajouter une ligne ci-dessous : { chapitre: "S&Px", label: "Nom affiché", page: "cle-url", template: "NomDuModule" }
// L'URL du module devient : <url du déploiement>?page=cle-url
const MODULES = [
  { chapitre: "Général",  label: "Auto-évaluation",       page: "competences",       template: "Index",            type: "page" },
  { chapitre: "Général",  label: "Codes Arduino / Python", page: "codes-source",      template: "CodesSource",      type: "page" },
  { chapitre: "Général",  label: "Outils",                page: "outils",            template: "Outils",           type: "page" },
  { chapitre: "Général",  label: "Tableau de bord (prof)", page: "tableau-de-bord",   template: "TableauDeBord",    type: "page" },
  { chapitre: "S&P2",     label: "Étalonnage Pt100",       page: "pt100-etalonnage",  template: "PT100Etalonnage",  type: "outil" },
];

function getModules() {
  return MODULES.map(m => ({ chapitre: m.chapitre, label: m.label, page: m.page, type: m.type }));
}

function getTemplateForPage(page) {
  const found = MODULES.find(m => m.page === page);
  return found ? found.template : "Index";
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Suivi SPCL")
    .addItem("Générer les codes élèves manquants", "genererCodesEleves")
    .addItem("Réinitialiser tous les codes (nouvelle année)", "reinitialiserCodes")
    .addItem("Initialiser le référentiel (S&P1 + S&P2)", "seedReferentiel")
    .addItem("Initialiser la bibliothèque de codes (S&P1 + S&P2)", "seedCodes")
    .addToUi();
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || "competences";
  const template = getTemplateForPage(page);
  const moduleInfo = MODULES.find(m => m.page === page);
  const tpl = HtmlService.createTemplateFromFile(template);
  tpl.scriptUrl = ScriptApp.getService().getUrl();
  tpl.currentPage = page;
  tpl.titre = moduleInfo ? moduleInfo.label : "Suivi SPCL";
  return tpl.evaluate()
    .setTitle("Suivi SPCL")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename, data) {
  const tmpl = HtmlService.createTemplateFromFile(filename);
  if (data) {
    Object.keys(data).forEach(function(k) { tmpl[k] = data[k]; });
  }
  return tmpl.evaluate().getContent();
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
    if (name === SHEET_CODES) sh.appendRow(["Chapitre", "Langage", "Titre", "Description", "Code"]);
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

function getCodes() {
  const data = getSheet(SHEET_CODES).getDataRange().getValues();
  return data.slice(1).filter(r => r[2]).map(r => ({
    chapitre: r[0], langage: r[1], titre: r[2], description: r[3], code: r[4],
  }));
}

function seedCodes() {
  const sh = getSheet(SHEET_CODES);
  const rows = [
    ["S&P1", "Arduino", "Mini-défi 1 — Clignoter une LED",
     "Fait clignoter la LED de la broche 13 toutes les 500 ms.",
     "void setup() {\n  pinMode(13, OUTPUT);\n}\n\nvoid loop() {\n  digitalWrite(13, HIGH);\n  delay(500);\n  digitalWrite(13, LOW);\n  delay(500);\n}"],
    ["S&P1", "Arduino", "Mini-défi 2 — Lire un bouton-poussoir",
     "Allume la LED (broche 13) quand le bouton (broche 2) est appuyé.",
     "void setup() {\n  pinMode(13, OUTPUT);\n  pinMode(2, INPUT);\n}\n\nvoid loop() {\n  int etat = digitalRead(2);\n  if (etat == HIGH) {\n    digitalWrite(13, HIGH);\n  } else {\n    digitalWrite(13, LOW);\n  }\n}"],
    ["S&P1", "Arduino", "Mini-défi 3 — Compteur sur le moniteur série",
     "Affiche un compteur qui s'incrémente toutes les secondes.",
     "int compteur = 0;\n\nvoid setup() {\n  Serial.begin(9600);\n}\n\nvoid loop() {\n  Serial.print(\"Compteur : \");\n  Serial.println(compteur);\n  compteur = compteur + 1;\n  delay(1000);\n}"],
    ["S&P1", "Arduino", "Mini-défi 4 — Lire un potentiomètre",
     "Affiche en continu la valeur lue sur l'entrée analogique A0 (0 à 1023).",
     "void setup() {\n  Serial.begin(9600);\n}\n\nvoid loop() {\n  int valeur = analogRead(A0);\n  Serial.println(valeur);\n  delay(200);\n}"],
    ["S&P1", "Arduino", "AE02 — Comparaison 10/8/6/4 bits",
     "Compare la résolution du CAN réel (10 bits) à des résolutions simulées plus faibles, à afficher dans le traceur série.",
     "void setup() {\n  Serial.begin(9600);\n}\n\nvoid loop() {\n  int brut = analogRead(A0);\n\n  int simule8bits = (brut >> 2) << 2;\n  int simule6bits = (brut >> 4) << 4;\n  int simule4bits = (brut >> 6) << 6;\n\n  Serial.print(brut);\n  Serial.print(\"\\t\");\n  Serial.print(simule8bits);\n  Serial.print(\"\\t\");\n  Serial.print(simule6bits);\n  Serial.print(\"\\t\");\n  Serial.println(simule4bits);\n}"],
    ["S&P2", "Arduino", "Exercice 4 — Voltmètre avec pont diviseur",
     "Affiche sur le moniteur série la tension mesurée en A0, convertie en volts.",
     "const int brochePont = A0;\n\nvoid setup() {\n  Serial.begin(9600);\n}\n\nvoid loop() {\n  int valeurBrute = analogRead(brochePont);\n  float tension = valeurBrute * 5.0 / 1023.0;\n\n  Serial.print(\"Tension : \");\n  Serial.print(tension);\n  Serial.println(\" V\");\n\n  delay(500);\n}"],
  ];
  const existing = sh.getDataRange().getValues().slice(1).map(r => r[2]);
  rows.forEach(r => {
    if (!existing.includes(r[2])) sh.appendRow(r);
  });
  SpreadsheetApp.getUi().alert(rows.length + " code(s) initialisé(s) dans l'onglet Codes.");
}

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

// ---------- Suivi élève (historique de ses propres niveaux) ----------

function getSuiviEleve(code) {
  const verif = verifierCode(code);
  if (!verif.ok) return { ok: false };
  const data = getSheet(SHEET_REPONSES).getDataRange().getValues();
  const rows = data.slice(1).filter(r => String(r[1]) === String(code));
  const latest = {};
  rows.forEach(r => {
    const key = r[3] + "|" + r[5];
    if (!latest[key] || new Date(r[0]) > new Date(latest[key][0])) {
      latest[key] = r;
    }
  });
  const resultats = Object.keys(latest).map(k => {
    const r = latest[k];
    return { chapitre: r[3], activite: r[4], code: r[5], intitule: r[6], niveau: r[7] };
  });
  resultats.sort((a, b) => (a.chapitre + a.code).localeCompare(b.chapitre + b.code));

  // Progression par chapitre : % de compétences au niveau A parmi TOUTES les compétences
  // du référentiel pour ce chapitre (une compétence pas encore évaluée compte pour 0).
  const referentiel = getSheet(SHEET_REFERENTIEL).getDataRange().getValues().slice(1).filter(r => r[2]);
  const chapitres = [];
  referentiel.forEach(r => { if (chapitres.indexOf(r[0]) === -1) chapitres.push(r[0]); });
  const progression = chapitres.map(ch => {
    const total = referentiel.filter(r => r[0] === ch).length;
    const evalues = resultats.filter(r => r.chapitre === ch);
    const nbA = evalues.filter(r => r.niveau === "A").length;
    return {
      chapitre: ch,
      pourcentageA: total > 0 ? Math.round((nbA / total) * 100) : 0,
      nbEvalues: evalues.length,
      nbTotal: total,
    };
  });

  return { ok: true, nom: verif.nom, resultats: resultats, progression: progression };
}

// ---------- Tableau de bord enseignant (protégé par mot de passe) ----------
// La toute première saisie de mot de passe le définit de façon permanente (rien à
// configurer à l'avance). Change-le en effaçant la propriété TEACHER_DASHBOARD_PASSWORD
// dans Projet > Propriétés du projet > Propriétés du script, si besoin de le réinitialiser.

const TEACHER_PASSWORD_PROP = "TEACHER_DASHBOARD_PASSWORD";

function verifierMotDePasseProf(motDePasse) {
  if (!motDePasse) return false;
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(TEACHER_PASSWORD_PROP);
  if (!stored) {
    props.setProperty(TEACHER_PASSWORD_PROP, motDePasse);
    return true;
  }
  return stored === motDePasse;
}

function getTableauDeBord(motDePasse) {
  if (!verifierMotDePasseProf(motDePasse)) {
    return { ok: false, message: "Mot de passe incorrect." };
  }
  const eleves = getSheet(SHEET_ELEVES).getDataRange().getValues().slice(1).filter(r => r[0] && r[2]);
  const referentiel = getSheet(SHEET_REFERENTIEL).getDataRange().getValues().slice(1).filter(r => r[2]);
  const reponses = getSheet(SHEET_REPONSES).getDataRange().getValues().slice(1);

  const latest = {};
  reponses.forEach(r => {
    const key = r[1] + "|" + r[5];
    if (!latest[key] || new Date(r[0]) > new Date(latest[key].date)) {
      latest[key] = { date: r[0], niveau: r[7] };
    }
  });

  const chapitres = [];
  referentiel.forEach(r => { if (chapitres.indexOf(r[0]) === -1) chapitres.push(r[0]); });

  const competencesParChapitre = {};
  chapitres.forEach(ch => {
    competencesParChapitre[ch] = referentiel.filter(r => r[0] === ch).map(r => ({ code: r[2], intitule: r[3] }));
  });

  const elevesData = eleves.map(e => {
    const code = String(e[2]);
    const nom = e[0] + " " + e[1];
    const niveaux = {};
    referentiel.forEach(r => {
      const key = code + "|" + r[2];
      niveaux[r[2]] = latest[key] ? latest[key].niveau : "";
    });
    return { nom: nom, niveaux: niveaux };
  });

  return { ok: true, chapitres: chapitres, competencesParChapitre: competencesParChapitre, eleves: elevesData };
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
