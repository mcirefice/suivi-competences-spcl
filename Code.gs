// ============================================================
// SUIVI COMPÉTENCES SPCL — Code.gs
// Géré via clasp + GitHub. Déploiement : clasp push, puis
// Déployer > Gérer les déploiements > modifier > nouvelle version.
// ============================================================

const SHEET_ELEVES = "Élèves";
const SHEET_REFERENTIEL = "Référentiel";
const SHEET_REPONSES = "Réponses";
const SHEET_CODES = "Codes";
const SHEET_QCM = "QCM";

// Liste des pages disponibles, organisées par chapitre pour la navigation.
// Pour ajouter un nouveau module :
// 1. Créer le fichier NomDuModule.html
// 2. Ajouter une ligne ci-dessous : { chapitre: "S&Px", label: "Nom affiché", page: "cle-url", template: "NomDuModule" }
// L'URL du module devient : <url du déploiement>?page=cle-url
const MODULES = [
  { chapitre: "Général",  label: "Auto-évaluation",       page: "competences",       template: "Index",            type: "page", profOnly: false },
  { chapitre: "Général",  label: "Codes Arduino / Python", page: "codes-source",      template: "CodesSource",      type: "page", profOnly: false },
  { chapitre: "Général",  label: "Outils",                page: "outils",            template: "Outils",           type: "page", profOnly: false },
  { chapitre: "Général",  label: "QCM de révision",       page: "qcm",               template: "QCM",              type: "page", profOnly: false },
  { chapitre: "Général",  label: "Tableau de bord (prof)", page: "tableau-de-bord",   template: "TableauDeBord",    type: "page", profOnly: true },
  { chapitre: "S&P2",     label: "Étalonnage Pt100",       page: "pt100-etalonnage",  template: "PT100Etalonnage",  type: "outil", profOnly: false },
];

function getModules(code) {
  const verif = verifierCode(code);
  const role = verif.ok ? verif.role : "Élève";
  return MODULES
    .filter(m => !m.profOnly || role === "Professeur")
    .map(m => ({ chapitre: m.chapitre, label: m.label, page: m.page, type: m.type }));
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
    .addItem("Initialiser les QCM (S&P1 + S&P2)", "seedQCM")
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
    if (name === SHEET_ELEVES) sh.appendRow(["Nom", "Prénom", "Code", "Rôle"]);
    if (name === SHEET_REFERENTIEL) sh.appendRow(["Chapitre", "Activité", "Code compétence", "Intitulé"]);
    if (name === SHEET_REPONSES) sh.appendRow(["Horodatage", "Code", "Nom complet", "Chapitre", "Activité", "Code compétence", "Intitulé", "Niveau"]);
    if (name === SHEET_CODES) sh.appendRow(["Chapitre", "Langage", "Titre", "Description", "Code"]);
    if (name === SHEET_QCM) sh.appendRow(["Chapitre", "Question", "OptionA", "OptionB", "OptionC", "OptionD", "BonneReponse", "Explication", "Niveau"]);
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

function getQCM(chapitre, niveau) {
  const data = getSheet(SHEET_QCM).getDataRange().getValues();
  return data.slice(1)
    .filter(r => r[0] === chapitre && r[1] && (String(r[8] || "Débutant") === niveau))
    .map(r => ({
      question: r[1], options: { A: r[2], B: r[3], C: r[4], D: r[5] }, bonneReponse: r[6], explication: r[7],
    }));
}

function getChapitresQCM() {
  const data = getSheet(SHEET_QCM).getDataRange().getValues();
  return [...new Set(data.slice(1).filter(r => r[0]).map(r => r[0]))];
}

function enregistrerResultatQCM(payload) {
  const verif = verifierCode(payload.code);
  if (!verif.ok) return { ok: false, message: "Code invalide." };
  const sh = getSheet("Réponses_Outils");
  if (sh.getLastRow() === 0) sh.appendRow(["Horodatage", "Code", "Nom complet", "Outil", "Données (JSON)"]);
  sh.appendRow([new Date(), payload.code, verif.nom, "qcm-" + payload.chapitre, JSON.stringify({ score: payload.score, total: payload.total, niveau: payload.niveau })]);
  return { ok: true };
}

function seedQCM() {
  const sh = getSheet(SHEET_QCM);

  // Migration : les lignes déjà présentes sans niveau renseigné deviennent "Débutant".
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const niveauCol = sh.getRange(2, 9, lastRow - 1, 1).getValues();
    niveauCol.forEach((v, i) => {
      if (!v[0]) sh.getRange(i + 2, 9).setValue("Débutant");
    });
  }

  const debutant = [
    ["S&P1", "Quelle fonction Arduino permet de mettre une broche à l'état HAUT ou BAS ?",
     "analogRead()", "digitalWrite()", "Serial.print()", "pinMode()", "B",
     "digitalWrite(broche, HIGH/LOW) impose un état logique (5V ou 0V) sur une broche numérique configurée en sortie. pinMode() sert uniquement à définir si la broche est en entrée ou en sortie, pas à changer son état.", "Débutant"],
    ["S&P1", "La fonction analogRead() sur une carte Arduino Uno renvoie une valeur comprise entre :",
     "0 et 5", "0 et 255", "0 et 1023", "0 et 1", "C",
     "Le CAN de l'Arduino Uno est un CAN 10 bits, donc 2¹⁰ = 1024 niveaux possibles, numérotés de 0 à 1023.", "Débutant"],
    ["S&P1", "Pour un CAN de n bits fonctionnant entre 0 et 5V, le quantum (plus petite variation détectable) est d'autant plus...",
     "grand que n est grand", "petit que n est grand", "indépendant de n", "grand que la fréquence d'échantillonnage est grande", "B",
     "q = Vmax / (2ⁿ - 1). Plus n est grand, plus le nombre de niveaux 2ⁿ est grand, donc plus q est petit : la résolution est plus fine.", "Débutant"],
    ["S&P1", "D'après le théorème de Shannon, pour reconstituer fidèlement un signal de fréquence maximale fmax, la fréquence d'échantillonnage fe doit vérifier :",
     "fe = fmax", "fe < fmax", "fe > 2×fmax", "fe < 2×fmax", "C",
     "Le théorème de Shannon impose fe > 2×fmax pour éviter le repliement de spectre (aliasing), qui déformerait le signal reconstitué à partir des échantillons.", "Débutant"],
    ["S&P1", "Quelle instruction affiche du texte suivi d'un retour à la ligne sur le moniteur série ?",
     "Serial.begin()", "Serial.print()", "Serial.println()", "Serial.read()", "C",
     "Serial.println() affiche le contenu ET ajoute un retour à la ligne, contrairement à Serial.print() qui reste sur la même ligne. Serial.begin() sert uniquement à initialiser la communication série.", "Débutant"],
    ["S&P2", "Un pont diviseur de tension permet de :",
     "amplifier un signal", "atténuer/adapter un signal", "inverser un signal", "filtrer les hautes fréquences", "B",
     "Un pont diviseur de tension ne peut qu'atténuer, jamais amplifier : la tension de sortie est toujours inférieure ou égale à la tension d'entrée. Si un capteur délivre un signal trop faible pour être bien exploité, le pont diviseur ne peut pas résoudre ce problème — il faudrait un véritable amplificateur (montage à AOP), qui est un autre sujet.", "Débutant"],
    ["S&P2", "Une sonde Pt100 est un capteur dont la grandeur électrique sensible à la température est :",
     "sa capacité", "sa résistivité", "sa tension de sortie directe", "son inductance", "B",
     "La Pt100 est un capteur résistif en platine (100 Ω à 0°C) : sa résistance varie avec la température, c'est pourquoi on la conditionne avec un pont diviseur de tension.", "Débutant"],
    ["S&P2", "La formule de la chaleur échangée par un corps de masse m et de chaleur massique c pour une variation de température ΔT est :",
     "Q = m/c × ΔT", "Q = m × c × ΔT", "Q = m × c / ΔT", "Q = m + c + ΔT", "B",
     "Q = m × c × ΔT, en joules si m est en kg, c en J.kg⁻¹.K⁻¹ et ΔT en K (ou °C, l'écart étant identique dans les deux échelles).", "Débutant"],
    ["S&P2", "Dans un système isolé (comme un calorimètre parfait), la somme des chaleurs échangées entre les corps vaut :",
     "toujours positive", "toujours négative", "0", "égale à la masse totale", "C",
     "Un système isolé n'échange rien avec l'extérieur : l'énergie thermique cédée par un corps est intégralement reçue par l'autre. D'où ΣQi = 0.", "Débutant"],
    ["S&P2", "Sur un graphique d'étalonnage, la méthode du « faisceau de droites » permet de déterminer :",
     "la valeur moyenne uniquement", "un encadrement (min/max) de la pente, donc son incertitude", "le nombre de points de mesure nécessaires", "la température ambiante", "B",
     "En traçant la droite la plus pentue et la moins pentue compatibles avec les rectangles d'incertitude des points extrêmes, on obtient un encadrement [amin ; amax] de la pente réelle — une estimation graphique de l'incertitude, sans aucun calcul.", "Débutant"],
  ];

  const avance = [
    ["S&P1", "Une autre carte utilise un CAN de 12 bits (au lieu de 10 bits sur l'Arduino Uno). Combien de niveaux de codage possibles cela représente-t-il ?",
     "4095", "1024", "2048", "4096", "D",
     "N = 2ⁿ = 2¹² = 4096 niveaux, numérotés de 0 à 4095. Attention à ne pas confondre le NOMBRE de niveaux (4096) avec la valeur MAXIMALE codée (4095).", "Avancé"],
    ["S&P1", "Pour ce CAN 12 bits, calibré sur Vmax = 3,3V, quel est le quantum (arrondi à 0,1 mV) ?",
     "0,8 mV", "3,3 mV", "1,2 mV", "4,89 mV", "A",
     "q = Vmax / (2ⁿ - 1) = 3300 / 4095 ≈ 0,806 mV, soit environ 0,8 mV. Attention, 4,89 mV est le quantum du CAN 10 bits de l'Arduino Uno sur 5V vu en cours — ce n'est pas la même carte ici !", "Avancé"],
    ["S&P1", "Entre le CAN 10 bits de l'Arduino Uno (sur 5V) et un CAN 8 bits (sur 5V également), lequel est le plus précis ?",
     "Le CAN 10 bits (quantum plus petit)", "Le CAN 8 bits (quantum plus petit)", "Les deux sont identiques", "Impossible à dire sans connaître fe", "A",
     "Plus n est grand, plus 2ⁿ est grand, donc plus le quantum q = Vmax/(2ⁿ-1) est petit : le CAN 10 bits distingue des écarts de tension plus fins que le CAN 8 bits, à Vmax identique.", "Avancé"],
    ["S&P1", "Un signal audio contient des fréquences jusqu'à fmax = 20 kHz (limite de l'audition humaine). D'après Shannon, la fréquence d'échantillonnage doit être strictement supérieure à :",
     "20 kHz", "30 kHz", "40 kHz", "60 kHz", "C",
     "fe > 2 × fmax = 2 × 20 = 40 kHz. C'est d'ailleurs pour rester au-dessus de ce seuil, avec une marge de sécurité, que les CD audio utilisent 44,1 kHz plutôt que tout juste 40 kHz.", "Avancé"],
    ["S&P1", "Dans le code de comparaison de résolution vu en AE02, l'instruction (brut >> 4) << 4 sert à dégrader une mesure 10 bits en 6 bits. Que fait concrètement l'opérateur >> 4 ?",
     "Multiplie la valeur par 4", "Divise par 16 en supprimant les 4 bits de poids faible", "Additionne 4 à la valeur", "Inverse tous les bits", "B",
     "Un décalage à droite de 4 bits (>>4) revient à diviser par 2⁴ = 16 en tronquant : les 4 bits de poids faible (les moins significatifs) sont perdus, ce qui dégrade la résolution.", "Avancé"],
    ["S&P2", "Une sonde Pt100 (R ≈ 107,79 Ω à 20°C) est conditionnée par un pont diviseur avec Rcond = 150 Ω et E = 5V, selon Vout = E × Rcond/(Rcond+Rcapteur). Quelle est la tension de sortie approximative ?",
     "1,45 V", "2,91 V", "3,50 V", "0,58 V", "B",
     "Vout = 5 × 150 / (150 + 107,79) = 5 × 150 / 257,79 ≈ 2,91 V.", "Avancé"],
    ["S&P2", "On chauffe un bloc de cuivre de 500 g de 20°C à 60°C (ccuivre ≈ 385 J.kg⁻¹.K⁻¹). Quelle quantité de chaleur a-t-il fallu fournir ?",
     "770 J", "3 850 J", "7 700 J", "15 400 J", "C",
     "Q = m × c × ΔT = 0,5 × 385 × 40 = 7700 J.", "Avancé"],
    ["S&P2", "Pour déterminer la chaleur massique d'un métal par la méthode des mélanges, quelle grandeur n'est PAS nécessaire ?",
     "La masse du métal", "La masse de l'eau", "La couleur du métal", "La température d'équilibre", "C",
     "La couleur du métal n'intervient dans aucune formule de calorimétrie — seules les masses, les températures et la capacité thermique du calorimètre sont nécessaires. Un bon réflexe d'esprit critique : identifier les grandeurs réellement utiles avant de se lancer dans un calcul.", "Avancé"],
    ["S&P2", "En passant de 3 à 5 points de mesure répartis sur toute la plage de température (même précision du thermomètre de référence), l'écart entre pente max et pente min obtenu par le faisceau de droites va généralement :",
     "Augmenter", "Diminuer", "Rester strictement identique", "Devenir nul", "B",
     "Plus les points extrêmes utilisés pour tracer le faisceau sont éloignés l'un de l'autre (grande plage de température couverte), moins l'incertitude ±u(T) sur chaque point pèse relativement sur la pente calculée : l'encadrement se resserre.", "Avancé"],
    ["S&P2", "Contrairement à la Pt100 (résistance qui augmente avec la température), une thermistance CTN a une résistance qui :",
     "Reste constante", "Augmente aussi avec la température", "Diminue avec la température", "Devient infinie", "C",
     "CTN signifie « Coefficient de Température Négatif » : sa résistance diminue quand la température augmente — un comportement inverse de celui de la Pt100.", "Avancé"],
  ];

  const existing = sh.getDataRange().getValues().slice(1).map(r => r[1]);
  debutant.concat(avance).forEach(r => {
    if (!existing.includes(r[1])) sh.appendRow(r);
  });
  SpreadsheetApp.getUi().alert("QCM initialisés : " + debutant.length + " question(s) Débutant, " + avance.length + " question(s) Avancé.");
}

function verifierCode(code) {
  const data = getSheet(SHEET_ELEVES).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(code)) {
      const role = (data[i][3] === "Professeur") ? "Professeur" : "Élève";
      return { ok: true, nom: data[i][0] + " " + data[i][1], role: role };
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

// ---------- Tableau de bord enseignant (accès réservé au rôle "Professeur") ----------
// Le rôle est défini dans l'onglet Élèves, colonne "Rôle" (mettre "Professeur" sur ta
// propre ligne). Aucun mot de passe séparé : c'est ton code personnel qui donne l'accès.

function getTableauDeBord(code) {
  const verif = verifierCode(code);
  if (!verif.ok || verif.role !== "Professeur") {
    return { ok: false, message: "Accès réservé au professeur." };
  }
  const eleves = getSheet(SHEET_ELEVES).getDataRange().getValues().slice(1).filter(r => r[0] && r[2]);
  const referentiel = getSheet(SHEET_REFERENTIEL).getDataRange().getValues().slice(1).filter(r => r[2]);
  const reponses = getSheet(SHEET_REPONSES).getDataRange().getValues().slice(1);
  const reponsesOutils = getSheet("Réponses_Outils").getDataRange().getValues().slice(1);

  // Dernier résultat de QCM par (code élève, chapitre)
  const qcmLatest = {};
  reponsesOutils.forEach(r => {
    const outil = String(r[3] || "");
    if (outil.indexOf("qcm-") !== 0) return;
    const chapitreQcm = outil.substring(4);
    const key = r[1] + "|" + chapitreQcm;
    if (!qcmLatest[key] || new Date(r[0]) > new Date(qcmLatest[key].date)) {
      try {
        const donnees = JSON.parse(r[4]);
        qcmLatest[key] = { date: r[0], score: donnees.score, total: donnees.total, niveau: donnees.niveau || "" };
      } catch (e) {}
    }
  });

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
    const qcm = {};
    chapitres.forEach(ch => {
      const key = code + "|" + ch;
      qcm[ch] = qcmLatest[key] ? { score: qcmLatest[key].score, total: qcmLatest[key].total, niveau: qcmLatest[key].niveau } : null;
    });
    return { nom: nom, niveaux: niveaux, qcm: qcm };
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
