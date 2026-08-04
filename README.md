# Suivi des compétences SPCL

Application Google Apps Script (attachée au Google Sheet de suivi), gérée via `clasp` et versionnée sur GitHub.

## Mise en place (à faire une seule fois)

### 1. Créer le projet Apps Script attaché au Sheet
- Ouvre ton Google Sheet (celui avec les onglets *Élèves / Référentiel / Réponses*).
- **Extensions → Apps Script**. Ça crée un projet Apps Script vide, automatiquement lié à ce Sheet.
- Dans l'éditeur Apps Script : **⚙️ Paramètres du projet** → copie le **Script ID**.

### 2. Installer clasp
```bash
npm install -g @google/clasp
clasp login
```
(Ça ouvre ton navigateur pour autoriser ton compte Google.)

### 3. Récupérer ce repo et le lier au projet Apps Script
```bash
git clone <URL_DE_CE_REPO>
cd suivi-competences-spcl
clasp clone <SCRIPT_ID_COPIÉ_ÉTAPE_1>
```
Si `clasp clone` proteste car des fichiers existent déjà, crée juste un fichier `.clasp.json` à la main :
```json
{
  "scriptId": "COLLE_TON_SCRIPT_ID_ICI",
  "rootDir": "."
}
```

### 4. Envoyer le code vers Apps Script
```bash
clasp push
```

### 5. Déployer en tant qu'application web
Dans l'éditeur Apps Script (`clasp open` pour l'ouvrir directement) :
**Déployer → Nouveau déploiement → Type : Application web**
- Exécuter en tant que : **Moi**
- Qui a accès : **Tout le monde**
- Copie l'URL générée : c'est celle que tu donnes à tes élèves (sur Google Classroom par exemple).

### 6. Préparer les données
Retourne sur le Google Sheet, rafraîchis la page — le menu **« Suivi SPCL »** apparaît.
1. **Initialiser le référentiel** (une fois, pré-remplit S&P1 + S&P2)
2. Remplis les colonnes Nom/Prénom dans l'onglet **Élèves**
3. **Générer les codes élèves manquants**

## Workflow au quotidien (ajouter un nouveau module)

1. Crée un nouveau fichier `.html` dans ce dossier (ex. `NouvelOutil.html`), sur le modèle de `PT100Etalonnage.html` (même bandeau, même nav).
2. Dans `Code.gs`, ajoute une ligne dans l'objet `PAGES` :
   ```js
   "nouvel-outil": "NouvelOutil",
   ```
3. `git add . && git commit -m "Ajout du module X" && git push`
4. `clasp push` pour envoyer vers Apps Script.
5. Dans Apps Script : **Déployer → Gérer les déploiements → ✏️ → Nouvelle version → Déployer** (l'URL ne change pas).

Le module est alors accessible via `<url du déploiement>?page=nouvel-outil`, et apparaît dans la barre de navigation de chaque page si tu l'ajoutes au `<div class="nav">`.

## Structure des fichiers

| Fichier | Rôle |
|---|---|
| `appsscript.json` | Manifeste Apps Script (obligatoire pour clasp) |
| `Code.gs` | Backend : routage des pages, gestion élèves/codes, lecture/écriture du Sheet |
| `Index.html` | Page d'auto-évaluation des compétences |
| `PT100Etalonnage.html` | Module : faisceau de droites pour l'étalonnage Pt100 (S&P2) |
| `CodesSource.html` | Module : bibliothèque de codes Arduino/Python, recherche + copie en un clic |
| `Nav.html` | Fragment de navigation partagé (onglets par chapitre), inclus dans chaque page |

## Base de données (le Google Sheet)

- **Élèves** : Nom, Prénom, Code — rempli à la main + codes générés par le menu.
- **Référentiel** : Chapitre, Activité, Code compétence, Intitulé — à compléter au fil des chapitres.
- **Réponses** : rempli automatiquement à chaque auto-évaluation.
- **Réponses_Outils** : rempli automatiquement par les modules type "outil" (ex. Pt100), au format horodatage / code / outil / données JSON.
- **Codes** : Chapitre, Langage (Arduino/Python), Titre, Description, Code — bibliothèque de codes affichée dans le module *Codes Arduino / Python*. **Tu peux ajouter une ligne directement dans cet onglet, aucun besoin de repasser par GitHub/clasp** — le module la récupère automatiquement à chaque ouverture. Pour le remplissage initial (S&P1 + S&P2), utilise le menu **Suivi SPCL → Initialiser la bibliothèque de codes**.
