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
2. Dans `Code.gs`, ajoute une ligne dans le tableau `MODULES` :
   ```js
   { chapitre: "S&Px", label: "Nom affiché", page: "cle-url", template: "NouvelOutil", type: "outil" },
   ```
   Utilise `type: "page"` seulement pour des pages générales (comme Auto-évaluation) qui doivent apparaître directement dans le menu principal. Pour un outil lié à une activité, utilise `type: "outil"` — il apparaîtra automatiquement dans le hub **Outils** plutôt que d'encombrer le menu principal.
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
| `PT100Etalonnage.html` | Outil : faisceau de droites pour l'étalonnage Pt100 (S&P2) |
| `Outils.html` | Hub qui liste tous les outils (type "outil"), groupés par chapitre |
| `TableauDeBord.html` | Vue classe entière pour toi, protégée par mot de passe (défini à la première connexion) |
| `QCM.html` | QCM de révision (5 questions/chapitre), correction détaillée immédiate après chaque réponse |
| `CodesSource.html` | Module : bibliothèque de codes Arduino/Python, recherche + copie en un clic |
| `Nav.html` | Fragment de navigation partagé (barre horizontale unique), inclus dans chaque page |
| `Style.html` | Styles CSS partagés (couleurs, fond, boutons...) — modifie ici pour changer le rendu de toutes les pages d'un coup |
| `Banniere.html` | Bandeau noir avec l'emblème du chapitre, inclus dans chaque page |

## Base de données (le Google Sheet)

- **Élèves** : Nom, Prénom, Code — rempli à la main + codes générés par le menu.
- **Référentiel** : Chapitre, Activité, Code compétence, Intitulé — à compléter au fil des chapitres.
- **Réponses** : rempli automatiquement à chaque auto-évaluation.
- **Réponses_Outils** : rempli automatiquement par les modules type "outil" (ex. Pt100), au format horodatage / code / outil / données JSON.
- **Codes** : Chapitre, Langage (Arduino/Python), Titre, Description, Code — bibliothèque de codes affichée dans le module *Codes Arduino / Python*. **Tu peux ajouter une ligne directement dans cet onglet, aucun besoin de repasser par GitHub/clasp** — le module la récupère automatiquement à chaque ouverture. Pour le remplissage initial (S&P1 + S&P2), utilise le menu **Suivi SPCL → Initialiser la bibliothèque de codes**.
- **QCM** : Chapitre, Question, OptionA-D, BonneReponse (A/B/C/D), Explication, **Niveau** (Débutant/Avancé, dernière colonne) — **tu peux ajouter des questions directement dans cet onglet**, aucun redéploiement nécessaire. Remplissage initial via **Suivi SPCL → Initialiser les QCM** (ça migre aussi automatiquement les anciennes questions vers "Débutant" si tu avais déjà lancé cette action avant l'ajout du niveau).

## Tableau de bord enseignant

Le module **Tableau de bord (prof)** te montre toute la classe, chapitre par chapitre, avec le dernier niveau (A/B/C/D) de chaque élève sur chaque compétence — sans avoir à ouvrir le Sheet.

Il est protégé par un **mot de passe séparé des codes élèves**. Ce mot de passe doit être **défini explicitement par toi**, depuis le Sheet : menu **Suivi SPCL → Définir le mot de passe du Tableau de bord**. Tant qu'il n'a pas été défini de cette façon, la page refuse tout accès — personne ne peut plus "prendre" le mot de passe en le devinant ou en tapant n'importe quoi en premier. Tu peux le changer à tout moment en relançant la même action de menu.

Les élèves, eux, ont un onglet **« Mon suivi »** dans la page Auto-évaluation (à côté de "Nouvelle auto-évaluation"), qui leur montre leur propre historique — mais uniquement leurs résultats, jamais ceux des autres.
