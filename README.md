## Idée de repas (Google Sheets) — site statique + API Apps Script

Ce repo contient:
- `index.html` : le site statique (GitHub Pages OK)
- `apps-script/` : un Google Apps Script à déployer en Web App pour lire la Google Sheet et renvoyer une recette aléatoire

### Pourquoi Google Sheets ?

- Tu modifies/ajoutes des recettes quand tu veux dans la Sheet
- Le site ne change plus (juste l’URL `/exec` à configurer une fois)

### Déploiement de l’API (Google Apps Script)

1. Va sur Google Apps Script et crée un nouveau projet.
2. Copie le contenu de `apps-script/Code.gs` dans ton `Code.gs`.
3. (Optionnel) copie `apps-script/appsscript.json` dans ton projet.
4. Dans Apps Script: **Project Settings** → **Script Properties**, ajoute:
   - `SPREADSHEET_ID` = l’ID de ta Google Sheet (dans l’URL)
   - `SHEET_NAME` = le nom de l’onglet (ex: `Recettes`)
5. Déploie: **Deploy** → **New deployment** → **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copie l’URL `/exec` donnée par Google.

### Configurer le site

Dans `index.html`, remplace:

- `API_BASE_URL = 'TON_URL_APPS_SCRIPT_ICI'`

par l’URL `/exec` de ton Apps Script.


