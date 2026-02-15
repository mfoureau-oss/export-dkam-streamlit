# Export Présentation DKAM (source)

Ce dossier est la **version source légère** de l’application (sans le Python embarqué à 40 000 fichiers).
Objectif : permettre à quelqu’un de modifier l’interface / le fonctionnement et de relancer l’app.

## 1) Pré-requis (1 seule fois sur le PC)
- Installer **Python 3.11 ou 3.12** (Windows).
  - Pendant l’installation : cocher **“Add Python to PATH”**.

## 2) Installation (1ère fois dans ce dossier)
Ouvrir l’invite de commandes dans le dossier (clic droit > “Ouvrir dans le Terminal”) puis exécuter :

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 3) Lancer l’application
Toujours dans le terminal :

```bat
.venv\Scripts\activate
streamlit run app\streamlit_app.py
```

Le navigateur s’ouvre sur l’application.

## 4) Où modifier quoi ?
- **Interface (textes, champs, boutons)** : `app/streamlit_app.py` (fonction `app_main()`).
- **Récupération Tableau** : classe `TableauSession` dans `app/streamlit_app.py`.
- **Récupération Looker** :
  - Gmail : fonctions `gmail_service_from_refresh()` / `fetch_latest_looker_pdf_bytes_gmail()`
  - URL : `fetch_looker_pdf_from_url()`
- **Mise en page PPT** : template `.pptx` dans `app/templates/` (placeholders `PH_TBL`, `PH_LKR_1`, etc.).

## 5) Secrets / identifiants (important)
Ne pas stocker d’identifiants en clair dans le code.
Deux options :

### Option A (recommandée) : variables d’environnement Windows
Définir ces variables (si vous utilisez Gmail) :
- `GMAIL_CLIENT_ID`
- `GMAIL_CLIENT_SECRET`
- `GMAIL_REFRESH_TOKEN`
- `GMAIL_USER` (optionnel)

Et (optionnel) réglages d’image :
- `TOPBAR_CROP_PCT`
- `IMAGE_FIT_MODE`
- `LKR_CROP_TOP`, `LKR_CROP_BOTTOM`, `LKR_CROP_LEFT`, `LKR_CROP_RIGHT`

### Option B : Streamlit secrets
Copier `.streamlit/secrets.toml.example` en `.streamlit/secrets.toml` puis remplacer `REPLACE_ME`.

⚠️ Ne pas commiter `.streamlit/secrets.toml` dans un repo.

## 6) Template PPT
Le template utilisé est dans `app/templates/`.
Vous pouvez le modifier via PowerPoint (positions des zones, textes, etc.).
