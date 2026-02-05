# TIFF to DOCX Converter

Convertisseur de fichiers TIFF multi-pages en documents Word (DOCX) avec OCR Google Vision.

## Fonctionnalités

- Upload de fichiers TIFF multi-pages
- Aperçu de toutes les pages en miniatures
- OCR avec Google Cloud Vision API
- Exclusion de pages de l'OCR (clic sur miniatures)
- Document DOCX formaté avec images et texte extrait
- Page de résumé avec informations de conversion
- Progression en temps réel

## Installation locale

```bash
# Cloner le repository
git clone <url-du-repo>
cd tiff_converter_app

# Créer un environnement virtuel
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/Mac

# Installer les dépendances
pip install -r requirements.txt

# Configurer les variables d'environnement
copy .env.example .env
# Éditer .env avec votre clé API Google Vision

# Lancer l'application
python app.py
```

Ouvrir http://localhost:5000

## Déploiement

### Option 1: Render (Recommandé - Gratuit)

1. Créer un compte sur [render.com](https://render.com)
2. Cliquer sur "New" → "Web Service"
3. Connecter votre repository GitHub
4. Configurer:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn --worker-class geventwebsocket.gunicorn.workers.GeventWebSocketWorker -w 1 --bind 0.0.0.0:$PORT app:app`
5. Ajouter les variables d'environnement:
   - `GOOGLE_API_KEY`: Votre clé API Google Vision
   - `SECRET_KEY`: Une clé secrète aléatoire
   - `ASYNC_MODE`: `gevent`

### Option 2: Railway

1. Créer un compte sur [railway.app](https://railway.app)
2. Cliquer sur "New Project" → "Deploy from GitHub repo"
3. Sélectionner votre repository
4. Ajouter les variables d'environnement dans "Variables"

### Option 3: Heroku

```bash
# Installer Heroku CLI
# Puis:
heroku login
heroku create nom-de-votre-app
heroku config:set GOOGLE_API_KEY=votre-cle-api
heroku config:set SECRET_KEY=votre-cle-secrete
heroku config:set ASYNC_MODE=gevent
git push heroku master
```

## Configuration Google Cloud Vision

1. Aller sur [Google Cloud Console](https://console.cloud.google.com)
2. Créer un projet ou en sélectionner un
3. Activer l'API "Cloud Vision API"
4. Créer une clé API dans "APIs & Services" → "Credentials"
5. Copier la clé dans la variable `GOOGLE_API_KEY`

## Variables d'environnement

| Variable | Description | Défaut |
|----------|-------------|--------|
| `SECRET_KEY` | Clé secrète Flask | `tiff-converter-secret-key-2024` |
| `GOOGLE_API_KEY` | Clé API Google Vision | - |
| `ASYNC_MODE` | Mode async (`threading` ou `gevent`) | `threading` |
| `UPLOAD_FOLDER` | Dossier des uploads | `uploads` |
| `OUTPUT_FOLDER` | Dossier des fichiers générés | `output` |

## Avertissement

Les fichiers sont envoyés à Google Cloud Vision pour l'OCR. Ne pas utiliser pour des documents sensibles ou confidentiels.

---

Développé par **Hervé Lenglin** | Propulsé par Google Cloud Vision
