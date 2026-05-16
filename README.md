# Gestion congés - Google OAuth + Gmail API

Compatible Python 3.13.

## Lancement
```powershell
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
python app.py
```

## Google Cloud
Type d'application: Application Web

Origine JavaScript autorisée:
http://127.0.0.1:5000

URI de redirection autorisée:
http://127.0.0.1:5000/oauth2callback

Scopes utilisés:
- openid
- userinfo.email
- userinfo.profile
- gmail.send

Mettre GOOGLE_CLIENT_ID et GOOGLE_CLIENT_SECRET dans `.env`.


## Base de données SQLite
La base est créée automatiquement ici : `instance/conges.db`.
Pour l’ouvrir : installe DB Browser for SQLite, puis Open Database > `instance/conges.db`.
Depuis l’application, une route de diagnostic existe : `/database-info`.

## Workflow directeur
Une demande créée reste `En attente`.
- `Approuver` : comptabilise les jours, ajoute la période dans le bilan et met à jour le solde.
- `Refuser` : ne comptabilise rien, aucun X/solde n’est modifié.
