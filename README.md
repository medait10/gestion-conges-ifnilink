# IFNILINK Congés Maroc Pro

Application Flask compatible Python 3.13.

## Lancement

```powershell
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
python app.py
```

Compte local initial : `admin / admin123`

Base SQLite créée ici :
`leave_maroc_pro\instance\conges_ifnilink.db`

## Google OAuth

Dans Google Cloud :
- OAuth redirect URI : `http://127.0.0.1:5000/oauth2callback`
- Activer Gmail API
- Activer Google Calendar API
- Ajouter l'utilisateur comme Test user si l'application est en mode Testing

## Jours fériés

L'application contient :
- jours fériés fixes marocains jusqu'à 2100
- fallback indicatif Hijri 2023-2026
- bouton de synchronisation depuis Google Calendar pour importer les jours fériés standard/hijri du calendrier public configuré.

## Règles implémentées

- 18 jours/an au départ = 1,5 jour/mois
- +1,5 jour/an chaque 5 ans d'ancienneté, plafonné à 30 jours/an
- week-end et jours fériés non comptés
- congés exceptionnels : naissance, décès, mariage, circoncision, opération familiale
- approbation/refus, historique, exports Excel/PDF, import Google Calendar


## Version WOW ajoutée
- Dashboard avec sélection d'année.
- Calendrier grégorien + Hijri approximatif affiché dans chaque jour.
- Jours fériés Hijri synchronisables depuis Google Calendar.
- Historique avec jours pris.
- Design glassmorphism moderne.
- Scope Google corrigé avec calendar.events.

En cas d'erreur "Scope has changed", supprime la base/token :
```powershell
del .\instance\conges_ifnilink.db
python app.py
```


## Ajouts version mode & historique
- Menu gauche avec scrollbar.
- Mode Dark / Light via le menu.
- Historique prérempli avec les congés déjà pris comme demandes approuvées.
- Les jours pris restent visibles dans l'historique.


## Ajouts V2
- Sélecteur Dark/Light en haut à droite.
- Liste années dynamique : de 2023 jusqu'à année courante + 1.
- Export Excel et PDF avec design amélioré.
- Type de demande : Repos maladie.
- Champ référence/note maladie.


## Ajouts V3 robustesse
- Si l’envoi Gmail API échoue, la demande est annulée et n’est pas enregistrée.
- Validation destinataire email obligatoire si l’option d’envoi est cochée.
- Pages d’erreur propres 403/404/500.
- Debug Flask désactivé par défaut pour éviter l’écran traceback en production locale.


## Ajouts V4 dynamique
- Input “Jours autorisés occasion / CNSS / loi”.
- Calcul automatique : jours ouvrables période - jours autorisés = jours à déduire du solde.
- Exemple : naissance, 5 jours ouvrables, 3 jours autorisés => 2 jours déduits.
- Historique affiche jours période, jours non déduits et jours déduits.
- Interface plus fluide avec prévisualisation instantanée côté navigateur.


## Ajouts V5 Calendriers Google
- Page Jours fériés avec sélection : Tous / Standard / Hijri.
- Synchronisation séparée depuis Google Calendar :
  - Standard Maroc : GOOGLE_HOLIDAY_CALENDAR_ID
  - Hijri Maroc : GOOGLE_HIJRI_CALENDAR_ID
- Les jours Hijri importés sont marqués comme Hijri dans la base.


## Correctif V6 OAuth
- Ajout des scopes courts Google `email` et `profile`.
- Ajout `OAUTHLIB_RELAX_TOKEN_SCOPE=1`.
- Callback Google sécurisé : retour propre au login avec message au lieu d'une erreur 500.


## Correctif V7 PKCE
- Correction de l'erreur Google OAuth : `(invalid_grant) Missing code verifier`.
- Génération du `code_verifier` au login.
- Envoi du `code_verifier` au callback OAuth.


## Ajouts V9
- Mariage du salarié : 2 jours offerts/non déduits par défaut, le reste est déduit du solde global.
- Exemple : période 4 jours ouvrables, 2 offerts => 2 jours déduits.
- Repos maladie : certificat médical/référence obligatoire.
- Le champ certificat médical est sécurisé côté UI et visible uniquement pour Repos maladie.


## Ajouts V10
- Possibilité d'annuler une demande en attente ou déjà approuvée.
- Si le congé approuvé était importé dans Google Calendar, l'événement est supprimé lors de l'annulation.
- Le solde est recalculé après annulation d'un congé approuvé.
- La page Demande contient une div avec la liste des dernières demandes sous le formulaire.


## Ajouts V11
- Ajout du menu gauche : Liste des demandes.
- Nouvelle page Liste des demandes.
- Synchronisation inverse depuis Google Calendar :
  - cherche les événements contenant “Congé”
  - ajoute dans la liste s’ils n’existent pas déjà
  - marque la demande comme approuvée et synchronisée.
- Annulation conserve la suppression Calendar si l’événement existe.


## Ajouts V12
- Bouton clair : Synchroniser avec calendrier.
- Tableaux et listes responsive PC/tablette/smartphone.
- Scroll horizontal sur PC si nécessaire.
- Transformation des tableaux en cartes sur smartphone.
