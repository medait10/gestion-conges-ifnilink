# MEDFLOW Congés Maroc Pro

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


## V13 Admin DB
- Menu Admin DB pour voir/modifier/supprimer users, leave_requests, monthly_balances, holidays depuis Render.
- Accès réservé admin.


## V14 Sécurité
- Champs sensibles masqués dans Admin DB : password_hash, google_token.
- Suppression de l’affichage du compte par défaut sur la page login.
- Sidebar rendue générique : pas d’informations personnelles de connexion.
- Headers sécurité ajoutés : X-Frame-Options, nosniff, no-referrer, no-store.
- Route `/admin/change_password` pour changer le mot de passe admin.
- Mot de passe admin initial configurable par Render Environment : `ADMIN_PASSWORD`.

Mot de passe admin fort proposé :
Adm-IFNI-2026!Qx7#M9v2

Sur Render, ajoute :
ADMIN_PASSWORD=Adm-IFNI-2026!Qx7#M9v2
FLASK_ENV=production


## V15 MEDFLOW
- Nom du site remplacé par MEDFLOW.
- Logo SVG sophistiqué ajouté.
- Page inscription utilisateur.
- Les utilisateurs créent leurs comptes et accèdent à leurs propres demandes.
- Exports Excel/PDF rebrandés MEDFLOW.


## V16 Subscription Platform
- Admin interface reserved only for ADMIN_EMAIL.
- Default admin email: aitelmalemmohamed@gmail.com.
- Other users can create accounts and must subscribe to use the application.
- Stripe Checkout subscription integration.
- Stripe webhook endpoint: /stripe_webhook.
- Environment variables needed:
  ADMIN_EMAIL
  STRIPE_PUBLISHABLE_KEY
  STRIPE_SECRET_KEY
  STRIPE_PRICE_ID
  STRIPE_WEBHOOK_SECRET
  APP_BASE_URL


## V17 Backup Strategy
- Menu admin `Backups`.
- Backup manuel SQLite.
- Téléchargement du fichier .db.
- Restauration contrôlée avec confirmation `RESTORE`.
- Backup automatique quotidien au démarrage de l’application.
- Conservation des 20 derniers backups.
- Important Render Free : télécharger régulièrement les backups, car le disque peut être éphémère.
- Pour production : migrer vers PostgreSQL avec backups automatiques.


## V18 User Email + Guide + Redesign
- Chaque utilisateur connecté Google envoie les emails depuis son propre compte Gmail autorisé.
- Profil affiche l’état Google et permet de déconnecter Google.
- Menu Guide utilisateur ajouté.
- Pages confidentialité et conditions ajoutées.
- Design plus universel et accessible.
- CSP et headers sécurité renforcés.


## V19 Owner + Trial + Multi-user clean data
- Owner/admin unique : Mohamed AIT ELMALEM / aitelmalemmohamed@gmail.com.
- Les interfaces admin sont réservées uniquement à ADMIN_EMAIL.
- Les données historiques préchargées concernent uniquement le owner.
- Chaque nouvel utilisateur démarre à zéro.
- Chaque nouvel utilisateur peut créer son compte par Google Gmail ou inscription manuelle.
- Nouveaux utilisateurs : période d’essai gratuite de 7 jours.
- Après expiration, abonnement requis.
- Dashboard et exports filtrés par utilisateur.


## V20 Secure Payment + Social Redesign
- Paiement sécurisé via Stripe Checkout hébergé.
- Billing portal Stripe pour gérer l’abonnement.
- Checkout Session renforcée : metadata, customer_email, billing address auto, promotion codes.
- Landing page moderne type SaaS/social app.
- Pricing page plus rassurante.
- Design responsive retravaillé.
- Variables Render à configurer :
  STRIPE_SECRET_KEY
  STRIPE_PUBLISHABLE_KEY
  STRIPE_PRICE_ID
  STRIPE_WEBHOOK_SECRET
  APP_BASE_URL


## V21 Multi-language
- Langues ajoutées : FR, EN, DE, ES, AR.
- Sélecteur de langue dans le topbar.
- Langue sauvegardée en session.
- Support RTL pour arabe.
- Dictionnaire simple dans app.py avec fonction `t(key)`.


## V22 Error i18n fix
- Correction `jinja2.exceptions.UndefinedError: 't' is undefined`.
- Ajout d'une fonction `safe_t`.
- Ajout de globals Jinja pour t/languages/current_lang/is_rtl.
- Page error.html rendue indépendante de base.html pour éviter les erreurs en cascade.


## V23 Social Design + Payment Setup + Better i18n
- Interface réorganisée façon grande plateforme sociale : topbar, left rail, content feed, right rail.
- Aucun nom/email affiché dans l’interface principale.
- Admin DB masque aussi email/full_name/tokens/password/Stripe IDs.
- Traduction automatique partielle du contenu via static/i18n.js selon langue choisie.
- Pricing mensuel et annuel :
  - 29 MAD/mois
  - 299 MAD/an
- Paiement via Stripe Checkout hébergé.
- Variables Render :
  STRIPE_MONTHLY_PRICE_ID
  STRIPE_ANNUAL_PRICE_ID
  STRIPE_SECRET_KEY
  STRIPE_WEBHOOK_SECRET
  APP_BASE_URL
- Pages ajoutées :
  /about
  /copyright


## V24 Route fix
- Correction endpoint Flask `about`.
- Correction endpoint Flask `copyright_page`.
- Évite l’erreur `Could not build url for endpoint 'about'`.


## V25 Responsive
- Sidebar mobile hamburger
- Responsive smartphone/tablette/desktop
- Tables scrollables
- KPI adaptatifs
- Cards fluides
- Topbar compacte
- Optimisation écrans < 860px / < 640px / < 420px


## MEDFLOW V26 Ultra UI
- Rebranding complet vers MEDFLOW.
- Logo premium M gradient bleu/indigo/cyan.
- Nouvelle landing page SaaS premium.
- Nouvelle page pricing premium.
- Palette couleurs moderne : Blue / Indigo / Cyan.
- UI inspirée des standards SaaS premium : Stripe, Linear, Notion.
- Responsive optimisé conservé.

## MEDFLOW V29
- Suppression du panneau droit.
- Traduction serveur globale via tt().
- Menus, tables, champs, statuts, mois et calendrier traduits selon langue choisie.
- JavaScript de traduction partielle désactivé pour éviter mélange de langues.
- Layout corrigé sans déformation.

## MEDFLOW V30 Premium i18n UI
- Traduction renforcée serveur + navigateur pour les textes restants.
- Menus, champs, placeholders, tables, statuts et pages principales traduits.
- Icônes premium CSS au lieu d’emojis.
- Scrollbars stylées.
- Animations premium avec respect prefers-reduced-motion.
- UI plus moderne et responsive.

## MEDFLOW V31 Instagram-inspired + Translation Fix
- Traduction renforcée des options de congé et textes visibles restants.
- `leave_label_i18n()` côté serveur pour traduire les types de congé.
- Traduction client des options select et placeholders restants.
- Style inspiré Instagram : gradient violet/rose/orange/bleu.
- Icônes et animations premium.
- Sélecteurs et scrollbars améliorés.

# MEDFLOW SaaS Ultimate V32

Version SaaS commercialisable basée sur V31 :
- UI premium Instagram/Linear-inspired
- Multilingue FR/EN/DE/ES/AR
- RTL arabe
- Responsive mobile/tablette/desktop
- Stripe subscriptions monthly/yearly
- Google OAuth Gmail/Calendar
- Owner/admin sécurisé
- Multi-users avec essai 7 jours
- Backup/restore SQLite
- Exports PDF/Excel
- Calendrier Maroc Standard & Hijri
- Dark/Light mode
- Scrollbars et animations premium

## MEDFLOW V33 Full i18n + Profiling
- Ajout page admin `/admin/profiling`.
- Mesure temps réponse par endpoint.
- Liste dernières requêtes.
- Traduction renforcée des textes FR restants côté serveur et navigateur.
- Ajout helper `tr_text()`.


## MEDFLOW V34
- Correction `tr_text is undefined`.
- Ajout global Jinja sécurisé.
- Page error.html indépendante.


## MEDFLOW V35 Full Language
- Pages publiques remplacées en 100% i18n serveur.
- Dictionnaire FR/EN/DE/ES/AR renforcé.
- Fallback navigateur pour textes anciens restants.
- Login/Register/About/Privacy/Terms/Copyright/Pricing/Home refaits en i18n.
