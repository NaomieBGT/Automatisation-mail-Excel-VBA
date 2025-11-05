# Automatisation-mail-Excel-VBA
Macro VBA pour Excel :automatise la création d’e-mails Outlook à partir d’une feuille Contacts. Le script génère des brouillons d’e-mails (ou envoie directement si modifié) avec un objet et un corps multilingue et la possibilité d’ajouter une pièce jointe. L’objet et le corps incluent automatiquement le mois et l’année du jour 
Cas d’usage

Idéal pour :

L’envoi de reportings périodiques clients.

La génération de messages répétitifs personnalisés par destinataire.

La préparation de brouillons pour vérification avant envoi.

Feuille attendue : "Contacts"

La macro lit chaque ligne à partir de la ligne 2. Les colonnes attendues sont :

Colonne	Contenu
A	langue (ex: fr pour français, toute autre valeur → anglais)
B	compte (nom du compte pour l’objet du mail)
C	destinataire (adresse e-mail principale)
D	cc1 (copie carbone)
E	cc2
F	cc3
G	responsable (nom)
H	mailResponsable (adresse e-mail)
I	pj (chemin complet vers la pièce jointe, facultatif)
Fonctionnement

Création d’une instance d’Outlook.

Parcours des lignes de la feuille Contacts.

Construction du sujet : "Reporting Client – <compte> – <mois année actuel>".

Construction du corps du mail :

Français si la colonne A = "fr", sinon anglais.

Ajout de la pièce jointe si la colonne I est renseignée.

Ouverture des messages en brouillon (.Display).

Remplacer par .Send pour un envoi automatique.

Installation / utilisation

Ouvrir le fichier Excel.

Menu Développeur → Visual Basic (ou Alt+F11).

Insérer un nouveau Module et coller le code VBA.

Sauvegarder.

Vérifier que la feuille Contacts existe et respecte la structure ci-dessus.

Lancer la macro EnvoiReportingClient (via l’éditeur VBA ou raccourci).

Autoriser Outlook si une fenêtre d’autorisation apparaît.

Bonnes pratiques et sécurité

Tester sur quelques lignes avant d’utiliser .Send.

Ne pas stocker ni partager d’informations sensibles dans le fichier sans contrôle d’accès.

Vérifier les chemins des pièces jointes : si un fichier est absent, la macro peut échouer.

Respecter la confidentialité des destinataires (utiliser BCC si nécessaire).
