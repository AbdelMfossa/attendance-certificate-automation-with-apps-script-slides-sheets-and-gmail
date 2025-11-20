# Automatisation de la Génération et de l'Envoi d' Attestations de Participation avec Google Apps Script

Ce projet implémente une solution robuste et automatisée pour la gestion des attestations de participation d'un événement, inspirée par le **DevFest Yaoundé**. Il permet de générer des identifiants uniques, de créer des certificats personnalisés à partir de modèles Google Slides, et de les envoyer par e-mail via Gmail, le tout géré depuis un Google Sheet.

## Fonctionnalités

  - **Génération d'Identifiants Uniques :** Attribue des IDs `devfest25-XXXXXX` à chaque participant manquant.
  - **Création Automatisée de Certificats :** Génère des attestations personnalisées (PDF) à partir de modèles Google Slides.
  - **Support Multi-Modèles :** Utilise différents modèles de certificats en fonction du rôle du participant (ex: "Attendee", "Organizer").
  - **Envoi d'E-mails Personnalisés :** Expédie les certificats en pièce jointe (PDF) via Gmail, avec un corps d'e-mail personnalisable.
  - **Suivi des Statuts :** Met à jour le Google Sheet avec l'emplacement des fichiers générés et le statut d'envoi (`CREATED`, `SENT`).
  - **Menu Personnalisé :** Intègre un menu "DevFest 2025" directement dans Google Sheets pour une exécution facile des scripts.

## Technologies utilisées

  - [Google Apps Script](https://script.google.com/)
  - [Google Sheets](https://sheets.google.com)
  - [Google Slides](https://slides.google.com)
  - [Gmail](https://mail.google.com)
  - [Google Drive](https://drive.google.com)

## Configuration requise

1.  **Créer une Google Sheet** avec les en-têtes de colonnes suivants (respecter la casse exacte) :

      * `Full Name`
      * `Date`
      * `Title` (ex: "Attendee", "Organizer")
      * `Email`
      * `Id`
      * `File Location`
      * `Status`

2.  **Préparer des Modèles Google Slides :**

      * Créez au moins un modèle de certificat dans Google Slides (par ex., un pour les "Attendees" et un pour les "Organizers").
      * Utilisez des "placeholders" pour les informations à personnaliser, par exemple : `{participant_name}`, `{role}`, `{issued_date}`, `{certificate_id}`.
      * Notez les **IDs de ces fichiers Slides** (`slideOVTemplateId` et `slideATemplateId` dans le code).

3.  **Créer un Dossier Temporaire dans Google Drive :**

      * Créez un dossier vide dans Google Drive pour y stocker temporairement les certificats générés.
      * Notez l'**ID de ce dossier** (`tempFolderId` dans le code).

4.  **Créer un Brouillon d'E-mail dans Gmail :**

      * Rédigez un brouillon d'e-mail avec le sujet exact : `Certificate of Participation – DevFest Yaoundé 2025`.
      * Dans le corps de l'e-mail, utilisez `{{fullName}}` pour le nom du participant.
      * Vous pouvez également ajouter des destinataires en copie (`Cc`) dans ce brouillon, ils seront repris automatiquement.

5.  **Copier le Code Apps Script :**

      * Ouvrez votre Google Sheet.
      * Allez dans `Extensions > App Script`.
      * Copiez le code source complet (`Code.gs`) dans l'éditeur.
      * **Mettez à jour les constantes** `slideOVTemplateId`, `slideATemplateId`, et `tempFolderId` avec les IDs que vous avez notés.

## Utilisation

Une fois le script configuré et enregistré, un nouveau menu **"DevFest 2025"** apparaîtra dans votre Google Sheet avec les options suivantes :

1.  **`Generate ID` :** Génère des identifiants uniques pour tous les participants n'en ayant pas encore un dans la colonne `Id`.
2.  **`Create certificates` :** Parcourt le Google Sheet, génère les certificats personnalisés (à partir des modèles Slides) pour les participants dont le statut est vide, et met à jour les colonnes `File Location` et `Status`.
3.  **`Send certificates` :** Envoie les e-mails personnalisés avec le certificat PDF en pièce jointe aux participants dont le statut est `CREATED`, puis met à jour leur statut à `SENT`.

## Limites et Quotas Gmail

Il est important de prendre en compte les quotas d'envoi de Gmail :

  * **Comptes Personnels (@gmail.com) :** Environ **100 e-mails par jour**.
  * **Comptes Google Workspace (professionnels) :** Jusqu'à **1 500 ou 2 000 e-mails par jour** selon l'édition.

Pour les grands volumes, l'utilisation d'un compte Google Workspace est fortement recommandée. Alternativement, le script peut être modifié pour envoyer les e-mails par lots sur plusieurs jours (nécessitant un déclencheur temporel).
