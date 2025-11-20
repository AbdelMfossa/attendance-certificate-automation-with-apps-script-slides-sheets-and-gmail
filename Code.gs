const slideOVTemplateId = 'XXXXXXXXXX'; // REMPLACER PAR L'ID DE VOTRE MODÈLE SLIDES ORGANISATEURS/SPEAKERS
const slideATemplateId = 'XXXXXXXXXX';   // REMPLACER PAR L'ID DE VOTRE MODÈLE SLIDES ATTENDEES
const tempFolderId = 'XXXXXXXXXX'; // REMPLACER PAR L'ID DE VOTRE DOSSIER TEMPORAIRE DANS GOOGLE DRIVE

/**
 * Crée un menu personnalisé "DevFest 2025" dans le tableur
 * avec des options pour générer des IDs, créer et envoyer des certificats.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('DevFest 2025')
      .addItem('1- Generate ID', 'generateUniqueIds')
      .addSeparator()
      .addItem('2- Create certificates', 'createCertificates')
      .addSeparator()
      .addItem('3- Send certificates', 'sendCertificates')
      .addToUi();
}

/**
 * Génère des identifiants uniques pour les participants qui n'en ont pas encore.
 * Les IDs sont de la forme 'devfest25-XXXXXX'.
 */
function generateUniqueIds() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = sheet.getDataRange().getValues();
  var idColumnIndex = 4; // Colonne 'Id' (E) - Note: les index commencent à 0, donc la 5ème colonne est l'index 4

  for (let i = 1; i < values.length; i++) { // Commence à la ligne 2 pour éviter la ligne d'en-tête (index 1)
    var cell = sheet.getRange(i + 1, idColumnIndex + 1); // i+1 car getRange est 1-basé
    if (!cell.getValue()) { // Vérifie si la cellule 'Id' est vide
      var uniqueId = 'devfest25-' + Utilities.getUuid().substring(0, 6); // Génère un ID unique
      cell.setValue(uniqueId);
    }
  }
  SpreadsheetApp.flush(); // Force la mise à jour du Sheet
  SpreadsheetApp.getUi().alert('Identifiants générés avec succès !');
}

/**
 * Crée un certificat personnalisé pour chaque participant dont le statut n'est pas "CREATED" ou "SENT".
 * Les certificats sont basés sur des modèles Google Slides et stockés dans un dossier Drive temporaire.
 */
function createCertificates() {
  let templateFile = "";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  // Recherche des index des colonnes par leur nom
  const fullNameIndex = headers.indexOf('Full Name');
  const dateIndex = headers.indexOf('Date');
  const titleIndex = headers.indexOf('Title');
  const idIndex = headers.indexOf('Id');
  const fileLocationIndex = headers.indexOf('File Location');
  const statusIndex = headers.indexOf('Status');

  const tempFolder = DriveApp.getFolderById(tempFolderId);

  for (let i = 1; i < values.length; i++) { // Itère sur chaque ligne de données (à partir de la 2ème ligne)
    const rowData = values[i];
    const fullName = rowData[fullNameIndex];
    const date = rowData[dateIndex];
    const title = rowData[titleIndex];
    const fileLocation = rowData[fileLocationIndex];
    const status = rowData[statusIndex];

    // Ne génère que si le certificat n'a pas déjà été créé ou envoyé
    if (!fileLocation || (status !== 'CREATED' && status !== 'SENT')) {
      if (title === "Attendee") {
        templateFile = DriveApp.getFileById(slideATemplateId); // Modèle pour les participants
      } else {
        templateFile = DriveApp.getFileById(slideOVTemplateId); // Modèle pour les autres rôles (Organizers, Speakers)
      }
      const id = rowData[idIndex];

      // Crée une copie du modèle Slides et la renomme
      const empSlideId = templateFile.makeCopy(tempFolder).setName(fullName + " " + id).getId();
      const empSlide = SlidesApp.openById(empSlideId).getSlides()[0]; // Ouvre la première diapositive

      // Remplace les placeholders dans le Slide
      empSlide.replaceAllText('{participant_name}', fullName);
      empSlide.replaceAllText('{role}', title);
      empSlide.replaceAllText('{issued_date}', Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM dd, yyyy'));
      empSlide.replaceAllText('{certificate_id}', id);

      // Met à jour le Google Sheet avec l'ID du nouveau Slide et le statut
      sheet.getRange(i + 1, fileLocationIndex + 1).setValue(empSlideId);
      sheet.getRange(i + 1, statusIndex + 1).setValue('CREATED');
      SpreadsheetApp.flush(); // Force la mise à jour du Sheet
    }
  }
  SpreadsheetApp.getUi().alert('Certificats créés avec succès !');
}

/**
 * Envoie un e-mail personnalisé à chaque participant avec son certificat PDF en pièce jointe.
 * Les e-mails sont basés sur un brouillon Gmail et le statut est mis à jour à "SENT".
 */
function sendCertificates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  const fullNameIndex = headers.indexOf('Full Name');
  const emailIndex = headers.indexOf('Email');
  const fileLocationIndex = headers.indexOf('File Location');
  const statusIndex = headers.indexOf('Status');

  // ---- Étape 1 : récupérer le brouillon modèle Gmail ----
  const subjectTemplate = 'Certificate of Participation – DevFest Yaoundé 2025';
  const drafts = GmailApp.getDrafts();
  const draft = drafts.find(d => d.getMessage().getSubject() === subjectTemplate);

  if (!draft) {
    throw new Error("❌ Aucun brouillon trouvé avec le sujet : " + subjectTemplate + ". Veuillez en créer un.");
  }
  const draftBody = draft.getMessage().getBody(); // Corps HTML du brouillon
  const senderName = 'GDG Yaoundé Team';
  const ccRecipients = draft.getMessage().getCc(); // Récupérer les destinataires Cc s'il y en a

  // ---- Étape 2 : parcourir les lignes du tableur ----
  for (let i = 1; i < values.length; i++) {
    const rowData = values[i];
    const fullName = rowData[fullNameIndex];
    const email = rowData[emailIndex];
    const slideId = rowData[fileLocationIndex];
    const status = rowData[statusIndex];

    // Sauter les lignes déjà envoyées ou incomplètes (pas d'ID de slide ou pas d'email)
    if (status === 'SENT' || !slideId || !email) {
      Logger.log(`Skipping row ${i+1}: Status=${status}, SlideId=${slideId ? 'Exists' : 'Missing'}, Email=${email ? 'Exists' : 'Missing'}`);
      continue;
    }

    try {
      const attachment = DriveApp.getFileById(slideId); // Récupère le Google Slide généré

      // ---- Étape 3 : personnaliser le corps du mail ----
      // Remplace le placeholder {{fullName}} par le nom réel du participant
      const personalizedBody = draftBody.replace(/{{\s*fullName\s*}}/gi, fullName);

      // ---- Étape 4 : envoyer le mail ----
      GmailApp.sendEmail(email, subjectTemplate, '', {
        htmlBody: personalizedBody,
        attachments: [attachment.getAs(MimeType.PDF)], // Convertit le Slide en PDF et l'attache
        cc: ccRecipients,
        name: senderName, // Nom de l'expéditeur qui apparaîtra
      });

      // Pause de 2 secondes pour éviter le throttling Gmail
      Utilities.sleep(2000);

      // ---- Étape 5 : marquer comme "SENT" ----
      sheet.getRange(i + 1, statusIndex + 1).setValue('SENT');
      SpreadsheetApp.flush(); // Force la mise à jour du Sheet
      Logger.log(`✅ Mail sent to ${fullName} (${email}).`);

    } catch (error) {
      Logger.log(`❌ Erreur lors de l'envoi du mail pour ${fullName} (${email}): ${error.message}`);
      // Vous pourriez vouloir mettre à jour le statut avec 'FAILED' ou 'ERROR' ici
      sheet.getRange(i + 1, statusIndex + 1).setValue('ERROR: ' + error.message);
      SpreadsheetApp.flush();
    }
  }
  Logger.log("✅ Tous les mails ont été traités (envoyés ou ignorés).");
  SpreadsheetApp.getUi().alert('Envoi des certificats terminé ! Veuillez vérifier les logs pour d\'éventuelles erreurs.');
}
