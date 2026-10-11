/*
 * Alerte au potager — réception des gagnants
 * Déploiement : Google Apps Script > Déployer > Nouvelle deployment > Application Web.
 * Exécuter en tant que : Moi ; Accès : Tout le monde.
 * Configurez les deux constantes ci-dessous avant le déploiement.
 */
const NOTIFICATION_EMAIL = "VOTRE_ADRESSE_EMAIL";
const SHEET_NAME = "Gagnants";

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const name = clean(data.name, 100);
    const email = clean(data.email, 254);
    const phone = clean(data.phone, 40);
    const consent = data.consent === true || data.consent === "true";
    if (!name || (!email && !phone) || !consent) {
      return jsonResponse({ ok: false, error: "Informations manquantes" });
    }

    const props = PropertiesService.getScriptProperties();
    let spreadsheetId = props.getProperty("SPREADSHEET_ID");
    let spreadsheet;
    if (spreadsheetId) {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } else {
      spreadsheet = SpreadsheetApp.create("Alerte au potager - Gagnants");
      props.setProperty("SPREADSHEET_ID", spreadsheet.getId());
      const sheet = spreadsheet.getSheets()[0];
      sheet.setName(SHEET_NAME);
      sheet.appendRow(["Date de réception", "Nom", "E-mail", "Téléphone", "Date du jeu"]);
      sheet.setFrozenRows(1);
    }
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      sheet.appendRow(["Date de réception", "Nom", "E-mail", "Téléphone", "Date du jeu"]);
      sheet.setFrozenRows(1);
    }
    // Préfixer les champs par une apostrophe pour éviter l'exécution de formules de tableur.
    sheet.appendRow([new Date(), safeCell(name), safeCell(email), safeCell(phone), clean(data.date, 10)]);

    if (NOTIFICATION_EMAIL && NOTIFICATION_EMAIL !== "VOTRE_ADRESSE_EMAIL") {
      MailApp.sendEmail({
        to: NOTIFICATION_EMAIL,
        subject: "Nouveau gagnant — Alerte au potager",
        body: "Une nouvelle participation gagnante a été enregistrée.\n\nNom : " + name +
          "\nE-mail : " + (email || "non renseigné") +
          "\nTéléphone : " + (phone || "non renseigné") +
          "\nDate du jeu : " + clean(data.date, 10) +
          "\n\nTableau des gagnants : " + spreadsheet.getUrl()
      });
    }
    return jsonResponse({ ok: true });
  } catch (err) {
    console.error(err);
    return jsonResponse({ ok: false, error: "Erreur interne" });
  }
}

function clean(value, maxLength) {
  return String(value == null ? "" : value).trim().replace(/[\u0000-\u001F\u007F]/g, "").slice(0, maxLength);
}
function safeCell(value) {
  const text = clean(value, 500);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
