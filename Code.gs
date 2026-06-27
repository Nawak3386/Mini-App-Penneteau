function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Feuille heures Penneteau');
}

function convertirHeureEnMinutes(valeur) {
  if (!valeur) return 0;

  let texte = String(valeur)
    .trim()
    .replace("➖", "-")
    .replace("➕", "")
    .replace("à devoir", "")
    .replace("à déduire", "")
    .replace("sup", "")
    .replace(/\s/g, "");

  const negatif = texte.startsWith("-");
  texte = texte.replace("-", "");

  const parts = texte.split("h");
  const h = Number(parts[0]) || 0;
  const m = Number(parts[1]) || 0;

  const total = h * 60 + m;
  return negatif ? -total : total;
}

function formatHeures(minutes) {
  minutes = Math.max(0, Number(minutes) || 0);
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return h + "h" + String(m).padStart(2, "0");
}

function getCompteurs(salarie) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Soldes");

  let aDevoir = 0;
  let heuresSup = 0;

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === salarie) {
      aDevoir = Math.abs(convertirHeureEnMinutes(values[i][1]));
      heuresSup = Math.abs(convertirHeureEnMinutes(values[i][2]));
      break;
    }
  }

  return { aDevoir, heuresSup };
}

function majSoldes(salarie, aDevoir, heuresSup) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Soldes");
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === salarie) {
      sheet.getRange(i + 1, 2).setValue(aDevoir === 0 ? "0h00" : "-" + formatHeures(aDevoir));
      sheet.getRange(i + 1, 3).setValue(formatHeures(heuresSup));
      return;
    }
  }
}

function appliquerEcart(compteurs, ecart) {
  if (ecart < 0) {
    compteurs.aDevoir += Math.abs(ecart);
    return;
  }

  if (ecart > 0) {
    if (compteurs.aDevoir > 0) {
      if (ecart <= compteurs.aDevoir) {
        compteurs.aDevoir -= ecart;
      } else {
        const reste = ecart - compteurs.aDevoir;
        compteurs.aDevoir = 0;
        compteurs.heuresSup += reste;
      }
    } else {
      compteurs.heuresSup += ecart;
    }
  }
}

function ajouterSeparateurSemaine(sheet) {
  sheet.appendRow([
    "--------------------",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ]);
}

function enregistrerSemaine(data) {
  clotureMensuelleHeuresSup();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Saisies V2");

  let compteurs = getCompteurs(data.salarie);

  data.lignes.forEach(function(ligne) {
    const ecart = convertirHeureEnMinutes(ligne.ecart || "0h00");

    appliquerEcart(compteurs, ecart);

    sheet.appendRow([
      data.salarie,
      ligne.date,
      ligne.jour,
      ligne.chantier,
      ligne.debut + " → " + ligne.fin,
      ligne.pause,
      ligne.totalJour,
      ligne.ecart,
      formatHeures(compteurs.aDevoir),
      formatHeures(compteurs.heuresSup),
      ligne.distanceKm,
      ligne.zone,
      ligne.note
    ]);
  });

  majSoldes(data.salarie, compteurs.aDevoir, compteurs.heuresSup);
  ajouterSeparateurSemaine(sheet);

  return "À devoir : " + formatHeures(compteurs.aDevoir) +
    " | Heures sup : " + formatHeures(compteurs.heuresSup);
}

function clotureMensuelleHeuresSup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const soldes = ss.getSheetByName("Soldes");
  const saisies = ss.getSheetByName("Saisies V2");

  if (!soldes || !saisies) return;

  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const isLastDayOfMonth = tomorrow.getMonth() !== today.getMonth();
  if (!isLastDayOfMonth) return;

  const timezone = Session.getScriptTimeZone();
  const mois = Utilities.formatDate(today, timezone, "MM/yyyy");
  const dateLigne = Utilities.formatDate(today, timezone, "yyyy-MM-dd");

  const props = PropertiesService.getScriptProperties();
  const cle = "cloture_heures_sup_" + mois;

  if (props.getProperty(cle)) return;

  const values = soldes.getDataRange().getValues();

  saisies.appendRow([
    "====================",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ]);

  for (let i = 1; i < values.length; i++) {
    const salarie = values[i][0];
    const solde = values[i][1];
    const heuresSup = values[i][2];

    if (!salarie) continue;

    saisies.appendRow([
      salarie,
      dateLigne,
      "TOTAL MENSUEL",
      "Clôture heures sup " + mois,
      "",
      "",
      "",
      "",
      solde || "0h00",
      heuresSup || "0h00",
      "",
      "",
      "Clôture automatique mensuelle"
    ]);

    soldes.getRange(i + 1, 3).setValue("0h00");
  }

  saisies.appendRow([
    "====================",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ]);

  props.setProperty(cle, "fait");
}