/**
 * Google Apps Script Web App (API Google Sheets)
 *
 * Objectif:
 * - Lire une Google Sheet contenant des recettes
 * - Appliquer des filtres (temps/prix/sans gluten/sans lactose)
 * - Retourner UNE recette aléatoire en JSON
 *
 * Colonnes attendues (ligne d'en-tête):
 * nom, temps_total_min, prix_par_personne, sans_gluten, sans_lactose, ingredients, instructions
 *
 * Paramètres GET acceptés:
 * - temps_max (number)
 * - prix_max (number)
 * - sans_gluten (true/false)
 * - sans_lactose (true/false)
 *
 * Réponses:
 * - { recipe: { ... } }
 * - si aucune: { success:false, message:"..." }
 */

function doGet(e) {
  try {
    var cfg = getConfig_();
    var filters = parseFilters_(e && e.parameter ? e.parameter : {});

    var sheet = openSheet_(cfg);
    var rows = sheet.getDataRange().getValues();
    if (!rows || rows.length < 2) {
      return json_({ success: false, message: "La feuille est vide (aucune recette)." });
    }

    var header = normalizeHeader_(rows[0]);
    var col = indexColumns_(header);

    var matches = [];
    for (var r = 1; r < rows.length; r++) {
      var recipe = rowToRecipe_(rows[r], col);
      if (!recipe || !recipe.nom) continue;
      if (matchesFilters_(recipe, filters)) matches.push(recipe);
    }

    if (matches.length === 0) {
      return json_({ success: false, message: "Aucune recette trouvée avec ces critères." });
    }

    var pick = matches[Math.floor(Math.random() * matches.length)];
    return json_({ recipe: pick });
  } catch (err) {
    return json_({ success: false, message: String(err && err.message ? err.message : err) });
  }
}

function getConfig_() {
  // Configure ces 2 valeurs dans Apps Script → Project Settings → Script properties
  // - SPREADSHEET_ID : l'ID du Google Sheet
  // - SHEET_NAME : le nom de l'onglet (ex: "Recettes")
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty("SPREADSHEET_ID");
  var sheetName = props.getProperty("SHEET_NAME") || "Recettes";

  if (!spreadsheetId) throw new Error("Config manquante: SPREADSHEET_ID (Script Properties)");
  return { spreadsheetId: spreadsheetId, sheetName: sheetName };
}

function openSheet_(cfg) {
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) throw new Error("Onglet introuvable: " + cfg.sheetName);
  return sheet;
}

function parseFilters_(p) {
  var tempsMax = p.temps_max !== undefined && p.temps_max !== "" ? Number(p.temps_max) : null;
  var prixMax = p.prix_max !== undefined && p.prix_max !== "" ? Number(p.prix_max) : null;

  return {
    tempsMax: isFinite(tempsMax) ? tempsMax : null,
    prixMax: isFinite(prixMax) ? prixMax : null,
    sansGluten: String(p.sans_gluten || "").toLowerCase() === "true",
    sansLactose: String(p.sans_lactose || "").toLowerCase() === "true"
  };
}

function normalizeHeader_(row) {
  return row.map(function (h) {
    return String(h || "").trim().toLowerCase();
  });
}

function indexColumns_(header) {
  function idx(name) {
    var i = header.indexOf(name);
    if (i === -1) throw new Error("Colonne manquante: " + name);
    return i;
  }

  return {
    nom: idx("nom"),
    temps_total_min: idx("temps_total_min"),
    prix_par_personne: idx("prix_par_personne"),
    sans_gluten: idx("sans_gluten"),
    sans_lactose: idx("sans_lactose"),
    ingredients: idx("ingredients"),
    instructions: idx("instructions")
  };
}

function toNumber_(v) {
  if (v === null || v === undefined || v === "") return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}

function toBool_(v) {
  if (v === true) return true;
  if (v === false) return false;
  var s = String(v || "").trim().toLowerCase();
  return s === "true" || s === "vrai" || s === "1" || s === "oui" || s === "yes";
}

function rowToRecipe_(row, col) {
  return {
    nom: String(row[col.nom] || "").trim(),
    temps_total_min: toNumber_(row[col.temps_total_min]),
    prix_par_personne: toNumber_(row[col.prix_par_personne]),
    sans_gluten: toBool_(row[col.sans_gluten]),
    sans_lactose: toBool_(row[col.sans_lactose]),
    ingredients: String(row[col.ingredients] || "").trim(),
    instructions: String(row[col.instructions] || "").trim()
  };
}

function matchesFilters_(recipe, filters) {
  if (filters.tempsMax !== null) {
    if (recipe.temps_total_min === null) return false;
    if (recipe.temps_total_min > filters.tempsMax) return false;
  }
  if (filters.prixMax !== null) {
    if (recipe.prix_par_personne === null) return false;
    if (recipe.prix_par_personne > filters.prixMax) return false;
  }
  if (filters.sansGluten && recipe.sans_gluten !== true) return false;
  if (filters.sansLactose && recipe.sans_lactose !== true) return false;
  return true;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


