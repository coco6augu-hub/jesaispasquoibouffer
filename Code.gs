/**
 * Google Apps Script Web App (API Google Sheets)
 *
 * Objectif:
 * - Lire une Google Sheet contenant des recettes
 * - Appliquer des filtres (temps/prix/sans gluten/sans lactose)
 * - Retourner UNE recette aléatoire en JSON
 *
 * Colonnes attendues (ligne d'en-tête):
 * nom, temps_total_min, prix_par_personne, sans_gluten, sans_lactose, sante, ingredients, instructions
 * (sante = "healthy" | "gras" | "normal" | "peu importe" ; "peu importe" dans la sheet = recette compte pour les 3 autres)
 *
 * Paramètres GET acceptés:
 * - temps_max (number)
 * - prix_min (number)
 * - prix_max (number)
 * - sans_gluten (true/false)
 * - sans_lactose (true/false)
 * - sante (string: "healthy" | "gras" | "normal" | "peu_importe" ; vide ou peu_importe = pas de filtre)
 * - count (number, optionnel, défaut: 1, max: 10)
 *
 * Réponses:
 * - { recipe: { ... } } (si count=1)
 * - { recipes: [ { ... }, ... ] } (si count>1)
 * - si aucune: { success:false, message:"..." }
 */

function doGet(e) {
  try {
    var cfg = getConfig_();
    var params = (e && e.parameter) ? e.parameter : {};
    var filters = parseFilters_(params);
    var count = parseCount_(params);

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

    if (count <= 1) {
      var pick = matches[Math.floor(Math.random() * matches.length)];
      return json_({ recipe: pick });
    }

    var picks = pickRandomDistinct_(matches, count);
    return json_({ recipes: picks });
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
  var prixMin = p.prix_min !== undefined && p.prix_min !== "" ? Number(p.prix_min) : null;
  var prixMax = p.prix_max !== undefined && p.prix_max !== "" ? Number(p.prix_max) : null;

  var santeRaw = String(p.sante || "").trim().toLowerCase();
  var sante = null;
  if (santeRaw === "healthy" || santeRaw === "gras" || santeRaw === "normal") sante = santeRaw;
  // "peu_importe" ou vide = pas de filtre

  return {
    tempsMax: isFinite(tempsMax) ? tempsMax : null,
    prixMin: isFinite(prixMin) ? prixMin : null,
    prixMax: isFinite(prixMax) ? prixMax : null,
    sansGluten: String(p.sans_gluten || "").toLowerCase() === "true",
    sansLactose: String(p.sans_lactose || "").toLowerCase() === "true",
    sante: sante
  };
}

function parseCount_(p) {
  var raw = p.count !== undefined && p.count !== "" ? Number(p.count) : 1;
  if (!isFinite(raw) || raw < 1) return 1;
  // sécurité: on limite pour éviter de renvoyer trop de données
  return Math.min(Math.floor(raw), 10);
}

function pickRandomDistinct_(arr, count) {
  // Fisher–Yates partiel (copie)
  var copy = arr.slice(0);
  for (var i = copy.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = copy[i];
    copy[i] = copy[j];
    copy[j] = tmp;
  }
  return copy.slice(0, Math.min(count, copy.length));
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
  function optIdx(name, alt) {
    var i = header.indexOf(name);
    if (i >= 0) return i;
    if (alt) return header.indexOf(alt);
    return -1;
  }

  return {
    nom: idx("nom"),
    temps_total_min: idx("temps_total_min"),
    prix_par_personne: idx("prix_par_personne"),
    sans_gluten: idx("sans_gluten"),
    sans_lactose: idx("sans_lactose"),
    sante: optIdx("sante", "santé"),
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

function normalizeSante_(v) {
  var s = String(v || "").trim().toLowerCase();
  if (s === "healthy" || s === "gras" || s === "normal" || s === "peu importe") return s;
  return "";
}

function rowToRecipe_(row, col) {
  var sante = col.sante >= 0 ? normalizeSante_(row[col.sante]) : "";
  return {
    nom: String(row[col.nom] || "").trim(),
    temps_total_min: toNumber_(row[col.temps_total_min]),
    prix_par_personne: toNumber_(row[col.prix_par_personne]),
    sans_gluten: toBool_(row[col.sans_gluten]),
    sans_lactose: toBool_(row[col.sans_lactose]),
    sante: sante,
    ingredients: String(row[col.ingredients] || "").trim(),
    instructions: String(row[col.instructions] || "").trim()
  };
}

function matchesFilters_(recipe, filters) {
  if (filters.tempsMax !== null) {
    if (recipe.temps_total_min === null) return false;
    if (recipe.temps_total_min > filters.tempsMax) return false;
  }
  if (recipe.prix_par_personne === null) {
    // Si le prix n'est pas renseigné, on exclut seulement si on a des filtres de prix
    if (filters.prixMin !== null || filters.prixMax !== null) return false;
  } else {
    if (filters.prixMin !== null) {
      if (recipe.prix_par_personne < filters.prixMin) return false;
    }
    if (filters.prixMax !== null) {
      if (recipe.prix_par_personne > filters.prixMax) return false;
    }
  }
  if (filters.sansGluten && recipe.sans_gluten !== true) return false;
  if (filters.sansLactose && recipe.sans_lactose !== true) return false;

  if (filters.sante) {
    if (!recipe.sante) return false;
    var rs = recipe.sante;
    if (rs !== filters.sante && rs !== "peu importe") return false;
  }
  return true;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


