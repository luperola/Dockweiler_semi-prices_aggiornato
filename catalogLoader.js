import xlsx from "xlsx";
import path from "path";
import fs from "fs";

const dataDir = path.join(process.cwd(), "data");
const straightPath = path.join(dataDir, "Straight_Items.xlsx");
const complexPath = path.join(dataDir, "Tees.xlsx"); // contiene i fogli "Tees" e "Reducers"
const coaxialPath = path.join(dataDir, "Coaxial.xlsx"); // Coassiali

// ---------- utils ----------
function parseNum(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  let s = String(v).trim();
  if (!s) return 0;
  const hasDot = s.includes(".");
  const hasComma = s.includes(",");
  if (hasDot && hasComma) {
    if (s.lastIndexOf(",") > s.lastIndexOf("."))
      s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (hasComma) s = s.replace(",", ".");
  const n = Number(s);
  return isFinite(n) ? n : 0;
}
function sheetJSON(wb, name) {
  return wb.Sheets[name]
    ? xlsx.utils.sheet_to_json(wb.Sheets[name], { defval: null })
    : null;
}
function pick(row, keys) {
  // prova più chiavi alternative (prima che esista)
  for (const k of keys) {
    if (Object.prototype.hasOwnProperty.call(row, k)) return row[k];
  }
  return null;
}

// ---------- mapping brand ----------
const PRICE_KEYS_TUBES = [
  ["VSR", ["VSR Code"], ["VSR €/m"], null],
  ["TCC", ["TCC code"], ["TCC €/m"], null],
  ["TCC.1", ["TCC.1 Code"], ["TCC.1 €/m"], null],
  ["Finetron", ["Finetron Code"], ["Finetron €/m"], null],
  ["Ultron", ["Ultron Code"], ["Ultron €/m"], null],
];

const PRICE_KEYS_FITTINGS = [
  ["TCC", ["TCC code", "TCC Codes"], null, ["TCC €/pc", "TCC €/piece"]],
  [
    "TCC.1",
    ["TCC.1 Code", "TCC.1 Codes"],
    null,
    ["TCC.1 €/pc", "TCC.1 €/piece"],
  ],
  [
    "Finetron",
    ["Finetron Code", "Finetron Codes"],
    null,
    ["Finetron €/pc", "Finetron €/piece"],
  ],
  [
    "Ultron",
    ["Ultron Code", "Ultron Codes"],
    null,
    ["Ultron €/pc", "Ultron €/piece"],
  ],
];

function extractBrandsFlexible(row, keyMap, per) {
  const out = {};
  for (const [brand, codeKeys, perMKeys, perPcKeys] of keyMap) {
    const code = pick(row, codeKeys);
    if (per === "m" && perMKeys) {
      const price = parseNum(pick(row, perMKeys));
      if (price)
        out[brand] = {
          code: code?.toString().trim() || null,
          pricePerM: price,
        };
    } else if (per === "pc" && perPcKeys) {
      const price = parseNum(pick(row, perPcKeys));
      if (price)
        out[brand] = {
          code: code?.toString().trim() || null,
          pricePerPc: price,
        };
    }
  }
  return out;
}

// ---------- loaders: Straight (tubes + elbows + caps) ----------
export function loadTubesCatalog() {
  if (!fs.existsSync(straightPath)) return [];
  const wb = xlsx.readFile(straightPath);
  const sheetName = wb.Sheets["Tubes"] ? "Tubes" : wb.SheetNames[0];
  const rows = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });
  return rows
    .filter((r) => r.OD != null && r.TH != null)
    .map((r) => ({
      itemType: "Tubes",
      OD: String(r.OD).trim(),
      TH: String(r.TH).trim(),
      pesoKgM: parseNum(r["Peso Kg/m"]),
      brands: extractBrandsFlexible(r, PRICE_KEYS_TUBES, "m"),
    }));
}

export function loadFittingsCatalog() {
  if (!fs.existsSync(straightPath)) return [];
  const wb = xlsx.readFile(straightPath);
  const pack = [];

  const e90 = sheetJSON(wb, "Elbows 90") || [];
  for (const r of e90) {
    if (!r.OD) continue;
    pack.push({
      itemType: "Elbows",
      angle: "90",
      OD: String(r.OD).trim(),
      TH: r.TH != null ? String(r.TH).trim() : "",
      brands: extractBrandsFlexible(r, PRICE_KEYS_FITTINGS, "pc"),
      uom: "pcs",
    });
  }

  const e45 = sheetJSON(wb, "Elbows 45") || [];
  for (const r of e45) {
    if (!r.OD) continue;
    pack.push({
      itemType: "Elbows",
      angle: "45",
      OD: String(r.OD).trim(),
      TH: r.TH != null ? String(r.TH).trim() : "",
      brands: extractBrandsFlexible(r, PRICE_KEYS_FITTINGS, "pc"),
      uom: "pcs",
    });
  }

  const caps = sheetJSON(wb, "End Caps") || [];
  for (const r of caps) {
    if (!r.OD) continue;
    pack.push({
      itemType: "Caps",
      OD: String(r.OD).trim(),
      TH: r.TH != null ? String(r.TH).trim() : "",
      brands: extractBrandsFlexible(r, PRICE_KEYS_FITTINGS, "pc"),
      uom: "pcs",
    });
  }

  return pack;
}

// ---------- loaders: Complex (tees + reducers) ----------
export function loadComplexCatalog() {
  // file Tees.xlsx con fogli: "Tees", "Reducers" e colonne: OD1, TH1, OD2, TH2 + brand €/pc
  if (!fs.existsSync(complexPath)) return [];
  const wb = xlsx.readFile(complexPath);
  const pack = [];

  const tees = sheetJSON(wb, "Tees") || [];
  for (const r of tees) {
    if (r.OD1 == null || r.OD2 == null) continue;
    pack.push({
      itemType: "Tees",
      OD1: String(r.OD1).trim(),
      TH1: r.TH1 != null ? String(r.TH1).trim() : "",
      OD2: String(r.OD2).trim(),
      TH2: r.TH2 != null ? String(r.TH2).trim() : "",
      brands: extractBrandsFlexible(r, PRICE_KEYS_FITTINGS, "pc"),
      uom: "pcs",
    });
  }

  const reducers = sheetJSON(wb, "Reducers") || [];
  for (const r of reducers) {
    if (r.OD1 == null || r.OD2 == null) continue;
    // Assunzione: OD1 = diametro maggiore, OD2 = minore (come richiesto)
    pack.push({
      itemType: "Reducers",
      OD1: String(r.OD1).trim(),
      TH1: r.TH1 != null ? String(r.TH1).trim() : "",
      OD2: String(r.OD2).trim(),
      TH2: r.TH2 != null ? String(r.TH2).trim() : "",
      brands: extractBrandsFlexible(r, PRICE_KEYS_FITTINGS, "pc"),
      uom: "pcs",
    });
  }

  return pack;
}

// ---------- loader entry ----------

// ---------- loaders: Coassiali (tubes/tees/elbows/sleeve/terminator/tee purge) ----------
export function loadCoaxialCatalog() {
  if (!fs.existsSync(coaxialPath)) return [];
  const wb = xlsx.readFile(coaxialPath);
  const allowedSheets = [
    "Tubes",
    "Tees",
    "Elbows 45",
    "Elbows 90",
    "Sleeve",
    "Terminator",
    "Tee Purge",
  ];
  const out = [];

  for (const sheetName of allowedSheets) {
    const rows = sheetJSON(wb, sheetName) || [];
    for (const r of rows) {
      const OD1 = r.OD1 != null ? String(r.OD1).trim() : null;
      const OD2 = r.OD2 != null ? String(r.OD2).trim() : null;
      if (!OD1 || !OD2) continue;
      const TH1 = r.TH1 != null ? String(r.TH1).trim() : "";
      const TH2 = r.TH2 != null ? String(r.TH2).trim() : "";

      const tccCode = pick(r, ["TCC Codes", "TCC Code"]);
      const ulCode = pick(r, ["Ultron Code", "Ultron Codes"]);
      const tccPrice = parseNum(
        pick(r, [sheetName === "Tubes" ? "TCC €/m" : "TCC €/pc"])
      );
      const ulPrice = parseNum(
        pick(r, [sheetName === "Tubes" ? "Ultron €/m" : "Ultron €/pc"])
      );
      const peso = sheetName === "Tubes" ? parseNum(pick(r, ["Peso Kg/m"])) : 0;

      const brands = {};
      /* if (sheetName === "Tubes" || sheetName.startsWith("Elbows")) {
        if (tccCode != null || tccPrice) {
          brands["TCC"] = { code: tccCode ? String(tccCode).trim() : null };
          if (sheetName === "Tubes") brands["TCC"].pricePerM = tccPrice;
          else brands["TCC"].pricePerPc = tccPrice;
        }
        if (ulCode != null || ulPrice) {
          brands["Ultron"] = { code: ulCode ? String(ulCode).trim() : null };
          if (sheetName === "Tubes") brands["Ultron"].pricePerM = ulPrice;
          else brands["Ultron"].pricePerPc = ulPrice;
        }
      } else {
        if (tccCode != null || tccPrice) {
          brands["TCC"] = {
            code: tccCode ? String(tccCode).trim() : null,
            pricePerPc: tccPrice,
          };
        }
      } */
      const isPerMeter = sheetName === "Tubes";
      const priceKey = isPerMeter ? "pricePerM" : "pricePerPc";

      const addBrand = (brand, codeRaw, price) => {
        const code = codeRaw != null ? String(codeRaw).trim() : "";
        const hasCode = Boolean(code);
        const hasPriceValue = Number.isFinite(price) && price > 0;
        const isZeroPrice = Number.isFinite(price) && price === 0;
        if (!hasCode && !hasPriceValue) return;

        const entry = {};
        if (hasCode) entry.code = code;
        if (hasPriceValue || (isZeroPrice && hasCode)) entry[priceKey] = price;
        if (Object.keys(entry).length) brands[brand] = entry;
      };

      addBrand("TCC", tccCode, tccPrice);
      addBrand("Ultron", ulCode, ulPrice);

      if (!Object.keys(brands).length) continue;

      const item = {
        itemType: "Coassiali",
        subtype: sheetName,
        OD1,
        TH1,
        OD2,
        TH2,
        brands,
      };
      if (sheetName === "Tubes") item.pesoKgM = peso;
      out.push(item);
    }
  }
  return out;
}

export function loadCatalog() {
  const tubes = loadTubesCatalog();
  const fittings = loadFittingsCatalog(); // Elbows / Caps
  const complex = loadComplexCatalog(); // Tees / Reducers (a doppio diametro)
  const coaxial = loadCoaxialCatalog();
  return { tubes, fittings, complex, coaxial };
}
