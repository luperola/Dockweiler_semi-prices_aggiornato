import express from "express";
import cors from "cors";
import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";
import { loadCatalog } from "./catalogLoader.js";
import multer from "multer";
import open from "open";
import fs from "fs";
import PizZip from "pizzip";
import XLSX from "xlsx";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);


// Arrotondamento decimale HALF-UP (come Excel visualizza: 38,485 -> 38,49)
function toDecimalString(value) {
  if (value === null || value === undefined || value === "") return "0";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "0";
    // Normalizza gli errori binari JS: 38.484999999999 -> 38.485
    return value.toFixed(12).replace(/0+$/, "").replace(/\.$/, "");
  }
  return String(value).trim().replace(/€/g, "").replace(/\s/g, "").replace(/,/g, ".") || "0";
}

function roundHalfUp(value, decimals = 2) {
  let s = toDecimalString(value);
  let sign = 1;
  if (s.startsWith("-")) { sign = -1; s = s.slice(1); }
  if (s.includes("e") || s.includes("E")) s = Number(s).toFixed(decimals + 8);
  let [intPart, fracPart = ""] = s.split(".");
  intPart = intPart.replace(/\D/g, "") || "0";
  fracPart = fracPart.replace(/\D/g, "");
  fracPart = fracPart.padEnd(decimals + 1, "0");
  const keep = fracPart.slice(0, decimals);
  const next = Number(fracPart[decimals] || "0");
  let scaled = BigInt(intPart + keep);
  if (next >= 5) scaled += 1n;
  return sign * Number(scaled) / 10 ** decimals;
}

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

let catalog = loadCatalog();
const upload = multer({ storage: multer.memoryStorage() });

// Endpoint base esistenti
app.get("/api/catalog/tubes", (req, res) => {
  res.json({ items: catalog.tubes || [] });
});

app.get("/api/catalog/fittings", (req, res) => {
  const { type, angle, OD, TH } = req.query;
  let items = catalog.fittings || [];
  if (type) items = items.filter((r) => r.itemType === type);
  if (angle) items = items.filter((r) => String(r.angle) === String(angle));
  if (OD) items = items.filter((r) => String(r.OD) === String(OD));
  if (TH) items = items.filter((r) => String(r.TH || "") === String(TH || ""));
  res.json({ items });
});

// NUOVO: Tees & Reducers (doppio diametro)
app.get("/api/catalog/complex", (req, res) => {
  const { type, OD1, TH1, OD2, TH2 } = req.query; // tutti opzionali
  let items = (catalog.complex || []).slice();

  if (type) items = items.filter((r) => r.itemType === type);

  // se passati, filtra anche per diametri/spessori
  if (OD1) items = items.filter((r) => String(r.OD1) === String(OD1));
  if (TH1)
    items = items.filter((r) => String(r.TH1 || "") === String(TH1 || ""));
  if (OD2) items = items.filter((r) => String(r.OD2) === String(OD2));
  if (TH2)
    items = items.filter((r) => String(r.TH2 || "") === String(TH2 || ""));

  res.json({ items });
});

// Export Excel

// Coassiali
app.get("/api/catalog/coaxial", (req, res) => {
  const { subtype, OD1, TH1, OD2, TH2 } = req.query;
  let items = (catalog.coaxial || []).slice();
  if (subtype) items = items.filter((r) => r.subtype === subtype);
  if (OD1) items = items.filter((r) => String(r.OD1) === String(OD1));
  if (TH1)
    items = items.filter((r) => String(r.TH1 || "") === String(TH1 || ""));
  if (OD2) items = items.filter((r) => String(r.OD2) === String(OD2));
  if (TH2)
    items = items.filter((r) => String(r.TH2 || "") === String(TH2 || ""));
  res.json({ items });
});

app.post("/api/export", async (req, res) => {
  try {
    const {
      rows = [],
      currency = "EUR",
      discountPercent = 0,
      transport,
      meta = {},
    } = req.body || {};
    const dp = Math.min(100, Math.max(0, Number(discountPercent) || 0)); // 0..100
    const tp = Math.max(0, Number(transport?.percent ?? 0)); // % trasporto

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Offerta");

    // Intestazioni base (A..K) - le mettiamo alla RIGA 2 (riga 1 riservata a meta/Sconto)
    ws.columns = [
      { header: "POS", key: "pos", width: 8 }, // A
      { header: "Descrizione e misura", key: "descr", width: 45 }, // B
      { header: "Codice", key: "code", width: 18 }, // C
      { header: "U.M.", key: "um", width: 8 }, // D
      { header: "Quantità", key: "qty", width: 10 }, // E
      { header: "Prezzo unitario base €/m o €/pz", key: "base", width: 24 }, // F
      { header: "Peso tubo Kg/m", key: "peso", width: 12 }, // G
      { header: "Alloy surcharge in €/kg", key: "asKg", width: 20 }, // H
      { header: "Alloy surcharge in €/mt", key: "asM", width: 20 }, // I
      { header: "Prezzo unitario", key: "pu", width: 16 }, // J
      { header: "Valore riga", key: "tot", width: 16 }, // K
    ];

    // Larghezze specifiche
    ws.getColumn(1).width = 5.27; // A ≈ 65px
    ws.getColumn(2).width = 29.18; // B ≈ 328px
    ws.getColumn(3).width = 10.27; // C ≈ 120px
    ws.getColumn(4).width = 5.91; // D ≈ 72px
    for (let c = 6; c <= 10; c++) ws.getColumn(c).width = 9; // F..J ≈ 106px

    // Inseriamo una riga sopra le intestazioni (diventerà la riga 1)
    ws.spliceRows(1, 0, []);

    // --- RIGA 1: metadati opzionali in celle contigue da B1 + sconto in O1 ---
    const richiesta = (meta.richiesta || "").trim();
    const cliente = (meta.cliente || "").trim();
    const cantiere = (meta.cantiere || "").trim();

    const pieces = [];
    if (richiesta) pieces.push(`Richiesta: ${richiesta}`);
    if (cliente) pieces.push(`Cliente: ${cliente}`);
    if (cantiere) pieces.push(`Cantiere: ${cantiere}`);

    let metaCol = 2; // B
    for (const txt of pieces) {
      ws.getCell(1, metaCol).value = txt;
      metaCol++;
    }

    // SCONTO sempre in O1
    ws.getCell(1, 15).value = `SCONTO: ${dp.toFixed(2)}%`; // O1
    ws.getCell(1, 15).font = { bold: true, color: { argb: "FFFF0000" } };

    // Header (ora alla riga 2): centrato + wrap + altezza 11.09 (~96.75pt)
    const headerRow = ws.getRow(2);
    headerRow.alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
    headerRow.height = 96.75;

    // === Dati: da riga 3 in poi
    rows.forEach((item, idx) => {
      const isTube =
        item.itemType === "Tubes" || item.itemType === "Coassiali Tubes";

      const pos = (idx + 1) * 100;
      const descr = item.description || "";
      const code = item.code || "";
      const um = isTube ? "mt" : "pz";
      const qty = Number(item.quantity || 0);

      const base = Number(
        (isTube ? item.basePricePerM : item.basePricePerPc) || 0,
      );
      const peso = isTube ? Number(item.pesoKgM || 0) : null;
      const asKg = isTube ? Number(item.alloySurchargePerKg || 0) : null;
      const asM = isTube ? Number((asKg || 0) * (peso || 0)) : null;
      const pu = isTube ? base + (asM || 0) : base;
      const tot = pu * qty;

      ws.addRow({
        pos,
        descr,
        code,
        um,
        qty,
        base,
        peso: isTube ? peso : "-",
        asKg: isTube ? asKg : "-",
        asM: isTube ? asM : "-",
        pu,
        tot,
      });
    });

    // Allineamenti/arrotondamenti iniziali (F..K) + E e D — RIGHE 3..rowCount
    for (let i = 3; i <= ws.rowCount; i++) {
      ws.getCell(i, 5).alignment = { horizontal: "right" }; // E Quantità
      ws.getCell(i, 4).alignment = { horizontal: "center" }; // D U.M.
      for (let j = 6; j <= 11; j++) {
        // F..K
        const cell = ws.getCell(i, j);
        if (typeof cell.value === "number") {
          cell.value = roundHalfUp(cell.value, 2);
          // € su F, J, K
          cell.numFmt =
            j === 6 || j === 10 || j === 11 ? "€ #,##0.00" : "#,##0.00";
          cell.alignment = { horizontal: "right" };
        } else if (cell.value === "-") {
          cell.alignment = { horizontal: "center" };
        }
      }
    }

    // === Nuove colonne M/N ===
    // Header M/N (riga 2)
    ws.getCell(2, 13).value = "Prezzo Unitario Base DA LISTINO (€/m o €/pz)"; // M
    ws.getCell(2, 14).value = "TOTALE LORDO DA LISTINO"; // N
    // Larghezze: M come F, N come K
    ws.getColumn(13).width = ws.getColumn(6).width; // M = F
    ws.getColumn(14).width = ws.getColumn(11).width; // N = K

    // Copia F -> M (rosso) e K -> N (rosso) per righe dati (3..rowCount)
    for (let i = 3; i <= ws.rowCount; i++) {
      const fVal = Number(ws.getCell(i, 6).value || 0); // F
      const kVal = Number(ws.getCell(i, 11).value || 0); // K

      const m = ws.getCell(i, 13);
      m.value = roundHalfUp(fVal, 2);
      m.numFmt = "#,##0.00";
      m.alignment = { horizontal: "right" };
      m.font = { color: { argb: "FFFF0000" } };

      const n = ws.getCell(i, 14);
      n.value = roundHalfUp(kVal, 2);
      n.numFmt = "#,##0.00";
      n.alignment = { horizontal: "right" };
      n.font = { color: { argb: "FFFF0000" } };
    }

    // === Applica sconto: M -> F scontato, poi ricalcoli J e K ===
    for (let i = 3; i <= ws.rowCount; i++) {
      const um = String(ws.getCell(i, 4).value || "").toLowerCase(); // "mt" => tubo
      const qty = Number(ws.getCell(i, 5).value || 0);

      // Nuova F = M * (1 - dp)
      const mVal = Number(ws.getCell(i, 13).value || 0);
      const fNew = roundHalfUp(mVal * (1 - dp / 100), 2);
      const cellF = ws.getCell(i, 6);
      cellF.value = fNew;
      cellF.numFmt = "€ #,##0.00";
      cellF.alignment = { horizontal: "right" };

      let jNew = fNew; // non tubi
      if (um === "mt") {
        // tubo: J = F (scontato) + I (AS €/mt)
        const iVal = Number(ws.getCell(i, 9).value || 0); // I
        jNew = fNew + iVal;
      }
      jNew = roundHalfUp(jNew, 2);
      const cellJ = ws.getCell(i, 10);
      cellJ.value = jNew;
      cellJ.numFmt = "€ #,##0.00";
      cellJ.alignment = { horizontal: "right" };

      const kNew = roundHalfUp(jNew * qty, 2);
      const cellK = ws.getCell(i, 11);
      cellK.value = kNew;
      cellK.numFmt = "€ #,##0.00";
      cellK.alignment = { horizontal: "right" };
    }

    // === Totali finali (senza ulteriori sconti) ===
    const dataStartRow = 3;
    const dataLastRow = ws.rowCount;

    // Altezza POS (3..dataLastRow): 31.5pt (~42px)
    for (let r = dataStartRow; r <= dataLastRow; r++)
      ws.getRow(r).height = 31.5;

    // helper bordo sottile nero
    const thinSide = { style: "thin", color: { argb: "FF000000" } };
    const setThinBorder = (cell) => {
      cell.border = {
        top: thinSide,
        left: thinSide,
        bottom: thinSide,
        right: thinSide,
      };
    };

    // === BORDI sulle INTESTAZIONI (riga 2): A..K + M..N ===
    for (let c = 1; c <= 11; c++) setThinBorder(ws.getCell(2, c)); // A..K
    setThinBorder(ws.getCell(2, 13)); // M
    setThinBorder(ws.getCell(2, 14)); // N

    // === BORDI: tutte le celle delle righe POS (A..K) ===
    for (let r = dataStartRow; r <= dataLastRow; r++) {
      for (let c = 1; c <= 11; c++) {
        setThinBorder(ws.getCell(r, c));
      }
    }

    // 1) Totale Items ex works (I..J uniti)
    const r1 = dataLastRow + 1;
    ws.mergeCells(r1, 9, r1, 10); // I..J
    const cellIJ1 = ws.getCell(r1, 9);
    cellIJ1.value = "Total Items ex works";
    cellIJ1.font = { bold: true };
    cellIJ1.alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
    ws.getRow(r1).height = 45; // ~60px

    // K r1: SOMMA K3..K{dataLastRow} (con €)
    const cellK1 = ws.getCell(r1, 11);
    cellK1.value = { formula: `SUM(K${dataStartRow}:K${dataLastRow})` };
    cellK1.numFmt = "€ #,##0.00";
    cellK1.font = { bold: true };
    cellK1.alignment = { horizontal: "right", vertical: "middle" };

    // N r1: SOMMA N3..N{dataLastRow} (rosso)
    const cellN1 = ws.getCell(r1, 14);
    cellN1.value = { formula: `SUM(N${dataStartRow}:N${dataLastRow})` };
    cellN1.numFmt = "#,##0.00";
    cellN1.alignment = { horizontal: "right", vertical: "middle" };
    cellN1.font = { color: { argb: "FFFF0000" } };

    // 2) Imballo e trasporto (I..J uniti) — K = K(r1) * trasporto%
    const r2 = r1 + 1;
    ws.mergeCells(r2, 9, r2, 10);
    const cellIJ2 = ws.getCell(r2, 9);
    cellIJ2.value = "Imballo e trasporto";
    cellIJ2.font = { bold: true };
    cellIJ2.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(r2).height = 45;

    const cellK2 = ws.getCell(r2, 11);
    const tpFrac = (tp / 100).toString().replace(",", "."); // es. 0.095
    cellK2.value = { formula: `${tpFrac}*K${r1}` };
    cellK2.numFmt = "€ #,##0.00";
    cellK2.alignment = { horizontal: "right", vertical: "middle" };
    cellK2.font = { bold: true }; // valore in grassetto

    // 3) Totale f.co destino (I..J uniti) — K = K(r1) + K(r2)
    const r3 = r2 + 1;
    ws.mergeCells(r3, 9, r3, 10);
    const cellIJ3 = ws.getCell(r3, 9);
    cellIJ3.value = "Totale f.co destino";
    cellIJ3.font = { bold: true };
    cellIJ3.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(r3).height = 45;

    const cellK3 = ws.getCell(r3, 11);
    cellK3.value = { formula: `K${r1}+K${r2}` };
    cellK3.numFmt = "€ #,##0.00";
    cellK3.alignment = { horizontal: "right", vertical: "middle" };
    cellK3.font = { bold: true }; // valore in grassetto

    // === BORDI: celle dei totali (I..J merge + K) per r1, r2, r3 ===
    for (const rr of [r1, r2, r3]) {
      for (let c = 9; c <= 10; c++) setThinBorder(ws.getCell(rr, c)); // I..J
      setThinBorder(ws.getCell(rr, 11)); // K
    }

    // Output
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader("Content-Disposition", 'attachment; filename="Offerta.xlsx"');
    return res.send(Buffer.from(buf));
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: "Errore in esportazione" });
  }
});

// Utility per ricaricare il catalogo (se cambi gli xlsx)
app.post("/api/reload", (req, res) => {
  catalog = loadCatalog();
  res.json({
    ok: true,
    counts: {
      tubes: catalog.tubes?.length || 0,
      fittings: catalog.fittings?.length || 0,
      complex: catalog.complex?.length || 0,
      coaxial: catalog.coaxial?.length || 0,
    },
  });
});

// ===== Clienti offerte: lettura/scrittura file data/Clienti_Offerte.xlsx =====
function getClientsFilePath() {
  const dataPath = path.join(__dirname, "data", "Clienti_Offerte.xlsx");
  if (fs.existsSync(dataPath)) return dataPath;
  return path.join(__dirname, "Clienti_Offerte.xlsx");
}

function normalizeClientId(value) {
  return (
    String(value || "")
      .trim()
      .toUpperCase()
      .replace(/[^A-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "") || `CLIENTE_${Date.now()}`
  );
}

function readClientsFromExcel() {
  const filePath = getClientsFilePath();
  if (!fs.existsSync(filePath)) return [];

  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return [];

  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  return rows
    .map((r) => ({
      id: String(r.ID || r.Id || r.id || "").trim(),
      name: String(
        r.Cliente || r["Nome Azienda"] || r["Ragione Sociale"] || "",
      ).trim(),
      address1: String(r["Indirizzo 1"] || r.Via || r.Indirizzo || "").trim(),
      address2: String(
        r["Indirizzo 2"] || r.Località || r.Localita || "",
      ).trim(),
      attention: String(r["C.A."] || r["CA"] || r["C.A"] || "").trim(),
    }))
    .filter((c) => c.id || c.name);
}

function sortClientsAZ(clients = []) {
  return [...clients].sort((a, b) =>
    String(a.name || "").localeCompare(String(b.name || ""), "it", {
      sensitivity: "base",
      ignorePunctuation: true,
      numeric: true,
    }),
  );
}

function writeClientsToExcel(clients) {
  let filePath = getClientsFilePath();
  const dataDir = path.join(__dirname, "data");
  if (!fs.existsSync(path.dirname(filePath)) && fs.existsSync(dataDir)) {
    filePath = path.join(dataDir, "Clienti_Offerte.xlsx");
  }
  fs.mkdirSync(path.dirname(filePath), { recursive: true });

  const sortedClients = sortClientsAZ(clients);

  const rows = sortedClients.map((c) => ({
    ID: c.id,
    Cliente: c.name,
    "Indirizzo 1": c.address1,
    "Indirizzo 2": c.address2,
    "C.A.": c.attention,
  }));

  const ws = XLSX.utils.json_to_sheet(rows, {
    header: ["ID", "Cliente", "Indirizzo 1", "Indirizzo 2", "C.A."],
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Clienti");
  XLSX.writeFile(wb, filePath);

  return sortedClients;
}

app.get("/api/clients", (req, res) => {
  try {
    res.json({ clients: readClientsFromExcel() });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Errore lettura clienti" });
  }
});

app.post("/api/clients", (req, res) => {
  try {
    const body = req.body || {};
    const client = {
      id: normalizeClientId(body.id || body.name),
      name: String(body.name || "").trim(),
      address1: String(body.address1 || "").trim(),
      address2: String(body.address2 || "").trim(),
      attention: String(body.attention || "").trim(),
    };

    if (!client.name)
      return res.status(400).json({ error: "Nome Azienda obbligatorio" });
    if (!client.address1)
      return res.status(400).json({ error: "Via obbligatoria" });
    if (!client.address2)
      return res.status(400).json({ error: "Località obbligatoria" });

    const clients = readClientsFromExcel();
    const idx = clients.findIndex((c) => c.id === client.id);
    if (idx >= 0) clients[idx] = client;
    else clients.push(client);

    const sortedClients = writeClientsToExcel(clients);

    res.json({ ok: true, client, clients: sortedClients });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Errore salvataggio cliente" });
  }
});

function formatItalianDate(date = new Date()) {
  const months = [
    "gennaio",
    "febbraio",
    "marzo",
    "aprile",
    "maggio",
    "giugno",
    "luglio",
    "agosto",
    "settembre",
    "ottobre",
    "novembre",
    "dicembre",
  ];
  const dd = String(date.getDate()).padStart(2, "0");
  return `${dd} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

function wp(text, opts = {}) {
  const align = opts.align
    ? `<w:pPr><w:jc w:val="${opts.align}"/></w:pPr>`
    : "";
  return `<w:p>${align}<w:r><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
}

function wpOfferLine(dateText) {
  return `
  <w:p>
    <w:pPr>
      <w:tabs><w:tab w:val="right" w:pos="14000"/></w:tabs>
    </w:pPr>
    <w:r><w:t xml:space="preserve">Oggetto: Ns. Offerta Dockweiler ----------------</w:t></w:r>
    <w:r><w:tab/></w:r>
    <w:r><w:t xml:space="preserve">Milano, ${xmlEscape(dateText)}</w:t></w:r>
  </w:p>`;
}

function buildClientHeaderXml(client = {}) {
  const name = String(client.name || "Nome Azienda").trim();
  const address1 = String(client.address1 || "Indirizzo").trim();
  const address2 = String(client.address2 || "Indirizzo").trim();
  const attention = String(client.attention || "-----------------").trim();
  const dateText = formatItalianDate();

  return [
    wp(""),
    wp("Spett.le"),
    wp(name),
    wp(address1),
    wp(address2),
    wp(""),
    wp(`C.A. ${attention}`),
    wpOfferLine(dateText),
  ].join("\n");
}

function paragraphPlainText(pXml) {
  return String(pXml || "")
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function replaceClientHeaderAndDate(documentXml, client) {
  if (!client || (!client.name && !client.address1 && !client.address2))
    return documentXml;
  const headerXml = buildClientHeaderXml(client);
  const matches = [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];

  const start = matches.find((m) =>
    paragraphPlainText(m[0]).includes("Spett.le"),
  );
  const end = matches.find((m) => {
    const txt = paragraphPlainText(m[0]);
    return txt.includes("Oggetto") && txt.includes("Offerta Dockweiler");
  });

  if (start && end && start.index <= end.index) {
    return (
      documentXml.slice(0, start.index) +
      headerXml +
      documentXml.slice(end.index + end[0].length)
    );
  }

  // Fallback: se il template cambia, sostituisce almeno la riga Oggetto con data aggiornata.
  for (const m of matches) {
    const txt = paragraphPlainText(m[0]);
    if (txt.includes("Oggetto") && txt.includes("Offerta Dockweiler")) {
      return (
        documentXml.slice(0, m.index) +
        wpOfferLine(formatItalianDate()) +
        documentXml.slice(m.index + m[0].length)
      );
    }
  }

  return documentXml;
}

// ===== Export Word: genera Offerta Word dal template e inserisce tabella A:K =====
function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function round2(n) {
  return roundHalfUp(n, 2);
}

function fmtWordNumber(n, euro = false) {
  const value = round2(n);

  const sign = value < 0 ? "-" : "";
  const abs = Math.abs(value);

  const fixed = abs.toFixed(2);
  const [intPart, decPart] = fixed.split(".");

  const intWithDots = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ".");

  const s = `${sign}${intWithDots},${decPart}`;

  return euro ? `€ ${s}` : s;
}

function buildOfferRowsForWord(
  rows = [],
  discountPercent = 0,
  transport = {},
  includeLeadTime = false,
) {
  const dp = Math.min(100, Math.max(0, Number(discountPercent) || 0));
  const tp = Math.max(0, Number(transport?.percent ?? 0));

  const header = [
    "POS",
    "Descrizione e misura",
    "Codice",
    "U.M.",
    "Q.tà",
    "Prezzo unitario base €/m o €/pz",
    "Peso tubo Kg/m",
    "Alloy surcharge in €/kg",
    "Alloy surcharge in €/mt",
    "Prezzo unitario",
    "Valore riga",
  ];
  if (includeLeadTime) header.push("Lead time");

  const tableRows = [header];
  let itemsTotal = 0;

  rows.forEach((item, idx) => {
    const isTube =
      item.itemType === "Tubes" || item.itemType === "Coassiali Tubes";
    const pos = (idx + 1) * 100;
    const descr = item.description || "";
    const code = item.code || "";
    const um = isTube ? "mt" : "pz";
    const qty = Number(item.quantity || 0);

    // IMPORTANTE: il Word deve copiare la stessa logica del file Excel esportato.
    // Nell'Excel le colonne F..K vengono prima arrotondate a 2 decimali,
    // poi F viene scontata partendo dal valore lordo arrotondato in M.
    // Se qui usiamo il valore grezzo del catalogo, es. 27,665, Word può dare
    // 17,75 mentre Excel mostra 17,76.
    const baseListinoRaw = Number(
      (isTube ? item.basePricePerM : item.basePricePerPc) || 0,
    );
    const baseListino = round2(baseListinoRaw);
    const baseScontato = round2(baseListino * (1 - dp / 100));
    const peso = isTube ? round2(Number(item.pesoKgM || 0)) : null;
    const asKg = isTube ? round2(Number(item.alloySurchargePerKg || 0)) : null;
    const asM = isTube ? round2((Number(item.alloySurchargePerKg || 0)) * (Number(item.pesoKgM || 0))) : null;
    const pu = isTube ? round2(baseScontato + (asM || 0)) : baseScontato;
    const tot = round2(pu * qty);
    itemsTotal += tot;

    const row = [
      String(pos),
      descr,
      code,
      um,
      String(Math.round(qty)),
      fmtWordNumber(baseScontato, true),
      isTube ? fmtWordNumber(peso, false) : "-",
      isTube ? fmtWordNumber(asKg, false) : "-",
      isTube ? fmtWordNumber(asM, false) : "-",
      fmtWordNumber(pu, true),
      fmtWordNumber(tot, true),
    ];
    if (includeLeadTime) row.push(String(item.leadTime || ""));
    tableRows.push(row);
  });

  const transportValue = round2((itemsTotal * tp) / 100);
  const totalDestino = round2(itemsTotal + transportValue);

  const total1 = [
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Total Items ex works",
    "",
    fmtWordNumber(itemsTotal, true),
  ];
  const total2 = [
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Imballo e trasporto",
    "",
    fmtWordNumber(transportValue, true),
  ];
  const total3 = [
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Totale f.co destino",
    "",
    fmtWordNumber(totalDestino, true),
  ];
  if (includeLeadTime) {
    total1.push("");
    total2.push("");
    total3.push("");
  }
  tableRows.push(total1);
  tableRows.push(total2);
  tableRows.push(total3);

  return tableRows;
}

function tc(text, opts = {}) {
  const bold = opts.bold ? "<w:b/>" : "";
  const align = opts.align || "center";
  const shaded = opts.shaded ? '<w:shd w:fill="D9EAF7"/>' : "";
  const width = opts.width || 900;
  const gridSpan = opts.gridSpan
    ? `<w:gridSpan w:val="${opts.gridSpan}"/>`
    : "";
  return `
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="${width}" w:type="dxa"/>
        ${gridSpan}
        <w:vAlign w:val="center"/>
        ${shaded}
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="${align}"/></w:pPr>
        <w:r>
          <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/>${bold}</w:rPr>
          <w:t xml:space="preserve">${xmlEscape(text).replace(/\n/g, '</w:t><w:br/><w:t xml:space="preserve">')}</w:t>
        </w:r>
      </w:p>
    </w:tc>`;
}

function tr(cells, rowIndex, includeLeadTime = false) {
  const isHeader = rowIndex === 0;
  const label = String(cells[8] || "");
  const isTotal =
    rowIndex > 0 &&
    (label === "Total Items ex works" ||
      label === "Imballo e trasporto" ||
      label === "Totale f.co destino");

  // A..K. POS e Valore riga sono stati allargati del 10%.
  const widths = includeLeadTime
    ? [787, 2300, 1050, 520, 760, 1050, 820, 850, 850, 1050, 1634, 1300]
    : [787, 2300, 1050, 520, 760, 1050, 820, 850, 850, 1050, 1634];

  if (isTotal) {
    const cellsXml = [];
    for (let i = 0; i < 8; i++) {
      cellsXml.push(tc(cells[i] || "", { width: widths[i], align: "center" }));
    }
    // Unisce colonna I + J: etichetta centrata su una cella unica.
    cellsXml.push(
      tc(label, {
        bold: true,
        align: "center",
        width: widths[8] + widths[9],
        gridSpan: 2,
      }),
    );
    cellsXml.push(
      tc(cells[10] || "", {
        bold: true,
        align: "right",
        width: widths[10],
      }),
    );
    if (includeLeadTime) {
      cellsXml.push(
        tc(cells[11] || "", {
          bold: true,
          align: "center",
          width: widths[11],
        }),
      );
    }
    return `<w:tr>${cellsXml.join("")}</w:tr>`;
  }

  return `<w:tr>${cells
    .map((c, i) =>
      tc(c, {
        bold: isHeader,
        shaded: isHeader,
        width: widths[i] || 900,
        align: isHeader
          ? "center"
          : i === 1
            ? "left"
            : i >= 4
              ? "right"
              : "center",
      }),
    )
    .join("")}</w:tr>`;
}

function buildWordTableXml(tableRows, includeLeadTime = false) {
  const widths = includeLeadTime
    ? [787, 2300, 1050, 520, 760, 1050, 820, 850, 850, 1050, 1634, 1300]
    : [787, 2300, 1050, 520, 760, 1050, 820, 850, 850, 1050, 1634];
  const grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
  return `
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
    <w:tblLayout w:type="fixed"/>
  </w:tblPr>
  ${grid}
  ${tableRows.map((r, idx) => tr(r, idx, includeLeadTime)).join("\n")}
</w:tbl>
<w:p/>`;
}

function insertTableAfterOfferParagraph(documentXml, tableXml) {
  const marker = "In allegato vi inviamo le specifiche del materiale offerto";
  const markerPos = documentXml.indexOf(marker);
  if (markerPos === -1)
    return documentXml.replace("</w:body>", `${tableXml}</w:body>`);
  const pEnd = documentXml.indexOf("</w:p>", markerPos);
  if (pEnd === -1)
    return documentXml.replace("</w:body>", `${tableXml}</w:body>`);
  const twoBlankLines = "<w:p/><w:p/>";
  return (
    documentXml.slice(0, pEnd + 6) +
    twoBlankLines +
    tableXml +
    documentXml.slice(pEnd + 6)
  );
}

app.post("/api/export-word", async (req, res) => {
  try {
    const {
      rows = [],
      discountPercent = 0,
      transport = {},
      client = null,
      includeLeadTime = false,
    } = req.body || {};
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).send("Nessuna riga da esportare in Word.");
    }

    const templatePath = path.join(__dirname, "Offerta Test.docx");
    if (!fs.existsSync(templatePath)) {
      return res
        .status(500)
        .send("Template Word non trovato: Offerta Test.docx");
    }

    const zip = new PizZip(fs.readFileSync(templatePath));
    const documentFile = zip.file("word/document.xml");
    if (!documentFile)
      return res
        .status(500)
        .send("document.xml non trovato nel template Word.");

    const tableRows = buildOfferRowsForWord(
      rows,
      discountPercent,
      transport,
      Boolean(includeLeadTime),
    );
    const tableXml = buildWordTableXml(tableRows, Boolean(includeLeadTime));
    const headerXml = replaceClientHeaderAndDate(documentFile.asText(), client);
    const newXml = insertTableAfterOfferParagraph(headerXml, tableXml);
    zip.file("word/document.xml", newXml);

    const buf = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    res.setHeader("Content-Disposition", 'attachment; filename="Offerta.docx"');
    return res.send(buf);
  } catch (err) {
    console.error(err);
    return res.status(500).send("Errore durante esportazione Word.");
  }
});

app.listen(3000, async () => {
  const url = "http://localhost:3000";

  console.log(`Server su ${url}`);

  try {
    await open(url, { app: { name: "chrome" } });
  } catch (err) {
    await open(url);
  }
});

// ===== Import Excel per ripopolare la tabella =================================
app.post("/api/import", upload.single("file"), async (req, res) => {
  try {
    if (!req.file || !req.file.buffer) {
      return res.status(400).send("Nessun file caricato.");
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);
    const ws = wb.getWorksheet("Offerta") || wb.worksheets[0];
    if (!ws) return res.json({ rows: [] });

    const rows = [];
    for (let r = 3; r <= ws.rowCount; r++) {
      const descr = ws.getCell(r, 2).value; // B
      const um = String(ws.getCell(r, 4).value || "")
        .trim()
        .toLowerCase(); // D
      const qty = ws.getCell(r, 5).value; // E
      const base = ws.getCell(r, 6).value; // F
      const peso = ws.getCell(r, 7).value; // G
      const asKg = ws.getCell(r, 8).value; // H
      const pu = ws.getCell(r, 10).value; // J
      const tot = ws.getCell(r, 11).value; // K
      const cellIJ = String(ws.getCell(r, 9).value || "").toLowerCase(); // I as text for 'Totale Items'

      const emptyDescr = descr == null || String(descr).trim() === "";
      const numbersAllZero =
        Number(qty || 0) === 0 &&
        Number(pu || 0) === 0 &&
        Number(tot || 0) === 0;

      if ((emptyDescr && numbersAllZero) || cellIJ.includes("totale items"))
        break;

      const code = ws.getCell(r, 3).value; // C
      const prezzoPieno = ws.getCell(r, 13).value; // M = Prezzo pieno (prima dello sconto)

      rows.push({
        description: descr ? String(descr) : "",
        code: code ? String(code) : "",
        um,
        quantity: Number(qty || 0),
        base: Number(base || 0),
        peso: um === "mt" ? Number(peso || 0) : null,
        asKg: um === "mt" ? Number(asKg || 0) : null,
        unitPrice: Number(prezzoPieno || 0), // usa SEMPRE colonna M come base
        pu_file: Number(pu || 0),
        tot_file: Number(tot || 0),
        prezzoPienoM: Number(prezzoPieno || 0),
        tot: Number(tot || 0),
        itemType: um === "mt" ? "Tubes" : "Imported",
      });
    }
    return res.json({ rows });
  } catch (err) {
    console.error(err);
    return res.status(500).send("Errore durante l'import.");
  }
});

// ======================================================
// WORD EXPORT TABLE STYLE
// Tutta la tabella Word deve usare:
// font: Calibri
// size: 10 pt
// ======================================================
