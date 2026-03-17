// api/actualizar.js
// Vercel Serverless Function — se ejecuta como cron o manualmente
// Descarga el Excel de OneDrive, procesa los datos y los guarda en GitHub

const https = require("https");
const http  = require("http");

// ── Helpers ──────────────────────────────────────────────────────────────────

function httpGet(url, redirects = 5) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith("https") ? https : http;
    lib.get(url, { headers: { "User-Agent": "AutoFacil-Dashboard/1.0" } }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location && redirects > 0) {
        return httpGet(res.headers.location, redirects - 1).then(resolve).catch(reject);
      }
      const chunks = [];
      res.on("data", c => chunks.push(c));
      res.on("end", () => resolve({ status: res.statusCode, body: Buffer.concat(chunks), headers: res.headers }));
      res.on("error", reject);
    }).on("error", reject);
  });
}

function githubRequest(method, path, token, body = null) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const opts = {
      hostname: "api.github.com",
      path,
      method,
      headers: {
        "Authorization": `token ${token}`,
        "Accept":        "application/vnd.github.v3+json",
        "User-Agent":    "AutoFacil-Dashboard/1.0",
        "Content-Type":  "application/json",
        ...(data ? { "Content-Length": Buffer.byteLength(data) } : {}),
      },
    };
    const req = https.request(opts, (res) => {
      const chunks = [];
      res.on("data", c => chunks.push(c));
      res.on("end", () => resolve({ status: res.statusCode, body: JSON.parse(Buffer.concat(chunks).toString() || "{}") }));
    });
    req.on("error", reject);
    if (data) req.write(data);
    req.end();
  });
}

// ── Procesamiento de Excel ────────────────────────────────────────────────────
// Parseo mínimo de XLSX sin dependencias externas
// El archivo es un ZIP con XMLs internos

const zlib = require("zlib");
const { promisify } = require("util");

async function parseXlsx(buffer) {
  // Usar child_process para llamar a Python (disponible en Vercel)
  const { execSync } = require("child_process");
  const fs = require("fs");
  const path = require("path");
  const tmpFile = "/tmp/excel_input.xlsm";
  const outFile = "/tmp/excel_output.json";

  fs.writeFileSync(tmpFile, buffer);

  const script = `
import openpyxl, json, sys
from collections import defaultdict

wb = openpyxl.load_workbook("${tmpFile}", read_only=True, keep_vba=True, data_only=True)
ws = wb["DETALLE"]

all_data = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None: continue
    mes = row[1]
    if not (mes and hasattr(mes,"year")): continue
    def n(v): return float(v) if v and isinstance(v,(int,float)) else 0.0
    def s(v): return str(v).strip() if v else ""
    prod = s(row[19]); fin_raw = s(row[7])
    if fin_raw=="AUTOFIN" or prod.startswith("AUTOFIN") or prod.startswith("AUTOFACIL"): inst="AUTOFIN"
    elif "UNIDAD" in fin_raw or prod.startswith("UNIDAD"): inst="UNIDAD DE CREDITO"
    else: inst="NO APLICA"
    all_data.append({
        "op":row[0],"mes":f"{mes.year}-{mes.month:02d}","ejecutivo":s(row[6]),
        "financiera":fin_raw,"institucion":inst,"automotora":s(row[8]),
        "estado_eval":s(row[13]),"estado_credito":s(row[16]),
        "saldo_precio":n(row[22]),"monto_financiado":n(row[38]),
        "tasa_cli":n(row[39]),"com_dealer":n(row[46]),"rentab_afa":n(row[52]),
        "com_seguros":n(row[93])+n(row[94])+n(row[95]),
        "com_parque":n(row[84]),
        "plazo":int(row[72]) if row[72] and isinstance(row[72],(int,float)) else 0,
        "mayor_menor":s(row[97]),
    })

with open("${outFile}","w") as f:
    json.dump(all_data, f, ensure_ascii=False)
print(len(all_data))
`;

  const tmpScript = "/tmp/parse_excel.py";
  fs.writeFileSync(tmpScript, script);
  execSync(`python3 ${tmpScript}`, { stdio: "inherit" });
  const result = JSON.parse(fs.readFileSync(outFile, "utf-8"));
  return result;
}

function procesarDatos(allData) {
  const mesesLabels = {
    "2025-01":"Ene 25","2025-02":"Feb 25","2025-03":"Mar 25","2025-04":"Abr 25",
    "2025-05":"May 25","2025-06":"Jun 25","2025-07":"Jul 25","2025-08":"Ago 25",
    "2025-09":"Sep 25","2025-10":"Oct 25","2025-11":"Nov 25","2025-12":"Dic 25",
    "2026-01":"Ene 26","2026-02":"Feb 26","2026-03":"Mar 26","2026-04":"Abr 26",
    "2026-05":"May 26","2026-06":"Jun 26","2026-07":"Jul 26","2026-08":"Ago 26",
    "2026-09":"Sep 26","2026-10":"Oct 26","2026-11":"Nov 26","2026-12":"Dic 26",
  };

  const mesesKeys = [...new Set(allData.map(r => r.mes))].sort();

  // Tendencia mensual
  const tendencia = mesesKeys.map(m => {
    const rows   = allData.filter(r => r.mes === m);
    const conFin = rows.filter(r => r.institucion === "AUTOFIN" || r.institucion === "UNIDAD DE CREDITO");
    const ot     = conFin.filter(r => r.estado_eval === "OTORGADO");
    return {
      mes:      mesesLabels[m] || m,
      mes_key:  m,
      total_ops: rows.length,
      otorgados: ot.length,
      rechazados: rows.filter(r => r.estado_eval === "RECHAZADO").length,
      saldo_ot:  ot.reduce((a, r) => a + r.saldo_precio, 0),
      com_dealer: ot.reduce((a, r) => a + r.com_dealer, 0),
      rentab_afa: ot.reduce((a, r) => a + r.rentab_afa, 0),
      com_seguros: ot.reduce((a, r) => a + r.com_seguros, 0),
      tasa_conversion: rows.length ? +(ot.length / rows.length * 100).toFixed(1) : 0,
    };
  });

  // Desempeño ejecutivos
  const ejMap = {};
  allData.forEach(r => {
    if (!r.ejecutivo) return;
    if (!ejMap[r.ejecutivo]) ejMap[r.ejecutivo] = {};
    if (!ejMap[r.ejecutivo][r.mes]) ejMap[r.ejecutivo][r.mes] = { ing:0,apro:0,ot:0,rec:0,monto_ot:0 };
    const d = ejMap[r.ejecutivo][r.mes];
    d.ing++;
    if (!["RECHAZADO","ANULADO"].includes(r.estado_eval)) d.apro++;
    if (r.estado_eval === "OTORGADO") { d.ot++; d.monto_ot += r.saldo_precio; }
    if (r.estado_eval === "RECHAZADO") d.rec++;
  });

  const ejecutivos = Object.keys(ejMap)
    .sort((a,b) => mesesKeys.reduce((s,m)=>(s+(ejMap[b][m]?.ot||0)),0) - mesesKeys.reduce((s,m)=>(s+(ejMap[a][m]?.ot||0)),0));

  const last12 = mesesKeys.slice(-12), last6 = mesesKeys.slice(-6), last3 = mesesKeys.slice(-3);

  const ejPerf = ejecutivos.map(ej => {
    const row = { nombre: ej, meses: {} };
    mesesKeys.forEach(m => {
      const d = ejMap[ej][m] || {};
      row.meses[m] = {
        ing:d.ing||0, apro:d.apro||0, ot:d.ot||0, rec:d.rec||0,
        tc: d.apro ? +((d.ot||0)/d.apro*100).toFixed(1) : 0,
        ta: d.ing  ? +((d.apro||0)/d.ing*100).toFixed(1) : 0,
        prom: d.ot  ? +((d.monto_ot||0)/d.ot/1e6).toFixed(2) : 0,
      };
    });
    [["p12",last12],["p6",last6],["p3",last3]].forEach(([tag, win]) => {
      const ti=win.reduce((s,m)=>s+(ejMap[ej][m]?.ing||0),0);
      const ta=win.reduce((s,m)=>s+(ejMap[ej][m]?.apro||0),0);
      const to=win.reduce((s,m)=>s+(ejMap[ej][m]?.ot||0),0);
      const tr=win.reduce((s,m)=>s+(ejMap[ej][m]?.rec||0),0);
      const tm=win.reduce((s,m)=>s+(ejMap[ej][m]?.monto_ot||0),0);
      row[tag] = {
        ing:+(ti/win.length).toFixed(1), apro:+(ta/win.length).toFixed(1),
        ot:+(to/win.length).toFixed(1),  rec:+(tr/win.length).toFixed(1),
        tc: ta?+(to/ta*100).toFixed(1):0, ta:ti?+(ta/ti*100).toFixed(1):0,
        prom: to?+(tm/to/1e6).toFixed(2):0,
      };
    });
    return row;
  });

  return {
    generado_en:     new Date().toISOString(),
    total_registros: allData.length,
    raw:       allData,
    tendencia,
    ej_perf: { meses: mesesKeys, meses_labels: mesesLabels, ejecutivos: ejPerf },
  };
}

// ── Handler principal ─────────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  // Seguridad: solo cron de Vercel o request con token correcto
  const authHeader = req.headers.authorization || "";
  const cronHeader = req.headers["x-vercel-cron"] || "";
  const cronSecret = process.env.CRON_SECRET || "";

  if (!cronHeader && authHeader !== `Bearer ${cronSecret}`) {
    return res.status(401).json({ error: "No autorizado" });
  }

  try {
    const ONEDRIVE_URL  = process.env.ONEDRIVE_URL;
    const GITHUB_TOKEN  = process.env.GITHUB_TOKEN;
    const GITHUB_REPO   = process.env.GITHUB_REPO;
    const GITHUB_FILE   = process.env.GITHUB_FILE || "data/dashboard_data.json";

    if (!ONEDRIVE_URL || !GITHUB_TOKEN || !GITHUB_REPO) {
      throw new Error("Faltan variables de entorno: ONEDRIVE_URL, GITHUB_TOKEN, GITHUB_REPO");
    }

    console.log("1. Descargando Excel...");
    const { body: excelBuffer } = await httpGet(ONEDRIVE_URL);

    console.log("2. Parseando Excel...");
    const allData = await parseXlsx(excelBuffer);

    console.log("3. Procesando datos...");
    const datos = procesarDatos(allData);

    console.log("4. Subiendo a GitHub...");
    const apiPath = `/repos/${GITHUB_REPO}/contents/${GITHUB_FILE}`;
    const existing = await githubRequest("GET", apiPath, GITHUB_TOKEN);
    const sha = existing.body?.sha;

    const contenidoB64 = Buffer.from(JSON.stringify(datos)).toString("base64");
    await githubRequest("PUT", apiPath, GITHUB_TOKEN, {
      message: `datos: actualización ${new Date().toISOString().slice(0,16)}`,
      content: contenidoB64,
      branch:  "main",
      ...(sha ? { sha } : {}),
    });

    console.log("✅ Completado");
    res.status(200).json({
      ok: true,
      registros: allData.length,
      generado_en: datos.generado_en,
    });

  } catch (err) {
    console.error("Error:", err.message);
    res.status(500).json({ error: err.message });
  }
};
