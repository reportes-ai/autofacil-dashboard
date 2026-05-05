"""
AutoFácil Dashboard — Actualizador de datos
============================================
Lee el Excel desde OneDrive y sube el JSON actualizado a GitHub.

Instalación:
    pip install requests openpyxl

Uso:
    python actualizar_datos.py

Variables de entorno necesarias (o editar directamente abajo):
    GITHUB_TOKEN   — Token personal de GitHub
    GITHUB_REPO    — Usuario/repositorio (ej: "juanperez/autofacil-dashboard")
    ONEDRIVE_URL   — URL compartida del Excel en OneDrive
"""

import requests
import openpyxl
import json
import base64
import os
import io
import sys
from collections import defaultdict
from datetime import datetime, timedelta
import re as _re

# ============================================================
# CONFIGURACIÓN — edita estos valores o usa variables de entorno
# ============================================================

ONEDRIVE_URL = os.getenv("ONEDRIVE_URL",
    "https://1drv.ms/x/c/2c6ba7ee2485da71/IQC8NJiCFIjLRo8EZDAXqLpWAX5U_yID0rVg88Cvimp2pLA?e=hEctPA"
)

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")        # ghp_xxxxxxxxxxxx
GITHUB_REPO  = os.getenv("GITHUB_REPO",  "")        # usuario/repositorio
GITHUB_FILE  = os.getenv("GITHUB_FILE",  "data/dashboard_data.json")

# ============================================================
# PASO 1: Descargar Excel desde OneDrive
# ============================================================

def descargar_excel(url: str) -> bytes:
    print("📥 Descargando Excel desde OneDrive...")

    # Convertir URL compartida a URL de descarga directa
    # Formato 1drv.ms → download directo
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }

    # Primer request para seguir redirecciones y obtener URL final
    r = requests.get(url, headers=headers, allow_redirects=True, timeout=30)

    if r.status_code != 200:
        raise Exception(f"Error descargando archivo: HTTP {r.status_code}")

    content = r.content

    # Si no es un ZIP válido (los xlsx son ZIPs), intentar con download=1
    if content[:4] != b"PK\x03\x04":
        download_url = (url + "&download=1") if "?" in url else (url + "?download=1")
        r2 = requests.get(download_url, headers=headers, allow_redirects=True, timeout=30)
        if r2.status_code == 200 and r2.content[:4] == b"PK\x03\x04":
            content = r2.content
        else:
            raise Exception("No se pudo obtener el archivo Excel. Verifica que el link sea publico.")

    print(f"   ✓ Archivo descargado ({len(content)/1024:.0f} KB)")
    return content


# ============================================================
# PASO 2: Procesar Excel y generar datos del dashboard
# ============================================================

def procesar_excel(contenido: bytes) -> dict:
    print("⚙️  Procesando hoja DETALLE...")

    wb = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True,
                                 keep_vba=True, data_only=True)
    ws = wb["DETALLE"]

    meses_labels = {
        "2025-01":"Ene 25","2025-02":"Feb 25","2025-03":"Mar 25","2025-04":"Abr 25",
        "2025-05":"May 25","2025-06":"Jun 25","2025-07":"Jul 25","2025-08":"Ago 25",
        "2025-09":"Sep 25","2025-10":"Oct 25","2025-11":"Nov 25","2025-12":"Dic 25",
        "2026-01":"Ene 26","2026-02":"Feb 26","2026-03":"Mar 26","2026-04":"Abr 26",
        "2026-05":"May 26","2026-06":"Jun 26","2026-07":"Jul 26","2026-08":"Ago 26",
        "2026-09":"Sep 26","2026-10":"Oct 26","2026-11":"Nov 26","2026-12":"Dic 26",
    }

    all_data = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        mes_raw = row[1]
        # Manejar cualquier formato de fecha en col B
        if mes_raw and hasattr(mes_raw, 'year'):
            pass  # datetime OK
        elif mes_raw and isinstance(mes_raw, str):
            m = _re.search(r'(\d{4})[-/](\d{1,2})', str(mes_raw))
            if m:
                mes_raw = datetime(int(m.group(1)), int(m.group(2)), 1)
            else:
                m2 = _re.search(r'(\d{1,2})[-/](\d{4})', str(mes_raw))
                if m2:
                    mes_raw = datetime(int(m2.group(2)), int(m2.group(1)), 1)
                else:
                    continue
        elif mes_raw and isinstance(mes_raw, (int, float)):
            mes_raw = datetime(1899, 12, 30) + timedelta(days=int(mes_raw))
        else:
            continue

        def n(v):
            if v is None: return 0.0
            if isinstance(v, (int, float)):
                # Cap: valores absurdos (fórmulas corruptas) → 0
                f = float(v)
                return f if abs(f) < 1e12 else 0.0
            try:
                f = float(str(v).replace('$','').replace(' ','').replace(',','.').strip())
                return f if abs(f) < 1e12 else 0.0
            except: return 0.0
        def s(v):  return str(v).strip() if v else ""

        prod      = s(row[19])   # T - PRODUCTO
        fin_raw   = s(row[7])    # H - FINANCIERA
        com_seg   = n(row[96])   # CS - COM.SEGUROS TOTAL

        # Derivar institución desde producto
        if fin_raw == "AUTOFIN" or prod.startswith("AUTOFIN") or prod.startswith("AUTOFACIL"):
            institucion = "AUTOFIN"
        elif "UNIDAD" in fin_raw or prod.startswith("UNIDAD"):
            institucion = "UNIDAD DE CREDITO"
        else:
            institucion = "NO APLICA"

        mes_key = f"{mes_raw.year}-{mes_raw.month:02d}"

        all_data.append({
            "op":          row[0],
            "mes":         mes_key,
            "rut":         s(row[3]),    # D - RUT
            "nombre":      s(row[4]),    # E - NOMBRE
            "ejecutivo":   s(row[6]),    # G - EJ. COMERCIAL
            "financiera":  fin_raw,      # H - FINANCIERA
            "institucion": institucion,
            "automotora":  s(row[8]),    # I - AUTOMOTORA
            "estado_eval": s(row[16]) if s(row[16]) in ["ANULADO","RECHAZADO"] else s(row[13]),  # Q / N
            "estado_credito": s(row[16]),   # Q - ESTADO CRÉDITO
            "saldo_precio":    n(row[22]),   # W - SALDO PRECIO
            "monto_financiado": n(row[38]),  # AM - MONTO FINANCIADO INDEXA
            "tasa_cli":    n(row[39]),   # AN - TASCLI REAL
            "com_dealer":  n(row[46]),   # AU - COMDEA $ REAL
            "rentab_afa":  n(row[62]),   # BK - MONTO DE PAGO COMISION FINAN.
            "com_seguros": com_seg,      # BC - COM SEGUROS
            "com_parque":  n(row[83]),   # CF - COM PARQUE
            "plazo":       int(row[72]) if row[72] and isinstance(row[72], (int, float)) else 0,  # BU
            "mayor_menor": s(row[97]),   # CT - MAYOR/MENOR
            "estado_sp":   s(row[36]),   # AK - ESTADO SP
            "fecha_ot":    row[17].strftime("%Y-%m-%d") if row[17] and hasattr(row[17],"day") else s(row[17]),  # R
            "fecha_estado": row[14].strftime("%Y-%m-%d") if row[14] and hasattr(row[14],"day") else "",  # O
            "dia_mes":     int(row[10]) if row[10] and isinstance(row[10], (int,float)) else 0,
            "dia_semana":  str(row[11]).strip().lower() if row[11] else "",
            # Campos exportación Saldo Precio
            "fecha_recepcion_fei":  row[33].strftime("%Y-%m-%d") if row[33] and hasattr(row[33],"day") else s(row[33]),  # AH
            "alerta_recep_fei":     s(row[34]),   # AI
            "alerta_pago_sp":       s(row[37]),   # AL
        })

    print(f"   ✓ {len(all_data)} registros procesados")

    # ---- Tendencia mensual ----
    meses_keys = sorted(set(r["mes"] for r in all_data))
    tendencia = []
    for m in meses_keys:
        rows = [r for r in all_data if r["mes"] == m]
        con_fin = [r for r in rows if r["institucion"] in ["AUTOFIN", "UNIDAD DE CREDITO"]]
        ot = [r for r in con_fin if r["estado_eval"] == "OTORGADO"]
        tendencia.append({
            "mes":       meses_labels.get(m, m),
            "mes_key":   m,
            "total_ops": len(rows),
            "otorgados": len(ot),
            "rechazados": len([r for r in rows if r["estado_eval"] == "RECHAZADO"]),
            "saldo_ot":  sum(r["saldo_precio"] for r in ot),
            "com_dealer": sum(r["com_dealer"] for r in ot),
            "rentab_afa": sum(r["rentab_afa"] for r in ot),
            "com_seguros": sum(r["com_seguros"] for r in ot),
            "tasa_conversion": round(len(ot)/len(rows)*100, 1) if rows else 0,
        })

    # ---- Desempeño por ejecutivo ----
    ej_data = defaultdict(lambda: defaultdict(lambda: {
        "ing":0,"apro":0,"ot":0,"rec":0,"monto_ot":0
    }))
    for r in all_data:
        ej = r["ejecutivo"]
        if not ej: continue
        m  = r["mes"]
        d  = ej_data[ej][m]
        d["ing"] += 1
        est = r["estado_eval"]
        if est not in ["RECHAZADO","ANULADO"]:
            d["apro"] += 1
        if est == "OTORGADO":
            d["ot"]       += 1
            d["monto_ot"] += r["saldo_precio"]
        if est == "RECHAZADO":
            d["rec"] += 1

    ejecutivos_sorted = sorted(
        ej_data.keys(),
        key=lambda e: sum(ej_data[e][m]["ot"] for m in meses_keys),
        reverse=True
    )

    def percentil(arr, p):
        s = sorted([v for v in arr if v > 0])
        if not s: return 0
        idx = max(0, int(len(s)*p/100)-1)
        return s[idx]

    last_12 = meses_keys[-12:]
    last_6  = meses_keys[-6:]
    last_3  = meses_keys[-3:]

    ej_perf = []
    for ej in ejecutivos_sorted:
        row_ej = {"nombre": ej, "meses": {}}
        for m in meses_keys:
            d = ej_data[ej][m]
            tc   = round(d["ot"]/d["apro"]*100,1)   if d["apro"] else 0
            ta   = round(d["apro"]/d["ing"]*100,1)   if d["ing"]  else 0
            prom = round(d["monto_ot"]/d["ot"]/1e6,2) if d["ot"]  else 0
            row_ej["meses"][m] = {
                "ing":d["ing"],"apro":d["apro"],"ot":d["ot"],
                "rec":d["rec"],"tc":tc,"ta":ta,"prom":prom
            }
        for tag, window in [("p12",last_12),("p6",last_6),("p3",last_3)]:
            ti  = sum(ej_data[ej][m]["ing"]      for m in window)
            ta_ = sum(ej_data[ej][m]["apro"]     for m in window)
            to  = sum(ej_data[ej][m]["ot"]       for m in window)
            tr  = sum(ej_data[ej][m]["rec"]      for m in window)
            tm  = sum(ej_data[ej][m]["monto_ot"] for m in window)
            row_ej[tag] = {
                "ing":  round(ti/len(window),1),
                "apro": round(ta_/len(window),1),
                "ot":   round(to/len(window),1),
                "rec":  round(tr/len(window),1),
                "tc":   round(to/ta_*100,1) if ta_ else 0,
                "ta":   round(ta_/ti*100,1) if ti  else 0,
                "prom": round(tm/to/1e6,2)  if to  else 0,
            }
        ej_perf.append(row_ej)

    resultado = {
        "generado_en":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_registros": len(all_data),
        "raw":      all_data,
        "tendencia": tendencia,
        "ej_perf":   {"meses": meses_keys, "meses_labels": meses_labels, "ejecutivos": ej_perf},
    }

    print(f"   ✓ {len(tendencia)} meses en tendencia")
    print(f"   ✓ {len(ej_perf)} ejecutivos procesados")
    return resultado


# ============================================================
# PASO 3: Subir JSON a GitHub
# ============================================================

def subir_a_github(datos: dict, token: str, repo: str, filepath: str):
    print(f"📤 Subiendo datos a GitHub ({repo}/{filepath})...")

    if not token:
        raise Exception("GITHUB_TOKEN no configurado")
    if not repo:
        raise Exception("GITHUB_REPO no configurado")

    contenido = json.dumps(datos, ensure_ascii=False, separators=(",",":"))
    contenido_b64 = base64.b64encode(contenido.encode()).decode()

    api_url = f"https://api.github.com/repos/{repo}/contents/{filepath}"
    headers = {
        "Authorization": f"token {token}",
        "Accept":        "application/vnd.github.v3+json",
        "Content-Type":  "application/json",
    }

    # Obtener SHA del archivo existente (necesario para actualizar)
    r = requests.get(api_url, headers=headers, timeout=15)
    sha = r.json().get("sha") if r.status_code == 200 else None

    payload = {
        "message": f"datos: actualización automática {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "content": contenido_b64,
        "branch":  "main",
    }
    if sha:
        payload["sha"] = sha

    r2 = requests.put(api_url, headers=headers,
                      data=json.dumps(payload), timeout=30)

    if r2.status_code in (200, 201):
        print(f"   ✓ JSON subido correctamente ({len(contenido)/1024:.0f} KB)")
        print(f"   ✓ Vercel redepoyará automáticamente en ~30 segundos")
    else:
        raise Exception(f"Error GitHub API: {r2.status_code} — {r2.text[:200]}")


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    print("=" * 55)
    print("  AutoFácil Dashboard — Actualizador de datos")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    try:
        # 1. Descargar
        excel_bytes = descargar_excel(ONEDRIVE_URL)

        # 2. Procesar
        datos = procesar_excel(excel_bytes)

        # 3. Subir a GitHub (si hay token configurado)
        if GITHUB_TOKEN and GITHUB_REPO:
            subir_a_github(datos, GITHUB_TOKEN, GITHUB_REPO, GITHUB_FILE)
        else:
            # Guardar localmente si no hay GitHub configurado
            with open("dashboard_data.json", "w", encoding="utf-8") as f:
                json.dump(datos, f, ensure_ascii=False, indent=2)
            print("💾 JSON guardado localmente en dashboard_data.json")
            print("   (configura GITHUB_TOKEN y GITHUB_REPO para subir a GitHub)")

        print()
        print("✅ ¡Completado sin errores!")

    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)
