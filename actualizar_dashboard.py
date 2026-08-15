"""
ACTUALIZADOR DE DASHBOARD - SEGUIMIENTO PRESUPUESTAL BRIAV32
============================================================
Lee TIP 2026.xlsx (hoja BD_Rellenar) y actualiza el bloque de datos
embebido en SEGUIMIENTO PRESUPUESTAL.html.

USO:
  python actualizar_dashboard.py

Luego para publicar en GitHub Pages:
  python publicar_dashboard.py
"""

import pandas as pd
import json, re, sys
from collections import defaultdict

XL    = "TIP 2026.xlsx"
SHEET = "BD_Rellenar"
HTML  = "SEGUIMIENTO PRESUPUESTAL.html"

# ── Indices de columnas (verificados contra header real de BD_Rellenar) ──
COL_RESERVA     = 1   # Reserva?
COL_UT          = 2   # UNIDAD TACTICA
COL_RUBRO       = 3   # RUBRO
COL_BIEN        = 4   # BIEN O SERVICIO (CONCEPTO)
COL_SOLO_VAL    = 8   # Solo valor  2026  (doble espacio) — filas padre
COL_ETAPA       = 10  # CONCEPTO GASTO — etapa: Contrato, Pre-Contractual, etc.
COL_DESC        = 11  # DESCRIPCION
COL_MONTO_2026  = 13  # Monto 2026
COL_DIAS        = 18  # DIAS FALTANTES
COL_ALERTA      = 19  # Alerta
COL_UASIG       = 30  # UNIDAD DONDE SE  ASIGNÓ EL PRESUPUESTO

# Orden de unidades igual al JS
UNITS = ['BAMAV1','BAMAV2','BAMAV3','BAMAV4','BAEMA','BAAAS','EM']

ETAPA_KEY = {
    'Contrato':                          'con',
    'Pre-Contractual':                   'pre',
    'Disponible':                        'dis',
    'Apalancamiento':                    'apa',
    'Recorte DAVAA para bolsa de Ppto.': 'rec',
}

# ── Helpers ──
def v(row, col):
    try:
        val = row[col]
        if pd.isna(val): return ""
        return val
    except: return ""

def num(row, col):
    try:
        val = row[col]
        if pd.isna(val): return 0
        return float(val)
    except: return 0

def es_padre(row):
    return str(v(row, COL_ETAPA)).strip() == ""

def normalizar_unidad(u):
    u = str(u).strip()
    return 'EM' if u.upper().startswith('EM') else u

def sev_key(a):
    if not a or str(a).strip() == "": return None
    s = str(a)
    if re.search(r'DEBIDO', s, re.I):      return 'DEBIDO'
    if re.search(r'menos de 20', s, re.I): return 'CRIT20'
    if '30' in s:                          return 'SEG30'
    if re.search(r'OK', s, re.I):         return 'OK'
    return None

def sev_rank(k):
    return {'DEBIDO':3,'CRIT20':2,'SEG30':1,'OK':0}.get(k, -1)

# ── Leer Excel ──
print(f"Leyendo {XL} ...")
try:
    raw = pd.read_excel(XL, sheet_name=SHEET, header=None)
except PermissionError:
    print(f"\nERROR: '{XL}' esta abierto en Excel. Ciérralo y vuelve a ejecutar.")
    sys.exit(1)
except FileNotFoundError:
    print(f"\nERROR: No se encontro '{XL}' en esta carpeta.")
    sys.exit(1)

df = raw.iloc[9:].reset_index(drop=True)
df.columns = range(len(df.columns))
df_no = df[df[COL_RESERVA].astype(str).str.strip().str.upper() == "NO"].copy()
print(f"Filas con Reserva=NO: {len(df_no)}")

# ── Procesar ──
units_map  = {u: {"u":u,"tot":0,"con":0,"pre":0,"dis":0,"apa":0,"rec":0,"nec":0} for u in UNITS}
line_groups  = {}
alert_groups = {}

for _, row in df_no.iterrows():
    uasig = normalizar_unidad(v(row, COL_UASIG))
    ut    = str(v(row, COL_UT)).strip()
    etapa = str(v(row, COL_ETAPA)).strip()
    bien  = str(v(row, COL_BIEN)).strip()
    desc  = str(v(row, COL_DESC)).strip()
    rubro = str(v(row, COL_RUBRO)).strip()
    monto = num(row, COL_MONTO_2026)
    alerta_raw = str(v(row, COL_ALERTA)).strip()
    if alerta_raw == "nan": alerta_raw = ""
    dias_raw = v(row, COL_DIAS)
    try: dias = int(float(dias_raw)) if not pd.isna(dias_raw) else None
    except: dias = None

    if es_padre(row):
        if uasig in units_map:
            units_map[uasig]["tot"] += num(row, COL_SOLO_VAL)
            units_map[uasig]["nec"] += 1
    else:
        ek = ETAPA_KEY.get(etapa)
        if uasig in units_map and ek:
            units_map[uasig][ek] += monto

        lkey = "||".join([ut, bien, etapa, desc])
        if lkey not in line_groups:
            line_groups[lkey] = {"ut":ut,"uasig":uasig,"bien":bien,"concepto":etapa,
                                 "desc":desc,"monto":0,"rubros":[],"_rank":-1,
                                 "alerta":None,"dias":None}
        lg = line_groups[lkey]
        lg["monto"] += monto
        lg["rubros"].append({"rubro":rubro,"monto":round(monto,2)})
        sk = sev_key(alerta_raw); rk = sev_rank(sk)
        if rk > lg["_rank"]:
            lg["_rank"] = rk
            lg["alerta"] = alerta_raw if alerta_raw else None
            lg["dias"]   = dias

        akey = "||".join([ut, bien])
        if akey not in alert_groups:
            alert_groups[akey] = {"ut":ut,"uasig":uasig,"bien":bien,"monto":0,
                                  "descs":set(),"_rank":-1,"severidad":"","sevkey":None,"dias":None}
        ag = alert_groups[akey]
        ag["monto"] += monto
        if desc: ag["descs"].add(desc)

# Segunda pasada para severidad de alertas
for _, row in df_no.iterrows():
    alerta_raw = str(v(row, COL_ALERTA)).strip()
    if alerta_raw in ("","nan"): continue
    ut    = str(v(row, COL_UT)).strip()
    bien  = str(v(row, COL_BIEN)).strip()
    uasig = normalizar_unidad(v(row, COL_UASIG))
    dias_raw = v(row, COL_DIAS)
    try: dias = int(float(dias_raw)) if not pd.isna(dias_raw) else None
    except: dias = None
    akey = "||".join([ut, bien])
    if akey not in alert_groups:
        alert_groups[akey] = {"ut":ut,"uasig":uasig,"bien":bien,"monto":0,
                              "descs":set(),"_rank":-1,"severidad":"","sevkey":None,"dias":None}
    ag = alert_groups[akey]
    sk = sev_key(alerta_raw); rk = sev_rank(sk)
    if rk > ag["_rank"]:
        ag["_rank"] = rk; ag["severidad"] = alerta_raw
        ag["sevkey"] = sk; ag["dias"] = dias

# ── Armar listas ──
lines = []
for lg in line_groups.values():
    lg["nrubros"] = len(lg["rubros"]); del lg["_rank"]
    lines.append(lg)
lines.sort(key=lambda x: x["monto"], reverse=True)

alerts = []
for ag in alert_groups.values():
    ag["ncontratos"] = len(ag["descs"]); del ag["descs"]; del ag["_rank"]
    alerts.append(ag)
sev_ord = {'DEBIDO':3,'CRIT20':2,'SEG30':1,'OK':0}
alerts.sort(key=lambda x: (-sev_ord.get(x["sevkey"],-1), -x["monto"]))

# ── KPIs ──
tot_aj=0; con_t=0; pre_t=0; dis_t=0; apa_t=0; rec_t=0; nec=0
for u in UNITS:
    r = units_map[u]
    tot_aj+=r["tot"]; con_t+=r["con"]; pre_t+=r["pre"]
    dis_t+=r["dis"];  apa_t+=r["apa"]; rec_t+=r["rec"]; nec+=r["nec"]

pct_com = round(con_t/tot_aj*1000)/10 if tot_aj else 0

D = {
    "kpis": {"tot_aj":round(tot_aj,4),"con_t":round(con_t,4),"pre_t":round(pre_t,4),
             "dis_t":round(dis_t,4),"apa_t":round(apa_t,4),"rec_t":round(rec_t,4),
             "pct_com":pct_com,"nec":nec},
    "units":  [units_map[u] for u in UNITS],
    "lines":  lines,
    "alerts": alerts
}

nuevo_json = json.dumps(D, ensure_ascii=False, separators=(',',':'))

# ── Inyectar en HTML ──
print(f"Actualizando {HTML} ...")
with open(HTML, encoding="utf-8") as f:
    content = f.read()

MARKER = "var D = {"
start_marker = content.find(MARKER)
if start_marker == -1:
    print("ERROR: no se encontro 'var D = {' en el HTML.")
    sys.exit(1)

brace_start = start_marker + len(MARKER) - 1
depth = 0; i = brace_start
while i < len(content):
    c = content[i]
    if c == "{": depth += 1
    elif c == "}":
        depth -= 1
        if depth == 0: end = i + 1; break
    i += 1

if "pct_com" not in content[brace_start:end]:
    print("ERROR: el bloque encontrado no es el bloque de datos.")
    sys.exit(1)

nuevo_content = content[:brace_start] + nuevo_json + content[end:]
with open(HTML, "w", encoding="utf-8") as f:
    f.write(nuevo_content)

# ── Resumen ──
print()
print("=" * 50)
print("  DASHBOARD ACTUALIZADO CORRECTAMENTE")
print("=" * 50)
print(f"  Total Ajustado  : ${tot_aj:>20,.0f}")
print(f"  Contrato        : ${con_t:>20,.0f}")
print(f"  Pre-Contractual : ${pre_t:>20,.0f}")
print(f"  Disponible      : ${dis_t:>20,.0f}")
print(f"  Apalancamiento  : ${apa_t:>20,.0f}")
print(f"  Recorte DAVAA   : ${rec_t:>20,.0f}")
print(f"  % Comprometido  : {pct_com}%")
print(f"  Necesidades     : {nec}")
print(f"  Lineas          : {len(lines)}")
print(f"  Alertas         : {len(alerts)}")
print("=" * 50)
print()
print("Para publicar en linea ejecuta:")
print("  python publicar_dashboard.py")
