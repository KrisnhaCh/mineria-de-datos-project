import openpyxl, glob, re, unicodedata
import pandas as pd, numpy as np

UP = "/mnt/user-data/uploads"
files = {}
for f in glob.glob(UP + "/*.xlsx"):
    for y in (2021, 2022, 2023):
        if f"Agricola_{y}_" in f: files[y] = f

def norm(s):
    if s is None: return ""
    s = str(s).replace("\xa0", " ").replace("\n", " ")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()

def num(v):
    if v is None or isinstance(v, bool): return np.nan
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).strip().replace(",", ""))
    except: return np.nan

DIST_FIX = {"C.G. ALBARRACIN L.": "CORONEL GREGORIO ALBARRACIN LANCHIPA"}
# Nombres vacios o ausentes en el original, inferidos (ver columna distrito_inferido)
INFER_SHEET = {(2022, "Table 24"): ("JORGE BASADRE", "ITE"),
               (2022, "Table 32"): ("TARATA", "TARUCACHI")}
INFER_BLOCK = {(2021, "Table 4", 1): ("TACNA", "CALANA"),     # 2.º bloque de la hoja de Tacna 2021
               (2021, "Table 4", 2): ("TACNA", "INCLAN")}     # 3.er bloque

def first_label(r, k=3):
    for j in range(min(k, len(r))):
        if r[j] not in (None, ""): return j, r[j]
    return None, None

def parse_name(r, j0, v0):
    """devuelve (nro, cultivo, indice_de_la_celda_del_nombre) o None"""
    if isinstance(v0, (int, float)) or (isinstance(v0, str) and re.fullmatch(r"\s*\d+\s*", v0)):
        for jj in range(j0 + 1, min(j0 + 3, len(r))):
            if isinstance(r[jj], str) and r[jj].strip():
                return int(float(v0)), r[jj], jj
    elif isinstance(v0, str):
        m = re.match(r"^\s*(\d+)\s+(.+)$", v0.replace("\xa0", " "))
        if m: return int(m.group(1)), m.group(2), j0
    return None

def coherent(prod, cos, rend, tol=0.10):
    if not (prod == prod and cos == cos and rend == rend and cos > 0): return None
    return abs(rend - prod * 1000 / cos) / (prod * 1000 / cos + 1e-9) <= tol

def token_parse(r, jname):
    """Alternativa por orden de aparicion: prod, total, cosechada, [crec], rend, precio"""
    toks = [c for c in r[jname + 1:] if c not in (None, "")]
    vals = [num(c) for c in toks]
    n = len(vals)
    if n >= 6: return dict(prod=vals[0], tot=vals[1], cos=vals[2], rend=vals[-2], precio=vals[-1])
    if n == 5: return dict(prod=vals[0], tot=vals[1], cos=vals[2], rend=vals[3], precio=vals[4])
    return None

recs, skipped, log = [], [], []
for year, f in sorted(files.items()):
    wb = openpyxl.load_workbook(f, data_only=True)
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        prov = dist = None; has_dist = False
        for r in rows[:8]:
            lab = norm(r[0]); vals = [norm(c) for c in r[1:] if c not in (None, "")]
            if lab.startswith("DISTRITO"): has_dist = True; dist = vals[0] if vals else None
            if lab.startswith("PROVINCIA"): prov = vals[0] if vals else None
        if not has_dist:
            skipped.append((year, ws.title)); continue
        inferred_sheet = False
        if not dist and (year, ws.title) in INFER_SHEET:
            prov, dist = INFER_SHEET[(year, ws.title)]; inferred_sheet = True
        dist = DIST_FIX.get(dist, dist)
        tds = [i for i, r in enumerate(rows) if norm(first_label(r)[1]).startswith("TOTAL DISTRITAL")]
        if not tds: skipped.append((year, ws.title)); continue
        # columnas segun encabezado
        cols = {}
        for r in rows[:tds[0]]:
            for j, c in enumerate(r):
                t = norm(c)
                if "PRODUCCION" in t: cols["prod"] = j
                elif t == "TOTAL": cols["tot"] = j
                elif t == "COSECHADA": cols["cos"] = j
                elif "RENDIMIENTO" in t: cols["rend"] = j
                elif "PRECIO" in t: cols["precio"] = j
        for b, td in enumerate(tds):
            end = tds[b + 1] if b + 1 < len(tds) else len(rows)
            blk = rows[td + 1:end]
            jt, _ = first_label(rows[td])
            tdn = [num(c) for c in rows[td][jt + 1:] if num(c) == num(c)]
            total_dist = tdn[0] if tdn else np.nan
            bprov, bdist, binf = prov, dist, inferred_sheet
            if (year, ws.title, b) in INFER_BLOCK:
                bprov, bdist = INFER_BLOCK[(year, ws.title, b)]; binf = True
            # filas de datos
            items = []; tipo = "TRANSITORIO"
            for r in blk:
                j0, v0 = first_label(r)
                if v0 is None: continue
                pn = parse_name(r, j0, v0)
                if pn is None:
                    t0 = norm(v0)
                    if "MIPERM" in t0: tipo = "SEMIPERMANENTE"
                    elif "PERMAN" in t0: tipo = "PERMANENTE"
                    elif "TRANSIT" in t0: tipo = "TRANSITORIO"
                    continue
                items.append((r, tipo, pn))
            # mapeo A: encabezado (con ajuste de columna vecina). mapeo B: posiciones de la fila TOTAL DISTRITAL
            def build(colmap):
                out = []
                for r, tipo_, (nro, name, jn) in items:
                    g = lambda k: num(r[colmap[k]]) if k in colmap and colmap[k] < len(r) else np.nan
                    out.append(dict(prod=g("prod"), tot=g("tot"), cos=g("cos"), rend=g("rend"), precio=g("precio")))
                return out
            def score(vals):
                ch = [coherent(v["prod"], v["cos"], v["rend"]) for v in vals]
                ch = [c for c in ch if c is not None]
                return (sum(ch) / len(ch)) if ch else 0
            candA = dict(cols)
            candB = dict(cols)
            pos = [j for j, c in enumerate(rows[td]) if j > jt and num(c) == num(c)]
            if len(pos) >= 3: candB.update(prod=pos[0], tot=pos[1], cos=pos[2])
            vA, vB = build(candA), build(candB)
            # candidato C: desplazar +-1 las columnas de encabezado
            cands = [(score(vA), candA, vA), (score(vB), candB, vB)]
            for d in (-1, 1):
                cm = {k: v + d for k, v in cols.items()}
                v = build(cm); cands.append((score(v), cm, v))
            cands.sort(key=lambda x: -x[0])
            sc, best_cols, vals = cands[0]
            nfix = 0
            for (r, tipo_, (nro, name, jn)), v in zip(items, vals):
                ch = coherent(v["prod"], v["cos"], v["rend"])
                if ch is not True:
                    alt = token_parse(r, jn)
                    if alt and coherent(alt["prod"], alt["cos"], alt["rend"]):
                        v = alt; nfix += 1
                alt = token_parse(r, jn)
                if alt and coherent(alt["prod"], alt["cos"], alt["rend"]) and alt["rend"] == v["rend"]:
                    if v["tot"] != v["tot"] and alt["tot"] == alt["tot"]: v = dict(v, tot=alt["tot"])
                    if v["precio"] != v["precio"] and alt["precio"] == alt["precio"] and alt["precio"] < 100:
                        v = dict(v, precio=alt["precio"])
                recs.append(dict(anio=year, provincia=norm(bprov), distrito=bdist, tipo_cultivo=tipo_, nro=nro,
                                 cultivo=norm(name), produccion_t=v["prod"], sup_total_ha=v["tot"],
                                 sup_cosechada_ha=v["cos"], rendimiento_kg_ha=v["rend"],
                                 precio_chacra_s_kg=v["precio"], hoja=ws.title, bloque=b,
                                 distrito_inferido=binf, total_distrital_t=total_dist))
            if sc < 0.9 or nfix:
                log.append((year, ws.title, bdist, f"coherencia={sc:.2f}", f"filas_corregidas_por_orden={nfix}"))
df = pd.DataFrame(recs)
df.to_pickle("/home/claude/raw.pkl")
print(df.shape)
print(df.groupby("anio").agg(filas=("cultivo","size"), distritos=("distrito","nunique")))
print("hojas omitidas:", skipped)
for l in log: print("AVISO", l)
print(df.tipo_cultivo.value_counts())
