#!/usr/bin/env python3
"""JARVIS · cálculos pesados del resumen diario (versión 1.0, 10/11/2026).

Uso:
  python3 jarvis_calc.py --fincas fincas.xlsx --ventas ventas.xlsx \
      [--minimos minimos.json] [--fecha AAAA-MM-DD] > calc.json

Entradas
- fincas.xlsx: exportación de la hoja «Stock Fincas» (pestañas Tabla principal,
  Stock y Presupuestos). La escribe la app Gestión fincas / Cuentas fincas.
- ventas.xlsx: exportación de la hoja «ventas FG Agro» (pestañas Ventas y Stock).
- minimos.json (opcional): lo que devuelve get_values de JARVIS!'Mínimos'!A1:D
  ({"values": [[Producto, Origen, Mínimo, Nota], ...]}) o directamente esa lista.
  Origen = «Venta» o «Tus fincas». Mínimo vacío = automático; 0 = no avisar.

Salida (JSON)
- stock_bajo: productos con stock por debajo del mínimo.
  Mínimo automático = lo que salió de ese producto en las 4 semanas siguientes
  a la misma fecha del año pasado (ventas: vendido; tus fincas: gastado en la
  Tabla principal, buscando el producto por su nombre en la descripción).
  Sin consumo el año pasado y sin mínimo manual → no avisa.
- fincas: gasto de cultivo de la campaña (gastos y tareas sin arrendamiento ni
  recolección) por finca: real a hoy, campaña anterior a la misma fecha y final,
  presupuesto, proyección al ritmo de la campaña anterior y alertas.
- avisos: datos que faltan o no cuadran.

Solo lee ficheros. No escribe nada en ningún sitio.
"""
import argparse
import datetime as dt
import json
import re
import sys
import unicodedata

import openpyxl

EXCLUIR_TIPOS = {"arrendamiento", "recoleccion"}
CONCEPTOS_GASTO_CULTIVO = {
    "fertilizantes", "fitosanitarios", "agua", "resto", "seguro", "otros_gastos",
    "tratamientos", "poda", "trituracion", "desbroce", "otras_tareas",
}
VENTANA_DIAS = 28


def norm(s):
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().strip()
    return re.sub(r"\s+", " ", s)


def num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("€", "").replace("%", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def fecha(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if v is None:
        return None
    s = str(v).strip()
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S.%fZ"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def hace_un_ano(d):
    try:
        return d.replace(year=d.year - 1)
    except ValueError:  # 29 de febrero
        return d.replace(year=d.year - 1, day=28)


def filas(ws):
    it = ws.iter_rows(values_only=True)
    cab = [norm(c) for c in next(it)]
    for r in it:
        if r is None or all(c is None or c == "" for c in r):
            continue
        yield {cab[i]: r[i] for i in range(min(len(cab), len(r)))}


def r2(x):
    return round(x + 0.0, 2)


def leer_minimos(ruta):
    if not ruta:
        return {}
    with open(ruta, encoding="utf-8") as f:
        d = json.load(f)
    vals = d.get("values", []) if isinstance(d, dict) else d
    out = {}
    for i, fila in enumerate(vals):
        if i == 0 and fila and norm(fila[0]) == "producto":
            continue
        if not fila or not str(fila[0]).strip():
            continue
        origen = norm(fila[1] if len(fila) > 1 else "")
        origen = "fincas" if "finca" in origen else "venta"
        m = fila[2] if len(fila) > 2 else ""
        if m is None or str(m).strip() == "":
            continue
        out[(origen, norm(fila[0]))] = num(m)
    return out


def nombre_base(p):
    """«Touchdown (glifosato)» → «touchdown»; «Complex AID E/10 L» → «complex aid»."""
    n = norm(p)
    n = re.sub(r"\(.*?\)", "", n)
    n = re.sub(r"\be/\s*[\d.,]+\s*(l|kg|g|ml|cc)\b", "", n)
    return re.sub(r"\s+", " ", n).strip()


def stock_bajo(wf, wv, hoy, minimos):
    ini = hace_un_ano(hoy)
    fin = ini + dt.timedelta(days=VENTANA_DIAS)
    res = []

    # Productos de venta
    vendido = {}
    for f in filas(wv["Ventas"]):
        d = fecha(f.get("fecha"))
        if d and ini < d <= fin:
            k = norm(f.get("producto"))
            vendido[k] = vendido.get(k, 0.0) + num(f.get("cantidad"))
    for f in filas(wv["Stock"]):
        if norm(f.get("activo")) not in ("si", ""):
            continue
        prod = f.get("producto")
        k = norm(prod)
        if not k or k == "total":
            continue
        stock = num(f.get("stock"))
        auto = vendido.get(k, 0.0)
        manual = minimos.get(("venta", k))
        minimo = manual if manual is not None else auto
        if minimo > 0 and stock < minimo:
            res.append({
                "origen": "Venta", "producto": prod, "stock": r2(stock),
                "minimo": r2(minimo), "fuente": "manual" if manual is not None else "auto",
                "salida_4s_ano_pasado": r2(auto), "proveedor": f.get("proveedor") or "",
            })

    # Productos de tus fincas
    gastos = []
    for f in filas(wf["Tabla principal"]):
        d = fecha(f.get("fecha"))
        if d and ini < d <= fin and norm(f.get("categoria")) == "gasto":
            gastos.append((norm(f.get("descripcion")), num(f.get("cantidad"))))
    for f in filas(wf["Stock"]):
        prod = f.get("producto")
        k = norm(prod)
        if not k or k == "total":
            continue
        base = nombre_base(prod)
        stock = num(f.get("stock"))
        auto = sum(q for desc, q in gastos if desc == k or (base and (desc == base or desc.startswith(base + " ") or desc.startswith(base + "("))))
        manual = minimos.get(("fincas", k))
        minimo = manual if manual is not None else auto
        if minimo > 0 and stock < minimo:
            res.append({
                "origen": "Tus fincas", "producto": prod, "stock": r2(stock),
                "minimo": r2(minimo), "fuente": "manual" if manual is not None else "auto",
                "salida_4s_ano_pasado": r2(auto),
            })
    res.sort(key=lambda x: (x["origen"], x["stock"] - x["minimo"]))
    return res


def fincas(wf, hoy):
    camp = hoy.year
    hoy_ant = hace_un_ano(hoy)
    datos = {}
    ultima = {}

    def acc(finca):
        return datos.setdefault(finca, {
            "real": 0.0, "ant_misma_fecha": 0.0, "ant_final": 0.0,
            "partidas_real": {}, "partidas_ant": {},
        })

    for f in filas(wf["Tabla principal"]):
        finca = f.get("finca")
        if not finca:
            continue
        d = fecha(f.get("fecha"))
        if d and (finca not in ultima or d > ultima[finca]):
            ultima[finca] = d
        cat = norm(f.get("categoria"))
        tipo = f.get("tipo") or "(sin tipo)"
        if cat not in ("gasto", "tarea") or norm(tipo) in EXCLUIR_TIPOS:
            continue
        c = int(num(f.get("campana")))
        imp = num(f.get("importe"))
        if c == camp and (d is None or d <= hoy):
            a = acc(finca)
            a["real"] += imp
            a["partidas_real"][tipo] = a["partidas_real"].get(tipo, 0.0) + imp
        elif c == camp - 1:
            a = acc(finca)
            a["ant_final"] += imp
            if d is not None and d <= hoy_ant:
                a["ant_misma_fecha"] += imp
                a["partidas_ant"][tipo] = a["partidas_ant"].get(tipo, 0.0) + imp

    pres = {}
    for f in filas(wf["Presupuestos"]):
        if int(num(f.get("campana"))) != camp:
            continue
        if norm(f.get("concepto")) in CONCEPTOS_GASTO_CULTIVO:
            pres[f.get("finca")] = pres.get(f.get("finca"), 0.0) + num(f.get("valor"))

    out = []
    for finca in sorted(set(datos) | set(pres)):
        a = datos.get(finca) or acc(finca)
        p = pres.get(finca)
        ritmo = a["ant_misma_fecha"] / a["ant_final"] if a["ant_final"] > 0 else None
        proy = a["real"] / ritmo if ritmo and ritmo >= 0.05 else None
        alertas = []
        if p is None:
            alertas.append(f"sin presupuesto {camp}")
        elif a["real"] > p:
            alertas.append(f"ya supera el presupuesto: {r2(a['real'])} € de {r2(p)} €")
        elif proy is not None and p > 0 and proy > p * 1.10 and proy - p >= 100:
            alertas.append(f"a este ritmo acabaría en {round(proy)} € (presupuesto {round(p)} €)")
        dif = a["real"] - a["ant_misma_fecha"]
        if a["ant_misma_fecha"] > 0 and dif >= 100 and dif > a["ant_misma_fecha"] * 0.10:
            alertas.append(f"+{round(dif)} € frente a la campaña anterior a la misma fecha")
        partidas = []
        for t in sorted(set(a["partidas_real"]) | set(a["partidas_ant"])):
            r, an = a["partidas_real"].get(t, 0.0), a["partidas_ant"].get(t, 0.0)
            if r - an >= 100 and (an == 0 or r > an * 1.10):
                partidas.append({"partida": t, "real": r2(r), "ant_misma_fecha": r2(an)})
        partidas.sort(key=lambda x: x["ant_misma_fecha"] - x["real"])
        u = ultima.get(finca)
        dias = (hoy - u).days if u else None
        if dias is not None and dias >= 30:
            alertas.append(f"{dias} días sin anotar nada")
        out.append({
            "finca": finca, "campana": camp, "real": r2(a["real"]),
            "presupuesto": r2(p) if p is not None else None,
            "pct_presupuesto": round(100 * a["real"] / p) if p else None,
            "ant_misma_fecha": r2(a["ant_misma_fecha"]), "ant_final": r2(a["ant_final"]),
            "proyeccion": round(proy) if proy is not None else None,
            "ultima_anotacion": u.isoformat() if u else None,
            "alertas": alertas, "partidas_por_encima": partidas[:4],
        })
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fincas", required=True)
    ap.add_argument("--ventas", required=True)
    ap.add_argument("--minimos")
    ap.add_argument("--fecha")
    a = ap.parse_args()
    hoy = dt.date.fromisoformat(a.fecha) if a.fecha else dt.date.today()
    wf = openpyxl.load_workbook(a.fincas, data_only=True, read_only=True)
    wv = openpyxl.load_workbook(a.ventas, data_only=True, read_only=True)
    avisos = []
    for nombre, wb, hojas in (("Stock Fincas", wf, ["Tabla principal", "Stock", "Presupuestos"]),
                              ("ventas FG Agro", wv, ["Ventas", "Stock"])):
        for h in hojas:
            if h not in wb.sheetnames:
                avisos.append(f"Falta la pestaña «{h}» en {nombre}")
    if avisos:
        json.dump({"fecha": hoy.isoformat(), "avisos": avisos}, sys.stdout, ensure_ascii=False, indent=1)
        sys.exit(1)
    res = {
        "version": "1.0",
        "fecha": hoy.isoformat(),
        "stock_bajo": stock_bajo(wf, wv, hoy, leer_minimos(a.minimos)),
        "fincas": fincas(wf, hoy),
        "avisos": avisos,
    }
    json.dump(res, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
