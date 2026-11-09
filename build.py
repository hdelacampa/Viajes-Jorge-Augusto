#!/usr/bin/env python3
"""Regenera viajes_jorge_augusto.html desde el export xlsx de la planilla "Viajes Jorge y Augusto".

Uso: python3 build.py [ruta.xlsx] [--force]
Salida (stdout, una línea): UNCHANGED ... | CHANGED ... | ERROR ...
Códigos: 0 = UNCHANGED o CHANGED, 2 = falló el auto-chequeo, 3 = error de datos/entrada.
Sin dependencias externas salvo openpyxl. Determinístico: mismos datos -> mismo HTML byte a byte.
"""
import sys, os, json, re, hashlib, datetime, urllib.request, urllib.parse

BUILD_VERSION = "4"  # subir si cambia la plantilla o las reglas de normalización
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "source.xlsx")
OUT = os.path.join(HERE, "viajes_jorge_augusto.html")
STATE = os.path.join(HERE, "state.json")
GEOCACHE = os.path.join(HERE, "geocache.json")
ASSETS = os.path.join(HERE, "assets")

MODES = ["Avión", "Crucero", "Tren", "Auto", "Bus"]
MODE_ALIASES = {"avion": 0, "avión": 0, "aereo": 0, "aéreo": 0, "vuelo": 0, "crucero": 1, "barco": 1, "ferry": 1,
                "tren": 2, "auto": 3, "coche": 3, "remis": 3, "bus": 4, "omnibus": 4, "ómnibus": 4, "micro": 4}
LINE_FIX = {"hidrofoli": "Hidrofoil", "costa luminoa": "Costa Luminosa", "gran celebration": "Grand Celebration",
            "festivaleuropeanstars": "Festival European Stars", "queen elizabeth ii": "Queen Elizabeth II",
            "msc musica": "MSC Musica", "msc  musica": "MSC Musica", "osterreichische": "Österreichische (ÖBB)",
            "supermirafiori": "Super Mirafiori", "song of america": "Song of America"}
DISP = {'London':'Londres','Paris':'París','Vienna':'Viena','Lisbon':'Lisboa','New York':'Nueva York','Napoli':'Nápoles','Ciudad de la Costa':'Montevideo','Puerto Iguazu':'Puerto Iguazú','Mexico City':'Ciudad de México','Amsterdam':'Ámsterdam','Maiquetía':'Caracas','Victoria Seychelles':'Victoria (Seychelles)','El Pireo':'Atenas (El Pireo)','Pireo (Atenas)':'Atenas (El Pireo)','Zaventem':'Bruselas','Beijing':'Pekín','Shanghai (Pudong)':'Shanghái','Geneva':'Ginebra','Frankfurt am Main':'Fráncfort','Ferno (VA)':'Milán (Malpensa)','Alejandria':'Alejandría','Munich':'Múnich','Manila (Pasay)':'Manila','Ho Chi Minh City':'Ciudad Ho Chi Minh','Tocumen':'Ciudad de Panamá','Venezia (VE)':'Venecia','Istanbul':'Estambul','Havana':'La Habana','Cordoba':'Córdoba','Los Angeles':'Los Ángeles','La Goulette':'Túnez','Tanger':'Tánger','La Coruna':'La Coruña','San José (Alajuela)':'San José (Costa Rica)','Buenos Aires (Ezeiza)':'Buenos Aires','Roma (Civitavecchia)':'Roma','Muara (Brunei)':'Muara (Brunéi)','Muara (Bandar Seri Begawan)':'Muara (Brunéi)','Santorini (Fira)':'Santorini','Perth (Fremantle)':'Perth','Goa (Mormugao)':'Goa','Hanga Roa (Isla de Pascua)':'Isla de Pascua','Dallas-Fort Worth':'Dallas','Ibiza (Eivissa)':'Ibiza','Funchal (Madeira)':'Funchal','Arrecife (Lanzarote)':'Lanzarote','Pointe-a-Pitre':'Pointe-à-Pitre'}
CN = {"AE":"Emiratos Árabes Unidos","AG":"Antigua y Barbuda","AR":"Argentina","AS":"Samoa Americana","AT":"Austria","AU":"Australia","AW":"Aruba","BE":"Bélgica","BN":"Brunéi","BO":"Bolivia","BR":"Brasil","BS":"Bahamas","CA":"Canadá","CH":"Suiza","CL":"Chile","CN":"China","CO":"Colombia","CR":"Costa Rica","CU":"Cuba","CW":"Curazao","CY":"Chipre","CZ":"República Checa","DE":"Alemania","DK":"Dinamarca","DO":"República Dominicana","EC":"Ecuador","EG":"Egipto","ES":"España","FI":"Finlandia","FR":"Francia","GB":"Reino Unido","GP":"Guadalupe","GR":"Grecia","HK":"Hong Kong","HR":"Croacia","HU":"Hungría","ID":"Indonesia","IE":"Irlanda","IL":"Israel","IN":"India","IS":"Islandia","IT":"Italia","JM":"Jamaica","JO":"Jordania","JP":"Japón","KE":"Kenia","KR":"Corea del Sur","LK":"Sri Lanka","MA":"Marruecos","MC":"Mónaco","MT":"Malta","MX":"México","MY":"Malasia","NL":"Países Bajos","NO":"Noruega","NZ":"Nueva Zelanda","OM":"Omán","PA":"Panamá","PE":"Perú","PF":"Polinesia Francesa","PH":"Filipinas","PL":"Polonia","PR":"Puerto Rico","PT":"Portugal","PY":"Paraguay","QA":"Catar","RU":"Rusia","SC":"Seychelles","SE":"Suecia","SG":"Singapur","SX":"Sint Maarten","SY":"Siria","TC":"Islas Turcas y Caicos","TH":"Tailandia","TN":"Túnez","TR":"Turquía","US":"Estados Unidos","UY":"Uruguay","VE":"Venezuela","VG":"Islas Vírgenes Británicas","VI":"Islas Vírgenes (EE. UU.)","VN":"Vietnam","ZA":"Sudáfrica"}


def die(code, msg):
    print(msg)
    sys.exit(code)


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def leg_days(cell, cruise):
    """Días de la planilla. Fuera de crucero se conserva la lectura histórica: una fecha no numérica vale 0.

    En crucero, Excel a veces guardó el número de días como fecha (2 → 1900-01-02). Ahí el número de
    serie es el valor real; sin eso el total de cruceros pierde esos días.
    """
    if cruise and isinstance(cell, (datetime.datetime, datetime.date, datetime.time)):
        from openpyxl.utils.datetime import to_excel
        return int(to_excel(cell))
    return int(num(cell) or 0)


def norm_boat(s):
    """Identidad del barco, alineada con la fórmula de Resumen y con los alias conocidos.

    TRIM + minúsculas, espacios internos repetidos colapsados (MSC Musica) y
    sustituciones exactas: hidrofoli, gran celebration, costa luminoa.
    No es un reemplazo por subcadena.
    """
    s = " ".join(str(s or "").strip().lower().split())
    return {
        "hidrofoli": "hidrofoil",
        "gran celebration": "grand celebration",
        "costa luminoa": "costa luminosa",
    }.get(s, s)


def _ymd(s):
    if not s or len(str(s)) < 8:
        return None
    s = str(s)
    return datetime.date(int(s[0:4]), int(s[4:6]), int(s[6:8]))


def group_cruise_legs(legs):
    """Agrupa tramos de crucero en viajes. Un viaje nuevo empieza si cambia el barco
    normalizado, si la salida es más de 7 días después del regreso anterior, si el
    comentario empieza con «embarque» o si el tramo anterior empieza con «desembarque».

    El mismo barco en otro año queda en otro viaje. Ida y vuelta dentro de la ventana
    (Hidrofoil Tarifa–Tánger, 16/4/1981 y 19/4/1981) es un solo viaje.
    `legs` es la lista ya normalizada: [fecha, medio, línea, origen, destino, millas,
    días, hotel, comentario, regreso, fila].
    """
    items = [(i, leg) for i, leg in enumerate(legs) if leg[1] == 1]
    items.sort(key=lambda it: (it[1][0], int(it[1][10])))
    groups, prev = [], None
    for i, leg in items:
        boat = norm_boat(leg[2])
        dep = _ymd(leg[0])
        fin = _ymd(leg[9]) or dep
        notes = str(leg[8] or "").strip().lower()
        start = prev is None
        if prev is not None:
            pleg = prev[1]
            pfin = _ymd(pleg[9]) or _ymd(pleg[0])
            pnotes = str(pleg[8] or "").strip().lower()
            start = (
                boat != norm_boat(pleg[2])
                or dep > pfin + datetime.timedelta(days=7)
                or re.match(r"^embarque\b", notes) is not None
                or re.match(r"^desembarque\b", pnotes) is not None
            )
        if start:
            groups.append({
                "boat": leg[2] or boat,
                "date": leg[0],
                "end": leg[9] or leg[0],
                "mi": int(leg[5] or 0),
                "days": int(leg[6] or 0),
                "o": leg[3],
                "d": leg[4],
                "i": [i],
                "rows": [int(leg[10])],
            })
        else:
            g = groups[-1]
            g["i"].append(i)
            g["rows"].append(int(leg[10]))
            g["mi"] += int(leg[5] or 0)
            g["days"] += int(leg[6] or 0)
            g["d"] = leg[4]
            end = leg[9] or leg[0]
            if end > g["end"]:
                g["end"] = end
        prev = (i, leg)
    return groups


def cell(v):
    if v == "" or v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def line_name(s):
    s = (s or "").strip() if isinstance(s, str) else ("" if s is None else str(s).strip())
    k = s.lower()
    if k in LINE_FIX:
        return LINE_FIX[k]
    return s.upper() if len(s) <= 3 else s.title()


def clean_comment(c):
    return re.sub(r"\s*\[[A-Z]\d+:[^\]]*\]", "", str(c)).strip() if c else ""


def nominatim(code, iso_hint):
    q = f"{code.upper()} airport" if len(code) == 3 else code
    params = {"format": "json", "limit": "1", "addressdetails": "1", "accept-language": "es", "q": q}
    if iso_hint:
        params["countrycodes"] = iso_hint.lower()
    try:
        req = urllib.request.Request("https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(params),
                                     headers={"User-Agent": "viajes-jorge-augusto-refresh/1.0"})
        d = json.load(urllib.request.urlopen(req, timeout=20))
        if d:
            a = d[0].get("address", {})
            name = a.get("city") or a.get("town") or a.get("village") or d[0].get("name") or code.title()
            return name, (a.get("country_code") or iso_hint or "").upper(), float(d[0]["lat"]), float(d[0]["lon"])
    except Exception:
        pass
    return None


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv
    src = args[0] if args else SRC
    if not os.path.exists(src):
        die(3, f"ERROR no existe {src}")
    import openpyxl
    wb = openpyxl.load_workbook(src, data_only=True)
    if "Viajes" not in wb.sheetnames:
        die(3, "ERROR falta la pestaña Viajes")
    ws = wb["Viajes"]
    allrows = list(ws.iter_rows(values_only=True))
    hdr = next((i for i, r in enumerate(allrows) if len(r) > 1 and r[1] == "Salida"), None)
    if hdr is None:
        die(3, "ERROR no encuentro la fila de encabezados (Salida) en Viajes")
    rows = [(hdr + 2 + k, r) for k, r in enumerate(allrows[hdr + 1:]) if any(c not in (None, "") for c in r)]

    lookup = {}
    for tab, start in (("Aeropuertos", 3), ("Puertos crucero", 2)):
        if tab in wb.sheetnames:
            for r in wb[tab].iter_rows(min_row=start, values_only=True):
                if r and r[0] and num(r[4]) is not None and num(r[5]) is not None:
                    lookup.setdefault(str(r[0]).strip().lower(), [str(r[1]), str(r[2] or ""), str(r[3] or ""), float(r[4]), float(r[5])])
    iso_cont = {}
    for v in lookup.values():
        if v[1] and v[2]:
            iso_cont.setdefault(v[1], v[2])

    # hash of the relevant sheet data
    sig_rows = [json.dumps([cell(c) for c in r[1:23]], ensure_ascii=False, default=str) for _, r in rows]
    h = hashlib.sha256()
    h.update(BUILD_VERSION.encode())
    for s in sig_rows:
        h.update(s.encode() + b"\n")
    h.update(json.dumps(sorted(lookup.items()), ensure_ascii=False).encode())
    for fn in ("app.js", "style.css"):
        h.update(open(os.path.join(ASSETS, fn), "rb").read())
    digest = h.hexdigest()
    state = json.load(open(STATE)) if os.path.exists(STATE) else {}
    if not force and state.get("hash") == digest and os.path.exists(OUT):
        die(0, f"UNCHANGED legs={state.get('legs')} miles={state.get('miles')} hash={digest[:12]}")

    geocache = json.load(open(GEOCACHE)) if os.path.exists(GEOCACHE) else {}
    geocache_dirty = False
    places, plist, geocoded_used, unmapped = {}, [], set(), {}

    def place(code, la, lo, iso_col, cont_col, rn):
        nonlocal geocache_dirty
        if code is None or str(code).strip() == "":
            return -1
        c = str(code).strip().lower()
        info = lookup.get(c)
        from_cache = False
        if not info and c in geocache:
            info = geocache[c]
            from_cache = info is not None
        if not info and c not in geocache:
            if num(la) is not None and num(lo) is not None:
                info = [c.title(), str(iso_col or "").upper(), str(cont_col or ""), float(la), float(lo)]
            else:
                g = nominatim(c, str(iso_col or ""))
                if g:
                    name, iso, lat, lon = g
                    info = [name, iso, iso_cont.get(iso, str(cont_col or "")), lat, lon]
                    from_cache = True
                geocache[c] = info
                geocache_dirty = True
        if not info:
            unmapped.setdefault(c, []).append(rn)
            return -1
        if from_cache:
            geocoded_used.add(info[0])
        name = DISP.get(info[0], info[0])
        if num(la) is None or num(lo) is None:
            la, lo = info[3], info[4]
        if name not in places:
            places[name] = len(plist)
            plist.append([name, info[1], info[2], round(float(la), 3), round(float(lo), 3)])
        return places[name]

    legs, no_mode, no_route, skipped, bad_modes = [], [], [], [], {}
    raw_miles = 0.0
    for rn, r in rows:
        d = r[1]
        if not isinstance(d, (datetime.datetime, datetime.date)):
            skipped.append(rn)
            continue
        raw_miles += num(r[10]) or 0
        mode = r[2]
        if mode is None or str(mode).strip() == "":
            mi_mode = 0
            no_mode.append(rn)
        else:
            key = str(mode).strip().lower()
            if key not in MODE_ALIASES:
                bad_modes.setdefault(str(mode), []).append(rn)
                continue
            mi_mode = MODE_ALIASES[key]
        o = place(r[4], r[15], r[16], r[5], r[6], rn)
        de = place(r[7], r[17], r[18], r[8], r[9], rn)
        if o < 0 or de < 0:
            no_route.append(rn)
        rg = r[13].strftime("%Y%m%d") if hasattr(r[13], "strftime") else ""
        hotel = " / ".join(str(x).strip() for x in (r[11], r[12]) if x and str(x).strip())
        legs.append([d.strftime("%Y%m%d"), mi_mode, line_name(r[3]), o, de, int(num(r[10]) or 0), leg_days(r[14], mi_mode == 1),
                     hotel, clean_comment(r[19]), rg, rn])
    if bad_modes:
        die(3, "ERROR medios de transporte desconocidos: " + "; ".join(f"{k} (filas {v})" for k, v in bad_modes.items()))
    if geocache_dirty:
        json.dump(geocache, open(GEOCACHE, "w"), ensure_ascii=False, indent=1, sort_keys=True)
    if not legs:
        die(3, "ERROR no hay tramos con fecha")

    total_mi = sum(l[5] for l in legs)
    years = [int(l[0][:4]) for l in legs]
    fmt = lambda n: f"{n:,}".replace(",", ".")
    notes = [{"title": "Fuente", "detail": f"Planilla \"Viajes Jorge y Augusto\", pestaña Viajes: {len(legs)} tramos, {fmt(total_mi)} millas. Cada fila es un tramo; las millas son las de la planilla, sin recalcular."}]
    if geocoded_used:
        g = sorted(geocoded_used, key=lambda s: s.lower())
        notes.append({"title": "Ubicaciones", "detail": f"Coordenadas de la planilla y de sus pestañas Aeropuertos y Puertos crucero. {len(g)} lugares sin coordenadas se ubicaron con OpenStreetMap: {', '.join(g)}."})
    if no_route:
        notes.append({"title": "Sin ruta en el mapa", "detail": f"{len(no_route)} tramos no tienen origen o destino ubicable (filas {', '.join(map(str, no_route))}): suman en los totales pero no se dibujan."})
    if unmapped:
        notes.append({"title": "Lugares sin ubicar", "detail": ", ".join(f"{k} (filas {', '.join(map(str, v))})" for k, v in sorted(unmapped.items()))})
    adj = "Se unificaron mayúsculas y nombres de barcos (por ejemplo Gran/Grand Celebration)."
    if no_mode:
        adj = f"{len(no_mode)} tramos sin medio cargado (filas {', '.join(map(str, no_mode))}) se cuentan como avión. " + adj
    notes.append({"title": "Ajustes", "detail": adj})
    cruises = group_cruise_legs(legs)
    cruise_legs = [leg for leg in legs if leg[1] == 1]
    cruise_mi = sum(g["mi"] for g in cruises)
    cruise_days = sum(g["days"] for g in cruises)
    if sum(len(g["i"]) for g in cruises) != len(cruise_legs) or cruise_mi != sum(leg[5] for leg in cruise_legs) or cruise_days != sum(leg[6] for leg in cruise_legs):
        die(2, f"ERROR agrupación de cruceros: viajes={len(cruises)} tramos={len(cruise_legs)} millas={cruise_mi} días={cruise_days}")
    notes.append({"title": "Cruceros", "detail": f"{len(cruises)} viajes y {len(cruise_legs)} tramos, {fmt(cruise_mi)} millas y {cruise_days} días. Cada crucero es un viaje: se suman los tramos del mismo barco (también con errores de nombre conocidos) hasta un embarque, el desembarque del tramo anterior o más de 7 días desde el regreso. El mismo barco en otro año es otro viaje."})

    data = {"p": plist, "t": legs, "cruises": cruises, "cn": CN, "meta": {"y0": min(years), "y1": max(years), "notes": notes}}
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    css = open(os.path.join(ASSETS, "style.css"), encoding="utf-8").read()
    js = open(os.path.join(ASSETS, "app.js"), encoding="utf-8").read()
    if "groupCruiseLegs" not in js or "$n.cruises" not in js:
        die(3, "ERROR assets/app.js no agrupa cruceros: faltan groupCruiseLegs y $n.cruises. No se publica el HTML viejo.")
    html = ("<!DOCTYPE html>\n<html lang=\"es\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n"
            "<title>Viajes de Jorge y Augusto</title>\n<style>" + css + "</style></head>\n<body>"
            "<div id=\"nojs\" style=\"max-width:640px;margin:40px auto;padding:20px;border:1px solid #d5dbe2;border-radius:14px;font-family:system-ui,sans-serif;background:#fff\">"
            "<h1 style=\"font-size:22px;margin:0 0 8px\">Viajes de Jorge y Augusto</h1>"
            "<p style=\"margin:0\">Este visor no puede mostrar la página interactiva. Descargá el archivo y abrilo con un navegador "
            "(Chrome, Safari, Edge o Firefox). En el celular: Descargar, y después abrirlo con Chrome o Safari.</p></div>\n"
            "<div id=\"root\"></div>\n"
            "<script id=\"viajes-data\">window.__VIAJES__=" + payload + ";</script>\n<script>" + js + "</script>\n</body></html>\n")

    # self-check: re-read what will be published
    m = re.search(r'<script id="viajes-data">window.__VIAJES__=(.*?);</script>', html, re.S)
    back = json.loads(m.group(1).replace("<\\/", "</")) if m else None
    back_groups = group_cruise_legs(back["t"]) if back else []
    ok = (back is not None and len(back["t"]) == len(legs) == len(rows) - len(skipped)
          and sum(l[5] for l in back["t"]) == total_mi == int(round(raw_miles))
          and all(-1 <= l[3] < len(back["p"]) and -1 <= l[4] < len(back["p"]) for l in back["t"])
          and [g["rows"] for g in back.get("cruises", [])] == [g["rows"] for g in back_groups]
          and sum(g["mi"] for g in back.get("cruises", [])) == cruise_mi
          and sum(g["days"] for g in back.get("cruises", [])) == cruise_days)
    if not ok:
        die(2, f"ERROR auto-chequeo: html_legs={len(back['t']) if back else '?'} legs={len(legs)} filas={len(rows) - len(skipped)} millas={total_mi} millas_planilla={raw_miles} cruceros={len(cruises)} millas_crucero={cruise_mi} dias_crucero={cruise_days}")
    targets = [OUT]
    index_html = os.path.join(HERE, "index.html")
    if os.path.abspath(index_html) != os.path.abspath(OUT) and os.path.exists(index_html):
        targets.append(index_html)
    for path in targets:
        tmp = path + ".tmp"
        open(tmp, "w", encoding="utf-8").write(html)
        os.replace(tmp, path)

    prev = set(state.get("rows", []))
    new_rows = sum(1 for s in sig_rows if s not in prev) if prev else 0
    json.dump({"hash": digest, "legs": len(legs), "miles": total_mi, "rows": sig_rows}, open(STATE, "w"), ensure_ascii=False)
    um = ",".join(sorted(unmapped)) or "none"
    extra = f" skipped_rows={skipped}" if skipped else ""
    print(f"CHANGED legs={len(legs)} miles={total_mi} cruises={len(cruises)} cruise_legs={len(cruise_legs)} cruise_miles={cruise_mi} cruise_days={cruise_days} new_rows={new_rows} unmapped={um} no_route_rows={len(no_route)}{extra} hash={digest[:12]}")
    sys.exit(0)


if __name__ == "__main__":
    main()
