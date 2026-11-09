#!/usr/bin/env python3
"""Nada que no sea un crucero cambia respecto del sitio publicado en main.

La única diferencia admitida en los tramos es el día de crucero que Excel guardó
como fecha (fila 355: 0 → 2), necesaria para los 142 días. El vuelo de la fila 165
y el resto de avión, tren, auto y bus quedan iguales, campo por campo.
"""
import json, os, re, subprocess, sys
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# Sitio publicado antes de agrupar cruceros.
BASELINE = "8c6e4d4b6f12024875941864405fa6cfe1992b9f"
FIELDS = ("fecha", "medio", "linea", "origen", "destino", "millas", "dias", "hotel", "comentario", "regreso", "fila")
MODES = ("Avión", "Crucero", "Tren", "Auto", "Bus")


def load_raw(raw):
    blob = re.search(r'window\.__VIAJES__=(.*?);</script>', raw, re.S).group(1).replace("<\\/", "</")
    return json.loads(blob)


def load_baseline():
    raw = subprocess.check_output(["git", "show", f"{BASELINE}:index.html"], cwd=ROOT, text=True)
    return load_raw(raw)


def load_current():
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "index.html")
    return load_raw(open(path, encoding="utf-8").read()), path


def stats(legs):
    """Estadísticas de tramos que no son crucero."""
    out = {
        "tramos": 0,
        "millas": 0,
        "dias": 0,
        "por_medio": {},
        "por_anio": {},
        "por_destino": {},
    }
    for mode in range(5):
        if mode == 1:
            continue
        out["por_medio"][MODES[mode]] = {"tramos": 0, "millas": 0, "dias": 0}
    years = defaultdict(lambda: {"tramos": 0, "millas": 0, "dias": 0})
    dest = defaultdict(lambda: {"llegadas": 0, "millas": 0})
    for leg in legs:
        if leg[1] == 1:
            continue
        out["tramos"] += 1
        out["millas"] += leg[5]
        out["dias"] += leg[6]
        bucket = out["por_medio"][MODES[leg[1]]]
        bucket["tramos"] += 1
        bucket["millas"] += leg[5]
        bucket["dias"] += leg[6]
        y = years[leg[0][:4]]
        y["tramos"] += 1
        y["millas"] += leg[5]
        y["dias"] += leg[6]
        if leg[4] >= 0 and leg[3] != leg[4]:
            d = dest[leg[4]]
            d["llegadas"] += 1
            d["millas"] += leg[5]
    out["por_anio"] = dict(sorted(years.items()))
    out["por_destino"] = {str(k): dest[k] for k in sorted(dest)}
    return out


def main():
    old = load_baseline()
    new, path = load_current()
    assert old["p"] == new["p"], "cambiaron los lugares"
    assert old["cn"] == new["cn"], "cambiaron los países"
    assert old["meta"]["notes"] == new["meta"]["notes"][:len(old["meta"]["notes"])], "cambiaron las notas que no son de crucero"
    assert len(new["meta"]["notes"]) == len(old["meta"]["notes"]) + 1
    assert new["meta"]["notes"][-1]["title"] == "Cruceros"

    old_by = {leg[10]: leg for leg in old["t"]}
    new_by = {leg[10]: leg for leg in new["t"]}
    assert set(old_by) == set(new_by), "cambiaron las filas publicadas"
    assert [leg[10] for leg in old["t"]] == [leg[10] for leg in new["t"]], "cambió el orden de los tramos"

    mismatches = []
    cruise_day_diffs = []
    for row, before in old_by.items():
        after = new_by[row]
        if before[1] != 1:
            if before != after:
                fields = [FIELDS[i] for i in range(len(before)) if before[i] != after[i]]
                mismatches.append(f"fila {row} medio {MODES[before[1]]}: {fields} antes={before} ahora={after}")
        else:
            for i in range(len(before)):
                if before[i] != after[i]:
                    if i != 6:
                        mismatches.append(f"crucero fila {row} cambió {FIELDS[i]}: {before[i]!r} → {after[i]!r}")
                    else:
                        cruise_day_diffs.append((row, before[6], after[6]))

    assert not mismatches, "registros no-crucero distintos:\n" + "\n".join(mismatches)
    assert cruise_day_diffs == [(355, 0, 2)], cruise_day_diffs

    flight = new_by[165]
    assert flight[1] == 0 and flight[6] == 0 and flight[5] == 6268, flight

    assert stats(old["t"]) == stats(new["t"])
    non = stats(new["t"])
    assert non["tramos"] == 322
    assert non["millas"] == 1030376 - 43653
    assert non["dias"] == sum(leg[6] for leg in old["t"] if leg[1] != 1)
    assert non["por_medio"]["Avión"]["tramos"] == 292
    assert non["por_medio"]["Avión"]["dias"] == sum(leg[6] for leg in old["t"] if leg[1] == 0)

    cruise_days = sum(leg[6] for leg in new["t"] if leg[1] == 1)
    cruise_miles = sum(leg[5] for leg in new["t"] if leg[1] == 1)
    assert cruise_days == 142 and cruise_miles == 43653
    assert sum(g["days"] for g in new["cruises"]) == 142
    assert "groupCruiseLegs" in open(os.path.join(ROOT, "assets", "app.js"), encoding="utf-8").read()
    assert "$n.cruises" in open(os.path.join(ROOT, "assets", "app.js"), encoding="utf-8").read()

    print(f"OK no-crucero idéntico a {BASELINE[:12]} "
          f"({non['tramos']} tramos, {non['millas']} millas, {non['dias']} días); "
          f"crucero 142 días; única diferencia de tramo: fila 355 días 0→2 ({path})")


if __name__ == "__main__":
    main()
