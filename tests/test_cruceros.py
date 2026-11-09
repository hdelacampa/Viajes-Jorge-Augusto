#!/usr/bin/env python3
"""Los 14 cruceros auditados se conservan en el HTML publicado y en el generador.

No congela el futuro: build.py solo exige que millas y días se conserven al agrupar.
Este test fija la foto actual (filas, 43653 millas, 142 días).
"""
import json, os, re, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from build import group_cruise_legs  # noqa: E402

EXPECTED = [
    {"boat": "hidrofoil", "rows": [23, 24], "miles": 44, "days": 3, "date": "19810416"},
    {"boat": "tirrenia", "rows": [34], "miles": 198, "days": 0, "date": "19820822"},
    {"boat": "paquet", "rows": [82], "miles": 0, "days": 10, "date": "19870127"},
    {"boat": "queen elizabeth ii", "rows": [141], "miles": 3442, "days": 5, "date": "19930417"},
    {"boat": "song of america", "rows": [155, 156, 157, 158, 159, 160, 161, 162], "miles": 3306, "days": 8, "date": "19941230"},
    {"boat": "festival flamenco", "rows": [230, 231], "miles": 0, "days": 0, "date": "20010120"},
    {"boat": "festival european stars", "rows": [246, 247, 248, 249, 250, 251], "miles": 894, "days": 0, "date": "20031128"},
    {"boat": "costa atlantica", "rows": [267, 268, 269, 270, 271, 272, 273, 274], "miles": 3343, "days": 10, "date": "20051021"},
    {"boat": "costa allegra", "rows": [280, 281, 282, 283, 284, 285, 286, 287, 288], "miles": 0, "days": 1, "date": "20070519"},
    {"boat": "grand celebration", "rows": [309, 310, 311, 312, 313, 314, 315, 316, 317, 319], "miles": 4947, "days": 10, "date": "20120311"},
    {"boat": "msc musica", "rows": [325, 326, 327, 328, 329, 330], "miles": 1766, "days": 7, "date": "20130122"},
    {"boat": "grand celebration", "rows": [346, 347, 348, 349, 350, 351, 352, 353, 354, 355], "miles": 2803, "days": 11, "date": "20140623"},
    {"boat": "costa luminosa", "rows": list(range(383, 411)), "miles": 20360, "days": 69, "date": "20160215"},
    {"boat": "carnival splendor", "rows": [415, 416, 417, 418, 419], "miles": 2550, "days": 8, "date": "20160813"},
]


def load_html(path):
    html = open(path, encoding="utf-8").read()
    raw = re.search(r'window\.__VIAJES__=(.*?);</script>', html, re.S).group(1).replace("<\\/", "</")
    return json.loads(raw)


def norm(s):
    return " ".join(str(s or "").strip().lower().split())


def check(groups, label):
    assert len(groups) == 14, f"{label}: {len(groups)} viajes, se esperaban 14"
    assert sum(g["mi"] if "mi" in g else g["miles"] for g in groups) == 43653, label
    assert sum(g["days"] for g in groups) == 142, label
    for got, exp in zip(groups, EXPECTED):
        rows = got["rows"]
        miles = got["mi"] if "mi" in got else got["miles"]
        assert rows == exp["rows"], f"{label} {exp['boat']}: {rows} != {exp['rows']}"
        assert miles == exp["miles"], f"{label} {exp['boat']} millas {miles}"
        assert got["days"] == exp["days"], f"{label} {exp['boat']} días {got['days']}"
        assert norm(got["boat"]) == exp["boat"], f"{label} nombre {got['boat']}"
        assert got["date"] == exp["date"], f"{label} fecha {got['date']}"
    rows_2012 = set(groups[9]["rows"])
    rows_2014 = set(groups[11]["rows"])
    assert rows_2012.isdisjoint(rows_2014)
    assert 349 in rows_2014 and 349 not in rows_2012
    assert 330 in set(groups[10]["rows"]) and len(groups[10]["rows"]) == 6
    assert 406 in set(groups[12]["rows"]) and len([g for g in groups if "luminosa" in norm(g["boat"])]) == 1
    assert groups[0]["rows"] == [23, 24]


def js_groups(legs):
    app = open(os.path.join(ROOT, "assets", "app.js"), encoding="utf-8").read()
    start = app.find("function normBoat(")
    gstart = app.find("function groupCruiseLegs(")
    assert start >= 0 and gstart > start
    i = app.find("{", gstart)
    depth = 0
    for j in range(i, len(app)):
        if app[j] == "{":
            depth += 1
        elif app[j] == "}":
            depth -= 1
            if depth == 0:
                src = app[start:j + 1]
                break
    else:
        raise AssertionError("no se pudo extraer groupCruiseLegs")
    script = src + "\nprocess.stdout.write(JSON.stringify(groupCruiseLegs(JSON.parse(process.argv[1]))));"
    out = subprocess.check_output(["node", "-e", script, json.dumps(legs)], text=True)
    return json.loads(out)


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "index.html")
    data = load_html(path)
    assert len(data["t"]) == 419
    assert sum(leg[5] for leg in data["t"]) == 1030376
    assert sum(1 for leg in data["t"] if leg[1] == 0) == 292
    assert sum(1 for leg in data["t"] if leg[1] == 1) == 97
    check(data["cruises"], "html")
    check(group_cruise_legs(data["t"]), "python")
    js = js_groups(data["t"])
    check(js, "javascript")
    assert [g["rows"] for g in js] == [g["rows"] for g in data["cruises"]]
    # La fila 355 conserva los 2 días que Excel guardó como fecha.
    row355 = next(leg for leg in data["t"] if leg[10] == 355)
    assert row355[6] == 2 and row355[5] == 345
    print(f"OK 14 cruceros, 97 tramos, 43653 millas, 142 días ({path})")


if __name__ == "__main__":
    main()
