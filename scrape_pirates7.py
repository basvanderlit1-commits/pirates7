"""Haalt alle data van Pirates 7 (DBMN, via feeds.teambeheer.nl) op en schrijft
'pirates 7 resultaten.xlsx' en het dashboard in site/ (GitHub Pages).
Seizoen, team-id en divisie worden automatisch opgezocht; alleen TEAM moet kloppen met de naam op teambeheer."""
import hashlib
import json
import os
import re
import shutil
from collections import defaultdict
from pathlib import Path
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from datetime import date, datetime
from zoneinfo import ZoneInfo
from requests.adapters import HTTPAdapter, Retry
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

B = "https://feeds.teambeheer.nl"
D, TEAM = 41, "Pirates 7"  # D = DBMN op teambeheer
OUT = "pirates 7 resultaten.xlsx"
SITE_URL = "https://basvanderlit1-commits.github.io/pirates7/"
NTFY_TOPIC = os.environ.get("NTFY_TOPIC", "pirates7-dbmn-uitslagen")  # abonneren in de gratis app ntfy
http = requests.Session()
http.mount("https://", HTTPAdapter(max_retries=Retry(total=4, backoff_factor=2, status_forcelist=[429, 500, 502, 503, 504])))


def get(path, full=False):
    r = http.get(B + path, timeout=30)
    r.raise_for_status()
    raw = r.content
    try:
        html = raw.decode("utf-8")
    except UnicodeDecodeError:  # sommige namen staan er in cp1252 in
        html = raw.decode("cp1252")
    soup = BeautifulSoup(html, "html.parser")
    # desktop-weergave; de pagina bevat dezelfde tabellen nogmaals voor tablet/mobiel
    if full:
        return soup
    return next((r for r in soup.select(".computer.only.row") if "mobile" not in r["class"]), soup)


def txt(el):
    return " ".join(el.get_text(" ").split()) if el else ""


def rows(table):
    return [[txt(td) for td in tr.find_all(["td", "th"])] for tr in table.find_all("tr")]


def num(v):
    v = str(v).replace("%", "").strip()
    try:
        return int(v)
    except ValueError:
        try:
            return float(v)
        except ValueError:
            return v


def header_table(soup, first_col):
    for t in soup.find_all("table"):
        h = [txt(th) for th in t.find_all("th")]
        if h and h[0] == first_col:
            return t


# ---------- team opzoeken in het huidige seizoen ----------
link = get(f"/web/teams?d={D}", full=True).find("a", string=lambda t: t and t.strip() == TEAM)
if not link:
    raise SystemExit(f"Team '{TEAM}' niet gevonden op {B}/web/teams?d={D} - klopt de naam nog?")
TEAM_ID, S = re.search(r"t=(\d+)&s=([\d-]+)", link["href"]).groups()
team = get(f"/web/team?d={D}&t={TEAM_ID}&s={S}")
DIV = re.search(r"Divisie (\w+)", txt(team.find("a", href=re.compile(r"/web/stand")))).group(1)


def parse_team(page, name):
    """Wedstrijden, spelers en speellocatie van een teampagina (ons team of een tegenstander)."""
    ms = []
    for tr in header_table(page, "#").find("tbody").find_all("tr"):
        td = tr.find_all("td")
        if len(td) < 5:
            continue
        a = td[4].find("a")
        score = txt(a) if a else ""
        home, away = txt(td[2]), txt(td[3])
        res = ""
        if score:
            h, u = map(int, score.split("-"))
            mine, theirs = (h, u) if home == name else (u, h)
            res = "W" if mine > theirs else "V" if mine < theirs else "G"
        opp = (td[3] if home == name else td[2]).find("a")
        ms.append(dict(ronde=txt(td[0]), datum=txt(td[1]), thuis=home, uit=away, score=score,
                       uitslag=res, form=a["href"] if a else None,
                       tegen_id=re.search(r"t=(\d+)", opp["href"]).group(1) if opp else None,
                       vrij=not re.fullmatch(r"\d{2}-\d{2}-\d{4}", txt(td[1]))))  # bv. "Vrije week" in de beker
    players = []
    for tr in header_table(page, "Naam").find("tbody").find_all("tr"):
        td = tr.find_all("td")
        a = td[0].find("a")
        rol = txt(td[0].find("b"))
        players.append(dict(naam=txt(a), rol={"C": "Captain", "RC": "Reserve captain"}.get(rol, rol),
                            id=re.search(r"l=(\d+)", a["href"]).group(1), singles=num(txt(td[1])), winst=num(txt(td[2]))))
    loc = page.find(string="Locatie")
    return ms, players, clean_loc(loc.find_parent().find_next("p").stripped_strings) if loc else ""


def clean_loc(parts):
    """Adresregels zonder telefoonnummer, gescheiden door komma's."""
    return ", ".join(p.strip() for p in parts if p.strip() and not re.fullmatch(r"[\d\s\-+()]{6,}", p.strip()))


matches, spelers, locatie = parse_team(team, TEAM)

# ---------- wedstrijdformulieren ----------
games, bijz = [], []
for m in matches:
    if not m["form"]:
        continue
    f = get(m["form"])
    thuis = m["thuis"] == TEAM
    for tr in f.find("table").find("tbody").find_all("tr"):
        td = tr.find_all("td")
        sc = txt(td[-1]) if td else ""
        if len(td) < 4 or not re.fullmatch(r"\d+-\d+", sc):  # Totaal / niet gespeeld
            continue
        hp = [txt(a) for a in td[1].find_all("a")]
        up = [txt(a) for a in td[2].find_all("a")]
        h, u = map(int, sc.split("-"))
        mine, theirs = (h, u) if thuis else (u, h)
        games.append(dict(ronde=m["ronde"], datum=m["datum"], tegen=m["uit"] if thuis else m["thuis"],
                          onderdeel=txt(td[0]), wij=", ".join(hp if thuis else up) or "(team)",
                          zij=", ".join(up if thuis else hp) or "(team)", legs_wij=mine, legs_zij=theirs,
                          uitslag="W" if mine > theirs else "V"))
    loc = f.find(string="Locatie")
    m["locatie"] = clean_loc(" | ".join(loc.find_parent().find_next("a").parent.stripped_strings)
                             .split(" | Bijz.")[0].removeprefix("Locatie | ").split(" | ")) if loc else ""
    for item in f.select(".ui.list .item"):
        bijz.append(dict(ronde=m["ronde"], datum=m["datum"], speler=txt(item.select_one(".header")),
                         prestatie=txt(item.select_one(".content")).replace(txt(item.select_one(".header")), "", 1).strip()))
ours = {s["naam"] for s in spelers}
for b in bijz:
    b["team"] = TEAM if b["speler"] in ours else "tegenstander"

# ---------- per speler: statistiek uit wedstrijdformulieren + spelerpagina ----------
agg = defaultdict(lambda: defaultdict(int))
for g in games:
    o = g["onderdeel"]
    kind = "s" if o.startswith(("Single", "Singel")) else "k" if o.startswith("Koppel") else "rr" if o.startswith("RR") else None
    if not kind:
        continue
    for p in g["wij"].split(", "):
        a = agg[p]
        a[kind + "_gesp"] += 1
        a[kind + "_w"] += g["uitslag"] == "W"
        a[kind + "_lv"] += g["legs_wij"]
        a[kind + "_lt"] += g["legs_zij"]
for s in spelers:
    sp = get(f"/web/speler?d={D}&l={s['id']}&s={S}")
    stats = {txt(st.select_one(".label")): txt(st.select_one(".value")) for st in sp.select(".statistic")}
    s["site_singles_pct"] = next((v for k, v in stats.items() if "singles" in k), "")
    s["site_koppels_pct"] = next((v for k, v in stats.items() if "koppels" in k), "")

# ---------- stand, klassementen, bijzondere resultaten 4A ----------
stand_soup = get(f"/web/stand?d={D}&s={S}&div={DIV}", full=True)
stand = next(rows(t) for t in stand_soup.find_all("table") if "Wed" in [txt(th) for th in t.find_all("th")])

pk_single = rows(get(f"/web/scorelijst-pk/?d={D}&mt=1&s={S}&filter=P-{DIV}").find("table"))
pk_koppel = rows(get(f"/web/scorelijst-pk/?d={D}&mt=2&s={S}&filter=P-{DIV}").find("table"))
# punten per ronde -> onze positie na elke gespeelde ronde (benadering: telt alle gespeelde punten tot en met die ronde)
rs = next(rows(t) for t in stand_soup.find_all("table") if "Wed" not in [txt(th) for th in t.find_all("th")])
rcols = [i for i, h in enumerate(rs[0]) if h.isdigit()]
pts = {r[1]: [num(r[i]) if i < len(r) else "" for i in rcols] for r in rs[1:] if len(r) > 2}
cum, verloop = defaultdict(float), []
for k, i in enumerate(rcols):
    for t, p in pts.items():
        cum[t] += p[k] if isinstance(p[k], (int, float)) else 0
    if isinstance(pts.get(TEAM, [None] * len(rcols))[k], (int, float)):
        verloop.append(dict(ronde=int(rs[0][i]), pos=1 + sum(v > cum[TEAM] for t, v in cum.items() if t != TEAM)))

# tegenstanders: adres (route), vorm en beste spelers
tegenstanders = {}
for m in matches:
    naam = m["uit"] if m["thuis"] == TEAM else m["thuis"]
    if m["tegen_id"] and naam not in tegenstanders:
        o_ms, o_sp, o_loc = parse_team(get(f"/web/team?d={D}&t={m['tegen_id']}&s={S}"), naam)
        top = [p for p in o_sp if isinstance(p["singles"], int) and p["singles"] > 0 and isinstance(p["winst"], (int, float)) and p["winst"] > 0]
        tegenstanders[naam] = dict(locatie=o_loc, vorm=[x["uitslag"] for x in o_ms if x["uitslag"]][-5:],
                                   spelers=[dict(naam=p["naam"], singles=p["singles"], winst=p["winst"])
                                            for p in sorted(top, key=lambda p: (-p["winst"], -p["singles"]))[:3]])
for m in matches:  # locatie van nog te spelen wedstrijden: thuis bij ons, uit bij de tegenstander
    if not m.get("locatie") and not m["vrij"]:
        m["locatie"] = locatie if m["thuis"] == TEAM else tegenstanders.get(m["uit"] if m["thuis"] == TEAM else m["thuis"], {}).get("locatie", "")

bijz_lists = {}
for t, name in [(1, "180ers"), (2, "Hoogste finishes"), (3, "Snelste leg"), (4, "171ers")]:
    tab = get(f"/web/scorelijst-bijzres/?d={D}&t={t}&s={S}&filter=P-{DIV}").find("table")
    bijz_lists[name] = rows(tab) if tab else [["(geen data)"]]

# ---------- voorbereiden ----------
NOW = datetime.now(ZoneInfo("Europe/Amsterdam"))
for m in matches:
    thuis = m["thuis"] == TEAM
    m["tegen"], m["tu"] = (m["uit"], "Thuis") if thuis else (m["thuis"], "Uit")
    if m["score"]:
        h, u = m["score"].split("-")
        m["wijzij"] = f"{h} - {u}" if thuis else f"{u} - {h}"
role = lambda r: ", ".join(s["naam"] for s in spelers if s["rol"] == r) or "–"
ours_only = lambda tab, col: [r for r in tab[1:] if len(r) > col and r[col] == TEAM]
pk_count = lambda tab: sum(1 for r in tab[1:] if len(r) > 2)  # aantal spelers in het klassement
won = lambda r: round(num(r[3]) * num(r[6]) / 100)  # gewonnen partijen = gespeeld x winst%
kind = lambda o: "RR" if o.startswith("RR") else "Single" if o.startswith(("Single", "Singel")) else \
    "Koppel" if o.startswith("Koppel") else "Team"
site = Path("site")
site.mkdir(exist_ok=True)

# ---------- Excel: alleen de uitgelezen data (de presentatie zit in het dashboard) ----------
wb = Workbook()
wb.remove(wb.active)


def to_date(s):
    try:
        return datetime.strptime(s, "%d-%m-%Y").date()
    except (TypeError, ValueError):
        return s


def sheet(title, header, data):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in data:
        ws.append([num(v) if isinstance(v, str) else v for v in r])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, date):
                c.number_format = "dd-mm-yyyy"
    for i, col in enumerate(ws.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(50, max(len(str(c.value or "")) for c in col) + 2)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


sheet("Info", ["Veld", "Waarde"],
      [["Team", TEAM], ["Seizoen", S], ["Divisie", DIV], ["Speellocatie", locatie],
       ["Captain", role("Captain")], ["Reserve captain", role("Reserve captain")],
       ["Bijgewerkt", NOW.strftime("%d-%m-%Y %H:%M")], ["Bron", f"{B}/web/team?d={D}&t={TEAM_ID}&s={S}"]])
sheet("Stand", ["Positie", "Team", "Gespeeld", "Winst", "Verlies", "Punten", "Gemiddeld", "Strafpunten"],
      [r for r in stand[1:] if any(r)])
sheet("Programma", ["Ronde", "Datum", "Thuis", "Uit", "Tegenstander", "Thuis/Uit", "Punten wij", "Punten zij",
                    "Uitslag", "Vrije week", "Locatie", "Wedstrijdformulier"],
      [[m["ronde"], to_date(m["datum"]), m["thuis"], m["uit"], m["tegen"], m["tu"],
        int(m["wijzij"].split(" - ")[0]) if m["score"] else None, int(m["wijzij"].split(" - ")[1]) if m["score"] else None,
        m["uitslag"] or None, "ja" if m["vrij"] else None, m.get("locatie", "") or None,
        B + m["form"] if m["form"] else None] for m in matches])
sheet("Partijen", ["Ronde", "Datum", "Tegenstander", "Onderdeel", "Type", "Pirates 7", "Tegenstander(s)",
                   "Legs wij", "Legs zij", "Uitslag"],
      [[g["ronde"], to_date(g["datum"]), g["tegen"], g["onderdeel"], kind(g["onderdeel"]), g["wij"], g["zij"],
        g["legs_wij"], g["legs_zij"], g["uitslag"]] for g in games])
sheet("Spelers", ["Speler", "Rol",
                  "Singles gespeeld", "Singles gewonnen", "Singles legs voor", "Singles legs tegen",
                  "Koppels gespeeld", "Koppels gewonnen", "Koppels legs voor", "Koppels legs tegen",
                  "RR gespeeld", "RR gewonnen", "RR legs voor", "RR legs tegen",
                  "180's", "Finishes", "Winst % singles (site)", "Winst % koppels (site)"],
      [[s["naam"], s["rol"] or "Speler",
        *(agg[s["naam"]][k + x] for k in ("s", "k", "rr") for x in ("_gesp", "_w", "_lv", "_lt")),
        sum("180" in b["prestatie"] for b in bijz if b["speler"] == s["naam"]),
        ", ".join(b["prestatie"].replace(" finish", "") for b in bijz
                  if b["speler"] == s["naam"] and "finish" in b["prestatie"]) or None,
        s["site_singles_pct"] or None, s["site_koppels_pct"] or None] for s in spelers])
for label, tab in [("Singles", pk_single), ("Koppels", pk_koppel)]:
    sheet(f"Klassement {label.lower()}", ["Positie", "Spelers in klassement", "Speler", "Gespeeld", "Gewonnen", "Winst %"],
          [[r[0], pk_count(tab), r[1], r[3], won(r), r[6]] for r in ours_only(tab, 2)])
sheet("Bijzondere resultaten", ["Ronde", "Datum", "Speler", "Prestatie"],
      [[b["ronde"], to_date(b["datum"]), b["speler"], b["prestatie"]] for b in bijz if b["team"] == TEAM])
sheet("Klassering bijzonder", ["Lijst", "Positie", "Speler", "Aantal / waarde"],
      [[name, r[0], r[1], r[4]] for name, tab in bijz_lists.items() for r in ours_only(tab, 2)])

wb.save(site / "pirates7-resultaten.xlsx")  # download op de website
try:
    wb.save(OUT)
except PermissionError:
    print(f"LET OP: '{OUT}' staat open in Excel en is niet bijgewerkt. Sluit hem en draai opnieuw.")

# ---------- Dashboard (HTML) ----------
dash = dict(
    team=TEAM, div=DIV, seizoen=S, bijgewerkt=NOW.strftime("%d-%m-%Y %H:%M"), locatie=locatie,
    captain=role("Captain"), rc=role("Reserve captain"), bron=f"{B}/web/team?d={D}&t={TEAM_ID}&s={S}",
    stand=[dict(pos=num(r[0]), team=r[1], wed=num(r[2]), w=num(r[3]), v=num(r[4]), pnt=num(r[5]), gem=num(r[6]))
           for r in stand[1:] if any(r)],
    matches=[dict(ronde=m["ronde"], datum=m["datum"], tegen=m["tegen"], tu=m["tu"], uitslag=m["uitslag"], vrij=m["vrij"],
                  wij=int(m["wijzij"].split(" - ")[0]) if m["score"] else None,
                  zij=int(m["wijzij"].split(" - ")[1]) if m["score"] else None,
                  locatie=m.get("locatie", ""), form=B + m["form"] if m["form"] else None)
             for m in matches],
    games=[dict(ronde=g["ronde"], type=kind(g["onderdeel"]), onderdeel=g["onderdeel"], wij=g["wij"], zij=g["zij"],
                lw=g["legs_wij"], lz=g["legs_zij"], uitslag=g["uitslag"]) for g in games],
    spelers=[dict(naam=s["naam"], rol=s["rol"] or "Speler",
                  **{k: [agg[s["naam"]][k + "_gesp"], agg[s["naam"]][k + "_w"], agg[s["naam"]][k + "_lv"],
                         agg[s["naam"]][k + "_lt"]] for k in ("s", "k", "rr")},
                  n180=sum("180" in b["prestatie"] for b in bijz if b["speler"] == s["naam"]),
                  finishes=[b["prestatie"].replace(" finish", "") for b in bijz
                            if b["speler"] == s["naam"] and "finish" in b["prestatie"]]) for s in spelers],
    pk={label: [dict(pos=num(r[0]), naam=r[1], n=num(r[3]), w=won(r), pct=num(r[6]))
                for r in ours_only(tab, 2)] for label, tab in [("Singles", pk_single), ("Koppels", pk_koppel)]},
    pk_totaal={"Singles": pk_count(pk_single), "Koppels": pk_count(pk_koppel)},
    bijz=[dict(ronde=b["ronde"], datum=b["datum"], speler=b["speler"], prestatie=b["prestatie"])
          for b in bijz if b["team"] == TEAM],
    bijzrank=[dict(lijst=name, pos=num(r[0]), speler=r[1], waarde=num(r[4]))
              for name, tab in bijz_lists.items() for r in ours_only(tab, 2)],
    verloop=verloop, tegenstanders=tegenstanders, ntfy=NTFY_TOPIC,
)
(site / "data.json").write_text(json.dumps(dash, ensure_ascii=False), encoding="utf-8")  # vorige stand voor meldingen

# ---------- melding bij een nieuwe uitslag (ntfy.sh; alleen in GitHub Actions, met de vorige data.json) ----------
prev = Path("prev.json")
if os.environ.get("MELDINGEN") == "aan" and prev.exists():
    oud = {m["ronde"] for m in json.loads(prev.read_text(encoding="utf-8"))["matches"] if m["uitslag"]}
    me = next(r for r in stand[1:] if r[1] == TEAM)
    for m in matches:
        if m["uitslag"] and m["ronde"] not in oud:
            score = m["wijzij"].replace(" ", "")
            titel = {"W": f"{TEAM} wint {score} van {m['tegen']}", "V": f"{TEAM} verliest {score} van {m['tegen']}",
                     "G": f"{TEAM} speelt {score} gelijk tegen {m['tegen']}"}[m["uitslag"]]
            wat = "Beker" if m["ronde"].startswith("b") else "Ronde " + m["ronde"]
            tekst = f"{wat} · {m['datum']} · {m['tu'].lower()}\nStand: {me[0]}e in {DIV} met {me[5]} punten"
            try:  # een storing bij ntfy mag de update van het dashboard nooit tegenhouden
                http.post("https://ntfy.sh/", json=dict(topic=NTFY_TOPIC, title=titel, message=tekst, click=SITE_URL,
                                                      tags=["dart", "trophy"] if m["uitslag"] == "W" else ["dart"]),
                          timeout=30).raise_for_status()
                print("Melding verstuurd:", titel)
            except requests.RequestException as e:
                print("Melding mislukt:", e)

tpl = (Path(__file__).parent / "dashboard_template.html").read_text(encoding="utf-8")
data = json.dumps(dash, ensure_ascii=False).replace("</", "<\\/")
html = tpl.replace("/*DATA*/null", data)
for f in (Path(__file__).parent / "web").iterdir():
    shutil.copy(f, site / f.name)
    # versienummer achter elk bestand, zodat browsers na een wijziging nooit een oude kopie tonen
    html = html.replace(f'"{f.name}"', f'"{f.name}?v={hashlib.md5(f.read_bytes()).hexdigest()[:8]}"')
(site / "index.html").write_text(html, encoding="utf-8")
print("OK", TEAM, S, DIV, "-", len(matches), "wedstrijden,", len(games), "partijen,", len(spelers), "spelers")
