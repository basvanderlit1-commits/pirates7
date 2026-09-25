"""Haalt alle data van Pirates 7 (DBMN, via feeds.teambeheer.nl) op en schrijft
'pirates 7 resultaten.xlsx' en het dashboard in site/ (GitHub Pages).
Seizoen, team-id en divisie worden automatisch opgezocht; alleen TEAM moet kloppen met de naam op teambeheer."""
import json
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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

B = "https://feeds.teambeheer.nl"
D, TEAM = 41, "Pirates 7"  # D = DBMN op teambeheer
OUT = "pirates 7 resultaten.xlsx"
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
wed_tab = team.find("table")
matches = []
for tr in wed_tab.find("tbody").find_all("tr"):
    td = tr.find_all("td")
    if len(td) < 5:
        continue
    a = td[4].find("a")
    score = txt(a) if a else ""
    home, away = txt(td[2]), txt(td[3])
    res = ""
    if score:
        h, u = map(int, score.split("-"))
        mine, theirs = (h, u) if home == TEAM else (u, h)
        res = "W" if mine > theirs else "V" if mine < theirs else "G"
    matches.append(dict(ronde=txt(td[0]), datum=txt(td[1]), thuis=home, uit=away, score=score,
                        uitslag=res, form=a["href"] if a else None,
                        vrij=not re.fullmatch(r"\d{2}-\d{2}-\d{4}", txt(td[1]))))  # bv. "Vrije week" in de beker

spelers = []
for tr in header_table(team, "Naam").find("tbody").find_all("tr"):
    td = tr.find_all("td")
    a = td[0].find("a")
    rol = txt(td[0].find("b"))
    spelers.append(dict(naam=txt(a), rol={"C": "Captain", "RC": "Reserve captain"}.get(rol, rol),
                        id=re.search(r"l=(\d+)", a["href"]).group(1), singles=num(txt(td[1])), winst=num(txt(td[2]))))
locatie = " | ".join(s.strip() for s in team.find(string="Locatie").find_parent().find_next("p").stripped_strings)

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
    m["locatie"] = " | ".join(loc.find_parent().find_next("a").parent.stripped_strings).split(" | Bijz.")[0].removeprefix("Locatie | ") if loc else ""
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
bijz_lists = {}
for t, name in [(1, "180ers"), (2, "Hoogste finishes"), (3, "Snelste leg"), (4, "171ers")]:
    tab = get(f"/web/scorelijst-bijzres/?d={D}&t={t}&s={S}&filter=P-{DIV}").find("table")
    bijz_lists[name] = rows(tab) if tab else [["(geen data)"]]

# ---------- Excel ----------
NAVY, BLUE, GOLD, GREY = "1F3864", "2F5597", "C9A227", "7F7F7F"
FONT = "Calibri"
fill = lambda c: PatternFill("solid", fgColor=c)
ZEBRA, OURS, TILE = fill("F2F5FA"), fill("FFF2CC"), fill("F2F5FA")
WV = {"W": (fill("E2F0D9"), "375623"), "V": (fill("FBE3E4"), "9C0006"), "G": (fill("FFF2CC"), "7F6000")}
LINE = Side(style="thin", color="D9DEE8")
NOW = datetime.now(ZoneInfo("Europe/Amsterdam"))
TODAY = NOW.strftime("%d-%m-%Y")
wb = Workbook()
wb.remove(wb.active)


def to_date(s):
    try:
        return datetime.strptime(s, "%d-%m-%Y").date()
    except (TypeError, ValueError):
        return s


def pct(v):  # 33.3 / "33.3 %" -> 0.333 (Excel-percentage)
    v = num(v)
    return v / 100 if isinstance(v, (int, float)) else None


def new_sheet(title, heading, widths):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    ws.column_dimensions["A"].width = 2  # marge
    for i, w in enumerate(widths, 2):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws["B1"], ws["B2"] = heading, f"{TEAM}  ·  Divisie {DIV}  ·  Seizoen {S}  ·  bijgewerkt {TODAY}"
    ws["B1"].font = Font(name=FONT, size=18, bold=True, color=NAVY)
    ws["B2"].font = Font(name=FONT, size=10, italic=True, color=GREY)
    ws.row_dimensions[1].height = 30
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def table(ws, row, header, data, left=(), fmt=None, groups=None, ours=None, band=None, wv=None, links=None, frozen=True):
    """Opgemaakte tabel vanaf kolom B. left: kolommen (0-based) links uitgelijnd, fmt: {kolom: numberformat},
    groups: [(label, van, tot)], ours(r): markeer rij, band(r): sleutel voor afwisselende achtergrond,
    wv: kolom met W/V, links: kolom met URL. Geeft de laatste rij terug."""
    fmt = fmt or {}
    if groups:
        for label, a, b in groups:
            ws.merge_cells(start_row=row, start_column=a + 2, end_row=row, end_column=b + 2)
            c = ws.cell(row, a + 2, label)
            c.font, c.alignment = Font(name=FONT, bold=True, color="FFFFFF"), Alignment(horizontal="center")
            for col in range(a + 2, b + 3):
                ws.cell(row, col).fill = fill(BLUE)
                ws.cell(row, col).border = Border(left=Side(style="thin", color="FFFFFF") if col == a + 2 else None)
        row += 1
    for j, h in enumerate(header):
        c = ws.cell(row, j + 2, h)
        c.font, c.fill = Font(name=FONT, bold=True, color="FFFFFF"), fill(NAVY)
        c.alignment = Alignment(horizontal="left" if j in left else "center", vertical="center", wrap_text=True,
                                indent=1 if j in left else 0)
    ws.row_dimensions[row].height = 32
    head, prev, zebra = row, object(), True
    for r in data:
        row += 1
        key = band(r) if band else row
        if key != prev:
            zebra, prev = not zebra, key
        mine = bool(ours and ours(r))
        for j, v in enumerate(r):
            c = ws.cell(row, j + 2, v)
            c.font = Font(name=FONT, bold=mine)
            c.alignment = Alignment(horizontal="left" if j in left else "center", vertical="center",
                                    indent=1 if j in left else 0)
            c.border = Border(bottom=LINE)
            c.fill = OURS if mine else ZEBRA if zebra else PatternFill()
            if j in fmt:
                c.number_format = fmt[j]
            if j == wv and v in WV:
                c.fill, color = WV[v]
                c.font = Font(name=FONT, bold=True, color=color)
            if j == links and v:
                c.value, c.hyperlink = "Bekijk ›", v
                c.font = Font(name=FONT, color="0563C1", underline="single")
        ws.row_dimensions[row].height = 20
    if frozen and data:  # alleen bij één tabel per blad
        ws.freeze_panes = ws.cell(head + 1, 2)
        ws.auto_filter.ref = f"B{head}:{get_column_letter(len(header) + 1)}{row}"
    return row


def section(ws, row, col, text, span=5):
    ws.cell(row, col, text).font = Font(name=FONT, size=13, bold=True, color=NAVY)
    for k in range(col, col + span):
        ws.cell(row, k).border = Border(bottom=Side(style="medium", color=NAVY))


def line(ws, row, cols, values, fonts=None):
    for i, (c, v) in enumerate(zip(cols, values)):
        cell = ws.cell(row, c, v)
        cell.font = (fonts or {}).get(i, Font(name=FONT))
        cell.border = Border(bottom=LINE)
        if isinstance(v, date):
            cell.number_format = "dd-mm-yyyy"
    return ws


# --- data voorbereiden ---
comp = [m for m in matches if m["uitslag"] and not m["ronde"].startswith("b")]
played = [m for m in matches if m["uitslag"]]
todo = [m for m in matches if not m["uitslag"] and not m["vrij"]]
my_row = next(r for r in stand[1:] if TEAM in r)
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

# --- Overzicht (dashboard) ---
ws = new_sheet("Overzicht", f"{TEAM}  –  Seizoensoverzicht", [17] + [13] * 11)
tiles = [("Positie", f"{my_row[0]}e"), ("Punten", num(my_row[5])), ("Gem. per avond", num(my_row[6])),
         ("Gewonnen / verloren", f"{sum(m['uitslag'] == 'W' for m in comp)} / {sum(m['uitslag'] == 'V' for m in comp)}"),
         ("Nog te spelen", len(todo)), ("180's", sum("180" in b["prestatie"] for b in bijz if b["team"] == TEAM))]
for i, (label, value) in enumerate(tiles):
    col = 2 + i * 2
    for rr in (4, 5, 6):
        ws.merge_cells(start_row=rr, start_column=col, end_row=rr, end_column=col + 1)
        for k in (col, col + 1):
            ws.cell(rr, k).fill = TILE
            ws.cell(rr, k).border = Border(top=Side(style="thick", color=GOLD) if rr == 4 else None,
                                           left=Side(style="thick", color="FFFFFF") if k == col else None)
    ws.cell(4, col, label).font = Font(name=FONT, size=10, color=GREY)
    ws.cell(5, col, value).font = Font(name=FONT, size=26, bold=True, color=NAVY)
    for rr in (4, 5):
        ws.cell(rr, col).alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[4].height, ws.row_dimensions[5].height = 22, 42

section(ws, 8, 2, "Teaminformatie")
info = [("Divisie", f"Divisie {DIV}"), ("Speellocatie", locatie.replace(" | ", ", ")), ("Captain", role("Captain")),
        ("Reserve captain", role("Reserve captain")), ("Spelers", len(spelers)),
        ("Bekerwedstrijden", len(played) - len(comp)), ("Bron", "teambeheer.nl (via dbmn.nl)")]
for i, (k, v) in enumerate(info, 9):
    ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=6)
    line(ws, i, range(2, 7), [k, v], {0: Font(name=FONT, bold=True, color=GREY)})
    ws.cell(i, 3).alignment = Alignment(horizontal="left")
ws.cell(15, 3).hyperlink = f"{B}/web/team?d={D}&t={TEAM_ID}&s={S}"
ws.cell(15, 3).font = Font(name=FONT, color="0563C1", underline="single")

GREYB = Font(name=FONT, bold=True, color=GREY)
section(ws, 8, 8, "Laatste uitslagen")
line(ws, 9, [8, 9, 11, 12], ["Datum", "Tegenstander", "T/U", "Uitslag"], {0: GREYB, 1: GREYB, 2: GREYB, 3: GREYB})
row = 9
for m in reversed(played[-5:]):
    row += 1
    ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
    line(ws, row, [8, 9, 11, 12], [to_date(m["datum"]), m["tegen"], m["tu"], f"{m['uitslag']}   {m['wijzij']}"])
    ws.cell(row, 8).alignment = Alignment(horizontal="left")
    c = ws.cell(row, 12)
    c.fill, color = WV[m["uitslag"]]
    c.font, c.alignment = Font(name=FONT, bold=True, color=color), Alignment(horizontal="center")

row += 2
section(ws, row, 8, "Volgende wedstrijden")
for m in todo[:3]:
    row += 1
    ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
    line(ws, row, [8, 9, 11, 12], [to_date(m["datum"]), m["tegen"], m["tu"], f"ronde {m['ronde']}"],
         {3: Font(name=FONT, color=GREY)})
    ws.cell(row, 12).alignment = Alignment(horizontal="center")
    ws.cell(row, 8).alignment = Alignment(horizontal="left")

# --- Stand ---
ws = new_sheet(f"Stand {DIV}", f"Stand Divisie {DIV}", [7, 32, 10, 9, 9, 9, 9, 12])
table(ws, 4, ["#", "Team", "Gespeeld", "Winst", "Verlies", "Punten", "Gem.", "Strafpunten"],
      [[num(v) for v in r] for r in stand[1:] if any(r)], left=(1,), fmt={6: "0.0"}, ours=lambda r: r[1] == TEAM)

# --- Programma ---
ws = new_sheet("Programma", "Programma & uitslagen", [7, 12, 30, 8, 11, 7, 60, 12])
end = table(ws, 4, ["Ronde", "Datum", "Tegenstander", "T/U", "Uitslag (wij-zij)", "W/V", "Locatie", "Formulier"],
            [[num(m["ronde"]), to_date(m["datum"]), m["tegen"], m["tu"], m.get("wijzij"), m["uitslag"] or None,
              m.get("locatie", "").replace(" | ", ", ") or None, B + m["form"] if m["form"] else None] for m in matches],
            left=(2, 6), fmt={1: "dd-mm-yyyy"}, wv=5, links=7)
for r in ws.iter_rows(min_row=5, max_row=end):  # nog te spelen: grijs
    if not r[6].value:
        for c in r[1:]:
            c.font = Font(name=FONT, color="A6A6A6")
ws.cell(end + 2, 2, "b1 = bekerwedstrijd  ·  grijs = nog te spelen").font = Font(name=FONT, size=9, italic=True, color=GREY)

# --- Wedstrijddetails ---
ws = new_sheet("Wedstrijddetails", "Uitslagen per avond", [7, 12, 26, 34, 36, 36, 8, 8, 7])
table(ws, 4, ["Ronde", "Datum", "Tegenstander", "Onderdeel", "Pirates 7", "Tegenstander(s)", "Legs wij", "Legs zij", "W/V"],
      [[num(g["ronde"]), to_date(g["datum"]), g["tegen"], g["onderdeel"], g["wij"], g["zij"], g["legs_wij"],
        g["legs_zij"], g["uitslag"]] for g in games],
      left=(2, 3, 4, 5), fmt={1: "dd-mm-yyyy"}, band=lambda r: r[0], wv=8)

# --- Spelers ---
ws = new_sheet("Spelers", "Spelersstatistieken", [22, 16] + [8.5] * 21 + [12])
sp_rows = []
for s in spelers:
    a = agg[s["naam"]]
    p = lambda w, n: w / n if n else None
    sp_rows.append([s["naam"], s["rol"] or "Speler", s["singles"], pct(s["winst"]), pct(s["site_singles_pct"]),
                    pct(s["site_koppels_pct"]),
                    a["s_gesp"], a["s_w"], a["s_gesp"] - a["s_w"], p(a["s_w"], a["s_gesp"]), a["s_lv"], a["s_lt"],
                    a["k_gesp"], a["k_w"], a["k_gesp"] - a["k_w"], p(a["k_w"], a["k_gesp"]), a["k_lv"], a["k_lt"],
                    a["rr_gesp"], a["rr_w"], a["rr_gesp"] - a["rr_w"], p(a["rr_w"], a["rr_gesp"]),
                    sum("180" in b["prestatie"] for b in bijz if b["speler"] == s["naam"]),
                    ", ".join(b["prestatie"].replace(" finish", "") for b in bijz
                              if b["speler"] == s["naam"] and "finish" in b["prestatie"]) or None])
end = table(ws, 4, ["Speler", "Rol", "Singles", "Winst%", "Singles%", "Koppels%",
                    "Gesp.", "W", "V", "W%", "Legs +", "Legs −", "Gesp.", "W", "V", "W%", "Legs +", "Legs −",
                    "Gesp.", "W", "V", "W%", "180's", "Finishes"], sp_rows, left=(0, 1),
            fmt={3: "0%", 4: "0%", 5: "0%", 9: "0%", 15: "0%", 21: "0%"},
            groups=[("Speler", 0, 1), ("Competitie (site)", 2, 5), ("Singles (incl. beker)", 6, 11),
                    ("Koppels (incl. beker)", 12, 17), ("Round robin 301", 18, 21), ("Bijzonder", 22, 23)])
ws.cell(end + 2, 2, "Blokken 'incl. beker' zijn berekend uit de wedstrijdformulieren; lege percentages staan niet op de site.") \
    .font = Font(name=FONT, size=9, italic=True, color=GREY)

# --- Persoonlijk klassement ---
ws = new_sheet("Persoonlijk klassement", "Persoonlijk klassement", [13, 26, 10, 9, 9, 10])
ws["B3"] = (f"Positie tussen alle spelers van divisie {DIV} (1e = beste). Alleen competitie, zonder beker. "
            "Bij gelijk winstpercentage staat wie meer partijen speelde hoger.")
ws["B3"].font = Font(name=FONT, size=10, color=GREY)
row = 5
for label, tab in [("Singles", pk_single), ("Koppels", pk_koppel)]:
    section(ws, row, 2, label, 6)
    rows_ = [[f"{num(r[0])}e van {pk_count(tab)}", r[1], num(r[3]), won(r), pct(r[6])] for r in ours_only(tab, 2)]
    row = table(ws, row + 1, ["Positie", "Speler", "Gespeeld", "Gewonnen", "Winst%"], rows_,
                left=(1,), fmt={4: "0%"}, frozen=False) + 2

# --- Bijzondere resultaten ---
ws = new_sheet("Bijzondere resultaten", "Bijzondere resultaten", [16, 12, 26, 16])
section(ws, 4, 2, "Per wedstrijd", 4)
end = table(ws, 5, ["Ronde", "Datum", "Speler", "Prestatie"],
            [[num(b["ronde"]), to_date(b["datum"]), b["speler"], b["prestatie"]] for b in bijz if b["team"] == TEAM],
            left=(2,), fmt={1: "dd-mm-yyyy"}, frozen=False)
section(ws, end + 2, 2, f"Positie in klassement {DIV}", 4)
table(ws, end + 3, ["Lijst", "Positie", "Speler", "Aantal / waarde"],
      [[name, num(r[0]), r[1], num(r[4])] for name, tab in bijz_lists.items() for r in ours_only(tab, 2)], left=(0, 2), frozen=False)

wb.active = 0
site = Path("site")
site.mkdir(exist_ok=True)
wb.save(site / "pirates7-resultaten.xlsx")  # download op de website
try:
    wb.save(OUT)
except PermissionError:
    print(f"LET OP: '{OUT}' staat open in Excel en is niet bijgewerkt. Sluit hem en draai opnieuw.")

# ---------- Dashboard (HTML) ----------
kind = lambda o: "RR" if o.startswith("RR") else "Single" if o.startswith(("Single", "Singel")) else \
    "Koppel" if o.startswith("Koppel") else "Team"
dash = dict(
    team=TEAM, div=DIV, seizoen=S, bijgewerkt=NOW.strftime("%d-%m-%Y %H:%M"), locatie=locatie.replace(" | ", ", "),
    captain=role("Captain"), rc=role("Reserve captain"), bron=f"{B}/web/team?d={D}&t={TEAM_ID}&s={S}",
    stand=[dict(pos=num(r[0]), team=r[1], wed=num(r[2]), w=num(r[3]), v=num(r[4]), pnt=num(r[5]), gem=num(r[6]))
           for r in stand[1:] if any(r)],
    matches=[dict(ronde=m["ronde"], datum=m["datum"], tegen=m["tegen"], tu=m["tu"], uitslag=m["uitslag"], vrij=m["vrij"],
                  wij=int(m["wijzij"].split(" - ")[0]) if m["score"] else None,
                  zij=int(m["wijzij"].split(" - ")[1]) if m["score"] else None,
                  locatie=m.get("locatie", "").replace(" | ", ", "), form=B + m["form"] if m["form"] else None)
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
)
tpl = (Path(__file__).parent / "dashboard_template.html").read_text(encoding="utf-8")
data = json.dumps(dash, ensure_ascii=False).replace("</", "<\\/")
(site / "index.html").write_text(tpl.replace("/*DATA*/null", data), encoding="utf-8")
for f in (Path(__file__).parent / "web").iterdir():
    shutil.copy(f, site / f.name)
print("OK", TEAM, S, DIV, "-", len(matches), "wedstrijden,", len(games), "partijen,", len(spelers), "spelers")
