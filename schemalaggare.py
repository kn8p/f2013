#!/usr/bin/env python3
"""
F2013 schemalaggare
===================
Laser F2013_schemaunderlag.xlsx (trupp, matcher, tillganglighet, installningar)
och raknar ut ett spelschema. Skriver ut en ny arbetsbok.

    pip install openpyxl pulp
    python schemalaggare.py F2013_schemaunderlag.xlsx

Frånvaro anges i bladet "Tillganglighet": tom cell = kan spela, allt annat = kan inte.
Gar det inte ihop krymper skriptet truppen (farre avbytare) i stallet for att ge upp,
och listar varje sadan match under "Varningar" i utdatafilen.
"""
import sys, datetime, collections
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
import pulp

# ---------------------------------------------------------------- inlasning
def read_input(path):
    wb = load_workbook(path, data_only=True)

    st = {}
    ws = wb["Inställningar"]
    for r in range(5, ws.max_row + 1):
        k = ws.cell(row=r, column=1).value
        if k: st[str(k).strip()] = ws.cell(row=r, column=2).value

    players = []
    ws = wb["Trupp"]
    for r in range(5, ws.max_row + 1):
        n = ws.cell(row=r, column=1).value
        if not n: continue
        if str(ws.cell(row=r, column=7).value).strip().upper() != "J": continue
        players.append(dict(
            namn=str(n).strip(),
            prim=str(ws.cell(row=r, column=2).value or "").strip(),
            sek=str(ws.cell(row=r, column=3).value or "").strip(),
            avg=float(ws.cell(row=r, column=4).value),
            grupp=str(ws.cell(row=r, column=5).value or "auto").strip(),
            mv=str(ws.cell(row=r, column=6).value or "N").strip().upper() == "J",
            ledare=str(ws.cell(row=r, column=8).value or "").strip()))

    for p in players:
        if p["grupp"].lower() == "auto":
            a = p["avg"]
            p["grupp"] = ("A" if a <= float(st["grupp_a_max_niva"]) else
                          "B" if a <= float(st["grupp_b_max_niva"]) else
                          "C" if a <= float(st["grupp_c_max_niva"]) else "D")
        p["grupp"] = p["grupp"].upper()

    units, matches = {}, []
    ws = wb["Matcher"]
    for r in range(5, ws.max_row + 1):
        uid = ws.cell(row=r, column=1).value
        if not uid: continue
        uid = str(uid).strip()
        m = dict(unit=uid, niva=str(ws.cell(row=r, column=2).value).strip(),
                 datum=str(ws.cell(row=r, column=3).value)[:10],
                 dag=ws.cell(row=r, column=4).value,
                 tid=str(ws.cell(row=r, column=5).value)[:5],
                 serie=ws.cell(row=r, column=6).value,
                 nr=ws.cell(row=r, column=7).value,
                 hemma=ws.cell(row=r, column=8).value,
                 borta=ws.cell(row=r, column=9).value,
                 plats=ws.cell(row=r, column=10).value)
        matches.append(m)
        u = units.setdefault(uid, dict(id=uid, niva=m["niva"], datum=m["datum"],
                                       dag=m["dag"], venue=m["plats"], n=0, matcher=[]))
        u["n"] += 1; u["matcher"].append(m)
    units = sorted(units.values(), key=lambda u: (u["datum"], u["matcher"][0]["tid"]))

    hemord = [w.strip() for w in str(st.get("hemmaplats_ord","Alingsås,Nolhagahallen")).split(",")]

    ws = wb["Tillgänglighet"]
    ucol = {}
    for c in range(4, ws.max_column + 1):
        v = ws.cell(row=7, column=c).value
        if v: ucol[str(v).strip()] = c
    unavail = set()
    for r in range(8, ws.max_row + 1):
        n = ws.cell(row=r, column=1).value
        if not n: continue
        n = str(n).strip()
        for uid, c in ucol.items():
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                unavail.add((n, uid))
    return st, players, units, matches, unavail, hemord

# ---------------------------------------------------------------- roller
def primrole(p):
    t = p["prim"]
    if p["mv"]: return "MV"
    if t in ("9M", "H9", "V9", "9"): return "nia"
    if t == "M6": return "sexa"
    if t in ("V6", "H6", "Kant"): return "kant"
    return "ovrig"

def caps(p, kind):
    s = set()
    for x in (p["prim"], p["sek"]):
        for tok in str(x).replace("/", " ").replace("+", " ").split():
            if tok in ("9M", "H9", "V9", "9"): s.add("nia")
            elif tok == "M6": s.add("sexa")
            elif tok in ("V6", "H6", "Kant"): s.add("kant")
    return kind in s

# ---------------------------------------------------------------- losning
def solve(st, players, units, unavail):
    I = lambda k: int(st[k]); Fl = lambda k: float(st[k])
    GK = [p for p in players if p["mv"]]
    FP = [p for p in players if not p["mv"]]
    ok = lambda p, u: (p["namn"], u["id"]) not in unavail

    hardu = [u for u in units if u["niva"] == "Svår"]
    easyu = [u for u in units if u["niva"] != "Svår"]
    want = {u["id"]: (I("svar_pa_plan") + I("svar_avbytare") if u["niva"] == "Svår"
                      else I("latt_pa_plan") + I("latt_avbytare")) for u in units}
    floor = {u["id"]: (I("svar_pa_plan") + I("svar_avbytare_min") if u["niva"] == "Svår"
                       else I("latt_pa_plan") + I("latt_avbytare_min")) for u in units}
    games = {u["id"]: u["n"] for u in units}

    prob = pulp.LpProblem("schema", pulp.LpMinimize)
    x, short_f, short_g = {}, {}, {}
    for p in players:
        for u in units:
            v = pulp.LpVariable(f"x_{abs(hash((p['namn'],u['id'])))}", cat="Binary")
            x[p["namn"], u["id"]] = v
            if not ok(p, u): prob += v == 0

    for u in units:
        uid = u["id"]
        short_f[uid] = pulp.LpVariable(f"sf_{uid}", lowBound=0,
                                       upBound=want[uid] - floor[uid], cat="Integer")
        short_g[uid] = pulp.LpVariable(f"sg_{uid}", lowBound=0, upBound=1, cat="Integer")
        prob += pulp.lpSum(x[p["namn"], uid] for p in GK) == 1 - short_g[uid]
        prob += pulp.lpSum(x[p["namn"], uid] for p in FP) == want[uid] - short_f[uid]

        hard = (u["niva"] == "Svår")
        pre = "svar" if hard else "latt"
        for role, key in (("nia", f"{pre}_min_nior"), ("sexa", f"{pre}_min_sexor"),
                          ("kant", f"{pre}_min_kanter")):
            avail = sum(1 for p in FP if primrole(p) == role and ok(p, u))
            prob += pulp.lpSum(x[p["namn"], uid] for p in FP if primrole(p) == role) \
                    >= min(I(key), avail)
        av = sum(1 for p in FP if caps(p, "nia") and ok(p, u))
        prob += pulp.lpSum(x[p["namn"], uid] for p in FP if caps(p, "nia")) >= min(3, av)

    # om ingen behorig malvakt ar tillganglig for en enhet slapps nivaspärren
    # for malvakter just den enheten (hellre fel nivå an ingen malvakt alls)
    mv_override = set()
    for u in units:
        hard = (u["niva"] == "Svår")
        elig = [p for p in GK if ok(p, u) and
                (p["grupp"] in ("A", "B", "C") if hard else p["grupp"] in ("B", "C", "D"))]
        if not elig and any(ok(p, u) for p in GK):
            mv_override.add(u["id"])

    # nivagrupper
    for p in players:
        g, n = p["grupp"], p["namn"]
        hg = pulp.lpSum(x[n, u["id"]] * games[u["id"]] for u in hardu)
        lg = pulp.lpSum(x[n, u["id"]] * games[u["id"]] for u in easyu)
        if g == "A":
            for u in easyu:
                if not (p["mv"] and u["id"] in mv_override): prob += x[n, u["id"]] == 0
        elif g == "D":
            for u in hardu:
                if not (p["mv"] and u["id"] in mv_override): prob += x[n, u["id"]] == 0
        elif g == "B":
            cap = I("mv_grupp_b_latta_max") if p["mv"] else I("b_latta_max")
            prob += lg <= cap
            if not p["mv"]:
                prob += lg >= min(I("b_latta_min"),
                                  sum(games[u["id"]] for u in easyu if ok(p, u)))
        elif g == "C":
            prob += hg <= I("c_svara_max")
            prob += hg >= min(I("c_svara_min"), sum(1 for u in hardu if ok(p, u)))

    # minst en ledares barn i varje trupp
    ledarbarn = [p["namn"] for p in players if p["ledare"]]
    for u in units:
        avail = [n for n in ledarbarn if (n, u["id"]) not in unavail]
        if avail: prob += pulp.lpSum(x[n, u["id"]] for n in avail) >= 1

    # ingen dubbelbokning samma dag
    if str(st.get("tillat_dubbel_samma_dag", "N")).strip().upper() != "J":
        byday = collections.defaultdict(list)
        for u in units: byday[u["datum"]].append(u)
        for dd, us in byday.items():
            if len(us) < 2: continue
            for p in players:
                prob += pulp.lpSum(x[p["namn"], u["id"]] for u in us) <= 1

    # rattvisa, viktad mot hur manga enheter spelaren faktiskt kan
    tot = {p["namn"]: pulp.lpSum(x[p["namn"], u["id"]] * games[u["id"]] for u in units)
           for p in players}
    fmax = pulp.LpVariable("fmax"); fmin = pulp.LpVariable("fmin")
    gmax = pulp.LpVariable("gmax"); gmin = pulp.LpVariable("gmin")
    for p in FP:
        prob += tot[p["namn"]] <= fmax; prob += tot[p["namn"]] >= fmin
    for p in GK:
        prob += tot[p["namn"]] <= gmax; prob += tot[p["namn"]] >= gmin

    prob += (1000 * pulp.lpSum(short_g.values()) + 100 * pulp.lpSum(short_f.values())
             + 3 * (fmax - fmin) + 3 * (gmax - gmin))
    status = prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=I("tidsgrans_sek")))
    if pulp.LpStatus[status] not in ("Optimal", "Not Solved"):
        raise SystemExit(f"Ingen losning: {pulp.LpStatus[status]}. "
                         "For manga franvarande eller for hardaanded installningar.")

    trupp = {u["id"]: [p["namn"] for p in players
                       if x[p["namn"], u["id"]].value() and x[p["namn"], u["id"]].value() > 0.5]
             for u in units}
    warn = []
    for uid in sorted(mv_override):
        u = next(z for z in units if z["id"] == uid)
        gk = [n for n in trupp[uid] if any(p["mv"] and p["namn"] == n for p in players)]
        warn.append((uid, u["datum"],
                     "Ingen målvakt i rätt nivågrupp var tillgänglig – "
                     f"{gk[0] if gk else 'ingen'} tar matchen utanför sin nivågrupp"))
    for u in units:
        uid = u["id"]
        if short_g[uid].value() and short_g[uid].value() > 0.5:
            warn.append((uid, u["datum"], "INGEN MÅLVAKT tillgänglig"))
        s = int(round(short_f[uid].value() or 0))
        if s > 0:
            warn.append((uid, u["datum"], f"{s} avbytare saknas (trupp {want[uid]-s} utespelare)"))
    return trupp, warn, pulp.LpStatus[status]

# ---------------------------------------------------------------- uppdrag
def is_home(venue, hemord):
    return any(w and w in str(venue) for w in hemord)

def solve_duties(st, players, units, trupp, hemord):
    """Fordelar foraldrauppdrag. Bara foraldrar vars barn spelar; ledare undantagna."""
    I = lambda k: int(st[k])
    HOME = {"Kiosk": I("uppdrag_kiosk"), "Matchvärd": I("uppdrag_matchvard"),
            "Sekretariat": I("uppdrag_sekretariat")}
    AWAY = {"Körning": I("uppdrag_korning")}
    HOME = {k: v for k, v in HOME.items() if v > 0}
    AWAY = {k: v for k, v in AWAY.items() if v > 0}
    TAK  = {"Kiosk": I("uppdrag_tak_kiosk"), "Matchvärd": I("uppdrag_tak_matchvard"),
            "Sekretariat": I("uppdrag_tak_sekretariat"), "Körning": I("uppdrag_tak_korning")}

    ledarbarn = {p["namn"] for p in players if p["ledare"]}
    fam = sorted({p["namn"] for p in players} - ledarbarn)
    if not fam: return {}, [("-", "inga behöriga föräldrar")]

    need = {u["id"]: (HOME if is_home(u["venue"], hemord) else AWAY) for u in units}
    elig = {u["id"]: [n for n in trupp[u["id"]] if n not in ledarbarn] for u in units}
    total = sum(sum(v.values()) for v in need.values())
    target = total / len(fam)

    prob = pulp.LpProblem("uppdrag", pulp.LpMinimize)
    z, warn = {}, []
    for u in units:
        uid = u["id"]
        for r in need[uid]:
            for n in elig[uid]:
                z[n, uid, r] = pulp.LpVariable(f"z_{abs(hash((n,uid,r)))}", cat="Binary")

    short = {}
    for u in units:
        uid = u["id"]
        cap = len(elig[uid])
        want = sum(need[uid].values())
        for r, k in need[uid].items():
            kk = k
            if cap < want:                       # for fa foraldrar: krymp jamnt
                kk = max(0, int(round(k * cap / want)))
                short[uid] = (cap, want)
            prob += pulp.lpSum(z[n, uid, r] for n in elig[uid]) == min(kk, cap)
        for n in elig[uid]:
            prob += pulp.lpSum(z[n, uid, r] for r in need[uid]) <= 1

    dev = {}
    for n in fam:
        tot = pulp.lpSum(z[n, u["id"], r] for u in units for r in need[u["id"]]
                         if (n, u["id"], r) in z)
        dev[n] = pulp.LpVariable(f"d_{abs(hash(n))}", lowBound=0)
        prob += dev[n] >= tot - target
        prob += dev[n] >= target - tot
        for r in TAK:
            terms = [z[n, u["id"], r] for u in units
                     if r in need[u["id"]] and (n, u["id"], r) in z]
            if terms: prob += pulp.lpSum(terms) <= TAK[r]
    hi = pulp.LpVariable("uhi"); lo = pulp.LpVariable("ulo")
    for n in fam:
        tot = pulp.lpSum(z[n, u["id"], r] for u in units for r in need[u["id"]]
                         if (n, u["id"], r) in z)
        prob += tot <= hi; prob += tot >= lo
    prob += 10 * pulp.lpSum(dev.values()) + (hi - lo)
    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=I("tidsgrans_sek")))

    out = {}
    for u in units:
        uid = u["id"]
        out[uid] = {r: sorted(n for n in elig[uid] if z[n, uid, r].value() and
                              z[n, uid, r].value() > .5) for r in need[uid]}
    for uid, (cap, want) in short.items():
        u = next(x for x in units if x["id"] == uid)
        warn.append((uid, u["datum"],
                     f"bara {cap} behöriga föräldrar i truppen, {want} uppdrag behövs"))
    return out, warn

# ---------------------------------------------------------------- utdata
F = "Arial"
navy = PatternFill("solid", fgColor="1F3864")
hardf = PatternFill("solid", fgColor="FCE4D6"); easyf = PatternFill("solid", fgColor="E2EFDA")
gkf = PatternFill("solid", fgColor="FFF2CC"); warnf = PatternFill("solid", fgColor="FFC7CE")
absf = PatternFill("solid", fgColor="D9D9D9"); altf = PatternFill("solid", fgColor="F7F7F7")
thin = Side(style="thin", color="BFBFBF")
bd = Border(left=thin, right=thin, top=thin, bottom=thin)

def hdr(c, rot=0):
    c.font = Font(name=F, size=9, bold=True, color="FFFFFF"); c.fill = navy
    c.alignment = Alignment(horizontal="center", vertical="center",
                            wrap_text=(rot == 0), text_rotation=rot)
    c.border = bd

def write_output(path, players, units, matches, trupp, warn, unavail, st,
                 duties=None, hemord=("Alingsås",)):
    duties = duties or {}
    pm = {p["namn"]: p for p in players}
    names = [p["namn"] for p in sorted(players, key=lambda p: (not p["mv"], p["avg"], p["namn"]))]
    umap = {u["id"]: u for u in units}
    rows = sorted(matches, key=lambda m: (m["datum"], m["tid"]))
    wb = Workbook()

    # --- Schema
    ws = wb.active; ws.title = "Schema"
    ws["A1"] = "Spelschema F2013"; ws["A1"].font = Font(name=F, size=14, bold=True)
    ws["A2"] = ("MV = målvakt · X = i truppen · grå = otillgänglig · tom = ledig. "
                f"Genererat {datetime.datetime.now():%Y-%m-%d %H:%M}")
    ws["A2"].font = Font(name=F, size=9, italic=True, color="808080")
    INFO = ["Datum", "Dag", "Tid", "Nivå", "Match", "Spelplats", "Antal"]
    NI = len(INFO); HR = 6
    for i, n in enumerate(names):
        col = NI + 1 + i
        c = ws.cell(row=4, column=col, value=pm[n]["prim"])
        c.font = Font(name=F, size=8, color="808080"); c.alignment = Alignment(horizontal="center")
        c = ws.cell(row=5, column=col, value=pm[n]["avg"])
        c.font = Font(name=F, size=8, color="808080"); c.alignment = Alignment(horizontal="center")
        hdr(ws.cell(row=HR, column=col, value=n), rot=90)
        ws.column_dimensions[get_column_letter(col)].width = 4.2
    for i, h in enumerate(INFO): hdr(ws.cell(row=HR, column=i + 1, value=h))
    lastcol = get_column_letter(NI + len(names))
    r = first = HR + 1
    for m in rows:
        u = umap[m["unit"]]
        for i, v in enumerate([m["datum"], m["dag"], m["tid"], m["niva"],
                               f'{m["hemma"]} – {m["borta"]}', m["plats"]]):
            c = ws.cell(row=r, column=i + 1, value=v)
            c.font = Font(name=F, size=9); c.border = bd
        ws.cell(row=r, column=7,
                value=f'=COUNTIF(H{r}:{lastcol}{r},"X")+COUNTIF(H{r}:{lastcol}{r},"MV")')
        for i, n in enumerate(names):
            c = ws.cell(row=r, column=NI + 1 + i)
            if n in trupp[m["unit"]]:
                c.value = "MV" if pm[n]["mv"] else "X"
                c.font = Font(name=F, size=9, bold=pm[n]["mv"])
                if pm[n]["mv"]: c.fill = gkf
            elif (n, m["unit"]) in unavail:
                c.value = "-"; c.fill = absf
                c.font = Font(name=F, size=9, color="808080")
            c.alignment = Alignment(horizontal="center"); c.border = bd
        for c in range(1, 5): ws.cell(row=r, column=c).fill = hardf if m["niva"] == "Svår" else easyf
        for c in (2, 3, 4, 7): ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
        r += 1
    last = r - 1
    tr = last + 1
    for lbl, off in (("Svåra", 0), ("Lätta", 1), ("Totalt", 2)):
        ws.cell(row=tr + off, column=NI, value=lbl).font = Font(name=F, size=9, bold=True)
    for i, n in enumerate(names):
        L = get_column_letter(NI + 1 + i)
        ws.cell(row=tr, column=NI+1+i,
                value=f'=COUNTIFS($D${first}:$D${last},"Svår",{L}${first}:{L}${last},"X")'
                      f'+COUNTIFS($D${first}:$D${last},"Svår",{L}${first}:{L}${last},"MV")')
        ws.cell(row=tr+1, column=NI+1+i,
                value=f'=COUNTIFS($D${first}:$D${last},"Lätt",{L}${first}:{L}${last},"X")'
                      f'+COUNTIFS($D${first}:$D${last},"Lätt",{L}${first}:{L}${last},"MV")')
        ws.cell(row=tr+2, column=NI+1+i, value=f'={L}{tr}+{L}{tr+1}')
        for rr in (tr, tr+1, tr+2):
            c = ws.cell(row=rr, column=NI+1+i)
            c.font = Font(name=F, size=9, bold=(rr == tr+2))
            c.alignment = Alignment(horizontal="center"); c.border = bd
    lo = int(st["latt_pa_plan"]) + int(st["latt_avbytare_min"])
    ws.conditional_formatting.add(f"G{first}:G{last}",
        CellIsRule(operator="lessThan", formula=[str(lo + 1)], fill=warnf))
    for col, w in zip("ABCDEFG", (11, 9, 6, 6, 34, 27, 7)): ws.column_dimensions[col].width = w
    ws.row_dimensions[HR].height = 105
    ws.freeze_panes = ws.cell(row=HR + 1, column=NI + 1)

    # --- Per match
    ws2 = wb.create_sheet("Per match")
    ws2["A1"] = "Trupp per match"; ws2["A1"].font = Font(name=F, size=14, bold=True)
    nmax = max(len(v) for v in trupp.values())
    cols = (["Datum", "Dag", "Tid", "Nivå", "Matchnr", "Match", "Spelplats", "Målvakt"]
            + [f"Spelare {i}" for i in range(1, nmax + 1)] + ["Antal"])
    for i, t in enumerate(cols): hdr(ws2.cell(row=3, column=i + 1, value=t))
    r = 4
    for k, m in enumerate(rows):
        sq = trupp[m["unit"]]
        gk = [n for n in sq if pm[n]["mv"]]
        fp = sorted([n for n in sq if not pm[n]["mv"]], key=lambda n: (primrole(pm[n]), n))
        vals = ([m["datum"], m["dag"], m["tid"], m["niva"], m["nr"],
                 f'{m["hemma"]} – {m["borta"]}', m["plats"], gk[0] if gk else "SAKNAS"]
                + fp + [None] * (nmax - len(fp)))
        for i, v in enumerate(vals):
            c = ws2.cell(row=r, column=i + 1, value=v)
            c.font = Font(name=F, size=9); c.border = bd
            if k % 2: c.fill = altf
        e = get_column_letter(8 + nmax)
        ws2.cell(row=r, column=9 + nmax, value=f'=COUNTA(H{r}:{e}{r})')
        ws2.cell(row=r, column=9 + nmax).font = Font(name=F, size=9)
        ws2.cell(row=r, column=9 + nmax).border = bd
        ws2.cell(row=r, column=4).fill = hardf if m["niva"] == "Svår" else easyf
        ws2.cell(row=r, column=8).fill = warnf if not gk else gkf
        r += 1
    ws2.auto_filter.ref = f"A3:{get_column_letter(9+nmax)}{r-1}"
    ws2.freeze_panes = "A4"
    for col, w in zip("ABCDEFG", (11, 9, 6, 6, 15, 34, 27)): ws2.column_dimensions[col].width = w
    for i in range(8, 9 + nmax): ws2.column_dimensions[get_column_letter(i)].width = 17

    # --- Uppdrag
    ws5 = wb.create_sheet("Uppdrag")
    ws5["A1"] = "Föräldrauppdrag"; ws5["A1"].font = Font(name=F, size=14, bold=True)
    ws5["A2"] = ("Namnet är barnets. Bara föräldrar vars barn spelar, ledarnas familjer undantagna.")
    ws5["A2"].font = Font(name=F, size=9, italic=True, color="808080")
    roles = ["Kiosk", "Matchvärd", "Sekretariat", "Körning"]
    for i, t in enumerate(["Truppenhet", "Datum", "Hemma/borta", "Match", "Spelplats"] + roles):
        hdr(ws5.cell(row=4, column=i + 1, value=t))
    r = 5
    for k, u in enumerate(units):
        home = is_home(u["venue"], hemord)
        du = duties.get(u["id"], {})
        vals = [u["id"], u["datum"], "Hemma" if home else "Borta",
                " / ".join(f'{m["hemma"]} – {m["borta"]}' for m in u["matcher"]), u["venue"]] + \
               [", ".join(du.get(x, [])) for x in roles]
        for i, v in enumerate(vals):
            c = ws5.cell(row=r, column=i + 1, value=v)
            c.font = Font(name=F, size=9); c.border = bd
            if k % 2: c.fill = altf
        ws5.cell(row=r, column=3).fill = hardf if home else easyf
        for i, x in enumerate(roles):
            if x in du and not du[x]: ws5.cell(row=r, column=6 + i).fill = warnf
        r += 1
    ws5.auto_filter.ref = f"A4:I{r-1}"; ws5.freeze_panes = "A5"
    for col, w in zip("ABCDEFGHI", (12, 11, 12, 34, 27, 26, 18, 26, 30)):
        ws5.column_dimensions[col].width = w

    # --- Per spelare
    ws3 = wb.create_sheet("Per spelare")
    ws3["A1"] = "Speltid per spelare"; ws3["A1"].font = Font(name=F, size=14, bold=True)
    h3 = ["Spelare", "Primär pos", "Nivå medel", "Grupp", "Svåra", "Lätta", "Totalt",
          "Frånvarande enheter", "Möjliga matcher", "Andel spelade", "Ledare", "Uppdrag"]
    for i, t in enumerate(h3): hdr(ws3.cell(row=3, column=i + 1, value=t))
    r = 4
    for k, n in enumerate(names):
        p = pm[n]; L = get_column_letter(NI + 1 + names.index(n))
        nabs = sum(1 for u in units if (n, u["id"]) in unavail)
        poss = sum(u["n"] for u in units if (n, u["id"]) not in unavail
                   and not (p["grupp"] == "A" and u["niva"] != "Svår")
                   and not (p["grupp"] == "D" and u["niva"] == "Svår"))
        nduty = sum(1 for du in duties.values() for lst in du.values() if n in lst)
        vals = [n, p["prim"], p["avg"], p["grupp"], f"=Schema!{L}{tr}", f"=Schema!{L}{tr+1}",
                f"=E{r}+F{r}", nabs, poss, f'=IF(I{r}=0,"",G{r}/I{r})',
                p["ledare"] or "", "" if p["ledare"] else nduty]
        for i, v in enumerate(vals):
            c = ws3.cell(row=r, column=i + 1, value=v)
            c.font = Font(name=F, size=10); c.border = bd
            if i >= 2: c.alignment = Alignment(horizontal="center")
            if k % 2: c.fill = altf
        ws3.cell(row=r, column=10).number_format = "0%"
        r += 1
    ws3.auto_filter.ref = f"A3:L{r-1}"; ws3.freeze_panes = "A4"
    for col, w in zip("ABCDEFGHIJKL", (24, 12, 12, 8, 9, 9, 9, 18, 16, 14, 12, 10)):
        ws3.column_dimensions[col].width = w

    # --- Varningar
    ws4 = wb.create_sheet("Varningar")
    ws4["A1"] = "Varningar"; ws4["A1"].font = Font(name=F, size=14, bold=True)
    if not warn:
        ws4["A3"] = "Inga. Alla trupper kunde fyllas med önskat antal avbytare."
        ws4["A3"].font = Font(name=F, size=11)
    else:
        for i, t in enumerate(["Truppenhet", "Datum", "Problem"]):
            hdr(ws4.cell(row=3, column=i + 1, value=t))
        for k, (uid, dt, msg) in enumerate(warn):
            for i, v in enumerate((uid, dt, msg)):
                c = ws4.cell(row=4 + k, column=i + 1, value=v)
                c.font = Font(name=F, size=10); c.border = bd; c.fill = warnf
    for col, w in zip("ABC", (14, 13, 70)): ws4.column_dimensions[col].width = w
    wb.save(path)

# ---------------------------------------------------------------- main
def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "F2013_schemaunderlag.xlsx"
    st, players, units, matches, unavail, hemord = read_input(src)
    print(f"{len(players)} spelare · {len(units)} truppenheter · {len(matches)} matcher "
          f"· {len(unavail)} frånvaromarkeringar")
    trupp, warn, status = solve(st, players, units, unavail)
    duties, dwarn = solve_duties(st, players, units, trupp, hemord)
    warn = warn + dwarn
    out = sys.argv[2] if len(sys.argv) > 2 else \
        f"F2013_spelschema_{datetime.datetime.now():%Y-%m-%d_%H%M}.xlsx"
    write_output(out, players, units, matches, trupp, warn, unavail, st, duties, hemord)
    print(f"Status: {status}")
    for w in warn: print("  VARNING:", *w)
    print("Skrev", out)

if __name__ == "__main__":
    main()
