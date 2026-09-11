#!/usr/bin/env python3
"""
Importera andringar fran webbappen
==================================
Tar CSV-filen du laddar ner med "Ladda ner alla trupper (CSV)" i Tranarlaget
och skriver in trupperna och foraldrauppdragen i en ny Excel-arbetsbok med
exakt samma blad som schemalaggare.py producerar.

    python importera_andringar.py f2013_trupper.csv F2013_schemaunderlag.xlsx

Skriptet raknar INTE om nagot. Det tar appen som facit och kontrollerar bara
att reglerna halls, sa att du ser vad handpalaggningen har kostat.
"""
import sys, csv, datetime, collections
import schemalaggare as S

def read_csv(path):
    trupp = collections.defaultdict(list)
    duties = collections.defaultdict(lambda: collections.defaultdict(list))
    pub = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            uid, namn = row["Truppenhet"].strip(), row["Spelare"].strip()
            if not uid or not namn: continue
            if namn not in trupp[uid]: trupp[uid].append(namn)
            roll = (row.get("Uppdrag") or "").strip()
            if roll: duties[uid][roll].append(namn)
            if row.get("Publicerad"):
                pub[row["Datum"][:7]] = row["Publicerad"].strip().lower().startswith("j")
    return dict(trupp), {k: dict(v) for k, v in duties.items()}, pub

def check(st, players, units, trupp, duties, hemord):
    I = lambda k: int(st[k])
    pm = {p["namn"]: p for p in players}
    ledarbarn = {p["namn"] for p in players if p["ledare"]}
    HOME = {"Kiosk": I("uppdrag_kiosk"), "Matchvärd": I("uppdrag_matchvard"),
            "Sekretariat": I("uppdrag_sekretariat")}
    AWAY = {"Körning": I("uppdrag_korning")}
    warn = []
    byday = collections.defaultdict(list)
    for u in units: byday[u["datum"]].append(u["id"])

    for u in units:
        uid, sq = u["id"], trupp.get(u["id"], [])
        hard = (u["niva"] == "Svår")
        want = (I("svar_pa_plan") + I("svar_avbytare") if hard
                else I("latt_pa_plan") + I("latt_avbytare"))
        okand = [n for n in sq if n not in pm]
        if okand: warn.append((uid, u["datum"], "okända namn: " + ", ".join(okand)))
        sq = [n for n in sq if n in pm]
        gk = [n for n in sq if pm[n]["mv"]]
        fld = [n for n in sq if not pm[n]["mv"]]
        if not gk: warn.append((uid, u["datum"], "ingen målvakt"))
        if len(fld) != want:
            warn.append((uid, u["datum"], f"{len(fld)} utespelare, förväntat {want}"))
        if len(set(sq)) != len(sq):
            warn.append((uid, u["datum"], "samma spelare listad två gånger"))
        if not [n for n in sq if n in ledarbarn]:
            warn.append((uid, u["datum"], "ingen ledare i truppen"))
        for grp, lvl, txt in (("A", "Lätt", "nivå 1 i lätt match"),
                              ("D", "Svår", "nivå 4 i svår match")):
            fel = [n for n in sq if pm[n]["grupp"] == grp and u["niva"] == lvl]
            if fel: warn.append((uid, u["datum"], txt + ": " + ", ".join(fel)))

        need = HOME if S.is_home(u["venue"], hemord) else AWAY
        du = duties.get(uid, {})
        taken = []
        for r, k in need.items():
            got = du.get(r, [])
            if len(got) != k:
                warn.append((uid, u["datum"], f"{r}: {len(got)} av {k} tillsatta"))
            for n in got:
                if n not in sq: warn.append((uid, u["datum"], f"{n} har {r} men spelar inte"))
                if n in ledarbarn: warn.append((uid, u["datum"], f"{n} är ledarfamilj, ska slippa {r}"))
                if n in taken: warn.append((uid, u["datum"], f"{n} har två uppdrag samma dag"))
                taken.append(n)
        for r in du:
            if r not in need: warn.append((uid, u["datum"], f"uppdraget {r} hör inte till denna match"))

    for d, ids in byday.items():
        if len(ids) < 2: continue
        seen = collections.Counter(n for i in ids for n in trupp.get(i, []))
        for n, c in seen.items():
            if c > 1: warn.append((ids[0], d, f"{n} är uttagen i två matcher samma dag"))
    return warn

def main():
    if len(sys.argv) < 2:
        raise SystemExit("Använd: python importera_andringar.py <csv> [underlag.xlsx] [ut.xlsx]")
    csv_path = sys.argv[1]
    src = sys.argv[2] if len(sys.argv) > 2 else "F2013_schemaunderlag.xlsx"
    out = sys.argv[3] if len(sys.argv) > 3 else \
        f"F2013_spelschema_{datetime.datetime.now():%Y-%m-%d_%H%M}_manuell.xlsx"

    st, players, units, matches, unavail, hemord = S.read_input(src)
    trupp, duties, pub = read_csv(csv_path)

    kanda = {u["id"] for u in units}
    saknas = kanda - set(trupp)
    extra = set(trupp) - kanda
    if saknas: print("Saknas i CSV:", ", ".join(sorted(saknas)))
    if extra:  print("Okända truppenheter i CSV:", ", ".join(sorted(extra)))
    for uid in saknas: trupp[uid] = []

    warn = check(st, players, units, trupp, duties, hemord)
    if pub:
        publ = sorted(k for k, v in pub.items() if v)
        utk = sorted(k for k, v in pub.items() if not v)
        print("Publicerat för föräldrarna:", ", ".join(publ) if publ else "inget")
        if utk: print("Fortfarande utkast:", ", ".join(utk))
        warn = [("—", "publicering",
                 "Publicerat: " + (", ".join(publ) if publ else "inget") +
                 (" · utkast: " + ", ".join(utk) if utk else ""))] + warn
    S.write_output(out, players, units, matches, trupp, warn, unavail, st, duties, hemord)

    spel = collections.Counter()
    upp = collections.Counter()
    gm = {u["id"]: u["n"] for u in units}
    for uid, ns in trupp.items():
        for n in ns: spel[n] += gm.get(uid, 1)
    for uid, du in duties.items():
        for r, ns in du.items():
            for n in ns: upp[n] += 1
    print(f"{len(trupp)} truppenheter inlästa · "
          f"matcher per spelare {min(spel.values())}–{max(spel.values())} · "
          f"uppdrag per familj {min(upp.values())}–{max(upp.values())}")
    if warn:
        print(f"{len(warn)} avvikelser (se bladet Varningar):")
        for w in warn[:12]: print("   ", *w)
        if len(warn) > 12: print("    ...")
    else:
        print("Inga avvikelser mot reglerna.")
    print("Skrev", out)

if __name__ == "__main__":
    main()
