#!/usr/bin/env python3
"""Controllo qualità dell'anagrafica Snoots (clienti + animali).

Input: i file JSON restituiti dagli strumenti MCP Snoots
  search-customers (s=1000, tutte le pagine) -> file con "customers" nel nome
  search-patients  (s=1000, tutte le pagine) -> file con "patients"  nel nome

Uso:
  python3 controllo_anagrafica.py --input <cartella_json> --out <cartella_output>

Produce:
  Controllo_Anagrafica_<data>.xlsx   report dettagliato (un foglio per controllo)
  dashboard_data.json                riepilogo per la dashboard (con storico mensile)
"""
import argparse
import collections
import datetime as dt
import difflib
import glob
import json
import os
import re
import unicodedata

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TODAY = dt.date.today()

# ---------------------------------------------------------------- caricamento


def load(input_dir, kind):
    docs = {}
    for f in sorted(glob.glob(os.path.join(input_dir, f"*{kind}*"))):
        data = json.load(open(f, encoding="utf-8"))
        for d in data["docs"]:
            docs[d["_id"]] = d
    return docs


# ---------------------------------------------------------------- normalizzazione


def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def norm_text(s):
    s = strip_accents((s or "").upper())
    return re.sub(r"[^A-Z0-9 ]", " ", s).split()


def norm_name(s):
    return " ".join(norm_text(s))


def norm_phone(p):
    d = re.sub(r"\D", "", p or "")
    if d.startswith("0039"):
        d = d[4:]
    elif d.startswith("39") and len(d) >= 11 and d[2] in "03":
        d = d[2:]
    return d


def phone_valid(d):
    if not d:
        return True
    if d.startswith("3"):
        return len(d) in (9, 10)
    if d.startswith("0"):
        return 6 <= len(d) <= 11
    return False


EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$")
EMAIL_TYPOS = {
    "gmial.com", "gmai.com", "gmail.it", "gamil.com", "gmail.con", "gmail.co", "gmal.com",
    "hotmail.con", "hotmai.it", "hotmial.it", "libero.ti", "libero.com", "libro.it",
    "liberi.it", "yahoo.con", "yaho.it", "virgilio.com", "tiscali.com", "alice.com",
}


def email_problem(e):
    if not e:
        return None
    e = e.strip()
    if not EMAIL_RE.match(e):
        return "formato non valido"
    dom = e.lower().rsplit("@", 1)[1]
    if dom in EMAIL_TYPOS:
        return f"dominio sospetto ({dom})"
    return None


PLACEHOLDER = {"", ".", "..", "-", "--", "X", "XX", "XXX", "?", "NN", "N D", "ND", "''", '""', "0", "1",
               "SCONOSCIUTO", "NESSUNO", "TEST", "PROVA"}


def is_placeholder(s):
    v = (s or "").strip().upper()
    return v in PLACEHOLDER or len(re.sub(r"[^A-Z]", "", strip_accents(v))) < 2


# ---------------------------------------------------------------- codice fiscale

CF_RE = re.compile(r"^[A-Z]{6}[0-9LMNPQRSTUV]{2}[ABCDEHLMPRST][0-9LMNPQRSTUV]{2}[A-Z][0-9LMNPQRSTUV]{3}[A-Z]$")
_ODD = dict(zip("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ",
                [1, 0, 5, 7, 9, 13, 15, 17, 19, 21, 1, 0, 5, 7, 9, 13, 15, 17, 19, 21, 2, 4, 18, 20, 11, 3, 6, 8,
                 12, 14, 16, 10, 22, 25, 24, 23]))
_EVEN = {c: (int(c) if c.isdigit() else ord(c) - 65) for c in "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"}
_OMO = dict(zip("LMNPQRSTUV", "0123456789"))


def cf_checksum_ok(cf):
    s = sum(_ODD[c] if i % 2 == 0 else _EVEN[c] for i, c in enumerate(cf[:15]))
    return chr(65 + s % 26) == cf[15]


def _cons_vow(s):
    s = re.sub(r"[^A-Z]", "", strip_accents((s or "").upper()))
    return [c for c in s if c not in "AEIOU"], [c for c in s if c in "AEIOU"]


def cf_surname_code(s):
    c, v = _cons_vow(s)
    return "".join((c + v + ["X"] * 3)[:3])


def cf_name_code(s):
    c, v = _cons_vow(s)
    if len(c) >= 4:
        return c[0] + c[2] + c[3]
    return "".join((c + v + ["X"] * 3)[:3])


def cf_info(cf):
    """Ritorna (problema, sesso_da_cf, data_nascita_da_cf)."""
    if not CF_RE.match(cf):
        return "formato non valido", None, None
    if not cf_checksum_ok(cf):
        return "carattere di controllo errato", None, None
    day = int("".join(_OMO.get(c, c) for c in cf[9:11]))
    sex = "female" if day > 40 else "male"
    return None, sex, day


def piva_ok(p):
    p = re.sub(r"^IT", "", (p or "").upper().strip())
    if not re.fullmatch(r"\d{11}", p):
        return False
    s = 0
    for i, ch in enumerate(p[:10]):
        n = int(ch)
        if i % 2 == 1:
            n *= 2
            if n > 9:
                n -= 9
        s += n
    return (10 - s % 10) % 10 == int(p[10])


# ---------------------------------------------------------------- helpers


def cust_label(c):
    if c.get("businessName"):
        return c["businessName"]
    return " ".join(x for x in [c.get("lastName"), c.get("firstName")] if x) or "(senza nome)"


def addr(c):
    a = c.get("address") or {}
    return a.get("address") or "", a.get("city") or "", a.get("zipCode") or "", a.get("state") or ""


def pdate(s):
    if not s:
        return None
    try:
        return dt.date.fromisoformat(s[:10])
    except ValueError:
        return None


class UF:
    def __init__(self):
        self.p = {}

    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a, b):
        self.p[self.find(a)] = self.find(b)


LEVEL_ORDER = {"Alta": 0, "Media": 1, "Bassa": 2}
IMPORTER = "API IMPORTER"


def fmt_date(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}"


def last_invoice_by_customer(I):
    """Data dell'ultima fattura emessa per cliente e data della fattura più vecchia presente."""
    last, first = {}, None
    for inv in I.values():
        cid, d = (inv.get("customer") or {}).get("id"), inv.get("emittedAt") or ""
        if not cid or not d:
            continue
        if d > last.get(cid, ""):
            last[cid] = d
        if first is None or d < first:
            first = d
    return last, first

# ---------------------------------------------------------------- controlli clienti


def check_customers(C, P, last_inv, inv_since):
    issues = []  # (cliente, categoria, gravità, campo, dettaglio)

    def add(c, cat, grav, campo, det=""):
        issues.append((c, cat, grav, campo, det))

    pets_by_owner = collections.defaultdict(list)
    for p in P.values():
        pets_by_owner[(p.get("owner") or {}).get("_id")].append(p)

    for c in C.values():
        t = c.get("type")
        street, city, zipc, prov = addr(c)
        ssn = (c.get("ssn") or "").strip().upper()
        if t == "individual":
            if is_placeholder(c.get("lastName")):
                add(c, "Dato mancante", "Alta", "Cognome", repr(c.get("lastName") or ""))
            if is_placeholder(c.get("firstName")):
                add(c, "Dato mancante", "Alta", "Nome", repr(c.get("firstName") or ""))
            if not ssn and not c.get("withoutSsn"):
                add(c, "Dato mancante", "Alta", "Codice fiscale")
        else:
            if is_placeholder(c.get("businessName")):
                add(c, "Dato mancante", "Alta", "Ragione sociale",
                    f"tipo '{t}' senza ragione sociale e senza nome")
            if not c.get("vatNumber") and t == "business":
                add(c, "Dato mancante", "Alta", "Partita IVA")
            if not ssn:
                add(c, "Dato mancante", "Media", "Codice fiscale")
        if not c.get("phone"):
            add(c, "Dato mancante", "Alta", "Telefono")
        if not c.get("email"):
            add(c, "Dato mancante", "Media", "Email")
        if is_placeholder(street.strip(" ,")):
            add(c, "Dato mancante", "Media", "Indirizzo (via)")
        if not city:
            add(c, "Dato mancante", "Media", "Città")
        if not zipc:
            add(c, "Dato mancante", "Bassa", "CAP")
        if not prov:
            add(c, "Dato mancante", "Bassa", "Provincia")
        if t == "individual" and c.get("gender") in (None, "unknown"):
            add(c, "Dato mancante", "Bassa", "Sesso")

        # --- coerenza / formato
        if ssn:
            if len(ssn) == 16:
                prob, sex, _ = cf_info(ssn)
                if prob:
                    add(c, "Dato errato", "Alta", "Codice fiscale", f"{ssn}: {prob}")
                elif t == "individual":
                    exp = cf_surname_code(c.get("lastName")) + cf_name_code(c.get("firstName"))
                    if not is_placeholder(c.get("lastName")) and not is_placeholder(c.get("firstName")) \
                            and ssn[:6] != exp:
                        swapped = cf_surname_code(c.get("firstName")) + cf_name_code(c.get("lastName"))
                        if ssn[:6] == swapped:
                            add(c, "Dato errato", "Media", "Nome/Cognome",
                                "nome e cognome scambiati tra i due campi (verificato col CF)")
                        elif ssn[:3] == exp[:3]:
                            add(c, "Dato errato", "Alta", "Codice fiscale",
                                f"{ssn}: il nome non corrisponde (atteso {exp}…): nome scritto diversamente "
                                f"o CF di un familiare")
                        else:
                            add(c, "Dato errato", "Alta", "Codice fiscale",
                                f"{ssn} non corrisponde a nome/cognome (atteso {exp}…): CF di un'altra persona?")
                    g = c.get("gender")
                    if g in ("male", "female") and g != sex:
                        add(c, "Dato errato", "Media", "Sesso",
                            f"in anagrafica '{'M' if g == 'male' else 'F'}', dal CF risulta "
                            f"'{'M' if sex == 'male' else 'F'}'")
            elif len(ssn) == 11:
                if not piva_ok(ssn):
                    add(c, "Dato errato", "Alta", "Codice fiscale", f"{ssn}: numerico non valido")
                elif t == "individual":
                    add(c, "Dato errato", "Media", "Tipo cliente",
                        f"privato con CF numerico ({ssn}): probabilmente è un'azienda/associazione")
            else:
                add(c, "Dato errato", "Alta", "Codice fiscale", f"{ssn}: lunghezza {len(ssn)}")
        if c.get("vatNumber") and not piva_ok(c["vatNumber"]):
            add(c, "Dato errato", "Alta", "Partita IVA", f"{c['vatNumber']}: non valida")
        if t != "individual" and len(ssn) == 16 and not c.get("businessName"):
            add(c, "Dato errato", "Media", "Tipo cliente",
                f"registrato come '{t}' con CF di persona fisica e senza ragione sociale")
        for f in ("phone", "alternatePhone"):
            d = norm_phone(c.get(f))
            if c.get(f) and not phone_valid(d):
                add(c, "Dato errato", "Alta", "Telefono" if f == "phone" else "Telefono alternativo",
                    f"{c.get(f)}: numero non valido")
        for f in ("email", "alternateEmail"):
            pr = email_problem(c.get(f))
            if pr:
                add(c, "Dato errato", "Alta", "Email", f"{c.get(f)}: {pr}")
        if zipc and not re.fullmatch(r"\d{5}", zipc):
            add(c, "Dato errato", "Media", "CAP", f"{zipc}: deve avere 5 cifre")
        if prov and not re.fullmatch(r"[A-Za-z]{2}", prov):
            add(c, "Dato errato", "Bassa", "Provincia", f"{prov}: sigla non valida")

        if not pets_by_owner.get(c["_id"]):
            if c["_id"] in last_inv:
                add(c, "Dato mancante", "Media", "Animali",
                    f"cliente attivo (ultima fattura {fmt_date(last_inv[c['_id']])}) senza animali registrati")
            elif (c.get("createdBy") or {}).get("name") != IMPORTER:
                add(c, "Dato mancante", "Media", "Animali", "cliente inserito in clinica senza animali registrati")
            else:
                since = f" dal {fmt_date(inv_since)}" if inv_since else ""
                add(c, "Cliente dormiente", "Bassa", "Animali",
                    f"cliente storico senza animali registrati e senza fatture{since}: da richiamare")
    return issues, pets_by_owner


def find_customer_duplicates(C):
    recs = []
    for c in C.values():
        ln, fn = norm_name(c.get("lastName")), norm_name(c.get("firstName"))
        full = norm_name(c.get("businessName")) if c.get("businessName") else f"{ln} {fn}".strip()
        street, city, _, _ = addr(c)
        recs.append({
            "c": c, "id": c["_id"], "full": full, "ln": ln, "fn": fn,
            "named": not (is_placeholder(full) or (not c.get("businessName") and
                                                    (is_placeholder(ln) or is_placeholder(fn)))),
            "ssn": (c.get("ssn") or "").strip().upper(),
            "vat": re.sub(r"^IT", "", (c.get("vatNumber") or "").upper().strip()),
            "phones": {norm_phone(c.get(f)) for f in ("phone", "alternatePhone") if norm_phone(c.get(f))},
            "emails": {(c.get(f) or "").strip().lower() for f in ("email", "alternateEmail") if c.get(f)},
            "street": norm_name(street), "city": norm_name(city),
        })
    by_id = {r["id"]: r for r in recs}
    pairs = {}  # (a,b) -> [motivi], livello

    def cf_near(x, y):
        # stessa persona con CF scritto male: cambia solo nome/cognome o solo il comune di nascita
        return len(x) == len(y) == 16 and (
            (x[6:15] == y[6:15] and sum(p != q for p, q in zip(x, y)) <= 3) or x[:11] == y[:11])

    def conflict(a, b):
        """Due schede che sono certamente persone/aziende diverse."""
        if len(a["ssn"]) == 16 and len(b["ssn"]) == 16 and a["ssn"] != b["ssn"] and not cf_near(a["ssn"], b["ssn"]):
            return True
        return bool(a["vat"] and b["vat"] and a["vat"] != b["vat"])

    def similar(a, b):
        return a["named"] and b["named"] and (
            a["full"] == b["full"] or difflib.SequenceMatcher(None, a["full"], b["full"]).ratio() >= 0.85
            or (a["ln"] and a["ln"] == b["ln"] and a["fn"][:1] == b["fn"][:1]))

    def add_pair(a, b, reason, level):
        if a == b:
            return
        k = tuple(sorted((a, b)))
        cur = pairs.setdefault(k, {"reasons": set(), "level": "Bassa"})
        cur["reasons"].add(reason)
        if LEVEL_ORDER[level] < LEVEL_ORDER[cur["level"]]:
            cur["level"] = level

    def index(keyfn):
        idx = collections.defaultdict(list)
        for r in recs:
            for k in keyfn(r):
                if k:
                    idx[k].append(r["id"])
        return idx

    shared_contacts = []  # contatti condivisi da troppi clienti
    for ssn, ids in index(lambda r: [r["ssn"]]).items():
        for i in ids[1:]:
            add_pair(ids[0], i, "stesso codice fiscale", "Alta")
    for vat, ids in index(lambda r: [r["vat"]]).items():
        for i in ids[1:]:
            add_pair(ids[0], i, "stessa partita IVA", "Alta")

    # CF quasi uguali: stessa data e comune di nascita, differenza di 1-2 caratteri (errore di battitura)
    for k, ids in index(lambda r: [r["ssn"][6:15], r["ssn"][:11]] if len(r["ssn"]) == 16 else []).items():
        for i in range(len(ids)):
            for j in range(i + 1, len(ids)):
                a, b = by_id[ids[i]], by_id[ids[j]]
                if a["ssn"] != b["ssn"] and cf_near(a["ssn"], b["ssn"]) and similar(a, b):
                    add_pair(a["id"], b["id"], "codici fiscali quasi uguali (errore di battitura?)", "Alta")

    for label, keyfn in (("telefono", lambda r: r["phones"]), ("email", lambda r: r["emails"])):
        for k, ids in index(keyfn).items():
            if len(ids) < 2:
                continue
            if len(ids) > 6:
                shared_contacts.append((label, k, ids))
                continue
            for i in range(len(ids)):
                for j in range(i + 1, len(ids)):
                    a, b = by_id[ids[i]], by_id[ids[j]]
                    if conflict(a, b):
                        add_pair(a["id"], b["id"], f"stesso {label} ma codici fiscali diversi (familiari?)", "Bassa")
                    elif a["full"] == b["full"] and a["named"]:
                        add_pair(a["id"], b["id"], f"stesso {label} e stesso nome", "Alta")
                    elif similar(a, b):
                        add_pair(a["id"], b["id"], f"stesso {label} e nome simile", "Media")
                    else:
                        add_pair(a["id"], b["id"], f"stesso {label}, nome diverso (familiare?)", "Bassa")

    # nomi uguali o simili (blocco per iniziali cognome)
    blocks = collections.defaultdict(list)
    for r in recs:
        if r["named"]:
            blocks[r["full"][:2]].append(r)
    for blk in blocks.values():
        for i in range(len(blk)):
            for j in range(i + 1, len(blk)):
                a, b = blk[i], blk[j]
                same_place = bool(a["street"]) and a["street"] == b["street"]
                if a["full"] == b["full"]:
                    if conflict(a, b):
                        add_pair(a["id"], b["id"], "stesso nome, codici fiscali diversi (omonimi)", "Bassa")
                    elif same_place:
                        add_pair(a["id"], b["id"], "stesso nome e stesso indirizzo", "Alta")
                    else:
                        add_pair(a["id"], b["id"], "stesso nome e cognome", "Media")
                elif abs(len(a["full"]) - len(b["full"])) <= 2 and not conflict(a, b) and \
                        difflib.SequenceMatcher(None, a["full"], b["full"]).ratio() >= 0.92:
                    add_pair(a["id"], b["id"], "nome quasi identico (errore di battitura?)", "Media")
    # nome/cognome invertiti
    swapped = collections.defaultdict(list)
    for r in recs:
        if r["named"] and r["ln"] and r["fn"]:
            swapped[(r["ln"], r["fn"])].append(r["id"])
    for r in recs:
        if r["named"] and r["ln"] and r["fn"] and r["ln"] != r["fn"]:
            for o in swapped.get((r["fn"], r["ln"]), []):
                if not conflict(r, by_id[o]):
                    add_pair(r["id"], o, "nome e cognome invertiti", "Media")

    # gruppi: solo coppie Alta/Media, mai unendo persone con CF/P.IVA in conflitto
    uf = UF()
    members_of = collections.defaultdict(set)
    ambiguous = collections.defaultdict(set)
    for (a, b), v in sorted(pairs.items(), key=lambda kv: LEVEL_ORDER[kv[1]["level"]]):
        if v["level"] == "Bassa":
            continue
        ra, rb = uf.find(a), uf.find(b)
        if ra == rb:
            continue
        ga, gb = members_of[ra] or {a}, members_of[rb] or {b}
        if any(conflict(by_id[x], by_id[y]) for x in ga for y in gb):
            ambiguous[a].add(b)
            ambiguous[b].add(a)
            continue
        uf.union(a, b)
        members_of[uf.find(a)] = ga | gb
    group_list = []
    for ms in members_of.values():
        if len(ms) < 2 or uf.find(next(iter(ms))) not in members_of or members_of[uf.find(next(iter(ms)))] != ms:
            continue
        ms = sorted(ms)
        reasons, level = set(), "Media"
        for i in range(len(ms)):
            for j in range(i + 1, len(ms)):
                v = pairs.get((ms[i], ms[j]))
                if v and v["level"] != "Bassa":
                    reasons |= v["reasons"]
                    if v["level"] == "Alta":
                        level = "Alta"
        amb = {o for m in ms for o in ambiguous.get(m, ()) if o not in ms}
        if amb:
            reasons.add("attenzione: somiglia anche a " + ", ".join(sorted(C[o]["number"] for o in amb))
                        + " (persona diversa per CF): verificare a quale appartiene")
        group_list.append({"members": ms, "reasons": sorted(reasons), "level": level})
    group_list.sort(key=lambda g: (LEVEL_ORDER[g["level"]], -len(g["members"])))
    low_pairs = [(a, b, v) for (a, b), v in pairs.items() if v["level"] == "Bassa"]
    return group_list, low_pairs, shared_contacts


DORMANT_DAYS = 365
MOTIVI_DORMIENTI = ["Cliente storico senza animali", "Con animali, nessuna fattura in Snoots",
                    "Con animali, ultima fattura oltre 12 mesi fa"]


def find_dormant(C, pets_by_owner, last_inv, no_pets_dormant, dup_groups):
    """Clienti da richiamare: storici senza animali, o con animali vivi e nessuna fattura da oltre 12 mesi.
    Esclusi i clienti con soli animali deceduti e quelli nuovi inseriti in clinica senza fatture.
    Le schede doppie contano come una persona: vale l'ultima fattura e gli animali di tutto il gruppo."""
    cutoff = (TODAY - dt.timedelta(days=DORMANT_DAYS)).isoformat()
    group = {m: g["members"] for g in dup_groups for m in g["members"]}
    out, deceased_only, seen = {}, 0, set()
    for cid, c in C.items():
        members = group.get(cid, [cid])
        if seen & set(members):
            continue
        pets = [p for m in members for p in pets_by_owner.get(m, [])]
        li = max((last_inv.get(m, "") for m in members), default="")
        if pets and cid in no_pets_dormant:
            continue  # gli animali sono sulla sua scheda doppia: la valuta il ramo sotto
        if cid in no_pets_dormant and li:
            continue  # fatture sulla scheda doppia: è attivo
        seen.update(members)
        if cid in no_pets_dormant:
            motivo = MOTIVI_DORMIENTI[0]
        elif not pets or (li and li >= cutoff):
            continue
        elif not li and (c.get("createdBy") or {}).get("name") != IMPORTER:
            continue
        elif all(p.get("deceased") for p in pets):
            deceased_only += 1
            continue
        else:
            motivo = MOTIVI_DORMIENTI[2] if li else MOTIVI_DORMIENTI[1]
        alive = [p.get("name") or "?" for p in pets if not p.get("deceased")]
        out[cid] = {"motivo": motivo, "ultima": li, "animali": ", ".join(alive)}
    order = {m: i for i, m in enumerate(MOTIVI_DORMIENTI)}
    return dict(sorted(out.items(), key=lambda kv: (order[kv[1]["motivo"]], kv[1]["ultima"], cust_label(C[kv[0]])))), \
        deceased_only


# ---------------------------------------------------------------- controlli animali

GENERIC_PET = {"CANE", "GATTO", "GATTA", "GATTINO", "GATTINA", "CUCCIOLO", "CUCCIOLA", "CONIGLIO",
               "SENZA NOME", "SENZANOME", "NN", "NO NAME", "MICIO", "MICIA", "CAGNOLINO", "RANDAGIO"}
SPECIES_IT = {"dog": "Cane", "cat": "Gatto", "rabbit": "Coniglio", "other": "Altro",
              "reptile-tortoise": "Tartaruga"}


def pet_breed(p):
    if p.get("breed"):
        return p["breed"], "campo razza"
    m = re.search(r"Razza: *(.*)", p.get("note") or "")
    if m and m.group(1).strip():
        return m.group(1).strip(), "solo nelle note"
    return "", ""


def check_patients(P, C):
    issues = []

    def add(p, cat, grav, campo, det=""):
        issues.append((p, cat, grav, campo, det))

    for p in P.values():
        name = (p.get("name") or "").strip()
        owner = (p.get("owner") or {}).get("_id")
        if owner not in C:
            add(p, "Proprietario", "Alta", "Proprietario", "proprietario non presente nell'anagrafica clienti")
        if is_placeholder(name.strip("'\"")):
            add(p, "Dato mancante", "Alta", "Nome animale", repr(name))
        elif norm_name(name) in GENERIC_PET:
            add(p, "Dato generico", "Media", "Nome animale", f"nome generico: '{name}'")
        bday = pdate(p.get("birthday"))
        if not bday:
            add(p, "Dato mancante", "Media", "Data di nascita")
        else:
            age = (TODAY - bday).days / 365.25
            if bday > TODAY:
                add(p, "Dato errato", "Alta", "Data di nascita", f"{bday} nel futuro")
            elif not p.get("deceased") and ((p["species"] in ("dog", "cat") and age > 25)
                                             or (p["species"] == "rabbit" and age > 15)):
                add(p, "Dato errato", "Media", "Data di nascita",
                    f"{bday}: età {age:.0f} anni non plausibile (o animale deceduto non segnato)")
        if p.get("sex") in (None, "", "unknown"):
            add(p, "Dato mancante", "Media", "Sesso")
        chip = re.sub(r"\s", "", p.get("microchip") or "")
        if not chip:
            if p["species"] == "dog" and not p.get("deceased"):
                add(p, "Dato mancante", "Alta", "Microchip", "cane senza microchip (obbligatorio per legge)")
        elif not re.fullmatch(r"\d{15}", chip):
            add(p, "Dato errato", "Alta", "Microchip", f"{chip}: deve avere 15 cifre")
        if p.get("sterilized") is None:
            add(p, "Dato mancante", "Bassa", "Sterilizzato")
        breed, where = pet_breed(p)
        if not breed and p["species"] in ("dog", "cat"):
            add(p, "Dato mancante", "Bassa", "Razza")
        if p.get("deceased") and not p.get("deceasedAt"):
            add(p, "Dato mancante", "Bassa", "Data decesso", "deceduto senza data")
    return issues


def find_patient_duplicates(P, cust_group_of):
    groups = []
    seen = set()

    def push(ids, reason, level):
        k = tuple(sorted(ids))
        if k in seen:
            return
        seen.add(k)
        groups.append({"members": list(k), "reasons": [reason], "level": level})

    by_chip = collections.defaultdict(list)
    by_owner_name = collections.defaultdict(list)
    by_name_bday = collections.defaultdict(list)
    for p in P.values():
        chip = re.sub(r"\s", "", p.get("microchip") or "")
        if chip:
            by_chip[chip].append(p["_id"])
        nm = norm_name(p.get("name"))
        if not nm or is_placeholder(nm):
            continue
        owner = (p.get("owner") or {}).get("_id")
        by_owner_name[(owner, nm, p["species"])].append(p["_id"])
        og = cust_group_of.get(owner)
        if og:
            by_owner_name[("G" + og, nm, p["species"])].append(p["_id"])
        if p.get("birthday"):
            by_name_bday[(nm, p["species"], p["birthday"][:10])].append(p["_id"])
    for chip, ids in by_chip.items():
        if len(ids) > 1:
            push(ids, f"stesso microchip {chip}", "Alta")
    for (owner, nm, sp), ids in by_owner_name.items():
        if len(ids) > 1:
            if owner.startswith("G"):
                push(ids, f"stesso nome '{nm}' sotto clienti doppioni", "Alta")
            else:
                push(ids, f"stesso proprietario, stesso nome '{nm}'", "Alta")
    for (nm, sp, b), ids in by_name_bday.items():
        if len(ids) > 1:
            owners = {(P[i].get("owner") or {}).get("_id") for i in ids}
            if len(owners) > 1:
                push(ids, f"stesso nome '{nm}' e stessa data di nascita {b}, proprietari diversi", "Media")
    groups.sort(key=lambda g: LEVEL_ORDER[g["level"]])
    return groups


# ---------------------------------------------------------------- Excel

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(color="FFFFFF", bold=True)
LEVEL_FILL = {"Alta": PatternFill("solid", fgColor="F8CBAD"), "Media": PatternFill("solid", fgColor="FFE699"),
              "Bassa": PatternFill("solid", fgColor="E2EFDA")}
GROUP_FILLS = [PatternFill("solid", fgColor="FFFFFF"), PatternFill("solid", fgColor="EDF2F9")]


def write_sheet(wb, title, headers, rows, widths=None, level_col=None, group_col=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    last_group, shade = None, 0
    for r in rows:
        ws.append(r)
        row = ws.max_row
        if group_col is not None:
            if r[group_col] != last_group:
                shade, last_group = 1 - shade, r[group_col]
            for cell in ws[row]:
                cell.fill = GROUP_FILLS[shade]
        if level_col is not None and r[level_col] in LEVEL_FILL:
            ws.cell(row=row, column=level_col + 1).fill = LEVEL_FILL[r[level_col]]
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(h, max(12, min(45, len(h) + 4)))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def cust_row(c):
    street, city, zipc, prov = addr(c)
    return [c["number"], cust_label(c), c.get("type"), c.get("ssn") or "", c.get("vatNumber") or "",
            c.get("phone") or "", c.get("email") or "", street, city, zipc, prov]


CUST_HEAD = ["Codice cliente", "Cliente", "Tipo", "Codice fiscale", "Partita IVA", "Telefono", "Email",
             "Indirizzo", "Città", "CAP", "Prov."]


def build(input_dir, out_dir):
    C = load(input_dir, "customers")
    P = load(input_dir, "patients")
    I = load(input_dir, "invoices")
    last_inv, inv_since = last_invoice_by_customer(I)
    cust_issues, pets_by_owner = check_customers(C, P, last_inv, inv_since)
    no_pets_dormant = {i[0]["_id"] for i in cust_issues if i[1] == "Cliente dormiente"}
    active_no_pets = {i[0]["_id"] for i in cust_issues if i[3] == "Animali" and i[1] == "Dato mancante"}
    dup_groups, low_pairs, shared = find_customer_duplicates(C)
    cust_group_of = {}
    for gi, g in enumerate(dup_groups, 1):
        for m in g["members"]:
            cust_group_of[m] = f"C{gi}"
    dormant, deceased_only = find_dormant(C, pets_by_owner, last_inv, no_pets_dormant, dup_groups)
    pet_issues = check_patients(P, C)
    pet_dups = find_patient_duplicates(P, cust_group_of)

    wb = Workbook()
    wb.remove(wb.active)

    # --- riepilogo
    ws = wb.create_sheet("Riepilogo")
    ws["A1"] = f"Controllo anagrafica Snoots — {TODAY.strftime('%d/%m/%Y')}"
    ws["A1"].font = Font(bold=True, size=14)
    n_cust_with_issue = {i[0]["_id"] for i in cust_issues if i[1] in ("Dato mancante", "Dato errato")}
    n_pet_with_issue = {i[0]["_id"] for i in pet_issues if i[2] in ("Alta", "Media")}
    dup_cust_records = sum(len(g["members"]) for g in dup_groups)
    summary = [
        ("Clienti totali", len(C)),
        ("Animali totali", len(P)),
        ("Gruppi di clienti doppioni (alta+media)", len(dup_groups)),
        ("  di cui confidenza alta", sum(1 for g in dup_groups if g["level"] == "Alta")),
        ("Schede cliente coinvolte nei doppioni", dup_cust_records),
        ("Schede in più da eliminare/unire", dup_cust_records - len(dup_groups)),
        ("Coppie con contatti condivisi (da verificare, spesso familiari)", len(low_pairs)),
        ("Clienti con almeno un dato mancante o errato", len(n_cust_with_issue)),
        ("Clienti completi e corretti", len(C) - len(n_cust_with_issue)),
        ("Clienti senza animali", sum(1 for c in C if not pets_by_owner.get(c))),
        ("  di cui dormienti da richiamare (nessuna fattura)", len(no_pets_dormant)),
        ("  di cui attivi o nuovi senza animale registrato", len(active_no_pets)),
        ("Clienti dormienti da richiamare (totale)", len(dormant)),
        *[(f"  {m.lower()}", n) for m, n in collections.Counter(d["motivo"] for d in dormant.values()).most_common()],
        ("  esclusi: clienti con soli animali deceduti", deceased_only),
        ("Gruppi di animali doppioni", len(pet_dups)),
        ("Animali con problemi (gravità alta/media)", len(n_pet_with_issue)),
    ]
    for i, (k, v) in enumerate(summary, 3):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=v).font = Font(bold=True)
    r = len(summary) + 5
    ws.cell(row=r, column=1, value="Problemi per campo").font = Font(bold=True, size=12)
    r += 1
    for h, col in (("Entità", 1), ("Tipo problema", 2), ("Campo", 3), ("N. record", 4)):
        c = ws.cell(row=r, column=col, value=h)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
    by_field = collections.Counter(("Cliente", i[1], i[3]) for i in cust_issues if i[1] != "Cliente dormiente")
    by_field.update(("Animale", i[1], i[3]) for i in pet_issues)
    for (ent, cat, campo), n in sorted(by_field.items(), key=lambda x: (x[0][0], -x[1])):
        r += 1
        for col, v in enumerate((ent, cat, campo, n), 1):
            ws.cell(row=r, column=col, value=v)
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 12

    # --- doppioni clienti
    rows = []
    for gi, g in enumerate(dup_groups, 1):
        for m in g["members"]:
            c = C[m]
            rows.append([f"C{gi}", g["level"], "; ".join(g["reasons"])] + cust_row(c) +
                        [len(pets_by_owner.get(m, [])), (c.get("createdAt") or "")[:10]])
    write_sheet(wb, "Doppioni clienti", ["Gruppo", "Confidenza", "Motivo"] + CUST_HEAD + ["N. animali", "Creato il"],
                rows, {"Motivo": 45, "Cliente": 30, "Email": 30, "Indirizzo": 30}, level_col=1, group_col=0)

    rows = []
    for a, b, v in sorted(low_pairs, key=lambda x: sorted(x[2]["reasons"])):
        rows.append([C[a]["number"], cust_label(C[a]), C[b]["number"], cust_label(C[b]),
                     "; ".join(sorted(v["reasons"])), C[a].get("phone") or "", C[a].get("email") or ""])
    for label, k, ids in shared:
        rows.append(["—", f"{len(ids)} clienti", "—", ", ".join(C[i]["number"] for i in ids[:15]),
                     f"stesso {label} usato da {len(ids)} clienti (numero/email generico?)", k if label == "telefono" else "",
                     k if label == "email" else ""])
    write_sheet(wb, "Contatti condivisi", ["Codice A", "Cliente A", "Codice B", "Cliente B", "Motivo",
                                           "Telefono", "Email"], rows,
                {"Cliente A": 30, "Cliente B": 30, "Motivo": 45, "Email": 30})

    # --- problemi clienti
    sev = {"Alta": 0, "Media": 1, "Bassa": 2}
    for title, cats in (("Clienti - dati mancanti", {"Dato mancante"}),
                        ("Clienti - dati errati", {"Dato errato"})):
        rows = [[i[0]["number"], cust_label(i[0]), i[2], i[3], i[4]] + cust_row(i[0])[2:]
                for i in sorted(cust_issues, key=lambda i: (sev[i[2]], i[3], cust_label(i[0])))
                if i[1] in cats]
        write_sheet(wb, title, ["Codice cliente", "Cliente", "Gravità", "Campo", "Dettaglio"] + CUST_HEAD[2:], rows,
                    {"Cliente": 30, "Dettaglio": 50, "Email": 30, "Indirizzo": 30}, level_col=2)

    # vista per cliente: una riga per cliente con l'elenco dei problemi
    per_c = collections.defaultdict(list)
    for i in cust_issues:
        if i[1] != "Cliente dormiente":
            per_c[i[0]["_id"]].append(f"{i[3]}{' (' + i[1].lower() + ')' if i[1] != 'Dato mancante' else ''}")
    rows = []
    for cid, probs in sorted(per_c.items(), key=lambda x: -len(x[1])):
        c = C[cid]
        rows.append([c["number"], cust_label(c), len(probs), ", ".join(probs), len(pets_by_owner.get(cid, [])),
                     cust_group_of.get(cid, "")])
    write_sheet(wb, "Clienti da sistemare", ["Codice cliente", "Cliente", "N. problemi", "Cosa sistemare",
                                             "N. animali", "Gruppo doppione"], rows,
                {"Cliente": 32, "Cosa sistemare": 80})

    def contact(c):
        ph = norm_phone(c.get("phone"))
        ways = (["telefono"] if ph and phone_valid(ph) else []) + (["email"] if c.get("email") and not email_problem(c["email"]) else [])
        return ", ".join(ways) or "nessun contatto valido"

    rows = [[C[cid]["number"], cust_label(C[cid]), d["motivo"], fmt_date(d["ultima"]) if d["ultima"] else "",
             d["animali"], C[cid].get("phone") or "", C[cid].get("email") or "", addr(C[cid])[1], contact(C[cid]),
             "Sì" if C[cid].get("optIn") else "No", cust_group_of.get(cid, "")]
            for cid, d in dormant.items()]
    write_sheet(wb, "Clienti dormienti", ["Codice cliente", "Cliente", "Motivo", "Ultima fattura", "Animali vivi",
                                          "Telefono", "Email", "Città", "Contattabile via",
                                          "Consenso marketing (da verificare)", "Gruppo doppione"], rows,
                {"Cliente": 30, "Motivo": 38, "Animali vivi": 30, "Email": 30, "Contattabile via": 22,
                 "Consenso marketing (da verificare)": 20})

    # --- animali
    def pet_row(p):
        o = C.get((p.get("owner") or {}).get("_id")) or p.get("owner") or {}
        breed, _ = pet_breed(p)
        return [p["number"], p.get("name") or "", SPECIES_IT.get(p["species"], p["species"]), breed,
                (p.get("birthday") or "")[:10], {"male": "M", "female": "F"}.get(p.get("sex"), ""),
                p.get("microchip") or "", "Sì" if p.get("deceased") else "No",
                o.get("number", ""), cust_label(o) if o else ""]

    PET_HEAD = ["Codice animale", "Nome", "Specie", "Razza", "Nascita", "Sesso", "Microchip", "Deceduto",
                "Codice proprietario", "Proprietario"]
    rows = []
    for gi, g in enumerate(pet_dups, 1):
        for m in g["members"]:
            rows.append([f"A{gi}", g["level"], "; ".join(g["reasons"])] + pet_row(P[m]))
    write_sheet(wb, "Doppioni animali", ["Gruppo", "Confidenza", "Motivo"] + PET_HEAD, rows,
                {"Motivo": 50, "Proprietario": 30, "Microchip": 18}, level_col=1, group_col=0)
    rows = [[i[0]["number"], i[0].get("name") or "", i[2], i[1], i[3], i[4]] + pet_row(i[0])[2:]
            for i in sorted(pet_issues, key=lambda i: (sev[i[2]], i[3], i[0].get("name") or ""))]
    write_sheet(wb, "Animali - problemi", ["Codice animale", "Nome", "Gravità", "Tipo", "Campo", "Dettaglio"]
                + PET_HEAD[2:], rows, {"Dettaglio": 45, "Proprietario": 30, "Microchip": 18}, level_col=2)

    breed_notes = sum(1 for p in P.values() if pet_breed(p)[1] == "solo nelle note")
    imported = [p for p in P.values() if (p.get("createdBy") or {}).get("name") == "API IMPORTER"]
    imp_f = sum(1 for p in imported if p.get("sex") == "female")
    imp_m = sum(1 for p in imported if p.get("sex") == "male")
    comma = sum(1 for c in C.values() if addr(c)[0].rstrip().endswith(","))
    sistemici = []
    if imported and imp_f > 3 * max(imp_m, 1):
        sistemici.append(("Sesso degli animali importati probabilmente errato",
                          f"{imp_f} femmine contro {imp_m} maschi tra i {len(imported)} animali importati: "
                          "l'importazione ha impostato 'femmina' come valore predefinito. Il sesso va "
                          "ricontrollato alla prima visita."))
    if breed_notes:
        sistemici.append(("Razza scritta solo nelle note",
                          f"{breed_notes} animali hanno la razza nel campo note ('Razza: …') ma non nel campo "
                          "Razza: non compare nei filtri e nelle statistiche."))
    if comma:
        sistemici.append(("Indirizzi con virgola finale",
                          f"{comma} indirizzi terminano con ',' (residuo dell'importazione): da ripulire, "
                          "compare su lettere e fatture."))
    imp_c = [c for c in C.values() if (c.get("createdBy") or {}).get("name") == IMPORTER]
    opt = sum(1 for c in imp_c if c.get("optIn"))
    if imp_c and opt == len(imp_c):
        sistemici.append(("Consenso marketing da verificare",
                          f"Tutti i {len(imp_c)} clienti importati risultano con consenso marketing: è il valore "
                          "predefinito dell'importazione, non un consenso raccolto. Va verificato prima di usare "
                          "la lista per promozioni o richiami commerciali."))
    ster_f = sum(1 for p in imported if p.get("sterilized") is False)
    if imported and ster_f > 0.9 * len(imported):
        sistemici.append(("Sterilizzazione da verificare",
                          f"{ster_f} animali importati su {len(imported)} risultano 'non sterilizzati': "
                          "probabile valore predefinito dell'importazione."))
    ws = wb["Riepilogo"]
    r = ws.max_row + 2
    ws.cell(row=r, column=1, value="Problemi sistematici (dovuti all'importazione)").font = Font(bold=True, size=12)
    for t, d in sistemici:
        r += 1
        ws.cell(row=r, column=1, value=t).font = Font(bold=True)
        ws.cell(row=r, column=2, value=d)

    os.makedirs(out_dir, exist_ok=True)
    xlsx = os.path.join(out_dir, f"Controllo_Anagrafica_{TODAY.isoformat()}.xlsx")
    wb.save(xlsx)

    # --- dati per la dashboard
    def ci(i):
        return {"n": i[0]["number"], "nome": cust_label(i[0]), "g": i[2], "cat": i[1], "campo": i[3],
                "det": i[4]}

    def pi(i):
        o = C.get((i[0].get("owner") or {}).get("_id")) or {}
        return {"n": i[0]["number"], "nome": i[0].get("name") or "",
                "specie": SPECIES_IT.get(i[0]["species"], i[0]["species"]), "own": o.get("number", ""),
                "ownNome": cust_label(o) if o else "", "g": i[2], "cat": i[1], "campo": i[3], "det": i[4]}

    snapshot = {
        "data": TODAY.isoformat(),
        "clienti": len(C), "animali": len(P),
        "clientiCompleti": len(C) - len(n_cust_with_issue),
        "gruppiDoppioniClienti": len(dup_groups),
        "gruppiDoppioniAlta": sum(1 for g in dup_groups if g["level"] == "Alta"),
        "schedeInPiu": dup_cust_records - len(dup_groups),
        "gruppiDoppioniAnimali": len(pet_dups),
        "clientiSenzaAnimali": sum(1 for c in C if not pets_by_owner.get(c)),
        "clientiDormienti": len(dormant),
        "dormientiSenzaAnimali": len(no_pets_dormant),
        "clientiAttiviSenzaAnimali": len(active_no_pets),
        "animaliConProblemi": len(n_pet_with_issue),
        "caniSenzaChip": sum(1 for i in pet_issues if i[3] == "Microchip" and i[1] == "Dato mancante"),
        "razzaSoloNote": breed_notes,
    }
    hist_path = os.path.join(out_dir, "dashboard_data.json")
    history = []
    if os.path.exists(hist_path):
        history = [h for h in json.load(open(hist_path, encoding="utf-8")).get("storico", [])
                   if h["data"] != snapshot["data"]]
    history.append(snapshot)
    data = {
        "aggiornato": TODAY.isoformat(),
        "kpi": snapshot,
        "storico": sorted(history, key=lambda h: h["data"]),
        "perCampo": [{"ent": e, "cat": c, "campo": f, "n": n} for (e, c, f), n in by_field.most_common()],
        "doppioniClienti": [{"gruppo": f"C{gi}", "livello": g["level"], "motivi": g["reasons"],
                             "membri": [{"n": C[m]["number"], "nome": cust_label(C[m]), "cf": C[m].get("ssn") or "",
                                         "tel": C[m].get("phone") or "", "email": C[m].get("email") or "",
                                         "citta": addr(C[m])[1], "animali": len(pets_by_owner.get(m, []))}
                                        for m in g["members"]]}
                            for gi, g in enumerate(dup_groups, 1)],
        "doppioniAnimali": [{"gruppo": f"A{gi}", "livello": g["level"], "motivi": g["reasons"],
                             "membri": [dict(zip(["n", "nome", "specie", "razza", "nascita", "sesso", "chip",
                                                  "deceduto", "own", "ownNome"], pet_row(P[m])))
                                        for m in g["members"]]}
                            for gi, g in enumerate(pet_dups, 1)],
        "problemiClienti": [ci(i) for i in cust_issues if i[1] != "Cliente dormiente"],
        "dormienti": [{"n": C[c]["number"], "nome": cust_label(C[c]), "citta": addr(C[c])[1],
                       "contatto": contact(C[c]), "motivo": d["motivo"], "ultima": d["ultima"][:10],
                       "animali": d["animali"]} for c, d in dormant.items()],
        "problemiAnimali": [pi(i) for i in pet_issues],
        "contattiCondivisi": len(low_pairs),
        "sistemici": [{"titolo": t, "dettaglio": d} for t, d in sistemici],
    }
    json.dump(data, open(hist_path, "w", encoding="utf-8"), ensure_ascii=False)
    export_dashboard(data, os.path.join(out_dir, "dashboard"))
    write_summary(data, os.path.join(out_dir, f"Riepilogo_{TODAY.isoformat()}.html"))
    return xlsx, hist_path, snapshot, summary


SENSITIVE = [(re.compile(r"\b[A-Z]{6}[0-9LMNPQRSTUV]{2}[A-Z][0-9LMNPQRSTUV]{2}[A-Z][0-9LMNPQRSTUV]{3}[A-Z]\b"), "CF"),
             (re.compile(r"[\w.%+'-]+@[\w.-]+"), "email"),
             (re.compile(r"\+?\d[\d/ ]{7,}\d"), "numero")]


def mask(text):
    """Toglie CF, email e numeri dai testi destinati alla dashboard (i dati completi restano nell'Excel)."""
    for rx, label in SENSITIVE:
        text = rx.sub(f"[{label}]", text)
    return text


SUMMARY_ROWS = [
    ("Clienti in anagrafica", "clienti", None),
    ("Animali in anagrafica", "animali", None),
    ("Clienti con scheda completa e corretta", "clientiCompleti", "up"),
    ("Gruppi di clienti doppioni", "gruppiDoppioniClienti", "down"),
    ("  di cui a confidenza alta", "gruppiDoppioniAlta", "down"),
    ("Schede cliente in più da unire", "schedeInPiu", "down"),
    ("Gruppi di animali doppioni", "gruppiDoppioniAnimali", "down"),
    ("Animali con problemi (gravità alta o media)", "animaliConProblemi", "down"),
    ("Cani senza microchip", "caniSenzaChip", "down"),
    ("Clienti dormienti da richiamare", "clientiDormienti", None),
    ("Clienti attivi senza animale registrato", "clientiAttiviSenzaAnimali", "down"),
]


def write_summary(data, path):
    """Riepilogo breve (senza dati personali) da salvare su Google Drive a ogni controllo."""
    from html import escape
    st = data["storico"]
    cur, prev = st[-1], (st[-2] if len(st) > 1 else None)
    it = lambda n: f"{n:,}".replace(",", ".")
    rows = []
    for label, k, good in SUMMARY_ROWS:
        if k not in cur:
            continue
        delta = ""
        if prev and prev.get(k) is not None:
            d = cur[k] - prev[k]
            delta = "invariato" if d == 0 else f"{'+' if d > 0 else ''}{it(d)}"
            if d and good:
                delta += " (meglio)" if (d > 0) == (good == "up") else " (peggio)"
        rows.append(f"<tr><td>{escape(label.strip())}</td><td>{it(cur[k])}</td><td>{delta}</td></tr>")
    prev_txt = (f"La colonna Variazione confronta con il controllo del {fmt_date(prev['data'])}." if prev
                else "È il primo controllo: le variazioni compariranno dal prossimo.")
    sys_items = "".join(f"<li><b>{escape(x['titolo'])}</b>: {escape(x['dettaglio'])}</li>" for x in data["sistemici"])
    html = f"""<html><body>
<h1>Controllo anagrafica Snoots del {fmt_date(cur['data'])}</h1>
<p>{prev_txt} Le liste complete con nomi e contatti sono nel file
Controllo_Anagrafica_{cur['data']}.xlsx; la dashboard è su
https://claude.ai/artifact/CV8aTA7Xg26R41AgbYBnJc</p>
<h2>Numeri principali</h2>
<table border="1" cellpadding="4"><tr><th>Voce</th><th>Valore</th><th>Variazione</th></tr>
{''.join(rows)}</table>
<h2>Problemi dovuti all'importazione</h2>
<ul>{sys_items}</ul>
</body></html>"""
    open(path, "w", encoding="utf-8").write(html)


def export_dashboard(data, out):
    """Un file JSON per ogni tabella della dashboard, senza CF/telefoni/email."""
    os.makedirs(out, exist_ok=True)
    tables = {
        "storico": data["storico"],
        "per_campo": data["perCampo"],
        "sistemici": data["sistemici"],
        "dormienti": [{"codice": d["n"], "nome": d["nome"], "citta": d["citta"], "contattabile": d["contatto"],
                       "motivo": d["motivo"], "ultima_fattura": d["ultima"], "animali": d["animali"]}
                      for d in data["dormienti"]],
        "doppioni_clienti": [
            {"gruppo": g["gruppo"], "livello": g["livello"], "motivi": mask("; ".join(g["motivi"])),
             "codice": m["n"], "nome": m["nome"], "citta": m["citta"], "animali": m["animali"],
             "ha_cf": bool(m["cf"]), "ha_tel": bool(m["tel"]), "ha_email": bool(m["email"])}
            for g in data["doppioniClienti"] for m in g["membri"]],
        "doppioni_animali": [
            {"gruppo": g["gruppo"], "livello": g["livello"], "motivi": g["motivi"][0] if len(g["motivi"]) == 1
             else "; ".join(g["motivi"]), "codice": m["n"], "nome": m["nome"], "specie": m["specie"],
             "nascita": m["nascita"], "ha_chip": bool(m["chip"]), "deceduto": m["deceduto"],
             "proprietario_codice": m["own"], "proprietario": m["ownNome"]}
            for g in data["doppioniAnimali"] for m in g["membri"]],
        "problemi_clienti": [
            {"id": f"{p['n']}|{p['campo']}|{i}", "codice": p["n"], "nome": p["nome"], "gravita": p["g"],
             "tipo": p["cat"], "campo": p["campo"], "dettaglio": mask(p["det"])}
            for i, p in enumerate(data["problemiClienti"])],
        "problemi_animali": [
            {"id": f"{p['n']}|{p['campo']}|{i}", "codice": p["n"], "nome": p["nome"], "specie": p["specie"],
             "proprietario_codice": p["own"], "proprietario": p["ownNome"], "gravita": p["g"], "tipo": p["cat"],
             "campo": p["campo"], "dettaglio": mask(p["det"])}
            for i, p in enumerate(data["problemiAnimali"])],
    }
    for name, rows in tables.items():
        json.dump(rows, open(os.path.join(out, f"{name}.json"), "w", encoding="utf-8"), ensure_ascii=False)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    xlsx, js, snap, summary = build(a.input, a.out)
    for k, v in summary:
        print(f"{k}: {v}")
    print("Excel:", xlsx)
    print("Dati dashboard:", js)
