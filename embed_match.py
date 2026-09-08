"""
News Event → Risk Register Matching Engine
Deterministic pipeline: embedding retrieval + rule-based contextual scoring.
No LLM anywhere in this process.

Usage (server subprocess):
    python embed_match.py --excel PATH --event-id ID [--cache PATH]
    python embed_match.py --excel PATH --rebuild-cache [--cache PATH]

Outputs JSON to stdout.
"""

import os, sys, json, re, math, argparse
import numpy as np
import openpyxl
from sentence_transformers import SentenceTransformer

# Force UTF-8 on stdout/stderr so Unicode (→, box-drawing chars, etc.) prints
# correctly on Windows, whose default cp1252 console codec cannot encode them.
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except AttributeError:
    # Python < 3.7 fallback
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION — all weights and mappings in one place
# ══════════════════════════════════════════════════════════════════════════════

EMBEDDING_MODEL = "all-MiniLM-L6-v2"
CACHE_FILENAME  = "register_embeddings_cache.json"

WEIGHTS = {
    "semantic":        0.50,
    "direction":       0.15,
    "geography":       0.10,
    "risk_theme":      0.10,
    "event_type":      0.05,
    "market":          0.05,
    "entity_sector":   0.05,
}

# ── Direction categories and keywords ────────────────────────────────────────
# Keywords are lowercased and matched against the combined event direction field + trigger text.
DIRECTION_KEYWORDS = {
    "RATE_HIKE":     ["rate hike","rate hikes","raises rate","raised rate","raises all","rate increase",
                      "increases rate","hiking rate","hawkish","tighter policy","policy tightening",
                      "monetary tightening","raises policy","raise rates","25 bps higher","50 bps higher"],
    "RATE_CUT":      ["rate cut","rate cuts","cuts rate","cut rate","lowers rate","lowered rate",
                      "rate decrease","rate reduction","monetary easing","dovish","accommodative",
                      "policy easing","easing cycle","cut cycle","25 bps lower","50 bps lower"],
    "TIGHTENING":    ["tightening","restriction","sanction","crackdown","enforcement","tighter standards",
                      "restrict","constrain","regulatory tighten"],
    "EASING":        ["easing","stimulus","looser","loosening","relaxing regulatory","relief","forbearance"],
    "INCREASE":      ["increase","increases","rising","surge","jump","spike","growth","expansion","upgrade",
                      "improvement","strengthening","appreciation","inflows"],
    "DECREASE":      ["decrease","decreases","falling","decline","drop","contraction","deterioration",
                      "downgrade","weakening","depreciation","outflows","compression"],
    "DISRUPTION":    ["disruption","outage","failure","cyber attack","cyber incident","hack","breach",
                      "closure","default","bankruptcy","collapse","shutdown","suspension",
                      "infrastructure failure","operational failure","clearing failure"],
    "STRESS":        ["stress","distress","crisis","shock","volatility","instability","pressure",
                      "systemic risk","contagion","liquidity stress","funding stress"],
    "NEUTRAL":       ["decision","announcement","meeting","review","guidance","publication","report"],
}

# Compatibility matrix — (cat_a, cat_b) → score
# Only non-symmetric or notable entries listed; handled symmetrically in code.
DIRECTION_COMPAT = {
    ("RATE_HIKE",  "RATE_HIKE"):  100,
    ("RATE_CUT",   "RATE_CUT"):   100,
    ("TIGHTENING", "TIGHTENING"): 100,
    ("EASING",     "EASING"):     100,
    ("INCREASE",   "INCREASE"):   100,
    ("DECREASE",   "DECREASE"):   100,
    ("DISRUPTION", "DISRUPTION"): 100,
    ("STRESS",     "STRESS"):     100,
    ("NEUTRAL",    "NEUTRAL"):    100,

    ("RATE_HIKE",  "TIGHTENING"): 95,
    ("RATE_CUT",   "EASING"):     95,
    ("DECREASE",   "STRESS"):     80,
    ("DISRUPTION", "STRESS"):     80,
    ("DISRUPTION", "DECREASE"):   70,
    ("STRESS",     "DECREASE"):   80,
    ("TIGHTENING", "DECREASE"):   75,
    ("EASING",     "INCREASE"):   75,

    ("RATE_HIKE",  "RATE_CUT"):  0,
    ("RATE_CUT",   "RATE_HIKE"): 0,
    ("TIGHTENING", "EASING"):    0,
    ("EASING",     "TIGHTENING"):0,
    ("INCREASE",   "DECREASE"):  0,
    ("DECREASE",   "INCREASE"):  0,
    ("RATE_HIKE",  "EASING"):    0,
    ("RATE_CUT",   "TIGHTENING"):0,
}
DIRECTION_COMPAT_DEFAULT_SAME   = 100
DIRECTION_COMPAT_DEFAULT_DIFF   = 40   # different but not opposite
DIRECTION_COMPAT_UNKNOWN        = 50

# ── Geography hierarchy ──────────────────────────────────────────────────────
# Each entry lists its parent chain in order (closest first).
GEO_HIERARCHY = {
    # North America
    "united states":       ["north america", "americas", "global"],
    "usa":                 ["north america", "americas", "global"],
    "us":                  ["north america", "americas", "global"],
    "canada":              ["north america", "americas", "global"],
    "mexico":              ["north america", "latin america", "americas", "global"],

    # Europe
    "united kingdom":      ["europe", "global"],
    "uk":                  ["europe", "global"],
    "britain":             ["europe", "global"],
    "germany":             ["euro area", "europe", "global"],
    "france":              ["euro area", "europe", "global"],
    "italy":               ["euro area", "europe", "global"],
    "spain":               ["euro area", "europe", "global"],
    "netherlands":         ["euro area", "europe", "global"],
    "belgium":             ["euro area", "europe", "global"],
    "austria":             ["euro area", "europe", "global"],
    "portugal":            ["euro area", "europe", "global"],
    "switzerland":         ["europe", "global"],
    "sweden":              ["europe", "global"],
    "norway":              ["europe", "global"],
    "denmark":             ["europe", "global"],

    # Euro Area / Europe regions
    "euro area":           ["europe", "global"],
    "euro zone":           ["europe", "global"],
    "eurozone":            ["europe", "global"],
    "europe":              ["global"],

    # Asia Pacific
    "japan":               ["asia pacific", "asia", "global"],
    "china":               ["asia pacific", "asia", "global"],
    "prc":                 ["asia pacific", "asia", "global"],
    "hong kong":           ["asia pacific", "asia", "global"],
    "singapore":           ["southeast asia", "asia pacific", "asia", "global"],
    "south korea":         ["asia pacific", "asia", "global"],
    "taiwan":              ["asia pacific", "asia", "global"],
    "india":               ["south asia", "asia", "global"],
    "australia":           ["asia pacific", "global"],
    "new zealand":         ["asia pacific", "global"],

    # Middle East
    "saudi arabia":        ["middle east", "global"],
    "uae":                 ["middle east", "global"],
    "united arab emirates":["middle east", "global"],
    "israel":              ["middle east", "global"],
    "turkey":              ["middle east", "europe", "global"],

    # Regions (self + parents)
    "north america":       ["americas", "global"],
    "latin america":       ["americas", "global"],
    "americas":            ["global"],
    "asia pacific":        ["asia", "global"],
    "southeast asia":      ["asia pacific", "asia", "global"],
    "south asia":          ["asia", "global"],
    "asia":                ["global"],
    "middle east":         ["global"],
    "africa":              ["global"],
    "various":             ["global"],
    "global":              [],
    "international":       ["global"],
    "cross-border":        ["global"],
}

GEO_SCORE = {
    "exact":         100,
    "sub_region":    90,
    "region":        80,
    "global_applicable": 70,
    "different":     20,
    "unknown":       50,
}

# ── Risk theme taxonomy synonyms ─────────────────────────────────────────────
# Maps synonyms/abbreviations → canonical L1 code
RISK_THEME_SYNONYMS = {
    "IRRBB":           ["irrbb","interest rate risk banking book","non-traded market","interest rate risk"],
    "MR":              ["mr","market risk","traded market","market risk traded"],
    "LIQ":             ["liq","liquidity","liquidity risk","funding risk","liquidity funding"],
    "CR":              ["cr","credit risk","counterparty credit","ccr","credit"],
    "CCR":             ["ccr","counterparty credit risk","counterparty risk","credit exposure"],
    "OPS":             ["ops","operational risk","ops risk","op risk"],
    "RESIL":           ["resil","resilience","operational resilience","business continuity","bcm"],
    "MODEL":           ["model","model risk","model validation"],
    "CON":             ["con","conduct risk","conduct","customer outcome"],
    "LEGAL":           ["legal","legal risk","litigation","litigation risk"],
    "REG":             ["reg","regulatory","compliance","regulatory risk","regulatory compliance"],
    "REP":             ["rep","reputational","reputational risk","reputation"],
    "STRAT":           ["strat","strategic","strategic risk","business model risk"],
    "CYBER":           ["cyber","cyber risk","cyber security","information security"],
    "FRAUD":           ["fraud","fraud risk","financial crime","financial fraud"],
    "THIRD":           ["third","third party risk","vendor","supplier risk","outsourcing risk"],

    # L2 synonyms → mapped to L1 parent for theme alignment
    "FIRESALE":        ["firesale","fire sale","mtm loss","mark to market","asset fire sale"],
    "PROVISION":       ["provision","ecl","expected credit loss","provisioning"],
    "SPREAD":          ["spread","credit spread","spread widening","credit spread widening"],
    "MODELINV":        ["modelinv","model invalidation","model error"],
    "REDRESS":         ["redress","remediation","customer redress"],
    "LIT":             ["lit","litigation","class action","lawsuit"],
    "FINE":            ["fine","fines","penalties","penalty"],
    "OUTAGE":          ["outage","service outage","system failure","service failure"],
    "FREEZE":          ["freeze","interbank freeze","funding freeze","market freeze"],
    "CONTAGION":       ["contagion","systemic","contagion risk"],
}

# L1 parent relationships for partial credit
L1_PARENTS = {
    "FIRESALE":  ["MR", "LIQ"],
    "PROVISION": ["CR"],
    "SPREAD":    ["CR", "MR"],
    "MODELINV":  ["MODEL"],
    "REDRESS":   ["CON", "LEGAL"],
    "LIT":       ["LEGAL"],
    "FINE":      ["REG", "CON"],
    "OUTAGE":    ["RESIL", "OPS"],
    "FREEZE":    ["LIQ"],
    "CONTAGION": ["LIQ", "CR"],
}

# ── Event type taxonomy ───────────────────────────────────────────────────────
EVENT_TYPE_FAMILIES = {
    "monetary_policy":     ["monetary policy","interest rate decision","rate decision","fomc","mpc decision",
                            "ecb decision","central bank decision","rate announcement","rate change"],
    "regulatory":          ["regulatory change","regulatory announcement","regulation","supervisory action",
                            "regulatory guidance","rule change","prudential","basel"],
    "enforcement":         ["enforcement","enforcement action","fine","penalty","settlement","sanction",
                            "regulatory sanction","legal action"],
    "credit_event":        ["credit event","default","bankruptcy","insolvency","credit deterioration",
                            "rating downgrade","sovereign default","counterparty failure"],
    "market_shock":        ["market shock","market dislocation","market disruption","flash crash",
                            "market event","market stress","exchange failure","clearing failure"],
    "cyber":               ["cyber incident","cyber attack","data breach","ransomware","cyber event",
                            "infrastructure attack","hacking"],
    "geopolitical":        ["geopolitical","war","conflict","sanctions","geopolitical event",
                            "political event","trade war","trade dispute"],
    "natural_disaster":    ["natural disaster","earthquake","flood","hurricane","pandemic","epidemic",
                            "climate event"],
    "macroeconomic":       ["macroeconomic","gdp","inflation","unemployment","economic data",
                            "economic release","recession","economic shock"],
    "corporate":           ["corporate event","merger","acquisition","restructuring","earnings",
                            "profit warning","corporate default"],
}

EVENT_TYPE_SCORE = {"exact": 100, "family": 80, "related": 60, "different": 20}

# ── Market taxonomy ──────────────────────────────────────────────────────────
MARKET_SYNONYMS = {
    "interest_rates":  ["interest rates","rates","short rates","overnight rate","policy rate","libor","sofr","sonia"],
    "fixed_income":    ["fixed income","bonds","government bonds","gilts","treasuries","corporate bonds","credit"],
    "fx":              ["fx","foreign exchange","currencies","currency","forex","usd","eur","gbp"],
    "equities":        ["equities","equity","stocks","shares","stock market","equity market"],
    "credit":          ["credit","credit markets","credit spread","cds","corporate credit"],
    "commodities":     ["commodities","commodity","oil","energy","metals","gold"],
    "payments":        ["payments","payments infrastructure","clearing","settlement","payment systems"],
    "derivatives":     ["derivatives","swaps","futures","options","structured products"],
    "money_markets":   ["money markets","repo","interbank","short term funding","commercial paper"],
}

MARKET_RELATED = {
    "interest_rates": ["fixed_income","fx","derivatives","money_markets"],
    "fixed_income":   ["interest_rates","credit","derivatives"],
    "fx":             ["interest_rates","derivatives"],
    "credit":         ["fixed_income","derivatives"],
    "equities":       ["derivatives"],
}

# ── Entity/sector taxonomy ───────────────────────────────────────────────────
ENTITY_SYNONYMS = {
    "banks":           ["banks","banking","bank","commercial bank","investment bank","retail bank","lender"],
    "central_banks":   ["central bank","central banks","federal reserve","ecb","bank of england","boe","fed"],
    "insurers":        ["insurer","insurers","insurance","insurance company"],
    "corporates":      ["corporates","corporate","companies","firms","non-financial","sme","businesses"],
    "households":      ["households","consumers","retail","individuals","personal"],
    "asset_managers":  ["asset manager","asset managers","fund","fund manager","investment manager","hedge fund"],
    "market_infra":    ["market infrastructure","exchange","clearinghouse","ccp","depository","custodian"],
    "regulators":      ["regulators","regulator","supervisory","regulatory authority","pra","fca","sec"],
    "sovereigns":      ["sovereign","government","treasury","public sector"],
    "energy":          ["energy","oil company","gas","utility","utilities"],
    "transport":       ["transport","airline","shipping","logistics"],
}

ENTITY_FAMILIES = {
    "banks":         ["banks","central_banks"],
    "central_banks": ["banks","central_banks","regulators"],
    "financial":     ["banks","asset_managers","insurers","market_infra"],
}

# ══════════════════════════════════════════════════════════════════════════════
# STAGE 1 — INITIAL CAUSE MATCHING TEXT
# ══════════════════════════════════════════════════════════════════════════════

def build_initial_cause_text(ev):
    """
    Actor + initiating action + direction + magnitude + immediate cause only.
    No downstream markets, L1/L2 risk labels, credit/liquidity consequences.
    """
    parts = []
    headline = (ev.get("Headline") or "").strip()
    if headline:
        parts.append(headline)

    summary = (ev.get("Event Summary") or "").strip()
    if summary:
        # First sentence only — the initiating action
        first = re.split(r'(?<=[.!?])\s+', summary)[0].strip()
        if first and first.lower() != headline.lower():
            parts.append(first)

    direction = (ev.get("Direction / Nature of Change") or "").strip()
    if direction:
        parts.append(direction)

    authority = (ev.get("Regulator / Authority / Organisation") or "").strip()
    if authority:
        parts.append(authority)

    subcat = (ev.get("Event Sub-Category") or ev.get("Event Category") or "").strip()
    if subcat:
        parts.append(subcat)

    return ". ".join(parts)


# ══════════════════════════════════════════════════════════════════════════════
# STAGE 2 — EMBEDDING RETRIEVAL
# ══════════════════════════════════════════════════════════════════════════════

_model = None
def get_model():
    global _model
    if _model is None:
        _model = SentenceTransformer(EMBEDDING_MODEL)
    return _model


def load_register(excel_path):
    wb   = openpyxl.load_workbook(excel_path, data_only=True)
    ws   = wb["Register"]
    rows = list(ws.iter_rows(values_only=True))
    hdrs = [str(h).strip() if h else "" for h in rows[0]]
    data = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        data.append({hdrs[i]: (str(row[i]).strip() if row[i] is not None else "")
                     for i in range(len(hdrs))})
    return data


def load_events(excel_path):
    wb   = openpyxl.load_workbook(excel_path, data_only=True)
    ws   = wb["External News Capture"]
    rows = list(ws.iter_rows(values_only=True))
    hdrs = [str(h).strip() if h else "" for h in rows[0]]
    data = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        d = {hdrs[i]: (str(row[i]).strip() if row[i] is not None else "")
             for i in range(len(hdrs))}
        if d.get("Event ID") or d.get("Headline"):
            data.append(d)
    return data


def get_or_build_cache(register_rows, cache_path):
    """
    Load cached trigger embeddings or recompute.
    Cache invalidated if Register row count changes or model changes.
    """
    triggers = [r.get("Trigger (initial cause)", "") for r in register_rows]
    ids      = [r.get("ID", "") for r in register_rows]

    if os.path.exists(cache_path):
        try:
            with open(cache_path) as f:
                c = json.load(f)
            if (c.get("model") == EMBEDDING_MODEL
                    and c.get("count") == len(triggers)
                    and c.get("ids") == ids):
                embs = np.array(c["embeddings"], dtype=np.float32)
                return embs
        except Exception:
            pass

    model = get_model()
    embs  = model.encode(triggers, normalize_embeddings=True, batch_size=64, show_progress_bar=False)
    cache = {
        "model":      EMBEDDING_MODEL,
        "count":      len(triggers),
        "ids":        ids,
        "embeddings": embs.tolist(),
    }
    with open(cache_path, "w") as f:
        json.dump(cache, f)
    return embs.astype(np.float32)


def semantic_retrieval(initial_cause_text, register_embs, register_rows, top_n=5):
    model    = get_model()
    ev_emb   = model.encode([initial_cause_text], normalize_embeddings=True)[0]
    sims     = (register_embs @ ev_emb).tolist()
    indexed  = sorted(enumerate(sims), key=lambda x: -x[1])
    top      = indexed[:top_n]
    sep      = round(top[0][1] - top[1][1], 4) if len(top) > 1 else None
    results  = []
    for rank, (idx, sim) in enumerate(top, 1):
        results.append({
            "semantic_rank": rank,
            "idx":           idx,
            "sim":           round(sim, 4),
            "sim_display":   round(sim * 100, 1),
            "reg":           register_rows[idx],
        })
    return results, round(sep, 4) if sep is not None else None


# ══════════════════════════════════════════════════════════════════════════════
# STAGE 3 — DETERMINISTIC CONTEXTUAL SCORING
# ══════════════════════════════════════════════════════════════════════════════

def _normalize(text):
    return text.lower().strip() if text else ""


# ── Direction alignment ───────────────────────────────────────────────────────

def _detect_direction(text):
    t = _normalize(text)
    for cat, keywords in DIRECTION_KEYWORDS.items():
        for kw in keywords:
            if kw in t:
                return cat
    return "UNKNOWN"


def direction_alignment(event, reg_row):
    # Build direction source text for event
    ev_dir_field = (event.get("Direction / Nature of Change") or "")
    ev_headline  = (event.get("Headline") or "")
    ev_summary   = (event.get("Event Summary") or "")
    ev_text      = " ".join([ev_dir_field, ev_headline, ev_summary[:200]])

    # Build direction source for trigger
    tr_text = (reg_row.get("Trigger (initial cause)") or "") + " " + (reg_row.get("L1 risk event & description") or "")[:100]

    ev_cat = _detect_direction(ev_text)
    tr_cat = _detect_direction(tr_text)

    if ev_cat == "UNKNOWN" or tr_cat == "UNKNOWN":
        return DIRECTION_COMPAT_UNKNOWN

    key_ab = (ev_cat, tr_cat)
    key_ba = (tr_cat, ev_cat)
    if key_ab in DIRECTION_COMPAT:
        return DIRECTION_COMPAT[key_ab]
    if key_ba in DIRECTION_COMPAT:
        return DIRECTION_COMPAT[key_ba]
    if ev_cat == tr_cat:
        return DIRECTION_COMPAT_DEFAULT_SAME
    return DIRECTION_COMPAT_DEFAULT_DIFF


# ── Geography alignment ───────────────────────────────────────────────────────

def _normalize_geo(text):
    """Return sorted list of lowercase geography tokens from a text."""
    t = _normalize(text)
    # Extract known geo terms
    found = []
    for key in GEO_HIERARCHY:
        if key in t:
            found.append(key)
    return found


def _geo_ancestors(geo_key):
    chain = [geo_key]
    chain.extend(GEO_HIERARCHY.get(geo_key, []))
    return chain


def geography_alignment(event, reg_row):
    ev_text  = _normalize(" ".join([
        event.get("Geography / Region") or "",
        event.get("Country") or "",
    ]))
    tr_text  = _normalize(reg_row.get("Region") or "")

    if not tr_text:
        return GEO_SCORE["unknown"]

    ev_geos = _normalize_geo(ev_text)
    tr_geos = _normalize_geo(tr_text)

    if not ev_geos:
        return GEO_SCORE["unknown"]

    # Check for "global" or "various" in Register
    if any(g in ["global", "various", "international"] for g in tr_geos):
        return GEO_SCORE["global_applicable"]

    # Build full ancestor chains for both sides
    ev_chain = set()
    for g in ev_geos:
        ev_chain.update(_geo_ancestors(g))
    tr_chain = set()
    for g in tr_geos:
        tr_chain.update(_geo_ancestors(g))

    # Exact match
    if set(ev_geos) & set(tr_geos):
        return GEO_SCORE["exact"]

    # Sub-region: one is ancestor of another within 1 hop
    for eg in ev_geos:
        parents = GEO_HIERARCHY.get(eg, [])
        if any(p in tr_geos for p in parents[:1]):
            return GEO_SCORE["sub_region"]
    for tg in tr_geos:
        parents = GEO_HIERARCHY.get(tg, [])
        if any(p in ev_geos for p in parents[:1]):
            return GEO_SCORE["sub_region"]

    # Same region: shared ancestor within 2 hops
    if ev_chain & tr_chain - {"global", "various"}:
        return GEO_SCORE["region"]

    return GEO_SCORE["different"]


# ── Risk theme alignment ──────────────────────────────────────────────────────

def _canonicalize_themes(text):
    """Map text to set of canonical L1 codes."""
    t = _normalize(text)
    codes = set()
    for code, synonyms in RISK_THEME_SYNONYMS.items():
        for syn in synonyms:
            if syn in t:
                # If it's an L2 code, also add its L1 parents
                if code in L1_PARENTS:
                    codes.update(L1_PARENTS[code])
                else:
                    codes.add(code)
    return codes


def risk_theme_alignment(event, reg_row):
    ev_text  = " ".join([
        event.get("Potential L1 Risk Themes") or "",
        event.get("Potential L2 Risk Themes") or "",
    ])
    tr_text  = " ".join([
        reg_row.get("L1 code(s)") or "",
        reg_row.get("L1 risk event & description") or "",
        reg_row.get("L2 code(s)") or "",
        reg_row.get("L2 secondary risk events") or "",
    ])

    ev_codes = _canonicalize_themes(ev_text)
    tr_codes = _canonicalize_themes(tr_text)

    if not ev_codes or not tr_codes:
        return 50   # unknown

    exact  = ev_codes & tr_codes
    if exact:
        ratio = len(exact) / max(len(ev_codes), len(tr_codes))
        if ratio >= 0.5:
            return 100
        return 85

    # Partial overlap — any shared L1 code
    shared = ev_codes & tr_codes
    if shared:
        return 70

    return 20   # no meaningful overlap


# ── Event type alignment ──────────────────────────────────────────────────────

def _detect_event_family(text):
    t = _normalize(text)
    for family, keywords in EVENT_TYPE_FAMILIES.items():
        for kw in keywords:
            if kw in t:
                return family
    return "UNKNOWN"


def event_type_alignment(event, reg_row):
    ev_text = " ".join([
        event.get("Event Category") or "",
        event.get("Event Sub-Category") or "",
        event.get("Event Type") or "",
        event.get("Headline") or "",
    ])
    tr_text = (reg_row.get("Trigger (initial cause)") or "")

    ev_fam = _detect_event_family(ev_text)
    tr_fam = _detect_event_family(tr_text)

    if ev_fam == "UNKNOWN" or tr_fam == "UNKNOWN":
        return 50
    if ev_fam == tr_fam:
        return EVENT_TYPE_SCORE["exact"]

    # Related families
    related_pairs = {
        ("monetary_policy", "macroeconomic"),
        ("monetary_policy", "market_shock"),
        ("credit_event", "market_shock"),
        ("regulatory", "enforcement"),
        ("cyber", "market_shock"),
    }
    pair = (ev_fam, tr_fam)
    rpair = (tr_fam, ev_fam)
    if pair in related_pairs or rpair in related_pairs:
        return EVENT_TYPE_SCORE["related"]

    return EVENT_TYPE_SCORE["different"]


# ── Market alignment ──────────────────────────────────────────────────────────

def _detect_markets(text):
    t = _normalize(text)
    found = set()
    for mkt, keywords in MARKET_SYNONYMS.items():
        if any(kw in t for kw in keywords):
            found.add(mkt)
    return found


def market_alignment(event, reg_row):
    ev_text = " ".join([
        event.get("Affected Market / Asset Class") or "",
        event.get("Headline") or "",
    ])
    tr_text = " ".join([
        reg_row.get("Trigger (initial cause)") or "",
        reg_row.get("L1 risk event & description") or "",
    ])

    ev_mkts = _detect_markets(ev_text)
    tr_mkts = _detect_markets(tr_text)

    if not ev_mkts or not tr_mkts:
        return 50

    exact = ev_mkts & tr_mkts
    if exact:
        ratio = len(exact) / max(len(ev_mkts), len(tr_mkts))
        return round(60 + ratio * 40)

    # Related markets
    related_found = 0
    for em in ev_mkts:
        related = set(MARKET_RELATED.get(em, []))
        if related & tr_mkts:
            related_found += 1
    if related_found:
        return 60

    return 20


# ── Entity/sector alignment ───────────────────────────────────────────────────

def _detect_entities(text):
    t = _normalize(text)
    found = set()
    for ent, keywords in ENTITY_SYNONYMS.items():
        if any(kw in t for kw in keywords):
            found.add(ent)
    return found


def entity_sector_alignment(event, reg_row):
    ev_text = " ".join([
        event.get("Affected Entity / Counterparty") or "",
        event.get("Industry / Sector") or "",
    ])
    tr_text = " ".join([
        reg_row.get("Sector / domain") or "",
        reg_row.get("Impacted entities") or "",
    ])

    ev_ents = _detect_entities(ev_text)
    tr_ents = _detect_entities(tr_text)

    if not ev_ents or not tr_ents:
        return 50

    exact = ev_ents & tr_ents
    if exact:
        ratio = len(exact) / max(len(ev_ents), len(tr_ents))
        return round(60 + ratio * 40)

    # Same sector family
    financial = {"banks", "central_banks", "insurers", "asset_managers", "market_infra"}
    ev_fin = bool(ev_ents & financial)
    tr_fin = bool(tr_ents & financial)
    if ev_fin and tr_fin:
        return 70

    return 20


# ══════════════════════════════════════════════════════════════════════════════
# FINAL SCORE
# ══════════════════════════════════════════════════════════════════════════════

def calculate_final_score(sim_display, ctx_scores):
    s = (
        WEIGHTS["semantic"]      * sim_display +
        WEIGHTS["direction"]     * ctx_scores["direction"] +
        WEIGHTS["geography"]     * ctx_scores["geography"] +
        WEIGHTS["risk_theme"]    * ctx_scores["risk_theme"] +
        WEIGHTS["event_type"]    * ctx_scores["event_type"] +
        WEIGHTS["market"]        * ctx_scores["market"] +
        WEIGHTS["entity_sector"] * ctx_scores["entity_sector"]
    )
    return round(s, 1)


# ══════════════════════════════════════════════════════════════════════════════
# DETERMINISTIC EXPLANATION — template-based, no LLM
# ══════════════════════════════════════════════════════════════════════════════

def _band(score, bands):
    """Return label for first band whose threshold score meets."""
    for threshold, label in bands:
        if score >= threshold:
            return label
    return bands[-1][1]

SEM_BANDS  = [(80,"Strong semantic match"),(60,"Moderate semantic match"),(40,"Partial semantic match"),(0,"Weak semantic match")]
DIR_BANDS  = [(95,"Direction is fully aligned"),(80,"Direction is broadly consistent"),(60,"Direction is compatible"),(25,"Direction is weakly aligned"),(0,"Direction is inconsistent — opposite change")]
GEO_BANDS  = [(90,"Geography is precisely matched"),(70,"Geography is applicable"),(50,"Geography is partially applicable"),(0,"Geography may not apply")]
THEME_BANDS= [(90,"Risk themes are strongly consistent"),(70,"Risk themes are broadly consistent"),(50,"Risk themes partially overlap"),(0,"Risk themes have little overlap")]
TYPE_BANDS = [(95,"Event type is an exact match"),(75,"Event type is in the same family"),(55,"Event type is related"),(0,"Event type differs")]
MKT_BANDS  = [(90,"Market alignment is strong"),(60,"Market alignment is partial"),(0,"Market alignment is weak")]
ENT_BANDS  = [(90,"Entity/sector alignment is strong"),(60,"Entity/sector alignment is partial"),(0,"Entity/sector alignment is weak")]


def generate_explanation(sim_display, ctx, final, reg_row):
    sem   = _band(sim_display,            SEM_BANDS)
    dire  = _band(ctx["direction"],        DIR_BANDS)
    geo   = _band(ctx["geography"],        GEO_BANDS)
    theme = _band(ctx["risk_theme"],       THEME_BANDS)
    etype = _band(ctx["event_type"],       TYPE_BANDS)
    mkt   = _band(ctx["market"],           MKT_BANDS)
    ent   = _band(ctx["entity_sector"],    ENT_BANDS)

    trigger = reg_row.get("Trigger (initial cause)", "")
    l1      = reg_row.get("L1 code(s)", "")
    region  = reg_row.get("Region", "")

    lines = [
        f"{sem} with Register trigger: \"{trigger}\".",
        f"{dire}.",
    ]
    if ctx["geography"] != 50:
        region_note = f" (Register: {region})" if region else ""
        lines.append(f"{geo}{region_note}.")
    lines.append(f"{theme}.")
    if l1:
        lines.append(f"Register L1 risk codes: {l1}.")
    if ctx["event_type"] < 80:
        lines.append(f"{etype}.")

    # Warning if direction opposite
    if ctx["direction"] == 0:
        lines.append("WARNING: Direction is opposite — verify this match carefully.")
    # Warning if semantic separation is very low (ambiguous retrieval)
    if sim_display < 40:
        lines.append("NOTE: Semantic similarity is low — this may be a marginal candidate.")

    return " ".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN MATCHING FUNCTION
# ══════════════════════════════════════════════════════════════════════════════

def match_event(event, register_rows, register_embs, top_n=5):
    initial_cause = build_initial_cause_text(event)

    # Semantic retrieval
    candidates, sep = semantic_retrieval(initial_cause, register_embs, register_rows, top_n)

    # Contextual scoring + final score
    for c in candidates:
        rr = c["reg"]
        ctx = {
            "direction":     direction_alignment(event, rr),
            "geography":     geography_alignment(event, rr),
            "risk_theme":    risk_theme_alignment(event, rr),
            "event_type":    event_type_alignment(event, rr),
            "market":        market_alignment(event, rr),
            "entity_sector": entity_sector_alignment(event, rr),
        }
        c["ctx"]   = ctx
        c["final"] = calculate_final_score(c["sim_display"], ctx)

    # Re-rank by final score; preserve semantic_rank
    candidates.sort(key=lambda x: -x["final"])
    for i, c in enumerate(candidates, 1):
        c["final_rank"] = i

    # Build output records
    results = []
    for c in candidates:
        rr   = c["reg"]
        expl = generate_explanation(c["sim_display"], c["ctx"], c["final"], rr)
        results.append({
            "event_id":            event.get("Event ID", ""),
            "initial_cause_text":  initial_cause,

            "register_id":         rr.get("ID", ""),
            "trigger":             rr.get("Trigger (initial cause)", ""),

            "semantic_rank":       c["semantic_rank"],
            "final_rank":          c["final_rank"],

            "semantic_score":      c["sim_display"],     # cosine × 100, not a probability
            "direction_score":     c["ctx"]["direction"],
            "geography_score":     c["ctx"]["geography"],
            "risk_theme_score":    c["ctx"]["risk_theme"],
            "event_type_score":    c["ctx"]["event_type"],
            "market_score":        c["ctx"]["market"],
            "entity_sector_score": c["ctx"]["entity_sector"],

            "final_score":         c["final"],
            "semantic_separation": sep if c["semantic_rank"] == 1 else None,

            "explanation":         expl,

            # Register pathway — pulled directly, not generated
            "l1_codes":            rr.get("L1 code(s)", ""),
            "l1_risk_event":       rr.get("L1 risk event & description", ""),
            "l2_codes":            rr.get("L2 code(s)", ""),
            "l2_risk_events":      rr.get("L2 secondary risk events", ""),
            "impacted":            rr.get("Impacted entities", ""),
            "severity":            rr.get("Severity (1-5)", ""),
            "transmission":        rr.get("Transmission note", ""),
            "crco_actions":        rr.get("CRCO actions / mitigants", ""),
            "basis":               rr.get("Basis", ""),
            "period":              rr.get("Period", ""),
            "region":              rr.get("Region", ""),
            "sector":              rr.get("Sector / domain", ""),
        })

    return results, sep


# ══════════════════════════════════════════════════════════════════════════════
# CLI ENTRY POINT (called by server subprocess)
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel",         required=True)
    parser.add_argument("--cache",         default=None)
    parser.add_argument("--event-id",      default=None)
    parser.add_argument("--rebuild-cache", action="store_true")
    args = parser.parse_args()

    excel_path = args.excel
    cache_path = args.cache or os.path.join(os.path.dirname(excel_path), CACHE_FILENAME)

    register_rows = load_register(excel_path)

    if args.rebuild_cache:
        if os.path.exists(cache_path):
            os.remove(cache_path)
        register_embs = get_or_build_cache(register_rows, cache_path)
        print(json.dumps({"status": "cache_built", "count": len(register_rows), "model": EMBEDDING_MODEL}))
        return

    register_embs = get_or_build_cache(register_rows, cache_path)

    events = load_events(excel_path)
    event = None
    if args.event_id:
        event = next((e for e in events if e.get("Event ID") == args.event_id), None)
    if event is None and events:
        event = events[0]
    if event is None:
        print(json.dumps({"error": "No event found"}))
        sys.exit(1)

    results, sep = match_event(event, register_rows, register_embs)
    out = {
        "event_id":           event.get("Event ID", ""),
        "headline":           event.get("Headline", ""),
        "initial_cause_text": results[0]["initial_cause_text"] if results else "",
        "semantic_separation": sep,
        "embedding_model":    EMBEDDING_MODEL,
        "weights":            WEIGHTS,
        "candidates":         results,
    }
    print(json.dumps(out, ensure_ascii=False))


# ══════════════════════════════════════════════════════════════════════════════
# INLINE TEST — ECB validation (run directly: python embed_match.py --test)
# ══════════════════════════════════════════════════════════════════════════════

ECB_TEST_EVENT = {
    "Event ID":                              "ECB-TEST-001",
    "Headline":                              "ECB raises all three key policy rates by 25 bps",
    "Event Summary":                         "The ECB Governing Council raised all three key interest rates by 25 basis points. The war in the Middle East was generating inflation pressures. The ECB projected higher inflation and weaker growth.",
    "Event Category":                        "Monetary Policy",
    "Event Sub-Category":                    "Interest Rate Decision",
    "Event Type":                            "Decision",
    "Direction / Nature of Change":          "Rate Hike / Monetary Policy Tightening",
    "Geography / Region":                    "Europe / Euro Area",
    "Country":                               "Euro Zone",
    "Regulator / Authority / Organisation":  "European Central Bank (ECB) / Governing Council",
    "Affected Market / Asset Class":         "Interest Rates / Fixed Income / FX / Equities / Credit",
    "Affected Entity / Counterparty":        "Banks / Financial Institutions / Corporates / Households",
    "Potential L1 Risk Themes":              "Market Risk; IRRBB; Liquidity Risk; Credit Risk",
    "Potential L2 Risk Themes":              "Interest Rate / Yield Curve Repricing; NII Impact; Valuation Risk",
}

if __name__ == "__main__":
    if "--test" in sys.argv:
        EXCEL = r"C:\Users\703313047\OneDrive - Genpact\Desktop\Projects\RiskIntelligencePlatform-v2\RiskIntelligencePlatform-v2\data\scenario\RiskRegisterwithExternalEvent.xlsx"
        CACHE = r"C:\Users\703313047\OneDrive - Genpact\Desktop\Projects\RiskIntelligencePlatform-v2\RiskIntelligencePlatform-v2\data\scenario\register_embeddings_cache.json"
        register_rows = load_register(EXCEL)
        register_embs = get_or_build_cache(register_rows, CACHE)
        results, sep  = match_event(ECB_TEST_EVENT, register_rows, register_embs)

        print("=" * 72)
        print(f"ECB VALIDATION TEST — {EMBEDDING_MODEL}")
        print("=" * 72)
        print(f"\nInitial Cause Text:")
        print(f'  "{results[0]["initial_cause_text"]}"')
        print(f"\nSemantic Separation (Rank#1 − Rank#2): {sep} ({sep*100:.1f} pts)")

        print(f"\n{'─'*72}")
        print(f"TOP 5 — Semantic Rank (before contextual re-ranking)")
        sem_sorted = sorted(results, key=lambda x: x["semantic_rank"])
        for r in sem_sorted:
            print(f"  #{r['semantic_rank']} [{r['register_id']}]  Sem={r['semantic_score']:.1f}  "
                  f"Trigger: {r['trigger'][:65]}")

        print(f"\n{'─'*72}")
        print(f"TOP 5 — Contextual Scores")
        print(f"  {'ID':<12} {'Sem':>5} {'Dir':>5} {'Geo':>5} {'Theme':>6} {'Type':>5} {'Mkt':>5} {'Ent':>5} {'Final':>7}")
        print(f"  {'-'*12} {'-'*5} {'-'*5} {'-'*5} {'-'*6} {'-'*5} {'-'*5} {'-'*5} {'-'*7}")
        for r in sem_sorted:
            print(f"  {r['register_id']:<12} {r['semantic_score']:>5.1f} "
                  f"{r['direction_score']:>5} {r['geography_score']:>5} {r['risk_theme_score']:>6} "
                  f"{r['event_type_score']:>5} {r['market_score']:>5} {r['entity_sector_score']:>5} "
                  f"{r['final_score']:>7.1f}%")

        print(f"\n{'─'*72}")
        print(f"TOP 5 — Final Ranking (after contextual re-ranking)")
        for r in results:
            flag = " ◄ ECB TARGET" if r["register_id"] == "TRR-0302" else ""
            print(f"  Final#{r['final_rank']} (was Sem#{r['semantic_rank']}) [{r['register_id']}]  "
                  f"Final={r['final_score']:.1f}%  Sem={r['semantic_score']:.1f}{flag}")
            print(f"    Trigger: {r['trigger']}")

        print(f"\n{'─'*72}")
        print(f"DETERMINISTIC EXPLANATION — Rank #1")
        top = results[0]
        print(f"  {top['explanation']}")

        print(f"\n{'─'*72}")
        print(f"REGISTER PATHWAY — {top['register_id']} (from Register, not generated)")
        print(f"  L1: {top['l1_codes']} — {top['l1_risk_event'][:80]}")
        print(f"  L2: {top['l2_codes']}")
        print(f"  L2 events: {top['l2_risk_events'][:100]}")
        print(f"  Impacted: {top['impacted'][:80]}")
        print(f"  Severity: {top['severity']}  Period: {top['period']}  Basis: {top['basis']}")
        print(f"  Transmission: {top['transmission'][:120]}")
        print(f"  CRCO actions: {top['crco_actions'][:100]}")
        print(f"\n{'='*72}")
        print(f"Cosine scores are NOT probabilities — labelled 'Semantic Similarity Score'")
        print(f"Weights: {WEIGHTS}")
    else:
        main()
