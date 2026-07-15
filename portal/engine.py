"""
Care Weekly Dashboard - Analysis Engine

Sheet layout expected in the uploaded Excel file:
  - Kap Tickets data  : Kapture ticket exports
  - SAP tickets        : SAP complaint/request exports
  - SO order           : Service Order (SO) request-type data
  - Other Enquiry      : Additional enquiry rows (counted as R+E, sub-bucket = Others)

Five dashboards produced:
  1. kapture       - Kapture Complaints vs Request+Enquiry breakdown
  2. sap_tickets   - SAP Complaints vs Request+Enquiry+Feedback breakdown
  3. so_output     - SO Request+Enquiry context breakdown
  4. overall_sap   - Combined SAP = SAP tickets + SO + Other Enquiry
  5. final_output  - Grand total = Kapture + Overall SAP
"""

from __future__ import annotations
import re
import io
import pandas as pd
from typing import Any

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SUB_BUCKETS = [
    "Refund",
    "Delivery related",
    "Return",
    "Invoice/Billing related",
    "Repair",
    "Demo & Installation",
    "PMS",
    "Warranty",
    "Others",
]

# SO output requires a specific column ordering (per logic doc)
SO_ORDER = [
    "Demo & Installation",
    "Repair",
    "Delivery related",
    "Invoice/Billing related",
    "Refund",
    "PMS",
    "Warranty",
    "Return",
    "Others",
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _norm(s: Any) -> str:
    """Normalise a value: strip, collapse spaces, lowercase."""
    return re.sub(r"\s+", " ", str(s).strip()).lower()


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Return the first DataFrame column whose normalised name matches any candidate."""
    normed = {_norm(c): c for c in df.columns}
    for cand in candidates:
        cand_n = _norm(cand)
        if cand_n in normed:
            return normed[cand_n]
        # partial / contains match
        for nc, orig in normed.items():
            if cand_n in nc:
                return orig
    return None


def _parse_tat_minutes(val: str) -> float | None:
    """Parse TAT strings like '1 Day(s) 9 Hr(s) 27 Min(s) 1 Sec(s)' into total minutes."""
    try:
        v = str(val).lower()
        days    = int(re.search(r"(\d+)\s*day",  v).group(1)) if re.search(r"\d+\s*day",  v) else 0
        hours   = int(re.search(r"(\d+)\s*hr",   v).group(1)) if re.search(r"\d+\s*hr",   v) else 0
        minutes = int(re.search(r"(\d+)\s*min",  v).group(1)) if re.search(r"\d+\s*min",  v) else 0
        seconds = int(re.search(r"(\d+)\s*sec",  v).group(1)) if re.search(r"\d+\s*sec",  v) else 0
        total   = days * 1440 + hours * 60 + minutes + seconds / 60
        return round(total, 2)
    except Exception:
        # Fallback: try plain numeric
        try:
            return float(val)
        except Exception:
            return None


def _pct(count: int, total: int) -> float:
    return round(count / total * 100, 2) if total else 0.0


# Canonical sub-category name normalization rules (pattern → canonical name)
_SUBCAT_NORM_RULES: list[tuple[list[str], str]] = [
    # Repair
    (["repair visit delayed", "delayed by engineer", "engineer delay"],    "Repair Visit Delayed by Engineer"),
    (["no update on resolution", "no update", "visit done no update"],     "No Update on Resolution"),
    (["part pending", "spare part"],                                        "Part Pending"),
    (["improper repair", "poor quality repair"],                           "Improper Repair"),
    (["extra charges for repair", "overcharging repair"],                  "Extra Charges for Repair"),
    (["repair order request", "repair service request", "repair request"], "Repair Request"),
    # Demo & Installation
    (["delay in engineer visit", "delay in installation"],                 "Delay in Engineer Visit"),
    (["installation order request", "installation service request",
      "installation request", "demo/installation request",
      "demo and installation request", "re-installation request",
      "reinstallation request"],                                           "Installation Request"),
    (["improper installation", "improper demo"],                           "Improper Installation/Demo"),
    (["extra charges for installation", "extra charges taken for install"],"Extra Charges for Installation"),
    # PMS
    (["pms visit delayed", "pms engineer delay", "pms delay"],            "PMS Visit Delayed by Engineer"),
    (["pms service request", "pms request"],                               "PMS Service Request"),
    # Delivery
    (["delay in delivery", "shipment delay", "in transit"],               "Delay in Delivery"),
    (["status marked", "order not received", "delivery status"],          "Order Not Received / Status"),
    (["accessories", "freebie", "booklet not received"],                   "Accessories/Freebie Not Received"),
    (["wrong product"],                                                    "Wrong Product Delivered"),
    # Invoice/Billing
    (["duplicate invoice", "invoice copy"],                               "Duplicate Invoice Copy"),
    (["gst issue", "gst"],                                                 "GST Issue"),
    # Refund
    (["refund not received", "delay in refund"],                           "Refund Not Received/Delayed"),
    (["refund request"],                                                   "Refund Request"),
    # Return/Exchange
    (["change of mind", "return request", "reverse pickup",
      "exchange request"],                                                 "Return/Exchange Request"),
    # Warranty
    (["incorrect warranty"],                                               "Incorrect Warranty Details"),
]

# Standard sub-category name for SO records, keyed by bucket
_SO_SUBCAT_NAMES: dict[str, str] = {
    "Repair":                 "Repair Request",
    "Demo & Installation":    "Installation Request",
    "PMS":                    "PMS Service Request",
    "Delivery related":       "Delivery Request",
    "Invoice/Billing related":"Invoice/Billing Request",
    "Return":                 "Return Request",
    "Warranty":               "Warranty Request",
    "Refund":                 "Refund Request",
    "Others":                 "Other Request",
}


def _normalize_subcat(name: str) -> str:
    """Map raw sub-category names to a canonical form for cross-source comparison."""
    n = name.lower().strip()
    for patterns, canonical in _SUBCAT_NORM_RULES:
        if any(p in n for p in patterns):
            return canonical
    return " ".join(w.capitalize() for w in name.strip().split())


def _collect_rawsubs(group_df: pd.DataFrame, group_total: int, top_n: int = 5) -> dict:
    """Return {bucket: [{name, count, pct}]} of top raw sub-categories per bucket."""
    result: dict = {}
    for bucket in SUB_BUCKETS:
        rows = group_df[group_df["_sub"] == bucket]
        subs = []
        for name, cnt in rows["_rawsub"].value_counts().head(top_n).items():
            n = str(name).strip()
            if n:
                subs.append({"name": n, "count": int(cnt),
                             "pct": _pct(int(cnt), group_total)})
        if subs:
            result[bucket] = subs
    return result


def _build_section(sub_col: pd.Series, order: list[str], total: int) -> dict:
    """Build {bucket: {count, pct}} from a classified sub-bucket Series."""
    counts = sub_col.value_counts().to_dict()
    return {
        bucket: {"count": int(counts.get(bucket, 0)),
                 "pct": _pct(int(counts.get(bucket, 0)), total)}
        for bucket in order
    }

# ---------------------------------------------------------------------------
# Kapture classification
# ---------------------------------------------------------------------------

def _kap_toplevel(vertical: str) -> str:
    return "Complaints" if _norm(vertical) == "complaint" else "Request + Enquiry"


def _kap_re_subtype(vertical: str) -> str:
    """Within Request+Enquiry rows: 'Request' if vertical says so, else 'Enquiry'."""
    return "Request" if "request" in _norm(vertical) else "Enquiry"


def _kap_subbucket(category: str, sub_category: str) -> str:
    c = _norm(category)
    s = _norm(sub_category)

    # --- Explicit high-priority category patterns ---
    if "resq app" in c:
        return "Others"
    # "Reject TR" / "Approved for TR" – TR = Transfer Request (return/exchange flow)
    if any(k in c for k in ["reject tr", "approved for tr", " tr"]):
        return "Return"
    # "Order Cancellation (Change Of Mind)" → Return
    if "change of mind" in c or "order cancellation" in c:
        return "Return"

    if "refund" in c or "reimburs" in c:
        return "Refund"
    # Check "return" BEFORE "resq" so "Return Related - resQ" → Return not Warranty
    if "return" in c:
        return "Return"
    if "repair" in c:
        return "Repair"
    # "Service related - resQ" = after-sales service request → Repair
    if "service related" in c:
        return "Repair"
    if any(k in c for k in ["demo", "installation", "uninstall", "reinstall"]):
        return "Demo & Installation"
    if "delivery" in c or "store - delivery" in c:
        return "Delivery related"
    if any(k in c for k in ["billing", "invoice"]):
        return "Invoice/Billing related"
    if any(k in c for k in ["pms", "preventive"]):
        return "PMS"
    if any(k in c for k in ["warranty", "resq", "insurance", "amtrust", "in-warranty"]):
        return "Warranty"

    # "store related" – context-sensitive via sub-category
    if "store related" in c or "store - " in c:
        if any(k in s for k in ["return", "exchange", "pick up", "pickup", "not picked"]):
            return "Return"
        if any(k in s for k in ["delivery", "wrong product", "dispatch"]):
            return "Delivery related"
        if "repair" in s or "not working" in s or "defective" in s:
            return "Repair"
        return "Others"

    # "product related" / "product related - dc" – context-sensitive
    if "product related" in c:
        if any(k in s for k in ["return", "exchange", "change of mind",
                                  "commercial return", "reverse pickup", "pickup"]):
            return "Return"
        if "wrong product" in s or "delivery" in s:
            return "Delivery related"
        if "warranty" in s:
            return "Warranty"
        return "Others"

    # Generic exchange
    if "exchange" in c:
        return "Return"

    # Sub-category fallback
    if "refund" in s or "reimburs" in s:
        return "Refund"
    if "repair" in s:
        return "Repair"
    if any(k in s for k in ["delivery", "dispatch", "shipment", "wrong product delivered"]):
        return "Delivery related"
    if any(k in s for k in ["demo", "installation"]):
        return "Demo & Installation"
    if any(k in s for k in ["billing", "invoice", "gst", "emi"]):
        return "Invoice/Billing related"
    if any(k in s for k in ["return", "exchange", "change of mind", "reverse pickup"]):
        return "Return"
    if "pms" in s:
        return "PMS"
    if "warranty" in s:
        return "Warranty"

    return "Others"


def analyze_kapture(df: pd.DataFrame) -> dict:
    vertical_col    = _find_col(df, ["vertical"])
    category_col    = _find_col(df, ["category"])
    subcategory_col = _find_col(df, ["sub category", "sub_category", "subcategory"])
    channel_col     = _find_col(df, ["channel", "source channel", "ticket channel", "source"])
    tat_col         = _find_col(df, ["tat", "tat (min)", "tat(min)",
                                      "resolution time", "handling time", "ticket tat"])

    if not vertical_col:
        raise ValueError("Kapture sheet: 'Vertical' column not found.")
    if not category_col:
        raise ValueError("Kapture sheet: 'Category' column not found.")

    # Exclude rows where the top-level vertical is blank (insufficient context)
    df = df[df[vertical_col].notna() & (df[vertical_col].astype(str).str.strip() != "")].copy()

    cat = df[category_col].fillna("").astype(str)
    sub = df[subcategory_col].fillna("").astype(str) if subcategory_col else pd.Series([""] * len(df))

    df["_top"]    = df[vertical_col].apply(_kap_toplevel)
    df["_sub"]    = [_kap_subbucket(c, s) for c, s in zip(cat, sub)]
    df["_rawsub"] = df[subcategory_col].fillna("").astype(str) if subcategory_col else ""

    comp = df[df["_top"] == "Complaints"]
    req  = df[df["_top"] == "Request + Enquiry"]

    # Request vs Enquiry split within R+E (based on vertical value)
    req_re_type = req[vertical_col].apply(_kap_re_subtype)
    kap_request_count = int((req_re_type == "Request").sum())
    kap_enquiry_count = int((req_re_type == "Enquiry").sum())

    # Top raw sub-categories per bucket (for KPI combined sub-cat analysis)
    re_top_subcats   = _collect_rawsubs(req,  len(req),  top_n=10)
    comp_top_subcats = _collect_rawsubs(comp, len(comp), top_n=10)

    # RD.IN channel metrics
    def _is_rdin(val: str) -> bool:
        return "rdin" in re.sub(r"[\s.\-_]", "", val.lower())

    rdin: dict = {"total": 0, "complaints": 0,
                  "tat_total_minutes": None, "tat_avg_minutes": None, "tat_avg_hours": None}
    if channel_col:
        rdin_mask    = df[channel_col].astype(str).apply(_is_rdin)
        rdin_df      = df[rdin_mask]
        rdin["total"]      = int(len(rdin_df))
        rdin["complaints"] = int((rdin_df["_top"] == "Complaints").sum())

        rdin_comp = rdin_df[rdin_df["_top"] == "Complaints"]
        if tat_col and len(rdin_comp) > 0:
            tat_vals = rdin_comp[tat_col].astype(str).apply(_parse_tat_minutes).dropna()
            if len(tat_vals) > 0:
                rdin["tat_total_minutes"] = round(float(tat_vals.sum()), 2)
                rdin["tat_avg_minutes"]   = round(float(tat_vals.mean()), 2)
                rdin["tat_avg_hours"]     = round(float(tat_vals.mean()) / 60, 2)

    return {
        "total": len(df),
        "complaints": {
            "total": len(comp),
            "breakdown": _build_section(comp["_sub"], SUB_BUCKETS, len(comp)),
        },
        "request_enquiry": {
            "total": len(req),
            "request_count": kap_request_count,
            "enquiry_count": kap_enquiry_count,
            "breakdown": _build_section(req["_sub"], SUB_BUCKETS, len(req)),
        },
        "re_top_subcats":   re_top_subcats,
        "comp_top_subcats": comp_top_subcats,
        "rdin":             rdin,
    }

# ---------------------------------------------------------------------------
# SAP ticket classification
# ---------------------------------------------------------------------------

# Presence of any of these patterns in SUB_CATEGORY → Request + Enquiry + Feedback
_SAP_REF_PATTERNS = [
    "duplicate invoice copy",
    "change of mind",
    "repair request",
    "pms service request",
    "demo/installation request",
    "demo and installation request",
    "installation request",
    "re-installation request",
    "reimbursement",
    "service request",
    "renewal of extended",
    " enquiry",
    "feedback",
]


def _sap_toplevel(sub_category: str) -> str:
    s = _norm(sub_category)
    if not s:
        return "Complaints"
    for pat in _SAP_REF_PATTERNS:
        if pat in s:
            return "Request + Enquiry + Feedback"
    return "Complaints"


def _sap_subbucket(category: str, sub_category: str) -> str:
    c = _norm(category)
    s = _norm(sub_category)

    if "refund related" in c:
        return "Refund"
    if "repair related" in c or "after sales service" in c:
        if "pms" in s:
            return "PMS"
        return "Repair"
    if any(k in c for k in ["demo & installation", "installation/ demo",
                              "demo and installation", "demo related"]):
        return "Demo & Installation"
    if "delivery related" in c:
        return "Delivery related"
    if "billing related" in c:
        return "Invoice/Billing related"
    if "pms related" in c:
        return "PMS"
    if any(k in c for k in ["resq care plan", "insurance related"]):
        return "Warranty"

    # In-store experience – context-sensitive
    if "in store" in c:
        if any(k in s for k in ["wrong product", "delivery",
                                  "order not received", "status marked"]):
            return "Delivery related"
        if any(k in s for k in ["exchange product not picked up", "not picked up"]):
            return "Return"
        return "Others"

    # Product related – context-sensitive
    if "product related" in c:
        if any(k in s for k in ["return", "exchange", "change of mind",
                                  "commercial return", "reverse pickup",
                                  "pickup for return"]):
            return "Return"
        if "warranty" in s:
            return "Warranty"
        return "Others"

    # Promotion / resQ app → Others
    if any(k in c for k in ["promotion", "offer related", "resq app"]):
        return "Others"

    # Generic fallbacks
    if "refund" in c or "reimburs" in c:
        return "Refund"
    if "repair" in c:
        return "Repair"
    if any(k in c for k in ["installation", "demo"]):
        return "Demo & Installation"
    if "delivery" in c:
        return "Delivery related"
    if any(k in c for k in ["billing", "invoice"]):
        return "Invoice/Billing related"
    if any(k in c for k in ["return", "exchange"]):
        return "Return"
    if "pms" in c:
        return "PMS"
    if any(k in c for k in ["warranty", "insurance", "resq"]):
        return "Warranty"

    return "Others"


def analyze_sap_tickets(df: pd.DataFrame) -> dict:
    cat_col = _find_col(df, ["category", "CATEGORY"])
    sub_col = _find_col(df, ["sub_category", "sub category", "SUB_CATEGORY"])

    if not cat_col:
        raise ValueError("SAP tickets sheet: 'Category' column not found.")

    df = df.copy()
    cat = df[cat_col].fillna("").astype(str)
    sub = df[sub_col].fillna("").astype(str) if sub_col else pd.Series([""] * len(df))

    df["_top"]    = sub.apply(_sap_toplevel)
    df["_sub"]    = [_sap_subbucket(c, s) for c, s in zip(cat, sub)]
    df["_rawsub"] = sub  # SAP sub_category as raw sub-cat name

    comp = df[df["_top"] == "Complaints"]
    ref  = df[df["_top"] == "Request + Enquiry + Feedback"]

    return {
        "total": len(df),
        "complaints": {
            "total": len(comp),
            "breakdown": _build_section(comp["_sub"], SUB_BUCKETS, len(comp)),
        },
        "request_enquiry_feedback": {
            "total": len(ref),
            "breakdown": _build_section(ref["_sub"], SUB_BUCKETS, len(ref)),
        },
        "re_top_subcats":   _collect_rawsubs(ref,  len(ref),  top_n=10),
        "comp_top_subcats": _collect_rawsubs(comp, len(comp), top_n=10),
    }

# ---------------------------------------------------------------------------
# SO order classification  (all rows = Request + Enquiry)
# ---------------------------------------------------------------------------

# Rules checked top-to-bottom; first match wins.
# PMS must come BEFORE Demo&Installation (both may share 'svc' substrings).
_SO_RULES: list[tuple[list[str], str]] = [
    (["pms", "health check ord", "with filter wp", "without filter wp",
      "clean svc", "health check", "gt pms", "resq gt pms"], "PMS"),
    (["gt installation", "std. inst", "inst@", "instltion", "uninstl",
      "uninstall", "reinstall", "re-installation", "installation",
      "install", "demo", "old product un-inst", "gt demo", "gt uninstallation",
      "resq gt inst", "resq gt uninstltion"], "Demo & Installation"),
    (["repair", "distrbutor rep", "distrib rep"], "Repair"),
    (["renewal of extended", "warranty renew"], "Warranty"),
    (["store defectives", "dc defectives", "defective"], "Return"),
]


def _so_subbucket(rtd: str) -> str:
    r = _norm(rtd)
    for patterns, bucket in _SO_RULES:
        if any(p in r for p in patterns):
            return bucket
    return "Others"


def analyze_so_order(df: pd.DataFrame) -> dict:
    rtd_col = _find_col(df, ["request type description", "request type",
                              "ticket context", "context description"])
    if not rtd_col:
        raise ValueError("SO order sheet: 'Request Type Description' column not found.")

    df = df.copy()
    df["_sub"] = df[rtd_col].fillna("").astype(str).apply(_so_subbucket)

    total = len(df)

    # SO records are all R+E; map each bucket to its canonical sub-category name
    top_subcats: dict = {}
    for bucket in SUB_BUCKETS:
        cnt = int((df["_sub"] == bucket).sum())
        if cnt > 0:
            canonical = _SO_SUBCAT_NAMES.get(bucket, bucket)
            top_subcats[bucket] = [{"name": canonical, "count": cnt,
                                     "pct": _pct(cnt, total)}]

    return {
        "total": total,
        "breakdown": _build_section(df["_sub"], SO_ORDER, total),
        "top_subcats": top_subcats,
    }

# ---------------------------------------------------------------------------
# Other Enquiry (all rows = R+E, sub-bucket = Others)
# ---------------------------------------------------------------------------

def analyze_other_enquiry(df: pd.DataFrame) -> dict:
    return {"total": len(df)}

# ---------------------------------------------------------------------------
# Combiners
# ---------------------------------------------------------------------------

def _add_breakdowns(a: dict, b: dict, order: list[str], total: int) -> dict:
    """Merge two sub-bucket breakdowns, recalculate percentages."""
    result = {}
    for bucket in order:
        cnt = a.get(bucket, {}).get("count", 0) + b.get(bucket, {}).get("count", 0)
        result[bucket] = {"count": cnt, "pct": _pct(cnt, total)}
    return result


def combine_overall_sap(sap: dict, so: dict, oe: dict) -> dict:
    """
    Overall SAP:
      Complaints          = SAP ticket Complaints
      Request+Enquiry+OE  = SAP R+E+F  +  SO  +  Other Enquiry (→ Others bucket)
    """
    oe_count = oe["total"]
    sap_ref_total = sap["request_enquiry_feedback"]["total"]
    so_total = so["total"]
    req_total = sap_ref_total + so_total + oe_count

    # Merge SAP R+E+F and SO breakdowns; OE rows land in Others
    sap_ref_bd = sap["request_enquiry_feedback"]["breakdown"]
    so_bd = {b: so["breakdown"].get(b, {"count": 0}) for b in SUB_BUCKETS}

    combined: dict = {}
    for bucket in SUB_BUCKETS:
        sap_cnt = sap_ref_bd.get(bucket, {}).get("count", 0)
        so_cnt = so_bd.get(bucket, {}).get("count", 0)
        oe_cnt = oe_count if bucket == "Others" else 0
        cnt = sap_cnt + so_cnt + oe_cnt
        combined[bucket] = {"count": cnt, "pct": _pct(cnt, req_total)}

    sap_comp = sap["complaints"]
    total = sap_comp["total"] + req_total

    return {
        "total": total,
        "other_enquiry_count": oe_count,
        "complaints": {
            "total": sap_comp["total"],
            "breakdown": sap_comp["breakdown"],
        },
        "request_enquiry_other": {
            "total": req_total,
            "breakdown": combined,
        },
    }


def combine_final_output(kapture: dict, overall_sap: dict) -> dict:
    """
    Final Output:
      Total       = Kapture  +  Overall SAP
      Complaints  = Kapture Complaints  +  SAP Complaints
      R+E         = Kapture R+E         +  Overall SAP R+E+OE
    """
    kap_comp = kapture["complaints"]
    sap_comp = overall_sap["complaints"]
    comp_total = kap_comp["total"] + sap_comp["total"]

    kap_re = kapture["request_enquiry"]
    sap_re = overall_sap["request_enquiry_other"]
    re_total = kap_re["total"] + sap_re["total"]

    comp_bd = _add_breakdowns(kap_comp["breakdown"], sap_comp["breakdown"],
                               SUB_BUCKETS, comp_total)
    re_bd = _add_breakdowns(kap_re["breakdown"], sap_re["breakdown"],
                             SUB_BUCKETS, re_total)

    return {
        "total": kapture["total"] + overall_sap["total"],
        "complaints": {"total": comp_total, "breakdown": comp_bd},
        "request_enquiry": {"total": re_total, "breakdown": re_bd},
    }

# ---------------------------------------------------------------------------
# KPI analysis
# ---------------------------------------------------------------------------

def analyze_kpi(kapture: dict, sap_tickets: dict, so_output: dict, final_output: dict) -> dict:
    """
    Assemble KPI output.
    R+E sub-cats: Kapture re_top_subcats + SAP re_top_subcats + SO top_subcats (combined)
    Complaints sub-cats: Kapture comp_top_subcats + SAP comp_top_subcats (combined)
    Names are normalized to canonical form before merging counts.
    """

    def _merge_subcats(sources: list[dict]) -> dict:
        """Normalize and merge {bucket: [{name,count,pct}]} from multiple sources."""
        merged: dict[str, dict[str, int]] = {}
        for source in sources:
            for bucket, subs in source.items():
                bucket_counts = merged.setdefault(bucket, {})
                for item in subs:
                    canonical = _normalize_subcat(item["name"])
                    bucket_counts[canonical] = bucket_counts.get(canonical, 0) + item["count"]
        return merged

    def _to_top_subcats(merged: dict, top_n: int = 3) -> dict:
        result = {}
        for bucket, counts in merged.items():
            bucket_total = sum(counts.values())
            top = sorted(counts.items(), key=lambda x: x[1], reverse=True)[:top_n]
            result[bucket] = [
                {"name": name, "count": cnt, "pct": _pct(cnt, bucket_total)}
                for name, cnt in top
            ]
        return result

    re_merged   = _merge_subcats([
        kapture.get("re_top_subcats", {}),
        sap_tickets.get("re_top_subcats", {}),
        so_output.get("top_subcats", {}),
    ])
    comp_merged = _merge_subcats([
        kapture.get("comp_top_subcats", {}),
        sap_tickets.get("comp_top_subcats", {}),
    ])

    re_top_subcats   = _to_top_subcats(re_merged)
    comp_top_subcats = _to_top_subcats(comp_merged)

    def _top_categories(breakdown: dict, top_subcats: dict, top_n: int = 5) -> list:
        entries = sorted(
            [{"category": cat, "count": v["count"], "pct": v["pct"]}
             for cat, v in breakdown.items()
             if cat != "Others" and v["count"] > 0],
            key=lambda x: x["count"],
            reverse=True,
        )[:top_n]
        for e in entries:
            e["top_sub_categories"] = top_subcats.get(e["category"], [])
        return entries

    return {
        "re_total":        final_output["request_enquiry"]["total"],
        "comp_total":      final_output["complaints"]["total"],
        "re_categories":   _top_categories(final_output["request_enquiry"]["breakdown"], re_top_subcats),
        "comp_categories": _top_categories(final_output["complaints"]["breakdown"], comp_top_subcats),
        "rdin":            kapture.get("rdin", {"total": 0, "complaints": 0,
                                                 "tat_total_minutes": None,
                                                 "tat_avg_minutes": None,
                                                 "tat_avg_hours": None}),
    }


# ---------------------------------------------------------------------------
# RD.IN Complaints Ageing analysis
# ---------------------------------------------------------------------------

_AGEING_BUCKETS = [
    {"label": "< 1 Day",    "min": 0,    "max": 1},
    {"label": "1–2 Days",   "min": 1,    "max": 2},
    {"label": "2–3 Days",   "min": 2,    "max": 3},
    {"label": "3–5 Days",   "min": 3,    "max": 5},
    {"label": "5–7 Days",   "min": 5,    "max": 7},
    {"label": "7+ Days",    "min": 7,    "max": None},
]


def _assign_bucket(days: float | None) -> str | None:
    if days is None or pd.isna(days):
        return None
    for b in _AGEING_BUCKETS:
        if b["max"] is None:
            if days >= b["min"]:
                return b["label"]
        elif b["min"] <= days < b["max"]:
            return b["label"]
    return _AGEING_BUCKETS[-1]["label"]


def _breakdown_rows(df: pd.DataFrame, col: str, total_tickets: int, top_n: int = 20) -> list:
    """Build per-group ageing breakdown for a given dimension column."""
    groups = df[col].fillna("(blank)").astype(str).value_counts().head(top_n).index.tolist()
    rows = []
    for grp in groups:
        mask = df[col].fillna("(blank)").astype(str) == grp
        grp_df  = df[mask]
        grp_tot = len(grp_df)
        resolved = grp_df["_tat_days"].notna()
        no_tat   = int((~resolved).sum())
        resolved_days = grp_df.loc[resolved, "_tat_days"]
        avg_resolved  = round(float(resolved_days.mean()), 2) if len(resolved_days) > 0 else None
        avg_all       = round(float(grp_df["_tat_days0"].mean()), 2)

        buckets_cnt = {}
        for b in _AGEING_BUCKETS:
            buckets_cnt[b["label"]] = int((grp_df["_bucket"] == b["label"]).sum())

        rows.append({
            "name":         grp,
            "total":        grp_tot,
            "pct_of_total": _pct(grp_tot, total_tickets),
            "resolved":     int(resolved.sum()),
            "no_tat":       no_tat,
            "avg_tat_resolved": avg_resolved,
            "avg_tat_all":  avg_all,
            "buckets":      buckets_cnt,
        })
    return rows


def analyze_rdin_ageing(df: pd.DataFrame) -> dict:
    """
    Filter Kapture sheet for RD.IN channel + Complaint vertical,
    then compute ticket ageing by TAT bucket, broken down by
    Category, Sub-category, and Queue Name.
    """
    channel_col  = _find_col(df, ["channel", "source channel", "source"])
    vertical_col = _find_col(df, ["vertical"])
    category_col = _find_col(df, ["category"])
    subcat_col   = _find_col(df, ["sub category", "sub_category", "subcategory"])
    queue_col    = _find_col(df, ["queue name", "queue"])
    tat_col      = _find_col(df, ["tat", "tat (min)", "tat(min)", "resolution time"])

    if not vertical_col or not tat_col:
        return {"error": "Required columns (Vertical, TAT) not found in Kapture sheet."}

    def _is_rdin(val: str) -> bool:
        return "rdin" in re.sub(r"[\s.\-_]", "", val.lower())

    # Filter RD.IN Complaints
    rdin_mask = (df[channel_col].fillna("").astype(str).apply(_is_rdin)
                 if channel_col else pd.Series([False] * len(df)))
    comp_mask = df[vertical_col].fillna("").astype(str).str.strip().str.lower() == "complaint"
    df = df[rdin_mask & comp_mask].copy()

    total = len(df)
    if total == 0:
        return {"total": 0, "resolved": 0, "open_no_tat": 0,
                "avg_tat_resolved": None, "avg_tat_all": 0,
                "max_tat_days": None, "bucket_summary": [], "bucket_labels": [],
                "by_category": [], "by_subcategory": [], "by_queue": []}

    # Parse TAT → days; blank/NaN → None (counted as No TAT / open)
    _BLANK_STRS = {"", "nan", "none", "nat", "n/a", "-"}

    def _tat_to_days(val) -> float | None:
        if pd.isna(val):
            return None
        s = str(val).strip().lower()
        if s in _BLANK_STRS:
            return None
        mins = _parse_tat_minutes(val)
        return round(mins / 1440, 4) if mins is not None else None

    df["_tat_days"] = df[tat_col].apply(_tat_to_days)
    df["_tat_days0"] = df["_tat_days"].fillna(0.0)   # no-TAT treated as 0
    df["_bucket"]    = df["_tat_days"].apply(_assign_bucket)

    resolved_mask = df["_tat_days"].notna()
    resolved      = int(resolved_mask.sum())
    open_no_tat   = total - resolved

    resolved_days = df.loc[resolved_mask, "_tat_days"]
    avg_resolved  = round(float(resolved_days.mean()), 2)  if resolved > 0  else None
    avg_all       = round(float(df["_tat_days0"].mean()), 2)
    max_days      = round(float(resolved_days.max()), 2)    if resolved > 0  else None

    # Bucket summary
    bucket_summary = []
    for b in _AGEING_BUCKETS:
        cnt = int((df["_bucket"] == b["label"]).sum())
        bucket_summary.append({
            "label":      b["label"],
            "count":      cnt,
            "pct_of_resolved": _pct(cnt, resolved),
            "pct_of_total":    _pct(cnt, total),
        })
    # Add No-TAT row
    bucket_summary.append({
        "label":      "No TAT (Open)",
        "count":      open_no_tat,
        "pct_of_resolved": 0.0,
        "pct_of_total":    _pct(open_no_tat, total),
    })

    # Dimension breakdowns
    by_category   = _breakdown_rows(df, category_col, total, top_n=20) if category_col else []
    by_subcategory= _breakdown_rows(df, subcat_col,   total, top_n=20) if subcat_col   else []
    by_queue      = _breakdown_rows(df, queue_col,    total, top_n=20) if queue_col    else []

    return {
        "total":            total,
        "resolved":         resolved,
        "open_no_tat":      open_no_tat,
        "avg_tat_resolved": avg_resolved,
        "avg_tat_all":      avg_all,
        "max_tat_days":     max_days,
        "bucket_labels":    [b["label"] for b in _AGEING_BUCKETS],
        "bucket_summary":   bucket_summary,
        "by_category":      by_category,
        "by_subcategory":   by_subcategory,
        "by_queue":         by_queue,
    }


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def _find_sheet(sheet_names: list[str], *candidates: str) -> str | None:
    """Return the first sheet whose name contains any candidate (case-insensitive)."""
    for cand in candidates:
        for s in sheet_names:
            if cand.lower() in s.lower():
                return s
    return None


def analyze_file(file_obj: io.BytesIO) -> dict:
    """
    Parse the uploaded Excel file and return all five dashboard results.

    Expected sheets: Kap Tickets data, SAP tickets, SO order, Other Enquiry
    """
    xl = pd.ExcelFile(file_obj)
    names = xl.sheet_names

    kap_sheet  = _find_sheet(names, "kap")
    sap_sheet  = _find_sheet(names, "sap ticket")
    so_sheet   = _find_sheet(names, "so order", "so ")
    oe_sheet   = _find_sheet(names, "other enquiry", "other enq")

    missing = []
    if not kap_sheet:  missing.append("Kapture tickets (expected sheet containing 'Kap')")
    if not sap_sheet:  missing.append("SAP tickets (expected sheet containing 'SAP ticket')")
    if not so_sheet:   missing.append("SO order (expected sheet containing 'SO order')")
    if not oe_sheet:   missing.append("Other Enquiry (expected sheet containing 'Other Enquiry')")

    if missing:
        raise ValueError(
            "The following required sheets were not found in the uploaded file:\n"
            + "\n".join(f"  • {m}" for m in missing)
            + f"\n\nSheets found: {', '.join(names)}"
        )

    kap_df = xl.parse(kap_sheet)
    sap_df = xl.parse(sap_sheet)
    so_df  = xl.parse(so_sheet)
    oe_df  = xl.parse(oe_sheet)

    kapture         = analyze_kapture(kap_df)
    sap_tickets     = analyze_sap_tickets(sap_df)
    so_output       = analyze_so_order(so_df)
    other_enquiry   = analyze_other_enquiry(oe_df)
    overall_sap     = combine_overall_sap(sap_tickets, so_output, other_enquiry)
    final_output    = combine_final_output(kapture, overall_sap)
    kpi             = analyze_kpi(kapture, sap_tickets, so_output, final_output)
    rdin_ageing     = analyze_rdin_ageing(kap_df)

    # Break-Up summary
    sap_request  = so_output["total"]
    sap_enquiry  = overall_sap["request_enquiry_other"]["total"] - sap_request
    kap_request  = kapture["request_enquiry"].get("request_count", 0)
    kap_enquiry  = kapture["request_enquiry"].get("enquiry_count", 0)
    break_up = {
        "kapture": {
            "total":           kapture["total"],
            "complaints":      kapture["complaints"]["total"],
            "request_enquiry": kapture["request_enquiry"]["total"],
        },
        "sap": {
            "total":           overall_sap["total"],
            "complaints":      overall_sap["complaints"]["total"],
            "request_enquiry": overall_sap["request_enquiry_other"]["total"],
        },
        "kapture_re": {
            "total":   kapture["request_enquiry"]["total"],
            "request": kap_request,
            "enquiry": kap_enquiry,
        },
        "sap_re": {
            "total":   overall_sap["request_enquiry_other"]["total"],
            "request": sap_request,
            "enquiry": sap_enquiry,
        },
        "overall_re": {
            "total":   final_output["request_enquiry"]["total"],
            "request": kap_request + sap_request,
            "enquiry": kap_enquiry + sap_enquiry,
        },
        "overall_complaints": {
            "total":   final_output["complaints"]["total"],
            "kapture": kapture["complaints"]["total"],
            "sap":     overall_sap["complaints"]["total"],
        },
        "overall_re_by_source": {
            "total":   final_output["request_enquiry"]["total"],
            "kapture": kapture["request_enquiry"]["total"],
            "sap":     overall_sap["request_enquiry_other"]["total"],
        },
        "overall": {
            "total":      final_output["total"],
            "complaints": final_output["complaints"]["total"],
            "request":    kap_request + sap_request,
            "enquiry":    kap_enquiry + sap_enquiry,
        },
    }

    return {
        "kapture":       kapture,
        "sap_tickets":   sap_tickets,
        "so_output":     so_output,
        "other_enquiry": other_enquiry,
        "overall_sap":   overall_sap,
        "final_output":  final_output,
        "kpi":           kpi,
        "rdin_ageing":   rdin_ageing,
        "break_up":      break_up,
        "meta": {
            "sheets": {
                "kapture": kap_sheet,
                "sap": sap_sheet,
                "so": so_sheet,
                "other_enquiry": oe_sheet,
            }
        }
    }
