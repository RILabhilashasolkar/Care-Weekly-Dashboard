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


def _pct(count: int, total: int) -> float:
    return round(count / total * 100, 2) if total else 0.0


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
    vertical_col = _find_col(df, ["vertical"])
    category_col = _find_col(df, ["category"])
    subcategory_col = _find_col(df, ["sub category", "sub_category", "subcategory"])

    if not vertical_col:
        raise ValueError("Kapture sheet: 'Vertical' column not found.")
    if not category_col:
        raise ValueError("Kapture sheet: 'Category' column not found.")

    # Exclude rows where the top-level vertical is blank (insufficient context)
    df = df[df[vertical_col].notna() & (df[vertical_col].astype(str).str.strip() != "")].copy()

    cat = df[category_col].fillna("").astype(str)
    sub = df[subcategory_col].fillna("").astype(str) if subcategory_col else pd.Series([""] * len(df))

    df["_top"] = df[vertical_col].apply(_kap_toplevel)
    df["_sub"] = [_kap_subbucket(c, s) for c, s in zip(cat, sub)]

    comp = df[df["_top"] == "Complaints"]
    req = df[df["_top"] == "Request + Enquiry"]

    return {
        "total": len(df),
        "complaints": {
            "total": len(comp),
            "breakdown": _build_section(comp["_sub"], SUB_BUCKETS, len(comp)),
        },
        "request_enquiry": {
            "total": len(req),
            "breakdown": _build_section(req["_sub"], SUB_BUCKETS, len(req)),
        },
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

    df["_top"] = sub.apply(_sap_toplevel)
    df["_sub"] = [_sap_subbucket(c, s) for c, s in zip(cat, sub)]

    comp = df[df["_top"] == "Complaints"]
    ref = df[df["_top"] == "Request + Enquiry + Feedback"]

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
    return {
        "total": total,
        "breakdown": _build_section(df["_sub"], SO_ORDER, total),
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

    kapture       = analyze_kapture(kap_df)
    sap_tickets   = analyze_sap_tickets(sap_df)
    so_output     = analyze_so_order(so_df)
    other_enquiry = analyze_other_enquiry(oe_df)
    overall_sap   = combine_overall_sap(sap_tickets, so_output, other_enquiry)
    final_output  = combine_final_output(kapture, overall_sap)

    return {
        "kapture":       kapture,
        "sap_tickets":   sap_tickets,
        "so_output":     so_output,
        "other_enquiry": other_enquiry,
        "overall_sap":   overall_sap,
        "final_output":  final_output,
        "meta": {
            "sheets": {
                "kapture": kap_sheet,
                "sap": sap_sheet,
                "so": so_sheet,
                "other_enquiry": oe_sheet,
            }
        }
    }
