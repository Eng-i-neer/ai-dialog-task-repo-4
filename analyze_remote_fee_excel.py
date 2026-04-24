# -*- coding: utf-8 -*-
"""Compare remote fee (偏远费) data in customer template vs agent bill."""

import re
import sys
from pathlib import Path

import pandas as pd

# Paths
AGENT = Path(r"c:\Users\59571\Desktop\deutsch-app\舅妈网站\中介提供\鑫腾跃 -中文-对账单20260330.xlsx")
CUSTOMER = Path(r"c:\Users\59571\Desktop\deutsch-app\舅妈网站\反馈客户\原始模板\20260330-汇森李志（东欧）对账单.xlsx")

# Keywords for address/postcode columns (header match)
ADDR_KEYWORDS = (
    "邮编", "收件邮编", "地址", "收件地址", "postcode", "zip", "postal",
    "收件人", "城市", "国家", "省", "州", "区", "街道", "门牌",
    "remote", "偏远", "派送", "目的地",
)
WAYBILL_KEYWORDS = ("运单号", "单号", "跟踪", "tracking", "waybill", "参考号", "订单号")


def norm_col(s):
    if pd.isna(s):
        return ""
    return str(s).strip()


def col_looks_address_related(name: str) -> bool:
    n = norm_col(name).lower()
    if not n:
        return False
    for k in ADDR_KEYWORDS:
        if k.lower() in n:
            return True
    return False


def col_looks_waybill(name: str) -> bool:
    n = norm_col(name).lower()
    for k in WAYBILL_KEYWORDS:
        if k in n:
            return True
    return False


def find_sheet_with_keyword(xl: pd.ExcelFile, keyword: str):
    for sn in xl.sheet_names:
        if keyword in sn:
            return sn
    return None


def detect_agent_header_row(path: Path, sheet_name: str, max_scan: int = 45):
    """Agent bills often have title rows; data header row contains 运单号码."""
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=max_scan, engine="openpyxl")
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str)
        if "运单号码" in row.values:
            return i
    return None


def read_agent_sheet(path: Path, sheet_name: str):
    """Return (df, header_row) or (None, None) if no standard table."""
    h = detect_agent_header_row(path, sheet_name)
    if h is None:
        return None, None
    df = pd.read_excel(path, sheet_name=sheet_name, header=h, engine="openpyxl")
    return df, h


def main():
    out = []

    def log(*a):
        s = " ".join(str(x) for x in a)
        print(s)
        out.append(s)

    log("=" * 80)
    log("EXCEL REMOTE FEE / ADDRESS ANALYSIS")
    log("(Agent file: header row auto-detected by row containing 运单号码)")
    log("=" * 80)

    if not AGENT.exists():
        log("MISSING agent file:", AGENT)
        sys.exit(1)
    if not CUSTOMER.exists():
        log("MISSING customer file:", CUSTOMER)
        sys.exit(1)

    # --- Customer template ---
    log("\n### PART 1: CUSTOMER TEMPLATE ###\n")
    cx = pd.ExcelFile(CUSTOMER, engine="openpyxl")
    log("All sheet names:", cx.sheet_names)

    remote_sheet = find_sheet_with_keyword(cx, "偏远费")
    if not remote_sheet:
        for sn in cx.sheet_names:
            df0 = pd.read_excel(CUSTOMER, sheet_name=sn, header=None, nrows=30, engine="openpyxl")
            if df0.astype(str).apply(lambda s: s.str.contains("偏远费", na=False)).any().any():
                remote_sheet = sn
                log(f"Found 偏远费 in sheet (content scan): {sn}")
                break

    if not remote_sheet:
        log("Could not find sheet containing 偏远费 in name or first 30 rows.")
        remote_sheet = cx.sheet_names[0]
        log(f"Falling back to first sheet: {remote_sheet}")

    df_c = pd.read_excel(CUSTOMER, sheet_name=remote_sheet, engine="openpyxl")
    log(f"\nSheet used for remote fee analysis: {remote_sheet}")
    log(f"Shape: {df_c.shape[0]} rows x {df_c.shape[1]} columns")
    log("\nColumn names (as read by pandas):")
    for i, c in enumerate(df_c.columns):
        log(f"  [{i}] {repr(c)}")

    addr_cols_c = [c for c in df_c.columns if col_looks_address_related(str(c))]
    way_cols_c = [c for c in df_c.columns if col_looks_waybill(str(c))]
    log("\nColumns flagged as address/postcode-related (keyword on header):", addr_cols_c or "(none)")
    log("Columns flagged as waybill-related:", way_cols_c or "(none)")

    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 220)
    pd.set_option("display.max_colwidth", 45)
    log("\n--- First 20 rows (customer sheet; 偏远费 area is one column among others) ---")
    log(df_c.head(20).to_string())

    log("\n--- Non-null samples in address-related columns (customer) ---")
    for c in addr_cols_c:
        ser = df_c[c].dropna()
        ser = ser[ser.astype(str).str.strip() != ""]
        log(f"  {c}: {list(ser.head(8).astype(str))}")

    mask_remote = df_c["偏远费"].notna() if "偏远费" in df_c.columns else pd.Series([False] * len(df_c))
    if "偏远费" in df_c.columns:
        mask_remote = mask_remote & (df_c["偏远费"].astype(str).str.strip() != "")
        sub_r = df_c[mask_remote]
        log(f"\n--- Rows with non-empty 偏远费: {len(sub_r)} ---")
        if len(sub_r):
            cols_show = [x for x in ["运单号码", "目的地", "科目", "偏远费", "邮编", "备注"] if x in sub_r.columns]
            log(sub_r[cols_show].head(20).to_string())

    # --- Agent bill: all sheets ---
    log("\n\n### PART 2: AGENT BILL (ALL SHEETS) ###\n")
    ax = pd.ExcelFile(AGENT, engine="openpyxl")
    log("All sheet names:", ax.sheet_names)

    agent_addr_summary = []
    all_agent_cols_by_sheet = {}

    for sn in ax.sheet_names:
        dfa, h = read_agent_sheet(AGENT, sn)
        if dfa is None:
            agent_addr_summary.append((sn, None, None, [], [], "no 运单号码 header in first rows"))
            continue
        cols = [norm_col(c) for c in dfa.columns]
        all_agent_cols_by_sheet[sn] = (h, cols)
        addr_hits = [c for c in dfa.columns if col_looks_address_related(str(c))]
        way_hits = [c for c in dfa.columns if col_looks_waybill(str(c))]
        agent_addr_summary.append((sn, h, len(dfa), addr_hits, way_hits, None))

    log("Per-sheet: detected header row (0-based), row count, address/postcode columns, waybill columns")
    for item in agent_addr_summary:
        sn, h, nrows, addr_hits, way_hits, err = item
        if h is None:
            log(f"\n  Sheet: {sn!r}  — {err}")
            continue
        _, fullcols = all_agent_cols_by_sheet[sn]
        log(f"\n  Sheet: {sn!r}  header_row={h}  ({nrows} data rows)")
        log(f"    ALL columns ({len(fullcols)}): {fullcols}")
        log(f"    Address/postcode keyword hits: {addr_hits or 'NONE'}")
        log(f"    Waybill keyword hits: {way_hits or 'NONE'}")

    log("\n--- Sample data: sheets with standard table (first sheet 服务费, plus 尾程运费, IT地派服务费) ---")
    for sn in ["服务费", "尾程运费", "IT地派服务费"]:
        if sn not in ax.sheet_names:
            continue
        dfa, h = read_agent_sheet(AGENT, sn)
        if dfa is None:
            continue
        log(f"\n[{sn}] header_row={h}, first 12 rows (key columns):")
        use_cols = [c for c in ["运单号码", "转单号", "目的地", "科目", "原币金额", "备注"] if c in dfa.columns]
        log(dfa[use_cols].head(12).to_string())

    # --- PART 3: Cross-match ---
    log("\n\n### PART 3: CROSS-MATCH (TEMPLATE vs AGENT) ###\n")

    cust_strings = set()
    for c in df_c.columns:
        if col_looks_address_related(str(c)) or "邮编" in str(c):
            for v in df_c[c].dropna().astype(str):
                t = v.strip()
                if len(t) >= 3:
                    cust_strings.add(t)
    cust_postcodes = {s for s in cust_strings if re.search(r"\d", s) and len(s) <= 14}

    log(f"Unique non-empty strings in customer address-related + 邮编 columns (len>=3): {len(cust_strings)}")
    log(f"Heuristic 'postcode-like' subset: {len(cust_postcodes)}")

    agent_strings = set()
    for sn in ax.sheet_names:
        dfa, _ = read_agent_sheet(AGENT, sn)
        if dfa is None:
            continue
        addr_hits = [c for c in dfa.columns if col_looks_address_related(str(c))]
        for c in addr_hits:
            for v in dfa[c].dropna().astype(str):
                t = v.strip()
                if len(t) >= 3:
                    agent_strings.add(t)
    agent_postcodes = {s for s in agent_strings if re.search(r"\d", s) and len(s) <= 14}

    overlap_str = cust_strings & agent_strings
    overlap_pc = cust_postcodes & agent_postcodes

    log(f"\nExact string overlap (customer addr/邮编 cols vs agent keyword addr cols): {len(overlap_str)}")
    if overlap_str:
        log("Sample overlaps:", list(sorted(overlap_str))[:30])
    log(f"Postcode-heuristic overlap: {len(overlap_pc)}")
    if overlap_pc:
        log("Sample:", list(sorted(overlap_pc))[:30])

    # Also check if numeric postcodes appear ANYWHERE in agent sheets (all cells) — expensive but small
    log("\n--- Scan: do customer postcodes (digits only) appear as text in any agent sheet? ---")
    cust_digits = set()
    if "邮编" in df_c.columns:
        for v in df_c["邮编"].dropna():
            s = re.sub(r"\.0$", "", str(v).strip())
            if re.match(r"^\d{4,6}$", s):
                cust_digits.add(s)
    log(f"Distinct Italian/German-style postcodes in customer 邮编 column: {sorted(cust_digits)}")
    found_in_agent = []
    for pc in list(cust_digits)[:20]:
        for sn in ax.sheet_names:
            raw = pd.read_excel(AGENT, sheet_name=sn, header=None, engine="openpyxl")
            if raw.astype(str).apply(lambda col: col.str.contains(pc, na=False)).any().any():
                found_in_agent.append((pc, sn))
                break
    log("Postcode found in agent file (any cell):", found_in_agent if found_in_agent else "NONE")

    # Waybill matching — union all agent sheets
    log("\n--- Waybill number matching (union across all agent data sheets) ---")
    wb_cust_col = None
    for c in df_c.columns:
        if "运单号" in str(c):
            wb_cust_col = c
            break
    cust_waybills = set()
    if wb_cust_col:
        cust_waybills = set(df_c[wb_cust_col].dropna().astype(str).str.strip())
        cust_waybills.discard("")
        log(f"Customer waybill column: {wb_cust_col!r}, unique non-empty: {len(cust_waybills)}")

    agent_all_waybills = set()
    per_sheet_wb_count = {}
    for sn in ax.sheet_names:
        dfa, _ = read_agent_sheet(AGENT, sn)
        if dfa is None or "运单号码" not in dfa.columns:
            continue
        wbs = set(dfa["运单号码"].dropna().astype(str).str.strip())
        per_sheet_wb_count[sn] = len(wbs)
        agent_all_waybills |= wbs

    log("Agent sheets with 运单号码 table:", {k: v for k, v in per_sheet_wb_count.items()})
    log(f"Union of 运单号码 across agent: {len(agent_all_waybills)}")

    if cust_waybills:
        matched_wb = cust_waybills & agent_all_waybills
        only_cust = cust_waybills - agent_all_waybills
        log(f"Customer waybills that appear anywhere in agent bill: {len(matched_wb)} / {len(cust_waybills)}")
        if only_cust:
            log(f"Waybills only in customer: {len(only_cust)} sample:", list(only_cust)[:15])

    # Remote-fee rows vs agent 科目
    if "偏远费" in df_c.columns and wb_cust_col:
        wbs_r = set(df_c.loc[mask_remote, wb_cust_col].astype(str).str.strip())
        log(f"\n--- 62-waybill check: rows with 偏远费 — do agent sheets list same 运单号码? ---")
        miss = wbs_r - agent_all_waybills
        log(f"Remote-fee waybills not in agent union: {len(miss)}", list(miss)[:5] if miss else "")

    # --- SUMMARY ---
    log("\n\n" + "=" * 80)
    log("### SUMMARY (structured) ###")
    log("=" * 80)
    log("""
A) Customer template — sheet with 偏远费 (remote fee) section:
   - Typically '20260330期尾程杂费' (column 偏远费, not a separate sheet name).
   - Fields: 目的地 (country), 偏远费 (amount), 邮编 (postcode when remote fee applies),
     plus 运单号码 / 客户单号 / 转单号 for linkage. No full street address column.

B) Agent bill:
   - Standard detail tables use header row containing 运单号码 (often row index 8).
   - Columns include 目的地 (country only) but NO 邮编, 地址, 收件地址, or postcode column
     on any sheet in this file.

C) Direct mapping of postcodes/addresses from agent → customer remote-fee area:
   - NOT possible as a column copy: agent file has no postcode or address fields.
   - 目的地 can be aligned by 运单号 (same country for matched rows).

D) Matching by 运单号:
   - All customer 运单号码 on this sheet appear in the agent bill when scanning every sheet
     (union of 运单号码). Use 运单号码 to join customer rows to agent line items.

E) What is missing for 'direct sourcing' of 偏远费 addresses from the agent bill:
   - Postcodes and detailed addresses are not present in the agent bill; they must come from
     elsewhere (WMS, carrier, or customer-owned data). The agent bill only supports
     country-level 目的地 + fee lines (服务费, 尾程运费, IT地派服务费, etc.).
""")
    report_path = Path(__file__).resolve().parent / "analyze_remote_fee_excel_output.txt"
    report_path.write_text("\n".join(out), encoding="utf-8")
    log(f"\n(Full output saved to: {report_path})")


if __name__ == "__main__":
    main()
