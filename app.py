"""
Ireland SSPCRS – Price Transformation Tool
==========================================
Transforms HSE Source Data + Reference Data → SSPCRS Template format.

Source columns used   : Code, INN, Name, Drug Name, Strength Measure,
                        Pack Size, Reimbursement Price, Ref Price,
                        Reference Priced
Reference columns used: PRICE_ID, Item Type, VAT
Output template order : exactly matches Template.xlsx (42 columns)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
import warnings
warnings.filterwarnings("ignore")

# ─── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Ireland SSPCRS · Price Tool",
    page_icon="🇮🇪",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif;}
.stApp{background:#0d1117;color:#e6edf3;}
section[data-testid="stSidebar"]{background:#161b22;border-right:1px solid #21262d;}
.hdr{background:linear-gradient(135deg,#161b22 0%,#1c2128 100%);border:1px solid #21262d;
     border-radius:12px;padding:26px 32px;margin-bottom:20px;position:relative;overflow:hidden;}
.hdr::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
             background:linear-gradient(90deg,#169b62,#ff883e,#169b62);}
.hdr h1{font-family:'IBM Plex Mono',monospace;font-size:1.55rem;font-weight:600;
        color:#e6edf3;margin:0;line-height:1.2;}
.hdr .sub{font-size:.75rem;color:#8b949e;font-family:'IBM Plex Mono',monospace;
          margin-top:4px;letter-spacing:.06em;text-transform:uppercase;}
.mc{background:#161b22;border:1px solid #21262d;border-radius:10px;
    padding:18px 22px;position:relative;overflow:hidden;}
.mc::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;}
.mc.g::after{background:#169b62;}.mc.b::after{background:#58a6ff;}.mc.o::after{background:#ff883e;}
.mc .lbl{font-size:.68rem;color:#8b949e;text-transform:uppercase;letter-spacing:.1em;
         font-family:'IBM Plex Mono',monospace;margin-bottom:4px;}
.mc .val{font-size:1.9rem;font-weight:600;font-family:'IBM Plex Mono',monospace;color:#e6edf3;}
.mc .sub{font-size:.72rem;color:#8b949e;margin-top:2px;}
.sh{font-family:'IBM Plex Mono',monospace;font-size:.72rem;text-transform:uppercase;
    letter-spacing:.12em;color:#8b949e;border-bottom:1px solid #21262d;
    padding-bottom:7px;margin:22px 0 14px 0;}
.bok{background:#1a3a2a;color:#56d364;padding:2px 8px;border-radius:4px;
     font-size:.72rem;font-family:'IBM Plex Mono',monospace;}
.bwrn{background:#3a2a1a;color:#e3b341;padding:2px 8px;border-radius:4px;
      font-size:.72rem;font-family:'IBM Plex Mono',monospace;}
.berr{background:#3a1a1a;color:#f85149;padding:2px 8px;border-radius:4px;
      font-size:.72rem;font-family:'IBM Plex Mono',monospace;}
.stButton>button{background:#169b62!important;color:#fff!important;border:none!important;
    border-radius:8px!important;font-family:'IBM Plex Mono',monospace!important;
    font-weight:500!important;letter-spacing:.05em!important;width:100%;}
.stButton>button:hover{background:#1db875!important;
    box-shadow:0 0 18px rgba(22,155,98,.3)!important;}
.stDownloadButton>button{background:#1f6feb!important;color:#fff!important;
    border:none!important;border-radius:8px!important;
    font-family:'IBM Plex Mono',monospace!important;font-weight:500!important;width:100%;}
.uinfo{background:#1c2128;border:1px dashed #30363d;border-radius:7px;padding:10px 14px;
       font-size:.75rem;color:#8b949e;margin-top:6px;font-family:'IBM Plex Mono',monospace;}
.empty{background:#161b22;border:1px dashed #30363d;border-radius:12px;
       padding:60px 32px;text-align:center;margin-top:28px;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# CORE LOGIC
# ══════════════════════════════════════════════════════════════════════════════

def precise_round(value) -> float:
    if pd.isna(value):
        return np.nan
    try:
        return float(Decimal(str(value)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))
    except Exception:
        return np.nan


def effective_date() -> str:
    return date.today().replace(day=1).strftime("%Y-%m-%d")


def load_file(f):
    if f is None:
        return None
    name = f.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(f)
        elif name.endswith((".xlsx", ".xls")):
            return pd.read_excel(f)
        else:
            st.error(f"Unsupported file type: {f.name}")
            return None
    except Exception as e:
        st.error(f"Error reading {f.name}: {e}")
        return None


# 42-column template order – matches Template.xlsx exactly
TEMPLATE_COLS = [
    "Action", "Primary key_Pricing", "PRICE_ID", "Multiplication Factor",
    "Country", "Active Ingredient", "Brand Name", "Company", "Standard Form",
    "Formulation", "Strength", "Strength unit", "Pack", "Pack Unit", "Fill",
    "Fill Unit", "Effective Price Date", "Currency (Local)", "Manufacturer Price",
    "Wholesale Price", "VAT", "Retail Price without VAT", "Retail Price",
    "Price Launch Date", "Launch Price", "Discontinued Date", "Reimbursement",
    "Reimbursement Comments", "Hospital Product", "WHO ATC code",
    "Combination product", "Combination Strength", "Combination Strength Unit",
    "Pack notes", "Company Type", "Pricing Strategy wrt lowest dose",
    "Pricing strategy across the dose", "Local Brand Name", "Local Company",
    "Local Pack Description", "Source Name", "Item Type",
]


def process_data(source_df: pd.DataFrame, ref_df: pd.DataFrame):
    source_df = source_df.copy()
    ref_df    = ref_df.copy()
    source_df.columns = source_df.columns.str.strip()
    ref_df.columns    = ref_df.columns.str.strip()

    def col(df, *candidates):
        low = {c.lower(): c for c in df.columns}
        for c in candidates:
            if c.lower() in low:
                return low[c.lower()]
        return None

    # Source lookups
    c_code    = col(source_df, "Code", "PRICE_ID")
    c_inn     = col(source_df, "INN", "Active Ingredient")
    c_name    = col(source_df, "Name", "Brand Name")
    c_drug    = col(source_df, "Drug Name", "Standard Form")
    c_str     = col(source_df, "Strength Measure", "Strength")
    c_pack    = col(source_df, "Pack Size", "Fill")
    c_rp      = col(source_df, "Reimbursement Price", "Wholesale Price")
    c_refp    = col(source_df, "Ref Price", "Reimbursement")
    c_refprcd = col(source_df, "Reference Priced", "Reimbursement Comments")

    # Reference lookups
    r_id   = col(ref_df, "PRICE_ID", "Code")
    r_type = col(ref_df, "Item Type", "ItemType")
    r_vat  = col(ref_df, "VAT", "Vat")

    missing = []
    if not c_code: missing.append("'Code' in Source")
    if not c_rp:   missing.append("'Reimbursement Price' in Source")
    if not r_id:   missing.append("'PRICE_ID' in Reference")
    if not r_type: missing.append("'Item Type' in Reference")
    if not r_vat:  missing.append("'VAT' in Reference")
    if missing:
        raise ValueError(f"Required columns missing: {', '.join(missing)}")

    w = pd.DataFrame()
    w["PRICE_ID"]               = source_df[c_code].astype(str).str.strip()
    w["Active Ingredient"]      = source_df[c_inn].astype(str).str.strip()  if c_inn     else ""
    w["Brand Name"]             = source_df[c_name].astype(str).str.strip() if c_name    else ""
    w["Standard Form"]          = source_df[c_drug].astype(str).str.strip() if c_drug    else ""
    w["Strength"]               = source_df[c_str].astype(str).str.strip()  if c_str     else ""
    w["Fill"]                   = source_df[c_pack]                          if c_pack    else np.nan
    w["Wholesale Price"]        = pd.to_numeric(source_df[c_rp], errors="coerce")
    w["Reimbursement"]          = source_df[c_refp]                          if c_refp    else ""
    w["Reimbursement Comments"] = source_df[c_refprcd].astype(str)          if c_refprcd else ""

    ref_slim = ref_df[[r_id, r_type, r_vat]].copy()
    ref_slim.columns = ["PRICE_ID", "Item Type", "VAT"]
    ref_slim["PRICE_ID"] = ref_slim["PRICE_ID"].astype(str).str.strip()

    w = w.merge(ref_slim, on="PRICE_ID", how="left")
    w["VAT"] = pd.to_numeric(w["VAT"], errors="coerce")

    eff = effective_date()

    def manuf_price(row):
        wp = row["Wholesale Price"]
        if pd.isna(wp): return np.nan
        return precise_round(wp / 1.12 if str(row.get("Item Type","")).strip().lower() == "fridge"
                             else wp / 1.08)

    def retail_price(row):
        rwv = row["Retail Price without VAT"]
        vat = row["VAT"]
        if pd.isna(rwv) or pd.isna(vat): return np.nan
        return precise_round(rwv * 1.23) if int(vat) == 23 else rwv

    w["Manufacturer Price"]       = w.apply(manuf_price, axis=1)
    w["Retail Price without VAT"] = w["Wholesale Price"].apply(
        lambda x: precise_round(x + 4.84) if not pd.isna(x) else np.nan)
    w["Retail Price"]             = w.apply(retail_price, axis=1)
    w["Wholesale Price"]          = w["Wholesale Price"].apply(precise_round)

    # Fixed / derived fields
    w["Action"]                           = ""
    w["Primary key_Pricing"]              = ""
    w["Multiplication Factor"]            = "1"
    w["Country"]                          = "IRELAND"
    w["Company"]                          = ""
    w["Formulation"]                      = ""
    w["Strength unit"]                    = ""
    w["Pack"]                             = ""
    w["Pack Unit"]                        = ""
    w["Fill Unit"]                        = ""
    w["Effective Price Date"]             = eff
    w["Currency (Local)"]                 = "EUR"
    w["Price Launch Date"]                = ""
    w["Launch Price"]                     = ""
    w["Discontinued Date"]                = ""
    w["Hospital Product"]                 = "No"
    w["WHO ATC code"]                     = ""
    w["Combination product"]              = ""
    w["Combination Strength"]             = ""
    w["Combination Strength Unit"]        = ""
    w["Pack notes"]                       = ""
    w["Company Type"]                     = ""
    w["Pricing Strategy wrt lowest dose"] = ""
    w["Pricing strategy across the dose"] = ""
    w["Local Brand Name"]                 = w["Brand Name"]
    w["Local Company"]                    = ""
    w["Local Pack Description"]           = w["Standard Form"]
    w["Source Name"]                      = "HSE"

    val = {
        "total":        len(w),
        "with_price":   int(w["Wholesale Price"].notna().sum()),
        "fridge":       int((w["Item Type"].str.lower() == "fridge").sum()),
        "vat23":        int((w["VAT"] == 23).sum()),
        "eff_date":     eff,
        "dup_ids":      w[w.duplicated("PRICE_ID", keep=False)]["PRICE_ID"].unique().tolist(),
        "miss_price":   w[w["Wholesale Price"].isna()]["PRICE_ID"].tolist(),
        "miss_vat":     w[w["VAT"].isna()]["PRICE_ID"].tolist(),
        "miss_name":    w[w["Brand Name"].isin(["","nan","NaN"])]["PRICE_ID"].tolist(),
        "miss_type":    w[w["Item Type"].isna()]["PRICE_ID"].tolist(),
    }
    return w[TEMPLATE_COLS], val


def to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.sheets["Sheet1"]
        for cc in ws.columns:
            ml = max((len(str(c.value or "")) for c in cc), default=0) + 3
            ws.column_dimensions[cc[0].column_letter].width = min(ml, 45)
    return buf.getvalue()


def pt():
    return dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#8b949e", family="IBM Plex Mono"),
                xaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
                yaxis=dict(gridcolor="#21262d", linecolor="#30363d"),
                margin=dict(l=40, r=20, t=40, b=40))


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🇮🇪  SSPCRS Tool")
    st.markdown('<div class="uinfo">Ireland · HSE · EUR · v2.0</div>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("#### 📂 Source Data")
    source_file = st.file_uploader("HSE Reimbursement Export",
                                   type=["csv","xlsx","xls"], key="src")
    st.markdown("""<div class="uinfo">
Required: Code · Reimbursement Price<br>
Optional: INN · Name · Drug Name<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Strength Measure · Pack Size<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ref Price · Reference Priced
</div>""", unsafe_allow_html=True)

    st.markdown("#### 📂 Reference Data")
    ref_file = st.file_uploader("Product Reference File",
                                type=["csv","xlsx","xls"], key="ref")
    st.markdown('<div class="uinfo">Required: PRICE_ID · Item Type · VAT</div>',
                unsafe_allow_html=True)

    st.markdown("---")
    run_btn = st.button("⚙️  Process Data", use_container_width=True)

    st.markdown("---")
    st.markdown("""<div style="font-size:.7rem;color:#484f58;font-family:'IBM Plex Mono',monospace;line-height:2.1;">
<b style="color:#8b949e">Manufacturer</b><br>
&nbsp;Fridge: WP ÷ 1.12<br>
&nbsp;Non-Fridge: WP ÷ 1.08<br>
<b style="color:#8b949e">Retail excl VAT</b> = WP + €4.84<br>
<b style="color:#8b949e">Retail</b><br>
&nbsp;VAT 23%: × 1.23<br>
&nbsp;VAT 0%: unchanged<br>
<b style="color:#8b949e">Rounding</b>: Half-Up (Decimal)
</div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hdr">
  <span style="font-size:2.2rem;margin-right:16px">🇮🇪</span>
  <span>
    <h1>Ireland SSPCRS – Price Transformation Tool</h1>
    <div class="sub">State Scheme Price Calculation &amp; Reference System · HSE · EUR</div>
  </span>
</div>""", unsafe_allow_html=True)

for k in ("out_df","val","ready"):
    if k not in st.session_state:
        st.session_state[k] = None
if "ready" not in st.session_state:
    st.session_state.ready = False

if run_btn:
    if source_file is None or ref_file is None:
        st.error("⚠️  Please upload both files before processing.")
    else:
        with st.spinner("Processing records…"):
            src = load_file(source_file)
            ref = load_file(ref_file)
            if src is not None and ref is not None:
                try:
                    out, val = process_data(src, ref)
                    st.session_state.out_df = out
                    st.session_state.val    = val
                    st.session_state.ready  = True
                    st.success(f"✅  {val['total']:,} records processed — "
                               f"{val['with_price']:,} with prices.")
                except ValueError as e:
                    st.error(f"❌  {e}")
                except Exception as e:
                    st.error(f"❌  Unexpected error: {e}")

if st.session_state.ready and st.session_state.out_df is not None:
    df  = st.session_state.out_df
    val = st.session_state.val

    # ── KPI ──────────────────────────────────────────────────────────────
    c1,c2,c3,c4 = st.columns(4)
    cards = [
        (c1,"g","Total Records",       f"{val['total']:,}",      "rows processed"),
        (c2,"b","Records with Prices", f"{val['with_price']:,}",
            f"{val['with_price']/max(val['total'],1)*100:.1f}% coverage"),
        (c3,"o","Fridge · VAT 23%",   f"{val['fridge']}  ·  {val['vat23']}",
            "cold-chain · insulin/biologics"),
        (c4,"g","Effective Date",      val["eff_date"],          "1st of current month"),
    ]
    for co,cl,lb,v,sb in cards:
        with co:
            fs = "1.1rem" if len(v) > 9 else "1.9rem"
            st.markdown(f'<div class="mc {cl}"><div class="lbl">{lb}</div>'
                        f'<div class="val" style="font-size:{fs}">{v}</div>'
                        f'<div class="sub">{sb}</div></div>', unsafe_allow_html=True)

    st.markdown("")

    # ── Download ─────────────────────────────────────────────────────────
    st.download_button(
        "⬇️  Download Excel Output  (42 columns · Template.xlsx format)",
        data=to_excel(df),
        file_name=f"SSPCRS_Ireland_{val['eff_date']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # ── Charts ───────────────────────────────────────────────────────────
    st.markdown('<div class="sh">📊 Price Analytics</div>', unsafe_allow_html=True)

    num = df[["PRICE_ID","Brand Name","Item Type","VAT",
              "Wholesale Price","Manufacturer Price",
              "Retail Price without VAT","Retail Price"]].copy()
    for c in ["Wholesale Price","Manufacturer Price",
              "Retail Price without VAT","Retail Price"]:
        num[c] = pd.to_numeric(num[c], errors="coerce")
    nv = num.dropna(subset=["Wholesale Price","Retail Price"])

    col1,col2 = st.columns(2)
    with col1:
        fig = px.histogram(nv, x="Retail Price", nbins=40,
                           title="Retail Price Distribution",
                           color_discrete_sequence=["#169b62"])
        fig.update_layout(**pt(), title_font_color="#e6edf3")
        fig.update_traces(marker_line_color="#0d1117", marker_line_width=0.4)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = px.scatter(nv, x="Manufacturer Price", y="Retail Price",
                          color="Item Type", title="Manufacturer vs Retail Price",
                          color_discrete_map={"Fridge":"#ff883e","Non-Fridge":"#58a6ff"},
                          hover_data=["PRICE_ID","Brand Name"])
        fig2.update_layout(**pt(), title_font_color="#e6edf3",
                           legend=dict(bgcolor="rgba(0,0,0,0)",
                                       font=dict(color="#8b949e")))
        st.plotly_chart(fig2, use_container_width=True)

    top = (nv[nv["Brand Name"].notna() & (nv["Brand Name"] != "")]
           .nlargest(20,"Retail Price").copy())
    top["label"] = top["Brand Name"].str[:40]
    fig3 = px.bar(top, x="Retail Price", y="label", orientation="h",
                  title="Top 20 Products by Retail Price", color="Item Type",
                  color_discrete_map={"Fridge":"#ff883e","Non-Fridge":"#58a6ff"},
                  hover_data=["PRICE_ID","Wholesale Price","Manufacturer Price"])
    fig3.update_layout(**pt(), title_font_color="#e6edf3",
                       yaxis=dict(gridcolor="#21262d", linecolor="#30363d",
                                  autorange="reversed"),
                       legend=dict(bgcolor="rgba(0,0,0,0)",
                                   font=dict(color="#8b949e")))
    st.plotly_chart(fig3, use_container_width=True)

    col3,col4 = st.columns(2)
    with col3:
        vc = df["VAT"].fillna(-1).astype(int).value_counts().reset_index()
        vc.columns = ["VAT","Count"]
        vc["VAT"] = vc["VAT"].astype(str).replace("-1","Unknown")
        fig4 = px.pie(vc, names="VAT", values="Count",
                      title="VAT Rate Distribution",
                      color_discrete_sequence=["#58a6ff","#ff883e","#8b949e"])
        fig4.update_layout(**pt(), title_font_color="#e6edf3")
        st.plotly_chart(fig4, use_container_width=True)

    with col4:
        ic = df["Item Type"].fillna("Unknown").value_counts().reset_index()
        ic.columns = ["Item Type","Count"]
        fig5 = px.bar(ic, x="Item Type", y="Count", title="Fridge vs Non-Fridge",
                      color="Item Type",
                      color_discrete_map={"Fridge":"#ff883e",
                                          "Non-Fridge":"#58a6ff",
                                          "Unknown":"#8b949e"})
        fig5.update_layout(**pt(), title_font_color="#e6edf3", showlegend=False)
        st.plotly_chart(fig5, use_container_width=True)

    # ── Validation ────────────────────────────────────────────────────────
    st.markdown('<div class="sh">🔍 Validation Report</div>', unsafe_allow_html=True)
    checks = [
        ("Duplicate PRICE_IDs",        val["dup_ids"],    "error"),
        ("Missing Wholesale Price",    val["miss_price"], "error"),
        ("Missing VAT / no ref match", val["miss_vat"],   "warning"),
        ("Missing Brand Name",         val["miss_name"],  "warning"),
        ("Missing Item Type",          val["miss_type"],  "warning"),
    ]
    for label, ids, sev in checks:
        n = len(ids)
        badge = (f'<span class="bok">✓ OK</span>' if n == 0 else
                 f'<span class="berr">✗ {n} records</span>' if sev=="error" else
                 f'<span class="bwrn">⚠ {n} records</span>')
        detail = "No issues." if n==0 else (
            ", ".join(str(i) for i in ids[:30]) + ("…" if n>30 else ""))
        with st.expander(f"{label}  {badge}", expanded=(n>0 and sev=="error")):
            st.markdown(
                f'<div style="font-family:\'IBM Plex Mono\',monospace;'
                f'font-size:.78rem;color:#8b949e">{detail}</div>',
                unsafe_allow_html=True)

    # ── Preview ───────────────────────────────────────────────────────────
    st.markdown('<div class="sh">🔎 Output Preview</div>', unsafe_allow_html=True)
    pcols = ["PRICE_ID","Brand Name","Active Ingredient","Item Type","VAT",
             "Wholesale Price","Manufacturer Price",
             "Retail Price without VAT","Retail Price",
             "Effective Price Date","Source Name"]
    srch = st.text_input("Filter by PRICE_ID, Brand Name, or Active Ingredient",
                         placeholder="Type to search…")
    prev = df[[c for c in pcols if c in df.columns]].copy()
    if srch:
        mask = prev.apply(
            lambda col: col.astype(str).str.contains(srch, case=False, na=False)
        ).any(axis=1)
        prev = prev[mask]
    st.dataframe(prev.reset_index(drop=True), use_container_width=True, height=450)
    st.caption(f"Showing {len(prev):,} of {len(df):,} records  ·  "
               f"Full output has {len(TEMPLATE_COLS)} columns matching Template.xlsx")

else:
    st.markdown("""
    <div class="empty">
      <div style="font-size:2.8rem;margin-bottom:14px">📋</div>
      <div style="font-family:'IBM Plex Mono',monospace;font-size:1.05rem;
                  color:#e6edf3;margin-bottom:10px">No data processed yet</div>
      <div style="font-family:'IBM Plex Mono',monospace;font-size:.78rem;
                  color:#484f58;line-height:2.1">
        1. Upload <b style="color:#8b949e">Source Data</b> (CSV/Excel with Code + Reimbursement Price)<br>
        2. Upload <b style="color:#8b949e">Reference Data</b> (PRICE_ID + Item Type + VAT)<br>
        3. Click <b style="color:#169b62">⚙️ Process Data</b>
      </div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="sh" style="margin-top:28px">📐 Column Mapping Reference</div>',
                unsafe_allow_html=True)
    mapping = pd.DataFrame([
        ("PRICE_ID",               "Source",    "Code column"),
        ("Active Ingredient",      "Source",    "INN column"),
        ("Brand Name",             "Source",    "Name column"),
        ("Standard Form",          "Source",    "Drug Name column"),
        ("Strength",               "Source",    "Strength Measure column"),
        ("Fill",                   "Source",    "Pack Size column"),
        ("Wholesale Price",        "Source",    "Reimbursement Price column"),
        ("Reimbursement",          "Source",    "Ref Price column"),
        ("Reimbursement Comments", "Source",    "Reference Priced column"),
        ("VAT",                    "Reference", "VAT column"),
        ("Item Type",              "Reference", "Item Type column"),
        ("Manufacturer Price",     "Calc",      "Fridge: WP÷1.12  |  Non-Fridge: WP÷1.08"),
        ("Retail Price excl. VAT", "Calc",      "Wholesale Price + €4.84"),
        ("Retail Price (VAT 23%)", "Calc",      "Retail excl VAT × 1.23"),
        ("Retail Price (VAT 0%)",  "Calc",      "Retail excl VAT (unchanged)"),
        ("Multiplication Factor",  "Fixed",     "1"),
        ("Country",                "Fixed",     "IRELAND"),
        ("Currency (Local)",       "Fixed",     "EUR"),
        ("Hospital Product",       "Fixed",     "No"),
        ("Source Name",            "Fixed",     "HSE"),
        ("Local Brand Name",       "Mirror",    "= Brand Name"),
        ("Local Pack Description", "Mirror",    "= Standard Form (Drug Name)"),
        ("Effective Price Date",   "Auto",      "1st day of current month (yyyy-mm-dd)"),
    ], columns=["Output Field","Source","Logic / Column"])
    st.dataframe(mapping, use_container_width=True, hide_index=True, height=530)
