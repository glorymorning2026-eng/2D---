import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from parser_2d import (
    COLUMNS, MIN_BET, MAX_BET,
    normalize_text, parse_bets_locally, process_and_add_bets, generate_styled_excel,
)

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(page_title="Myanmar 2D Dealer", page_icon="📒", layout="wide")
MM_TZ = ZoneInfo("Asia/Yangon")
DATA_FILE = Path("ledger_2d.csv")


def now_mm():
    return datetime.now(MM_TZ)


st.markdown(
    """
<style>
.main-header { border-left: 6px solid #D97706; padding-left: 15px; margin-bottom: 20px; }
.stButton>button { border-radius: 8px; font-weight: bold; }
div[data-testid="stMetricValue"] { color: #D97706; }
</style>
""",
    unsafe_allow_html=True,
)
st.markdown(
    """
<div class="main-header">
<h1>📒 မြန်မာ 2D Master Dealer App</h1>
<p>Unified Precision Parser • Preview & Edit • Warning System • Viber Statement • Excel Export</p>
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
# PERSISTENCE (auto-save CSV + manual backup/restore)
# =========================================================
def sanitize_ledger(df):
    """Clean any ledger DataFrame: valid 2D numbers, int amounts, filled metadata."""
    cleaned = []
    for _, row in df.iterrows():
        num = normalize_text(str(row.get("Number", "") or "")).strip()
        if num.endswith(".0"):
            num = num[:-2]
        num = num.zfill(2)
        try:
            amt = int(float(row.get("Amount")))
        except (ValueError, TypeError):
            continue
        if not (num.isascii() and num.isdigit() and len(num) == 2):
            continue
        if amt < MIN_BET or amt > MAX_BET:
            continue
        date = str(row.get("Date", "") or now_mm().strftime("%Y-%m-%d"))
        day = str(row.get("Day", "") or "")
        if not day:
            try:
                day = datetime.strptime(date, "%Y-%m-%d").strftime("%A").upper()
            except ValueError:
                day = "-"
        cust = str(row.get("Customer", "") or "Unknown")
        batch = str(row.get("BatchID", "") or f"{cust}(1)")
        cat = str(row.get("Category", "") or "main")
        desc = str(row.get("Description", "") or num)
        cleaned.append({
            "Date": date, "Day": day, "Customer": cust, "BatchID": batch,
            "Category": cat, "Description": desc, "Number": num, "Amount": amt,
        })
    return pd.DataFrame(cleaned, columns=COLUMNS)


def load_ledger():
    if DATA_FILE.exists():
        try:
            return sanitize_ledger(pd.read_csv(DATA_FILE, dtype={"Number": str}))
        except Exception:
            pass
    return pd.DataFrame(columns=COLUMNS)


def save_ledger(df):
    try:
        df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
    except Exception:
        pass


# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များ")
max_limit = st.sidebar.number_input("Limit (Max/number)", min_value=0, value=500000, step=50000)
payout_rate = st.sidebar.number_input("ဆုကြေးနှုန်း (ဆ)", min_value=0, value=80, step=5)
commission_pct = st.sidebar.number_input("ကော်မရှင် (%)", min_value=0, max_value=100, value=10, step=1)
winning_number = st.sidebar.text_input("ထွက်မည့် ပေါက်ဂဏန်း", value="")

raw_win = normalize_text(winning_number).strip()
win = ""
if raw_win:
    if raw_win.isdigit() and len(raw_win) <= 2:
        win = raw_win.zfill(2)
    else:
        st.sidebar.warning("ပေါက်ဂဏန်းသည် 00–99 အတွင်း ဖြစ်ရမည်။")

st.sidebar.markdown("---")
st.sidebar.subheader("⏰ Time Lock")
enable_time_lock = st.sidebar.checkbox("Enable Time Lock", value=False)
closing_time = st.sidebar.time_input("Closing Time", value=datetime.strptime("12:00", "%H:%M").time())
st.sidebar.caption(f"မြန်မာစံတော်ချိန်: {now_mm().strftime('%H:%M')}")

st.sidebar.markdown("---")
st.sidebar.subheader("🚫 Blocked Numbers")
blocked_input = st.sidebar.text_input("Blocked (comma-separated)", value="")
blocked_numbers = []
for n in normalize_text(blocked_input).split(","):
    n = n.strip()
    if n.isdigit() and 0 <= int(n) <= 99:
        blocked_numbers.append(n.zfill(2))

# =========================================================
# SESSION STATE
# =========================================================
if "data" not in st.session_state:
    st.session_state.data = load_ledger()
for k, v in [("pending_parse", None), ("confirm_clear", False),
             ("flash_success", None), ("flash_warning", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

if st.session_state.flash_success:
    st.success(st.session_state.flash_success)
    st.session_state.flash_success = None
if st.session_state.flash_warning:
    st.warning(st.session_state.flash_warning)
    st.session_state.flash_warning = None

# =========================================================
# ENTRY SECTION
# =========================================================
st.subheader("📝 စာရင်းသွင်းရန်")

is_closed = False
if enable_time_lock and now_mm().time() >= closing_time:
    is_closed = True
    st.error(f"⚠️ စာရင်းပိတ်ချိန် ({closing_time.strftime('%H:%M')}) ကျော်လွန်ပါပြီ။")

c1, c2 = st.columns(2)
with c1:
    entry_date = st.date_input("📆 ရက်စွဲ", value=now_mm().date())
with c2:
    day_opts = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "WEEKEND"]
    selected_day = st.selectbox("📅 နေ့", day_opts, index=min(entry_date.weekday(), 5))
date_str = entry_date.strftime("%Y-%m-%d")

col_a, col_b = st.columns([1, 2])
with col_a:
    cust = st.text_input("ဖောက်သည် အမည်", placeholder="ဥပမာ - သိမ့်")
with col_b:
    text_in = st.text_area(
        "ဖောက်သည် ပို့လိုက်သော စာသား",
        placeholder="61.2000R1000\n90.2000@.500\n40_2000@500\n4ထိတ်1000@500\n23.r500tt..\nအပူး500\n25@1000",
        height=180,
    )

if st.button("🔍 Parse & Preview", type="primary"):
    if is_closed:
        st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပါသည်။")
    elif not cust.strip() or not text_in.strip():
        st.warning("ဖောက်သည်အမည်နှင့် စာသားကို ဖြည့်ပါ။")
    else:
        res = parse_bets_locally(text_in)
        st.session_state.pending_parse = {
            "customer": cust.strip(), "date": date_str, "day": selected_day,
            "rows": res["rows"], "warnings": res["warnings"], "errors": res["errors"],
        }
        st.rerun()

# =========================================================
# PREVIEW SECTION
# =========================================================
p = st.session_state.pending_parse
if p:
    st.markdown("---")
    st.subheader(f"👁️ Preview — {p['customer']} ({p['date']} / {p['day']})")
    for e in p["errors"]:
        st.error(e)
    for w in p["warnings"]:
        st.warning(w)

    if p["rows"]:
        df_prev = pd.DataFrame(p["rows"])[["Number", "Amount", "Category", "Description"]]
        st.info(f"**စုစုပေါင်း အကွက်ရေ:** {len(df_prev)} ကွက် | "
                f"**စုစုပေါင်း ပမာဏ:** {df_prev['Amount'].sum():,.0f} ကျပ်")

        if win:
            win_matches = df_prev[df_prev["Number"] == win]
            if not win_matches.empty:
                st.warning(f"🏆 **Winning Number ({win}) ပါဝင်နေပါသည်။** "
                           f"ထိုးငွေ: {win_matches['Amount'].sum():,.0f} ကျပ်")
            else:
                st.success(f"✅ **Winning Number ({win}) မပါဝင်ပါ။**")

        edited = st.data_editor(
            df_prev[["Number", "Amount"]],
            hide_index=True, use_container_width=True, num_rows="dynamic",
            column_config={
                "Number": st.column_config.TextColumn("2D", max_chars=2),
                "Amount": st.column_config.NumberColumn(
                    "Amount", min_value=MIN_BET, max_value=MAX_BET, step=100),
            },
            key="parse_preview",
        )

        cs, cc = st.columns(2)
        with cs:
            if st.button("✅ Save to Ledger", type="primary", use_container_width=True):
                rows_fixed = []
                orig_meta = df_prev[["Category", "Description"]].to_dict("records")
                for idx, row in edited.iterrows():
                    num = normalize_text(str(row.get("Number", "") or "")).strip()
                    if not num or not num.isdigit() or len(num) > 2:
                        continue
                    amt_val = row.get("Amount")
                    if pd.isna(amt_val):
                        continue
                    try:
                        amt = int(amt_val)
                    except (ValueError, TypeError):
                        continue
                    if amt <= 0:
                        continue
                    meta = (orig_meta[idx] if idx < len(orig_meta)
                            else {"Category": "main", "Description": num.zfill(2)})
                    rows_fixed.append({
                        "Number": num.zfill(2), "Amount": amt,
                        "Category": meta["Category"], "Description": meta["Description"],
                    })
                existing = st.session_state.data[
                    st.session_state.data["Customer"] == p["customer"]]["BatchID"]
                v, blk, inv, tot = process_and_add_bets(
                    rows_fixed, p["date"], p["day"], p["customer"],
                    existing_batch_ids=existing.unique(), blocked_numbers=blocked_numbers,
                )
                if v:
                    st.session_state.data = pd.concat(
                        [st.session_state.data, pd.DataFrame(v)], ignore_index=True)
                    save_ledger(st.session_state.data)
                    msg = f"🎉 {len(v)} ကွက် / {tot:,} ကျပ် သိမ်းဆည်းပြီး"
                    if blk:
                        msg += f" | ပိတ်ဂဏန်း ဖယ်: {', '.join(sorted(set(blk)))}"
                    if inv:
                        msg += f" | မမှန်ကန်သော {inv} ကွက် ဖယ်ထားသည်"
                    st.session_state.flash_success = msg
                    st.session_state.pending_parse = None
                    st.rerun()
                else:
                    st.session_state.flash_warning = "သိမ်းဆည်းနိုင်သော ကွက် မရှိပါ။"
                    st.rerun()
        with cc:
            if st.button("❌ Cancel", use_container_width=True):
                st.session_state.pending_parse = None
                st.rerun()
    else:
        st.error("Parse မှ ရလဒ် မထွက်ပါ။ စာသားပုံစံကို စစ်ဆေးပါ။")
        if st.button("❌ Cancel"):
            st.session_state.pending_parse = None
            st.rerun()

# =========================================================
# DASHBOARD
# =========================================================
st.markdown("---")

if not st.session_state.data.empty:
    df = st.session_state.data.copy()
    tot = float(df["Amount"].sum())
    comm = tot * commission_pct / 100
    net = tot - comm
    matched = df[df["Number"] == win] if win else df.iloc[0:0]
    win_tot = float(matched["Amount"].sum())
    payout = win_tot * payout_rate
    pl = net - payout

    st.subheader("📊 အဓိက စာရင်းချုပ်")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("ရောင်းရငွေ", f"{tot:,.0f}")
    m2.metric(f"ကော်မရှင် {commission_pct}%", f"{comm:,.0f}")
    m3.metric("Net Sales", f"{net:,.0f}")
    m4.metric("အသားတင် P/L", f"{pl:,.0f}")

    st.markdown("---")
    st.subheader("🏆 ပေါက်ဂဏန်း တိုက်စစ်")
    if win:
        if not matched.empty:
            st.success(f"ပေါက်ဂဏန်း ({win}) ကံထူးရှင်များ:")
            st.dataframe(matched, use_container_width=True)
            cc1, cc2 = st.columns(2)
            cc1.metric("ဆုကြေး", f"{payout:,.0f}")
            cc2.metric("P/L", f"{pl:,.0f}")
        else:
            st.info(f"ပေါက်ဂဏန်း ({win}) ထိုးထားသူ မရှိ။ ဒိုင်အမြတ် = {net:,.0f}")
    else:
        st.info("Sidebar တွင် ပေါက်ဂဏန်း ထည့်ပါ။")

    st.markdown("---")
    with st.expander("✏️ Edit / Delete Bet Records"):
        edited_ledger = st.data_editor(
            st.session_state.data, num_rows="dynamic",
            use_container_width=True, key="edit_ledger",
        )
        ce1, ce2 = st.columns(2)
        with ce1:
            if st.button("💾 Apply Changes to Ledger"):
                cleaned = sanitize_ledger(edited_ledger)
                dropped = len(edited_ledger) - len(cleaned)
                st.session_state.data = cleaned
                save_ledger(cleaned)
                msg = "ပြင်ဆင်မှု သိမ်းပြီး"
                if dropped > 0:
                    msg += f" (မမှန်ကန်သော {dropped} ကြောင်း ဖယ်ထားသည်)"
                st.session_state.flash_success = msg
                st.rerun()
        with ce2:
            st.caption(f"စုစုပေါင်း Record: {len(st.session_state.data)}")

    st.markdown("---")
    st.subheader("📨 Viber Statements")
    col_rep1, col_rep2 = st.columns(2)

    with col_rep1:
        st.markdown("##### 📒 ဒိုင်ချုပ်")
        dsum = df.groupby(["Date", "Day"])["Amount"].sum().reset_index()
        st.dataframe(dsum, use_container_width=True)
        lines = ["*မြန်မာ 2D စာရင်းချုပ်အစီရင်ခံစာ*", "=" * 28]
        for r in dsum.itertuples():
            lines.append(f"{r.Day} ({r.Date}): {r.Amount:,.0f} ကျပ်")
        lines.extend(["-" * 28, f"ရောင်းရငွေ = {tot:,.0f}",
                      f"ကော်မရှင် ({commission_pct}%) = -{comm:,.0f}", f"Net = {net:,.0f}"])
        if win:
            lines.append(f"ပေါက်ဂဏန်း ({win}) = {win_tot:,.0f}")
            lines.append(f"ဆုကြေး ({payout_rate}ဆ) = -{payout:,.0f}")
        lines.append(f"အသားတင် = {pl:,.0f}")
        st.text_area("Dealer Viber Text", "\n".join(lines), height=220)

    with col_rep2:
        st.markdown("##### 👤 ဖောက်သည် ရှင်းတမ်း (အပိုင်းအလိုက် အသေးစိတ်)")
        custs = sorted(df["Customer"].unique())
        sel = st.selectbox("ဖောက်သည်", custs)
        if sel:
            cdf = df[df["Customer"] == sel]
            lines = [f"📊 *{sel}* ၏ စုစုပေါင်း ထိုးငွေအသေးစိတ်ရှင်းတမ်း", "=" * 32]
            batch_totals = []
            for b in cdf["BatchID"].unique():
                bdf = cdf[cdf["BatchID"] == b]
                bt = float(bdf["Amount"].sum())
                batch_totals.append((b, bt))
                lines.append(f"--- *{b}* ---")
                main_df = bdf[bdf["Category"] == "main"]
                if not main_df.empty:
                    lines.append("🔹 *နေ့ရှင်းနှင့် အဓိကထိုးငွေများ:*")
                    for desc_lbl, grp_d in main_df.groupby("Description", sort=False):
                        lines.append(f"  • {desc_lbl} = {grp_d['Amount'].sum():,.0f} ကျပ်")
                rpat_df = bdf[bdf["Category"] == "r_pat"]
                if not rpat_df.empty:
                    lines.append("\n🔹 *ပတ်လည်နှင့် ရိုဂဏန်းများ (R & Pat):*")
                    for desc_lbl, grp_d in rpat_df.groupby("Description", sort=False):
                        lines.append(f"  • {desc_lbl} = {grp_d['Amount'].sum():,.0f} ကျပ်")
                lines.append(f"\n*{b} စုစုပေါင်း* = {bt:,.0f} ကျပ်\n")
            ctot = float(cdf["Amount"].sum())
            ccomm = ctot * commission_pct / 100
            cnet = ctot - ccomm
            cwin = float(cdf[cdf["Number"] == win]["Amount"].sum() * payout_rate) if win else 0
            bal = cnet - cwin
            formula = " + ".join(f"[{b}={a:,.0f} ကျပ်]" for b, a in batch_totals)
            lines.extend(["=" * 32, f"{formula} = *({ctot:,.0f} ကျပ်)*",
                          f"ကော်မရှင် ({commission_pct}%) = -{ccomm:,.0f} ကျပ်",
                          f"စုစုပေါင်း ရှင်းတမ်း (Net Total) = *{cnet:,.0f} ကျပ်*"])
            if win:
                lines.append(f"ဆုကြေး ({win}) = +{cwin:,.0f} ကျပ်")
                lines.append(f"နောက်ဆုံး လက်ကျန်ရှင်းတမ်း = {bal:,.0f} ကျပ်")
            st.text_area(f"Viber ({sel})", "\n".join(lines), height=360)

    st.markdown("---")
    cl, cr = st.columns(2)
    with cl:
        st.subheader("🔢 ဂဏန်း Limit")
        grp = df.groupby("Number")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
        grp["Allowed"] = grp["Amount"].apply(lambda x: min(x, max_limit))
        grp["Excess"] = grp["Amount"] - grp["Allowed"]
        st.dataframe(grp, use_container_width=True)
    with cr:
        st.subheader("👥 Batch အလိုက်")
        cons = df.groupby(["Date", "Day", "Customer", "BatchID", "Number"])["Amount"].sum().reset_index()
        st.dataframe(cons, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Export & Management")
    ce, cx = st.columns([2, 1])
    with ce:
        xls = generate_styled_excel(df, commission_pct, payout_rate, win, max_limit,
                                    generated_at=now_mm())
        st.download_button(
            "📥 Excel Download", data=xls,
            file_name=f"2D_Report_{now_mm().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.download_button(
            "💾 Backup (CSV)", data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"2D_Backup_{now_mm().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )
        up = st.file_uploader("♻️ Backup ပြန်သွင်းရန် (CSV)", type=["csv"])
        if up is not None and st.button("📤 Restore Backup"):
            try:
                restored = sanitize_ledger(pd.read_csv(up, dtype={"Number": str}))
                st.session_state.data = restored
                save_ledger(restored)
                st.session_state.flash_success = f"Backup ပြန်သွင်းပြီး ({len(restored)} ကြောင်း)"
                st.rerun()
            except Exception as ex:
                st.error(f"Restore မအောင်မြင်ပါ: {ex}")
    with cx:
        if not st.session_state.confirm_clear:
            if st.button("🗑️ Clear All"):
                st.session_state.confirm_clear = True
                st.rerun()
        else:
            st.warning("သေချာပါသလား? ပြန်မရနိုင်ပါ။")
            c1x, c2x = st.columns(2)
            with c1x:
                if st.button("✅ Yes"):
                    st.session_state.data = pd.DataFrame(columns=COLUMNS)
                    save_ledger(st.session_state.data)
                    st.session_state.confirm_clear = False
                    st.rerun()
            with c2x:
                if st.button("❌ No"):
                    st.session_state.confirm_clear = False
                    st.rerun()
    st.caption("💡 Streamlit Cloud တွင် ဖိုင်သိမ်းခြင်းသည် app restart ဖြစ်လျှင် ပျက်နိုင်သဖြင့် "
               "နေ့စဉ် Backup (CSV) ဒေါင်းလုဒ်ဆွဲထားပါ။")
else:
    st.info("စာရင်းမရှိသေးပါ။ အထက်တွင် စတင်သွင်းပါ။")
