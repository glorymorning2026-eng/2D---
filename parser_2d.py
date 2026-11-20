"""parser_2d.py — Myanmar 2D bet parser + validator + Excel report (fixed & tested)."""
import io
import re
from datetime import datetime
from itertools import permutations, product

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd

COLUMNS = ["Date", "Day", "Customer", "BatchID", "Category", "Description", "Number", "Amount"]
MIN_BET = 100
MAX_BET = 500_000

BUR2ENG = str.maketrans("၀၁၂၃၄၅၆၇၈၉", "0123456789")


def normalize_text(text):
    """Myanmar digits -> ASCII, remove thousand separators (1,000 -> 1000)."""
    text = str(text).translate(BUR2ENG)
    text = re.sub(r"(?<=\d),(?=\d{3}(?!\d))", "", text)
    return text


def add_number(rows, number, amount, is_default=False, category="main", desc=""):
    n = str(number).zfill(2)
    if len(n) == 2 and n.isascii() and n.isdigit() and 0 <= int(n) <= 99:
        rows.append({
            "Number": n, "Amount": int(amount), "_default": is_default,
            "Category": category, "Description": desc,
        })
        return True
    return False


def get_amount(text):
    """Extracts the last number in a string as the amount."""
    matches = re.findall(r"(\d+(?:,\d+)*)(?:tt)?", text, re.IGNORECASE)
    if matches:
        return int(matches[-1].replace(",", ""))
    return 1000


def amount_after(line, end):
    """Amount AFTER the pattern keyword; else a >=3-digit number before it;
    else default 1000. Never returns the pattern digit itself."""
    tail = re.findall(r"\d[\d,]*", line[end:])
    if tail:
        return int(tail[-1].replace(",", "")), False
    head = [n for n in re.findall(r"\d[\d,]*", line[:end])
            if len(n.replace(",", "")) >= 3]
    if head:
        return int(head[-1].replace(",", "")), False
    return 1000, True


def _bk_for_digit(digit):
    d = int(digit)
    return [f"{i:02d}" for i in range(100) if (i // 10 + i % 10) % 10 == d]


def _pat_for_digit(digit):
    return [f"{i:02d}" for i in range(100) if digit in f"{i:02d}"]


# ထိပ်/ပိတ် + R/@ combined patterns (front_first=True -> group2=ထိပ်, group3=ပိတ်)
_COMBO_PATTERNS = [
    (r"(\d)\s*[\.\-\_\s]*\s*(?:ထိပ်စီး|ထိပ်|ထိတ်)\s*[\.\-\_\s]*\s*(\d+)\s*[\.\-\_\s]*[\@Rr][\.\-\_\s]*(\d+)", True),
    (r"(\d)\s*[\.\-\_\s]*\s*(?:နောက်ပိတ်|ပိတ်)\s*[\.\-\_\s]*\s*(\d+)\s*[\.\-\_\s]*[\@Rr][\.\-\_\s]*(\d+)", False),
    (r"(\d)\s*(?:ထိပ်စီး|ထိပ်|ထိတ်)\s*(\d+)\s*[\@Rr]\.?\s*(\d+)", True),
    (r"(\d)\s*(?:နောက်ပိတ်|ပိတ်)\s*(\d+)\s*[\@Rr]\.?\s*(\d+)", False),
]


def parse_line(raw_line):
    rows, warns = [], []
    line = re.sub(r"(?:tt|ks|mmk|ကျပ်|\.)+$", "", raw_line.strip(), flags=re.IGNORECASE).strip()
    if not line:
        return rows, warns

    # 0-2. ထိပ်/ပိတ် + R/@ (dot-separated or compact)
    for pat, front_first in _COMBO_PATTERNS:
        m = re.fullmatch(pat, line, re.I)
        if m:
            dig = m.group(1)
            a, b = int(m.group(2)), int(m.group(3))
            fa, ba = (a, b) if front_first else (b, a)
            for i in range(10):
                add_number(rows, f"{dig}{i}", fa, category="r_pat",
                           desc=f"{dig} ထိပ် (၁၀ ကွက် × {fa:,})")
            for i in range(10):
                add_number(rows, f"{i}{dig}", ba, category="r_pat",
                           desc=f"{dig} ပိတ် (၁၀ ကွက် × {ba:,})")
            return rows, warns

    # 3. ညီအကို
    if re.search(r"ညီအကို|ညီကို", line):
        amt = get_amount(line)
        added = set()
        for n in ["01", "12", "23", "34", "45", "56", "67", "78", "89", "90"]:
            for item in (n, n[::-1]):
                if item not in added:
                    added.add(item)
                    add_number(rows, item, amt, category="r_pat",
                               desc=f"ညီအကို (၂၀ ကွက် × {amt:,})")
        return rows, warns

    # 4. အပူး
    if re.search(r"အပူ|ပူး", line) and not re.search(r"ခွေ", line):
        amt = get_amount(line)
        for i in range(10):
            add_number(rows, f"{i}{i}", amt, category="r_pat",
                       desc=f"အပူး (၁၀ ကွက် × {amt:,})")
        return rows, warns

    # 5. နက္ခတ်
    if "နက္ခတ်" in line:
        amt = get_amount(line)
        for n in ["07", "70", "18", "81", "24", "42", "35", "53", "69", "96"]:
            add_number(rows, n, amt, category="r_pat", desc=f"နက္ခတ် (၁၀ ကွက် × {amt:,})")
        return rows, warns

    # 6. ပါဝါ
    if re.search(r"ပါဝါ|power", line, re.I):
        amt = get_amount(line)
        for n in ["05", "50", "16", "61", "27", "72", "38", "83", "49", "94"]:
            add_number(rows, n, amt, category="r_pat", desc=f"ပါဝါ (၁၀ ကွက် × {amt:,})")
        return rows, warns

    # 7. ခွေ
    kh = re.search(r"(\d{2,5})\s*(ခွေ|အခွေ)", line)
    if kh:
        digits = kh.group(1)
        amt, dflt = amount_after(line, kh.end())
        if dflt:
            warns.append(f"'{line}' → amount မပါ၊ 1,000 ဟု ယူဆသည်")
        inc = bool(re.search(r"အပူးပါ|ပူးပါ|အပူးအပါ", line))
        combos = list(product(set(digits), repeat=2)) if inc else list(permutations(digits, 2))
        added = set()
        for p in combos:
            n = "".join(p)
            if n not in added:
                added.add(n)
                add_number(rows, n, amt, category="r_pat", desc=f"{digits} ခွေ ({amt:,})")
        return rows, warns

    # 8. BK / ဘရိတ်
    bks = list(re.finditer(r"(\d)\s*(bk|ဘရိတ်|B)(?![a-zA-Z])", line, re.I))
    if bks:
        amt, dflt = amount_after(line, bks[-1].end())
        if dflt:
            warns.append(f"'{line}' → amount မပါ၊ 1,000 ဟု ယူဆသည်")
        bk_digits = [m.group(1) for m in bks]
        label = " / ".join(f"{d} bk" for d in bk_digits)
        desc_str = f"{label} (ဘရိတ်ဂဏန်း {len(bk_digits)} ခု × {amt:,})"
        for d in bk_digits:
            for n in _bk_for_digit(d):
                add_number(rows, n, amt, category="r_pat", desc=desc_str)
        return rows, warns

    # 9. ပတ်
    pa = re.search(r"((?:\d[\s\.\-\/]*)+)\s*(?:ပါတ်|ပတ်|ပက်|ပါ|ပ(?!ိတ်|ူး))", line)
    if pa:
        digits_list = []
        for d in re.findall(r"\d", pa.group(1)):
            if d not in digits_list:
                digits_list.append(d)
        amt, dflt = amount_after(line, pa.end())
        if dflt:
            warns.append(f"'{line}' → amount မပါ၊ 1,000 ဟု ယူဆသည်")
        inc_double = bool(re.search(r"အပူးပါ|အပူးအပါ|ပူးပါ|ပူးအပါ", line))
        for digit in digits_list:
            desc_str = (f"{digit} ပတ် အပူးပါ ({amt:,})" if inc_double
                        else f"{digit} ပတ် (၁၉ ကွက် × {amt:,})")
            for n in _pat_for_digit(digit):
                add_number(rows, n, amt, category="r_pat", desc=desc_str)
            if inc_double:
                double_num = f"{digit}{digit}"
                add_number(rows, double_num, amt, category="r_pat",
                           desc=f"{digit} ပတ် အပူးထပ်ဆောင်း ({double_num} = {amt:,})")
        return rows, warns

    # 10. ထိပ် (plain)
    th = re.search(r"(\d)\s*[\.\-\_\s]*\s*(?:ထိပ်စီး|ထိပ်|ထိတ်)", line)
    if th:
        digit = th.group(1)
        amt, dflt = amount_after(line, th.end())
        if dflt:
            warns.append(f"'{line}' → amount မပါ၊ 1,000 ဟု ယူဆသည်")
        for i in range(10):
            add_number(rows, f"{digit}{i}", amt, category="r_pat",
                       desc=f"{digit} ထိပ် (၁၀ ကွက် × {amt:,})")
        return rows, warns

    # 11. ပိတ် (plain)
    pt = re.search(r"(\d)\s*[\.\-\_\s]*\s*(?:နောက်ပိတ်|ပိတ်)", line)
    if pt:
        digit = pt.group(1)
        amt, dflt = amount_after(line, pt.end())
        if dflt:
            warns.append(f"'{line}' → amount မပါ၊ 1,000 ဟု ယူဆသည်")
        for i in range(10):
            add_number(rows, f"{i}{digit}", amt, category="r_pat",
                       desc=f"{digit} ပိတ် (၁၀ ကွက် × {amt:,})")
        return rows, warns

    # 12. R / @
    if re.search(r"[\@Rr]", line):
        typo = re.fullmatch(r"(\d{4})[\@Rr]\.?(\d+)", line)
        if typo:
            num = typo.group(1)[:2]
            base_amt = int(typo.group(1))
            rev_amt = int(typo.group(2))
            add_number(rows, num, base_amt, category="r_pat",
                       desc=f"{num} R ({base_amt:,}/{rev_amt:,})")
            if num != num[::-1]:
                add_number(rows, num[::-1], rev_amt, category="r_pat",
                           desc=f"{num} R ({base_amt:,}/{rev_amt:,})")
            warns.append(f"Ambiguous '{line}' → assumed {num}={base_amt}, {num[::-1]}={rev_amt}")
            return rows, warns

        parts = re.split(r"[\.\-\_\s]*[\@Rr][\.\-\_\s]*", line, maxsplit=1)
        if len(parts) == 2:
            left, right = parts
            rev_amt = get_amount(right)
            tokens = [t for t in re.split(r"[^\d]+", left) if t]
            base_amt = None
            nums = []
            for t in tokens:
                if len(t) <= 2:
                    nums.append(t.zfill(2))
                elif int(t) >= 100:
                    base_amt = int(t)
            if base_amt is None:
                base_amt = rev_amt
            num_label = " / ".join(nums)
            total_k = sum(2 if n != n[::-1] else 1 for n in nums)
            desc_str = (f"{num_label} R ({total_k} ကွက် × {rev_amt:,})"
                        if base_amt == rev_amt
                        else f"{num_label} R ({base_amt:,} / {rev_amt:,})")
            for n in nums:
                add_number(rows, n, base_amt, category="r_pat", desc=desc_str)
                if n != n[::-1]:
                    add_number(rows, n[::-1], rev_amt, category="r_pat", desc=desc_str)
            if not nums and tokens:
                warns.append(f"Could not extract numbers from '{left}'")
            return rows, warns

    # 13. Universal: 2-digit numbers + amount >= 100
    tokens = [t for t in re.split(r"[^\d]+", line) if t]
    if tokens:
        nums, amt = [], None
        for t in tokens:
            if len(t) <= 2:
                nums.append(t.zfill(2))
            elif int(t) >= 100:
                amt = int(t)
        if nums and amt is not None:
            num_label = ", ".join(nums)
            desc_str = f"{num_label} (တစ်ကွက် {amt:,})" if len(nums) > 1 else num_label
            for n in nums:
                add_number(rows, n, amt, category="main", desc=desc_str)
            return rows, warns
        elif nums and amt is None and re.fullmatch(r"[\d\s.,/-]+", line):
            num_label = ", ".join(nums)
            for n in nums:
                add_number(rows, n, 1000, is_default=True, category="main",
                           desc=f"{num_label} (ထပ်ဆောင်း)")
            return rows, warns

    warns.append(f"Unrecognized: '{line}'")
    return rows, warns


def split_mixed_line(raw_line):
    line = raw_line.strip()
    if "နေ့ရှင်း" in line:
        return [line]
    # "99*250" shorthand -> "99=250"
    line = re.sub(r"\b(\d{1,2})\*(\d{3,})\b", r"\1=\2", line)
    # Pure numeric '=' chains: 99=1000/33/11=500
    if "=" in line and re.fullmatch(r"[\d\s.,/\-=]+", line):
        groups = re.split(r"[/*\s]+", re.sub(r"\s*=\s*", "=", line))
        segs, buf = [], []
        for g in groups:
            if not g:
                continue
            if "=" in g:
                head, amt = g.rsplit("=", 1)
                for t in buf + [head]:
                    if t:
                        segs.append(f"{t}={amt}")
                buf = []
            else:
                buf.append(g)
        segs.extend(buf)
        if len(segs) > 1:
            return segs
    if re.search(r"(?:bk|ဘရိတ်|ပတ်|ပါတ်|ပက်|ခွေ|ထိပ်|ထိတ်|ပိတ်|[\@Rr])", line, re.I) \
            and re.search(r"[/*=]", line):
        segments = []
        for part in re.split(r"\*+", line):
            part = part.strip(" /")
            if not part:
                continue
            sub_parts = re.split(
                r"(?<=\d\d\d)(?=\d{1,2}\s*=)"
                r"|(?<=\d)(?=[0-9]\s*(?:bk|ဘရိတ်|B|ပတ်|ပါတ်|ပက်|ခွေ|ထိပ်|ထိတ်|ပိတ်))"
                r"|\s+(?=\d{1,2}\s*=)",
                part, flags=re.I)
            for sp in sub_parts:
                sp_clean = sp.strip(" /*")
                if not sp_clean:
                    continue
                if "/" in sp_clean and ("=" in sp_clean
                        or re.search(r"(?:bk|ဘရိတ်|B)", sp_clean, re.I)):
                    slash_tokens = [t.strip() for t in sp_clean.split("/") if t.strip()]
                    carry, fixed = None, []
                    for tok in reversed(slash_tokens):  # inherit NEAREST following amount
                        m_amt = (re.search(r"=\s*(\d[\d,]*)\s*$", tok)
                                 or re.search(r"(?:bk|ဘရိတ်|B)\s*(\d[\d,]*)\s*$", tok, re.I))
                        if m_amt:
                            carry = m_amt.group(1)
                            fixed.append(tok)
                        elif carry:
                            fixed.append(f"{tok} {carry}")
                        else:
                            fixed.append(tok)
                    segments.extend(reversed(fixed))
                else:
                    segments.append(sp_clean)
        if len(segments) > 1:
            return segments
    return [line]


def parse_bets_locally(text):
    text = normalize_text(text)
    rows, warnings, errors = [], [], []
    pending = []
    raw_lines = text.replace("၊", "\n").replace(";", "\n").replace("\r", "\n").split("\n")
    expanded_lines = []
    for rl in raw_lines:
        if rl.strip():
            expanded_lines.extend(split_mixed_line(rl))
    for ln, raw in enumerate(expanded_lines, 1):
        line = raw.strip()
        if not line:
            continue
        m = re.search(r"နေ့ရှင်း\s*([\d,]+)", line) or re.search(r"([\d,]+)\s*နေ့ရှင်း", line)
        if m:
            amt = int(m.group(1).replace(",", ""))
            rest = (line[:m.start()] + line[m.end():]).strip()
            for t in re.findall(r"(?<!\d)\d{1,2}(?!\d)", rest):
                add_number(rows, t, 0, is_default=True, category="main", desc="နေ့ရှင်း")
                pending.append(rows[-1])
            if not pending:
                add_number(rows, "11", amt, category="main", desc=f"{amt:,} နေ့ရှင်း")
                warnings.append(f"Line {ln}: နေ့ရှင်း {amt:,} (no preceding numbers → defaulted to 11)")
            else:
                per_num_amt = amt // len(pending)
                rem = amt % len(pending)
                for idx, r in enumerate(pending):
                    r["Amount"] = per_num_amt + (rem if idx == 0 else 0)
                    r["_default"] = False
                    r["Description"] = f"{amt:,} နေ့ရှင်း ({', '.join(x['Number'] for x in pending)})"
                warnings.append(f"Line {ln}: နေ့ရှင်း {amt:,} → {len(pending)} number(s)")
                pending.clear()
            continue
        try:
            new_rows, new_warns = parse_line(line)
            for r in new_rows:
                if r.get("_default"):
                    pending.append(r)
            rows.extend(new_rows)
            warnings.extend(f"Line {ln}: {w}" for w in new_warns)
        except Exception as e:
            errors.append(f"Line {ln}: {e}")
    for r in rows:
        r.pop("_default", None)
    return {"rows": rows, "warnings": warnings, "errors": errors}


def process_and_add_bets(extracted, date_str, day, customer,
                         existing_batch_ids=(), blocked_numbers=(),
                         min_bet=MIN_BET, max_bet=MAX_BET):
    valid, blocked = [], []
    invalid, total = 0, 0
    nums = [int(m.group(1)) for b in set(existing_batch_ids)
            if (m := re.search(r"\((\d+)\)", str(b)))]
    next_n = max(nums) + 1 if nums else 1
    batch_id = f"{customer}({next_n})"
    blocked_set = set(blocked_numbers)
    for row in extracted:
        try:
            amt = int(row.get("Amount", 0))
            num = normalize_text(str(row.get("Number", ""))).strip().zfill(2)
            cat = str(row.get("Category", "main") or "main")
            desc = str(row.get("Description", num) or num)
        except (ValueError, TypeError):
            invalid += 1
            continue
        if len(num) != 2 or not (num.isascii() and num.isdigit()) or amt < min_bet or amt > max_bet:
            invalid += 1
            continue
        if num in blocked_set:
            blocked.append(num)
            continue
        valid.append({
            "Date": date_str, "Day": day, "Customer": customer, "BatchID": batch_id,
            "Category": cat, "Description": desc, "Number": num, "Amount": amt,
        })
        total += amt
    return valid, blocked, invalid, total


def generate_styled_excel(df, comm_pct, payout, win_num_in, limit, generated_at=None):
    out = io.BytesIO()
    wb = openpyxl.Workbook()
    hdr_fill = PatternFill("solid", fgColor="1E293B")
    accent = PatternFill("solid", fgColor="FEF3C7")
    zebra = PatternFill("solid", fgColor="F8FAFC")
    profit = PatternFill("solid", fgColor="DCFCE7")
    loss = PatternFill("solid", fgColor="FEE2E2")
    f_title = Font(name="Pyidaungsu", size=14, bold=True, color="1E293B")
    f_hdr = Font(name="Pyidaungsu", size=11, bold=True, color="FFFFFF")
    f_body = Font(name="Pyidaungsu", size=11, color="0F172A")
    f_bold = Font(name="Pyidaungsu", size=11, bold=True, color="0F172A")
    bd = Border(left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
                top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"))

    def sheet_fmt(ws, widths, freeze="A4"):
        ws.views.sheetView[0].showGridLines = True
        for c, w in widths.items():
            ws.column_dimensions[c].width = w
        ws.freeze_panes = freeze

    def hdr(ws, row, headers):
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=i, value=h)
            cell.fill, cell.font = hdr_fill, f_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = bd

    win = str(win_num_in).strip().zfill(2) if win_num_in else ""
    tot = float(df["Amount"].sum())
    comm = tot * comm_pct / 100
    net = tot - comm
    win_tot = float(df[df["Number"] == win]["Amount"].sum()) if win else 0
    payout_amt = win_tot * payout
    pl = net - payout_amt

    ws = wb.active
    ws.title = "စာရင်းချုပ်အကျဉ်း"
    ws["A1"] = "မြန်မာနှစ်လုံးထီ (2D) ငွေစာရင်းချုပ်"
    ws["A1"].font = f_title
    ws["A2"] = f"ထုတ်ပြန်သည့်ရက်စွဲ - {(generated_at or datetime.now()).strftime('%Y-%m-%d %H:%M')}"
    ws["A2"].font = f_body
    summ = [
        ("အကြောင်းအရာ", "ပမာဏ", "မှတ်ချက်"),
        ("ရောင်းရငွေ", tot, f"အကွက် {len(df):,}"),
        (f"ကော်မရှင် ({comm_pct}%)", comm, ""),
        ("Net Sales", "=B5-B6", ""),
        ("ထွက်ဂဏန်း", win or "N/A", f"{payout} ဆ"),
        ("ပေါက်ဂဏန်းတန်ဖိုး", win_tot, ""),
        ("ဆုကြေး", f"=B9*{payout}", ""),
        ("အသားတင်", "=B7-B10", "(+) / (-)"),
    ]
    for i, vals in enumerate(summ):
        r = i + 4
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = bd
            if i == 0:
                c.fill, c.font = hdr_fill, f_hdr
                c.alignment = Alignment(horizontal="center")
            else:
                c.font = f_body
                if i == 3:
                    c.fill = accent
                if i == 7:
                    c.fill = profit if pl >= 0 else loss
                if j == 2 and i != 4:
                    c.number_format = "#,##0"
                    c.alignment = Alignment(horizontal="right")
    sheet_fmt(ws, {"A": 42, "B": 24, "C": 38}, freeze="A5")

    ws = wb.create_sheet("နေ့အလိုက်_အပတ်စဉ်")
    ws["A1"] = "နေ့အလိုက် စာရင်းချုပ်"
    ws["A1"].font = f_title
    dg = df.groupby(["Date", "Day"]).agg(Bets=("Number", "count"), Sales=("Amount", "sum")).reset_index()
    hdr(ws, 3, ["ရက်စွဲ", "နေ့", "အကွက်", "ထိုးကြေး", f"ကော်မရှင် ({comm_pct}%)", "Net"])
    for r, row in enumerate(dg.itertuples(), 4):
        vals = [row.Date, str(row.Day), int(row.Bets), float(row.Sales),
                f"=D{r}*{comm_pct/100}", f"=D{r}-E{r}"]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci >= 3:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra
    sheet_fmt(ws, {"A": 18, "B": 16, "C": 14, "D": 22, "E": 20, "F": 22})

    ws = wb.create_sheet("ဖောက်သည်အလိုက်")
    ws["A1"] = "ဖောက်သည်အလိုက် ရှင်းတမ်း"
    ws["A1"].font = f_title
    cg = df.groupby("Customer")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
    hdr(ws, 3, ["ဖောက်သည်", "ထိုးကြေး", f"ကော်မရှင် ({comm_pct}%)", "Net", f"ဆုကြေး ({payout}ဆ)", "အသားတင်"])
    for r, row in enumerate(cg.itertuples(), 4):
        cwin = 0
        if win:
            cwin = float(df[(df["Customer"] == row.Customer) & (df["Number"] == win)]["Amount"].sum()) * payout
        vals = [row.Customer, float(row.Amount), f"=B{r}*{comm_pct/100}", f"=B{r}-C{r}", cwin, f"=D{r}-E{r}"]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci > 1:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra
    sheet_fmt(ws, {"A": 26, "B": 22, "C": 20, "D": 24, "E": 24, "F": 24})

    ws = wb.create_sheet("ဂဏန်း_Limit")
    ws["A1"] = "ဂဏန်းအလိုက် Limit"
    ws["A1"].font = f_title
    ng = df.groupby("Number")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
    hdr(ws, 3, ["2D", "စုစုပေါင်း", "ဒိုင်လက်ခံ", "ကျော်လွန်"])
    for r, row in enumerate(ng.itertuples(), 4):
        nc = ws.cell(row=r, column=1, value=str(row.Number).zfill(2))
        nc.number_format, nc.alignment = "@", Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=float(row.Amount)).number_format = "#,##0"
        ws.cell(row=r, column=3, value=f"=MIN(B{r},{limit})").number_format = "#,##0"
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = "#,##0"
        for ci in range(1, 5):
            ws.cell(row=r, column=ci).font = f_body
            ws.cell(row=r, column=ci).border = bd
        if float(row.Amount) > limit:
            ws.cell(row=r, column=4).fill = loss
            ws.cell(row=r, column=4).font = f_bold
    sheet_fmt(ws, {"A": 16, "B": 24, "C": 24, "D": 26})

    ws = wb.create_sheet("အသေးစိတ်")
    ws["A1"] = "အသေးစိတ် မှတ်တမ်း"
    ws["A1"].font = f_title
    dg = df.groupby(["Date", "Day", "Customer", "BatchID", "Number"])["Amount"].sum().reset_index()
    hdr(ws, 3, ["ရက်စွဲ", "နေ့", "ဖောက်သည်", "Batch", "2D", "ထိုးငွေ"])
    for r, row in enumerate(dg.itertuples(), 4):
        vals = [row.Date, row.Day, row.Customer, row.BatchID, str(row.Number).zfill(2), float(row.Amount)]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci == 5:
                c.number_format = "@"
            if ci == 6:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra
    sheet_fmt(ws, {"A": 16, "B": 16, "C": 22, "D": 22, "E": 16, "F": 22})

    wb.save(out)
    return out.getvalue()
