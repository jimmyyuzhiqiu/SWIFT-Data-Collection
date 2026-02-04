
# swift_core.py
import os
import re
from datetime import datetime
import pandas as pd

# =========================
# 默认配置（按你的实际路径）
# =========================
DEFAULT_MSG_FOLDER = r"Z:\To Jimmy Yu\Swift Data Collection\Swift"
DEFAULT_OUTPUT_FOLDER = r"Z:\To Jimmy Yu\Swift Data Collection"

DEFAULT_MAPPING_FILE = r"Z:\To Jimmy Yu\Swift Data Collection\Swift Data Collection.xlsx"
DEFAULT_MAPPING_SHEET = "ACCT Mapping"


# -----------------------------
# Read .msg as text (2 modes)
# -----------------------------
def read_msg_text(path: str) -> str:
    # Try extract_msg (for Outlook .msg)
    try:
        import extract_msg  # pip install extract-msg
        msg = extract_msg.Message(path)
        text = (msg.body or "") + "\n" + (msg.subject or "")
        msg.close()
        if text.strip():
            return text
    except Exception:
        pass

    # Fallback: raw decode (for text-export .msg)
    with open(path, "rb") as f:
        data = f.read()
    for enc in ("utf-8", "utf-16", "latin1"):
        try:
            t = data.decode(enc, errors="ignore")
            if t.strip():
                return t
        except Exception:
            continue
    return data.decode("latin1", errors="ignore")


# -----------------------------
# Normalization helpers
# -----------------------------
def normalize_swift(raw: str) -> str:
    """Remove '-' and non-alnum, uppercase, pad to 11 with X."""
    if not raw:
        return ""
    s = re.sub(r"[^A-Za-z0-9]", "", raw).upper()
    if len(s) < 11:
        s += "X" * (11 - len(s))
    return s[:11]


def parse_amount_to_float(amount_str: str):
    """
    Parse EU/US formats:
      346.000,      -> 346000.00
      4.772.159,07  -> 4772159.07
      633.086,7     -> 633086.70
    """
    if not amount_str:
        return None
    s = amount_str.strip()
    s = re.sub(r"\(.*?\)", "", s)              # drop (011)
    s = re.sub(r"[^0-9,.\-]", "", s).strip()
    if not s:
        return None

    last_comma = s.rfind(",")
    last_dot = s.rfind(".")

    if last_comma != -1 and last_dot != -1:
        # rightmost is decimal
        if last_comma > last_dot:
            dec_sep, thou_sep = ",", "."
        else:
            dec_sep, thou_sep = ".", ","
        s = s.replace(thou_sep, "")
        s = s.replace(dec_sep, ".")
    else:
        sep = "," if last_comma != -1 else "."
        pos = last_comma if last_comma != -1 else last_dot
        digits_after = len(s) - pos - 1
        # 0/1/2 digits -> decimal (also supports trailing comma)
        if digits_after in (0, 1, 2):
            s = s.replace(sep, ".")
        else:
            s = s.replace(sep, "")

    try:
        return float(s)
    except Exception:
        return None


def format_amount(x):
    return "" if x is None else f"{x:,.2f}"


def format_date_to_iso(d: str) -> str:
    d = d.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(d, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return ""


# -----------------------------
# Direction detection (Step1)
# -----------------------------
def detect_direction(text: str) -> str:
    t = text.lower()
    if "destination" in t:
        return "OUT"
    if re.search(r"^\s*sender\s*:", text, flags=re.IGNORECASE | re.MULTILINE):
        return "IN"
    if "sender" in t and "destination" not in t:
        return "IN"
    return ""


# -----------------------------
# SWIFT block extraction
# -----------------------------
def extract_block_lines(text: str, tag: str) -> list[str]:
    """
    IMPORTANT FIX:
    - Do NOT treat the remainder after ':' as data because it is often just a title
      e.g. '50K : Ordering Customer' -> 'Ordering Customer' should be ignored.
    """
    lines = text.splitlines()
    start_pat = re.compile(rf"^\s*{re.escape(tag)}\s*:\s*(.*)$")
    next_tag_pat = re.compile(r"^\s*\d{2}[A-Z]?\s*:\s*")

    out = []
    in_block = False

    for line in lines:
        m = start_pat.match(line)
        if m:
            in_block = True
            # ignore m.group(1) (title), data comes from next lines
            continue

        if in_block:
            if next_tag_pat.match(line):
                break
            if line.strip().startswith("-" * 5):
                break
            cleaned = line.strip()
            if cleaned:
                out.append(cleaned)

    return out


def strip_star(s: str) -> str:
    return re.sub(r"^\*\s*", "", s).strip()


def cleaned_lines(lines: list[str]) -> list[str]:
    return [strip_star(x) for x in lines if strip_star(x)]


# -----------------------------
# Robust pickers for acct/name/swift/bank
# -----------------------------
def looks_like_account(s: str) -> bool:
    # accounts often numeric/iban-ish; allow letters+digits but must contain digits
    return bool(re.search(r"\d", s)) and len(s.replace(" ", "")) >= 6


def pick_account_line1(lines: list[str]) -> str:
    """
    For 50K/59/59F: pick first line that looks like an account (contains digits).
    """
    for s in cleaned_lines(lines):
        if looks_like_account(s):
            return s
    return ""


def pick_name_line2(lines: list[str]) -> str:
    """
    For 50K/59/59F: pick the first "name-like" line after the account line.
    """
    cl = cleaned_lines(lines)
    acct_idx = None
    for i, s in enumerate(cl):
        if looks_like_account(s):
            acct_idx = i
            break
    if acct_idx is not None and acct_idx + 1 < len(cl):
        return cl[acct_idx + 1]
    # fallback: first non-account line
    for s in cl:
        if not looks_like_account(s):
            return s
    return ""


def looks_like_swift(s: str) -> bool:
    """
    SWIFT/BIC-like: should have at least 4 letters and length between 6-20 (before normalization),
    may include hyphens.
    """
    raw = s.strip()
    letters = re.findall(r"[A-Za-z]", raw)
    return (len(letters) >= 4) and (6 <= len(re.sub(r"\s", "", raw)) <= 20)


def pick_swift(lines: list[str]) -> str:
    """
    For 52A/57A: pick first line that looks like swift, skip pure numeric account/ids.
    """
    for s in cleaned_lines(lines):
        if looks_like_account(s) and not looks_like_swift(s):
            continue
        if looks_like_swift(s):
            return normalize_swift(s)
    return ""


def pick_bank_name(lines: list[str]) -> str:
    """
    Prefer narrative lines that are not swift/account.
    Take first 2 such lines.
    """
    cl = cleaned_lines(lines)
    swift = pick_swift(lines)

    cand = []
    for s in cl:
        # skip swift line
        if swift and normalize_swift(s) == swift:
            continue
        # skip account-ish lines
        if looks_like_account(s) and not looks_like_swift(s):
            continue
        # keep bank narrative lines
        if re.search(r"[A-Za-z]", s):
            cand.append(s)

    return " / ".join(cand[:2]).strip()


def parse_32A(text: str):
    b = extract_block_lines(text, "32A")
    cl = cleaned_lines(b)

    date_iso = ""
    ccy = ""
    amt = ""

    # find date line
    for s in cl:
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", s.strip()):
            date_iso = format_date_to_iso(s)
            break

    # find amount line with CCY
    for s in cl:
        m = re.search(r"\b([A-Z]{3})\b\s*([0-9][0-9,.\s]*)(?:\(|$)", s.strip())
        if m:
            ccy = m.group(1).upper()
            amt = format_amount(parse_amount_to_float(m.group(2)))
            break

    return date_iso, ccy, amt


# -----------------------------
# Mapping loader (ACCT Mapping)
# -----------------------------
def load_acct_mapping(mapping_file: str, mapping_sheet: str):
    if not os.path.exists(mapping_file):
        raise FileNotFoundError(
            f"找不到 mapping 文件：{mapping_file}\n请确认 Z 盘已映射且有权限。"
        )

    df = pd.read_excel(mapping_file, sheet_name=mapping_sheet, engine="openpyxl")
    df.columns = [str(c).strip().upper() for c in df.columns]

    need = {"PRIMARY ID", "CCY", "R-TAG"}
    if not need.issubset(set(df.columns)):
        raise ValueError(f"ACCT Mapping sheet 需要列：{need}；当前列：{list(df.columns)}")

    map_by_acct_ccy = {}
    map_by_acct_only = {}

    for _, r in df.iterrows():
        prim = str(r["PRIMARY ID"]).strip()
        ccy  = str(r["CCY"]).strip().upper()
        acct = str(r["R-TAG"]).strip()

        if not acct or acct.lower() == "nan":
            continue
        if not prim or prim.lower() == "nan":
            continue

        map_by_acct_ccy[(acct, ccy)] = prim
        map_by_acct_only.setdefault(acct, prim)

    return map_by_acct_ccy, map_by_acct_only


# -----------------------------
# 59 段专用名字提取（OUT）
# 从账号行的下一行开始，一直到出现 ADD:/ADDRESS:/CITY:/COUNTRY: 之前
# 同时剔除任何 'IBAN:' 及其后内容；处理行内的 '- ADD:' 截断
# -----------------------------
def pick_beneficiary_name_until_add(lines: list[str]) -> str:
    cl = cleaned_lines(lines)

    # 找到账号行索引
    acct_idx = None
    for i, s in enumerate(cl):
        if looks_like_account(s):
            acct_idx = i
            break
    start = acct_idx + 1 if acct_idx is not None else 0

    name_parts = []
    for raw in cl[start:]:
        s = raw.strip()

        # 行内出现 '- ADD:'，只保留其左侧
        if re.search(r"-\s*ADD\s*:", s, flags=re.IGNORECASE):
            s = re.split(r"-\s*ADD\s*:", s, flags=re.IGNORECASE)[0].strip()

        # 遇到 ADD/ADDRESS/CITY/COUNTRY 开头，停止
        if re.match(r"^\s*(ADD|ADDRESS|CITY|COUNTRY)\s*:", s, flags=re.IGNORECASE):
            break

        # 删除 'IBAN: ...' 片段（不停止，只清理）
        s = re.sub(r"\bIBAN\b\s*:\s*[A-Z0-9\s]+", "", s, flags=re.IGNORECASE).strip()

        # 如果整行变空（例如只有 IBAN），则跳过
        if not s:
            continue

        name_parts.append(s)

    # 合并为单行名称
    name = " ".join(name_parts).strip()
    # 压缩多余空格
    name = re.sub(r"\s{2,}", " ", name)

    return name


# -----------------------------
# Per-message extraction
# -----------------------------
def extract_step3_record(text: str) -> dict:
    direction = detect_direction(text)
    date_iso, ccy, amt = parse_32A(text)

    b50k = extract_block_lines(text, "50K")
    b59  = extract_block_lines(text, "59")
    b59f = extract_block_lines(text, "59F")
    b52a = extract_block_lines(text, "52A")
    b57a = extract_block_lines(text, "57A")

    if direction == "OUT":
        client_acct = pick_account_line1(b50k)
        cp_name     = pick_beneficiary_name_until_add(b59)
        cp_acct     = pick_account_line1(b59)
        cp_swift    = pick_swift(b57a)
        cp_bank     = pick_bank_name(b57a)

    elif direction == "IN":
        client_acct = pick_account_line1(b59f) or pick_account_line1(b59)
        cp_name     = pick_name_line2(b50k)
        cp_acct     = pick_account_line1(b50k)
        cp_swift    = pick_swift(b52a)
        cp_bank     = pick_bank_name(b52a)

    else:
        client_acct = cp_name = cp_acct = cp_swift = cp_bank = ""

    return {
        "Client Acct": client_acct,
        "DATE": date_iso,
        "CCY": ccy,
        "AMT": amt,
        "CP NAME": cp_name,
        "CP A/C": cp_acct,
        "CP SWIFT": cp_swift,
        "CP BANK NAME": cp_bank,
        "DIRECTION": direction,
    }


# -----------------------------
# 自动列宽（openpyxl）
# -----------------------------
def autofit_worksheet(ws, df):
    for i, col in enumerate(df.columns, start=1):
        max_len = len(str(col))
        for cell in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=i, max_col=i):
            v = cell[0].value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max_len + 2, 80)


# =========================
# UI 调用入口：带进度/状态回调
# =========================
def run_swift_batch(
    input_dir: str,
    output_dir: str,
    mapping_file: str,
    mapping_sheet: str = DEFAULT_MAPPING_SHEET,
    skip_keywords=None,
    progress_callback=None,   # progress_callback(done:int, total:int, filename:str)
    status_callback=None      # status_callback(message:str)
) -> str:
    if skip_keywords is None:
        skip_keywords = ["FFD", "MT199"]

    if not os.path.exists(input_dir):
        raise FileNotFoundError(f"找不到 msg 文件夹：{input_dir}")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    map_by_acct_ccy, map_by_acct_only = load_acct_mapping(mapping_file, mapping_sheet)

    # 动态输出名：YYYYMMDD_Swift.xlsx
    today_str = datetime.now().strftime("%Y%m%d")
    output_path = os.path.join(output_dir, f"{today_str}_Swift.xlsx")

    files_all = [fn for fn in os.listdir(input_dir) if fn.lower().endswith(".msg")]
    files = [fn for fn in files_all if not any(k in fn.upper() for k in skip_keywords)]
    total = len(files)
    done = 0

    rows = []

    for fn in files:
        path = os.path.join(input_dir, fn)
        try:
            if status_callback:
                status_callback(f"解析中：{fn}")

            text = read_msg_text(path)
            rec = extract_step3_record(text)

            acct = rec["Client Acct"]
            ccy  = rec["CCY"]

            prim = map_by_acct_ccy.get((acct, ccy), "")
            if not prim and acct:
                prim = map_by_acct_only.get(acct, "")

            rec["PRIM ID"] = prim
            rec["FILE"] = fn
            rows.append(rec)

        except Exception as e:
            rows.append({
                "FILE": fn,
                "Client Acct": "",
                "PRIM ID": "",
                "DATE": "",
                "CCY": "",
                "AMT": "",
                "CP NAME": "",
                "CP A/C": "",
                "CP SWIFT": "",
                "CP BANK NAME": "",
                "DIRECTION": "",
                "ERROR": str(e)
            })

        done += 1
        if progress_callback:
            progress_callback(done, total, fn)

    df = pd.DataFrame(rows)

    step3_cols = [
        "Client Acct","PRIM ID","DATE","CCY","AMT",
        "CP NAME","CP A/C","CP SWIFT","CP BANK NAME","DIRECTION"
    ]
    step3 = df.reindex(columns=step3_cols)

    key_cols = ["Client Acct", "DATE", "CCY", "AMT", "CP A/C", "CP SWIFT"]
    mask_valid = step3[key_cols].apply(lambda s: s.astype(str).str.strip().ne("")).any(axis=1)
    step3_final = step3[mask_valid].copy()

    debug_cols = ["FILE","DIRECTION"] + step3_cols + (["ERROR"] if "ERROR" in df.columns else [])
    debug = df.reindex(columns=debug_cols)

    if status_callback:
        status_callback("写入 Excel 中...")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        step3_final.to_excel(writer, sheet_name="Step3_Final", index=False)
        debug.to_excel(writer, sheet_name="Debug", index=False)

        ws_final = writer.sheets["Step3_Final"]
        ws_debug = writer.sheets["Debug"]
        autofit_worksheet(ws_final, step3_final)
        autofit_worksheet(ws_debug, debug)

    if status_callback:
        status_callback(f"完成 ✅ 输出：{output_path}")

    return output_path


if __name__ == "__main__":
    out = run_swift_batch(
        input_dir=DEFAULT_MSG_FOLDER,
        output_dir=DEFAULT_OUTPUT_FOLDER,
        mapping_file=DEFAULT_MAPPING_FILE
    )
    print("输出文件：", out)
