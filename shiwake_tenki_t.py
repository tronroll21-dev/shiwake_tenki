"""
himoduke_export.py
Windows-only script.

Usage:
    python himoduke_export.py <output_dir>
"""

import csv
import os
import sys
from datetime import datetime

try:
    import tkinter
except ImportError:
    print("エラー: tkinter が見つかりません。Windows用Python標準インストーラーには含まれています。")
    sys.exit(1)

BUMON_CODE = "000004"
EXPECTED_TITLE = "売掛残高一覧表"
EXPECTED_HEADERS = ["コード","得意先名","前月売掛残高","入金合計","繰越残高","税抜売上額","消費税等","税込売上額","当月売掛残高"]
NUM_COLUMNS = len(EXPECTED_HEADERS)
HEADER_LINE_COUNT = 4  # lines 1-4 are preamble; line 5 (index 4) is the table header
NOT_FOUND_VALUE = "000000"
DEBIT_CODE_POSITIVE = 109  # 借方科目ｺｰﾄﾞ when 税込売上額 > 0
DEBIT_CODE_NEGATIVE = 601  # 借方科目ｺｰﾄﾞ when 税込売上額 < 0
FISCAL_CLOSING_DAY = 20  # month closes on the 20th
KARIKATA_ZEICODE_POSITIVE = "000"  # karikatazeicode when 税込売上額 > 0
KARIKATA_ZEICODE_NEGATIVE = "112"  # karikatazeicode when 税込売上額 < 0
KARIKATA_ZEIKUBUN_POSITIVE = 0  # karikatazeikubun when 税込売上額 > 0
KARIKATA_ZEIKUBUN_NEGATIVE = 1  # karikatazeikubun when 税込売上額 < 0
KASHIKATA_KAMOKU_CODE_FIRST_POSITIVE = 601  # kashikatakamokucode on the first positive-total line
KASHIKATA_KAMOKU_CODE_NEGATIVE = 109  # kashikatakamokucode on the first negative-total line
KASHIKATA_ZEICODE_FIRST_POSITIVE = "112"  # kashikatazeicode on the first positive-total line
KASHIKATA_ZEICODE_NEGATIVE = "000"  # kashikatazeicode on every negative-total line
KASHIKATA_ZEIKUBUN_FIRST_POSITIVE = 1  # kashikatazeikubun on the first positive-total line
KASHIKATA_ZEIKUBUN_NEGATIVE = 0  # kashikatazeikubun on every negative-total line

OUTPUT_HEADERS = [
    "伝票区分",       # denpyou_kugiri
    "会計年月",         # fiscal_month_str
    "日付",           # date_str
    "行",             # gyou
    "借方科目コード",   # karikata_kamoku_code
    "借方内訳コード",   # karikata_uchiwake_code
    "借方税コード",     # karikata_zeicode
    "借方税区分",       # karikata_zeikubun
    "借方部門コード",   # karikatabumoncode
    "借方金額",         # karikata_kingaku_gokei
    "貸方科目コード",   # kashikata_kamoku_code
    "貸方内訳コード",   # kashikata_uchiwake_code
    "貸方税コード",     # kashikatazeicode
    "貸方税区分",       # kashikatazeikubun
    "貸方部門コード",   # kashikatabumoncode
    "貸方金額",         # kashikata_kingaku_gokei
]

def load_himoduke_dict(path):
    """Read a 2-column CSV (key,value) into a dict. Abort if the file is missing."""
    rows = None
    last_error = None

    # Try UTF-8 first, then fall back to Shift-JIS (cp932), since we don't
    # know for certain which encoding himoduke.csv was saved in.
    for encoding in ("utf-8-sig", "cp932"):
        try:
            with open(path, "r", encoding=encoding, newline="") as f:
                reader = csv.reader(f)
                rows = list(reader)
            break
        except FileNotFoundError:
            print(f"エラー: ファイルが見つかりません: {path}")
            sys.exit(1)
        except UnicodeDecodeError as e:
            last_error = e
            continue

    if rows is None:
        print(f"エラー: ファイルの文字コードを判別できませんでした: {path} ({last_error})")
        sys.exit(1)

    mapping = {}
    for i, row in enumerate(rows, start=1):
        if not row:
            continue
        if len(row) < 2:
            print(f"警告: {i}行目が2列ではありません。スキップします: {row}")
            continue
        key, value = row[0].strip(), row[1].strip()
        mapping[key] = value

    if not mapping:
        print(f"エラー: {path} からデータを読み込めませんでした（空のファイルです）。")
        sys.exit(1)

    return mapping


def load_clipboard_text():
    """Read the raw clipboard text via tkinter (no external dependencies)."""
    try:
        root = tkinter.Tk()
        root.withdraw()
        text = root.clipboard_get()
        root.destroy()
    except tkinter.TclError as e:
        print(f"エラー: クリップボードの内容を読み取れませんでした（クリップボードが空か、テキストではありません）: {e}")
        sys.exit(1)

    return text


def split_clipboard_lines(text):
    """Split clipboard text into lines, tolerating \\r\\n and trailing blank lines."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    # Drop a single trailing empty line (common when Excel copy ends with a newline)
    if lines and lines[-1] == "":
        lines.pop()
    return lines


def validate_title(lines):
    """Check that line 1 is the expected report title."""
    if len(lines) < HEADER_LINE_COUNT + 1:  # need at least the 4 header lines + table header
        print(f"エラー: クリップボードのデータが短すぎます（{len(lines)}行）。処理を中止します。")
        sys.exit(1)

    title = lines[0].strip()
    if title != EXPECTED_TITLE:
        print(f"エラー: 1行目のタイトルが想定と一致しません。処理を中止します。")
        print(f"  想定: {EXPECTED_TITLE}")
        print(f"  実際: {title}")
        sys.exit(1)


def parse_row(line):
    """Split a tab-separated line into fields."""
    return [f.strip() for f in line.split("\t")]


def validate_table_header(lines):
    """Line 5 (index 4) must be the table header row with the expected column names."""
    header_line = lines[HEADER_LINE_COUNT]
    actual_headers = parse_row(header_line)
    if actual_headers != EXPECTED_HEADERS:
        print("エラー: 表のヘッダー行（5行目）が想定と一致しません。処理を中止します。")
        print(f"  想定: {EXPECTED_HEADERS}")
        print(f"  実際: {actual_headers}")
        sys.exit(1)


def clean_code(code):
    """Strip a single leading '<' and a single trailing '>' from the code, if present."""
    code = code.strip()
    if code.startswith("<"):
        code = code[1:]
    if code.endswith(">"):
        code = code[:-1]
    return code


def parse_amount(amount_str, line):
    """Strip thousands-separator commas and convert 税込売上額 to a number. Abort on failure."""
    cleaned = amount_str.replace(",", "").strip()
    try:
        return int(cleaned)
    except ValueError:
        try:
            return float(cleaned)
        except ValueError:
            print(f"エラー: 税込売上額を数値に変換できませんでした: '{amount_str}'。処理を中止します。")
            print(f"  行: {line}")
            sys.exit(1)


def build_output_rows(lines, himoduke_dict):
    """
    Aggregate 税込売上額 (numeric) by looked-up コード value.
    (コード itself is not written — its dictionary value is written instead,
    and codes sharing the same looked-up value are summed together.)
    Data starts the line after the table header (line 6 overall).
    """
    aggregated = {}  # looked_up_value -> running total
    missing_codes = set()
    not_found_codes = set()

    data_lines = lines[HEADER_LINE_COUNT + 1:]

    for line in data_lines:
        if line.strip() == "":
            continue  # skip stray blank lines

        fields = parse_row(line)
        if len(fields) != NUM_COLUMNS:
            print(f"エラー: データ行の列数が{NUM_COLUMNS}列ではありません（{len(fields)}列でした）。処理を中止します。")
            print(f"  行: {line}")
            sys.exit(1)

        code, _customer_name, _previous_balance, _total_receipts, _carried_forward_balance, \
            _sales_amount_excl_tax, _consumption_tax, sales_amount_incl_tax, _current_balance = fields

        if code.strip() == "":
            break  # reached the total row (blank code) — stop processing here

        code = clean_code(code)

        looked_up_value = himoduke_dict.get(code)
        if looked_up_value is None:
            missing_codes.add(code)
            looked_up_value = ""  # aggregate under blank if the code isn't found
        elif looked_up_value == NOT_FOUND_VALUE:
            not_found_codes.add(code)

        numeric_amount = parse_amount(sales_amount_incl_tax, line)

        if numeric_amount == 0:
            continue

        aggregated[looked_up_value] = aggregated.get(looked_up_value, 0) + numeric_amount

    if missing_codes:
        print("警告: 以下のコードが himoduke.csv に見つかりませんでした（空欄で出力されます）:")
        for c in sorted(missing_codes):
            print(f"  - {c}")

    if not_found_codes:
        print(f"警告: 以下のコードの紐付け値が「{NOT_FOUND_VALUE}」でした:")
        for c in sorted(not_found_codes):
            print(f"  - {c}")

    return aggregated


def order_rows_and_compute_totals(aggregated):
    """
    Split aggregated (looked_up_value -> total) rows into a positive-total
    group and a negative-total group, placing the negative group last.
    Also compute each group's grand total.
    Returns: (ordered_rows, num_positive_rows, total_positive, total_negative)
    """
    positive_rows = []
    negative_rows = []

    for looked_up_value, total in aggregated.items():
        if total > 0:
            positive_rows.append((looked_up_value, total))
        else:
            negative_rows.append((looked_up_value, total))

    total_positive = sum(total for _, total in positive_rows)
    total_negative = sum(total for _, total in negative_rows)

    ordered_rows = positive_rows + negative_rows
    return ordered_rows, len(positive_rows), total_positive, total_negative


def prompt_for_date():
    """Ask the user for a date (yyyymmdd), defaulting to today if left blank."""
    default_date = datetime.now()
    default_str = default_date.strftime("%Y%m%d")

    while True:
        entered = input(f"日付を入力してください (yyyymmdd) [{default_str}]: ").strip()
        if entered == "":
            return default_date
        try:
            return datetime.strptime(entered, "%Y%m%d")
        except ValueError:
            print("エラー: yyyymmdd形式で入力してください。")


def compute_fiscal_month(date_obj):
    """
    Derive the fiscal month (yyyymm) for a date, given a monthly closing
    on FISCAL_CLOSING_DAY: dates on or before the closing day belong to
    that calendar month; dates after it belong to the following month.
    """
    if date_obj.day <= FISCAL_CLOSING_DAY:
        fiscal_date = date_obj
    else:
        if date_obj.month == 12:
            fiscal_date = date_obj.replace(year=date_obj.year + 1, month=1, day=1)
        else:
            fiscal_date = date_obj.replace(month=date_obj.month + 1, day=1)

    return fiscal_date.strftime("%Y%m")


def finalize_rows(ordered_rows, num_positive_rows, total_positive, total_negative, date_str, fiscal_month_str):

    final_rows = []
    for i, (looked_up_value, total) in enumerate(ordered_rows):
        is_positive = total > 0
        abs_total = abs(total)
        is_first_negative = (i == num_positive_rows)
        is_first_positive = (i == 0 and num_positive_rows > 0)

        denpyou_kugiri = 1 if i == 0 else 0
        gyou = i + 1

        # karikata-side fields: every positive row, plus only the first
        # negative row — not repeated on later negative rows.
        if is_positive:
            karikatabumoncode = BUMON_CODE
            karikata_zeicode = KARIKATA_ZEICODE_POSITIVE
            karikata_zeikubun = KARIKATA_ZEIKUBUN_POSITIVE
            karikata_kamoku_code = DEBIT_CODE_POSITIVE
        elif is_first_negative:
            karikatabumoncode = BUMON_CODE
            karikata_zeicode = KARIKATA_ZEICODE_NEGATIVE
            karikata_zeikubun = KARIKATA_ZEIKUBUN_NEGATIVE
            karikata_kamoku_code = DEBIT_CODE_NEGATIVE
        else:
            karikatabumoncode = ""
            karikata_zeicode = ""
            karikata_zeikubun = ""
            karikata_kamoku_code = ""

        # kashikata-side fields: only the first positive row, plus every
        # negative row.
        if is_first_positive:
            kashikata_kamoku_code = KASHIKATA_KAMOKU_CODE_FIRST_POSITIVE
            kashikatabumoncode = BUMON_CODE
            kashikatazeicode = KASHIKATA_ZEICODE_FIRST_POSITIVE
            kashikatazeikubun = KASHIKATA_ZEIKUBUN_FIRST_POSITIVE
        elif is_first_negative:
            kashikata_kamoku_code = KASHIKATA_KAMOKU_CODE_NEGATIVE
            kashikatabumoncode = BUMON_CODE
            kashikatazeicode = KASHIKATA_ZEICODE_NEGATIVE
            kashikatazeikubun = KASHIKATA_ZEIKUBUN_NEGATIVE
        elif not is_positive:
            # later negative rows: no kamoku code, but bumon/zeicode repeat
            kashikata_kamoku_code = KASHIKATA_KAMOKU_CODE_NEGATIVE
            kashikatabumoncode = BUMON_CODE
            kashikatazeicode = KASHIKATA_ZEICODE_NEGATIVE
            kashikatazeikubun = KASHIKATA_ZEIKUBUN_NEGATIVE
        else:
            # later positive rows
            kashikatabumoncode = ""
            kashikatazeicode = ""
            kashikatazeikubun = ""
            kashikata_kamoku_code = ""

        if is_positive:
            karikata_uchiwake_code = looked_up_value
            kashikata_uchiwake_code = ""
        else:
            karikata_uchiwake_code = ""
            kashikata_uchiwake_code = looked_up_value

        # karikata_kingaku_gokei / kashikata_kingaku_gokei: default both to
        # blank, then let exactly one branch below set the applicable one.
        karikata_kingaku_gokei = ""
        kashikata_kingaku_gokei = ""

        if is_first_positive:
            karikata_kingaku_gokei = abs_total
            kashikata_kingaku_gokei = abs(total_positive)
        elif is_first_negative:
            karikata_kingaku_gokei = abs(total_negative)
            kashikata_kingaku_gokei = abs_total
        elif is_positive:
            karikata_kingaku_gokei = abs_total
        elif not is_positive:
            kashikata_kingaku_gokei = abs_total

        final_rows.append([
            denpyou_kugiri,
            fiscal_month_str,
            date_str,
            gyou,
            karikata_kamoku_code,
            karikata_uchiwake_code,
            karikata_zeicode,
            karikata_zeikubun,
            karikatabumoncode,
            karikata_kingaku_gokei,
            kashikata_kamoku_code,
            kashikata_uchiwake_code,
            kashikatazeicode,
            kashikatazeikubun,
            kashikatabumoncode,
            kashikata_kingaku_gokei,
        ])

    return final_rows


def write_output_csv(rows, output_dir):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_dir, f"out_{timestamp}.txt")

    try:
        with open(output_path, "w", encoding="cp932", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(OUTPUT_HEADERS)
            writer.writerows(rows)
    except UnicodeEncodeError as e:
        print(f"エラー: Shift-JISでエンコードできない文字が含まれています: {e}")
        sys.exit(1)
    except OSError as e:
        print(f"エラー: ファイルを書き込めませんでした ({output_path}): {e}")
        sys.exit(1)

    return output_path


def main():
    if len(sys.argv) != 2:
        print(f"使い方: python {os.path.basename(sys.argv[0])} <output_dir>")
        sys.exit(1)

    output_dir = sys.argv[1]

    himoduke_path = os.path.join(output_dir, f"himoduke_t.csv")

    himoduke_dict = load_himoduke_dict(himoduke_path)

    text = load_clipboard_text()
    lines = split_clipboard_lines(text)
    validate_title(lines)
    validate_table_header(lines)

    aggregated = build_output_rows(lines, himoduke_dict)
    if not aggregated:
        print("エラー: 出力対象のデータ行がありません。処理を中止します。")
        sys.exit(1)

    ordered_rows, num_positive_rows, total_positive, total_negative = order_rows_and_compute_totals(aggregated)

    date_obj = prompt_for_date()
    date_str = date_obj.strftime("%Y/%m/%d")
    fiscal_month_str = compute_fiscal_month(date_obj)

    final_rows = finalize_rows(
        ordered_rows, num_positive_rows, total_positive, total_negative, date_str, fiscal_month_str
    )

    output_path = write_output_csv(final_rows, output_dir)
    print(f"完了しました。出力先: {output_path} ({len(final_rows)}件)")


if __name__ == "__main__":
    main()