import re
import json
import base64
import datetime
import urllib.request
import urllib.error
import streamlit as st
import streamlit.components.v1 as components
import anthropic
import pandas as pd
from io import BytesIO
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

try:
    import jpholiday
    HAS_JPHOLIDAY = True
except ImportError:
    HAS_JPHOLIDAY = False

st.set_page_config(
    page_title="献立チェックシステム",
    layout="wide",
    page_icon="🍱"
)

BASE_DIR = Path(__file__).parent
RULES_FILE = BASE_DIR / "rules.txt"

GITHUB_OWNER = "fukiyamiyazaki-blip"
GITHUB_REPO = "-"
GITHUB_BRANCH = "main"
GITHUB_RULES_PATH = "rules.txt"

# Streamlit の C キーショートカット（キャッシュクリア）を無効化
components.html("""
<script>
try {
    function blockCKey(e) {
        if (e.key !== 'c' && e.key !== 'C') return;
        var el = window.parent.document.activeElement;
        if (el) {
            var tag = el.tagName.toUpperCase();
            if (tag === 'INPUT' || tag === 'TEXTAREA' || el.isContentEditable) return;
        }
        e.stopImmediatePropagation();
        e.stopPropagation();
    }
    window.parent.document.addEventListener('keydown', blockCKey, true);
    window.parent.document.addEventListener('keyup', blockCKey, true);
    window.parent.addEventListener('keydown', blockCKey, true);
    window.parent.addEventListener('keyup', blockCKey, true);
} catch(err) {}
</script>
""", height=0, scrolling=False)


def load_rules():
    if RULES_FILE.exists():
        return RULES_FILE.read_text(encoding="utf-8")
    return ""


def save_rules(text):
    RULES_FILE.write_text(text, encoding="utf-8")


# ─────────────────────────────────────────────
# Python事前計算用ヘルパー
# ─────────────────────────────────────────────

def _parse_date(ds, year):
    m = re.match(r'(\d+)/(\d+)\(', ds)
    if not m:
        return None
    try:
        return datetime.date(year, int(m.group(1)), int(m.group(2)))
    except ValueError:
        return None


def _dow(ds):
    """月=0 … 日=6"""
    m = re.search(r'\(([月火水木金土日])\)', ds)
    return "月火水木金土日".index(m.group(1)) if m else -1


def _is_holiday(d):
    if not HAS_JPHOLIDAY or d is None:
        return False
    try:
        return bool(jpholiday.is_holiday(d))
    except Exception:
        return False


def _split_ing(ing_text):
    """材料テキストを個別トークンに分割"""
    parts = re.split(r'[,、]', ing_text)
    result = []
    for part in parts:
        # 半角・全角カッコも区切り文字として扱う
        for token in re.split(r'[\s　・／/()（）]+', part.strip()):
            t = token.strip()
            if t and t not in ('nan', ''):
                result.append(t)
    return result


def _parse_structured(excel_text):
    """
    excel_text を日付ごとの構造体に変換。
    Returns: (year, month_num, sorted_dates, entries)
      entries[date_str] = {'lunch': str, 'snack': str, 'ingredients': str}
    """
    year, month_num = 0, 0
    ym = re.search(r'(\d{4})年(\d+)月', excel_text)
    if ym:
        year, month_num = int(ym.group(1)), int(ym.group(2))

    entries, current = {}, None
    for line in excel_text.split('\n'):
        dm = re.match(r'【(\d+/\d+\([月火水木金土日]\))】', line)
        if dm:
            current = dm.group(1)
            entries[current] = {'lunch': '', 'snack': '', 'ingredients': ''}
        elif current:
            if line.startswith('昼食:'):
                entries[current]['lunch'] = line[3:].strip()
            elif line.startswith('おやつ:'):
                entries[current]['snack'] = line[4:].strip()
            elif line.startswith('材料:'):
                entries[current]['ingredients'] = line[3:].strip()

    sorted_dates = sorted(
        entries.keys(),
        key=lambda ds: (_parse_date(ds, year) or datetime.date.max)
    )
    return year, month_num, sorted_dates, entries


# ─────────────────────────────────────────────
# チェック用定数
# ─────────────────────────────────────────────

# 2日連続チェック免除（完全一致）
_EXEMPT_EXACT = {
    '白米', 'しょうが', 'にんにく', 'みそ', '酢', '白ごま',
    '人参', '玉ねぎ', '鶏肉', '豚肉', '牛肉', 'ひき肉',
    '昆布', 'かつお',  # 天然だし(かつお・昆布) をカッコ分割した際の破片を免除
}
# 2日連続チェック免除（部分一致：これを含むトークンは免除）
_EXEMPT_SUB = ['醤油', '砂糖', 'みりん', '酒', '塩', '油',
               'だし', 'ごま', '水', '片栗粉', '小麦粉']
# 調味料独自チェックで管理するもの（汎用2日連続から除外）
_SEASONING_HANDLED = {'シャンタン', '中華だし', 'コンソメ', 'カレー', 'ソース', 'チーズ'}


def _is_exempt(token):
    if token in _EXEMPT_EXACT or token in _SEASONING_HANDLED:
        return True
    return any(s in token for s in _EXEMPT_SUB)


SEASONING_2DAY = ['シャンタン', '中華だし', 'コンソメ', 'カレー', 'ソース']
MEAT_3DAY      = ['豚肉', '鶏肉', '牛肉', 'ひき肉']
IMO_KW         = ['じゃが芋', 'さつま芋', '里芋', 'かぼちゃ']
NERIMONO_KW    = ['ちくわ', 'かにかま', '赤かまぼこ']
FISH_KW        = ['サケ', 'サーモン', 'サバ', 'サワラ', 'タラ', 'アジ', 'イワシ',
                  'ブリ', 'カレイ', 'メカジキ', 'タイ', 'マグロ', 'カツオ',
                  'シシャモ', 'ほっけ', '白身魚']
FISH_WITH_TUNA = FISH_KW + ['ツナ']
TOFU_KW        = ['木綿豆腐', '焼き豆腐', '厚揚げ', '油揚げ', '高野豆腐']
NOODLE_KW      = ['スパゲティ', 'うどん', 'めん', '麺', 'そば', '丼']
MUSHROOM_KW    = ['しめじ', 'エリンギ', 'えのき', 'なめこ']
MON_NG_ITEMS   = ['ロールパン', '食パン', 'りんご', 'バナナ', 'オレンジ',
                  '太もやし', '切干大根', 'ひじき', '高野豆腐']
PREP_KW        = [
    '玉ねぎ', '人参', '白菜', '焼き豆腐', 'かぼちゃ', '大根', 'いんげん',
    'えのき', 'キャベツ', 'じゃが芋', '里芋', '厚揚げ', 'しめじ', 'れんこん',
    '木綿豆腐', 'なめこ', 'エリンギ', 'ちくわ', '板こんにゃく', '糸こんにゃく',
    'さつま芋', 'ごぼう', 'ささがきごぼう', 'にら', 'マッシュルーム',
    'きゅうり', 'トマト', 'なす', '春雨', '切干大根', 'ロースハム',
    'ウインナー', 'たけのこ', 'アスパラガス',
]


def get_api_key():
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return ""


def get_github_token():
    try:
        return st.secrets.get("GITHUB_TOKEN", "")
    except Exception:
        return ""


def push_rules_to_github(text):
    token = get_github_token()
    if not token:
        return False, "GitHubトークンが未設定です"

    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
        f"/contents/{GITHUB_RULES_PATH}"
    )
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    }

    # 現在のSHAを取得
    req = urllib.request.Request(
        api_url + f"?ref={GITHUB_BRANCH}", headers=headers
    )
    try:
        with urllib.request.urlopen(req) as resp:
            sha = json.loads(resp.read())["sha"]
    except Exception as e:
        return False, f"GitHub取得エラー: {e}"

    # ファイルを更新
    payload = json.dumps({
        "message": "ルール管理から更新",
        "content": base64.b64encode(text.encode("utf-8")).decode("utf-8"),
        "sha": sha,
        "branch": GITHUB_BRANCH,
    }).encode("utf-8")

    req = urllib.request.Request(api_url, data=payload, headers=headers, method="PUT")
    try:
        with urllib.request.urlopen(req):
            return True, "保存しました（GitHub反映済み）"
    except Exception as e:
        return False, f"GitHub更新エラー: {e}"


def get_sheet_names(uploaded_file):
    engine = "xlrd" if uploaded_file.name.lower().endswith(".xls") else "openpyxl"
    try:
        xl = pd.ExcelFile(uploaded_file, engine=engine)
        return xl.sheet_names
    except Exception as e:
        st.error(f"ファイルの読み込みに失敗しました: {e}")
        return []


def excel_to_text(uploaded_file, sheet_name):
    engine = "xlrd" if uploaded_file.name.lower().endswith(".xls") else "openpyxl"
    df = pd.read_excel(
        uploaded_file, sheet_name=sheet_name, header=None, dtype=str, engine=engine
    )
    n_rows, n_cols = df.shape

    def cell_val(r, c):
        if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
            return ""
        v = str(df.iloc[r, c]).strip()
        return "" if v in ("nan", "", "None") else v

    # 日付行を全行から検索（「N日(曜)」パターンが3個以上ある行をブロック開始とする）
    blocks = []  # list of (date_row_idx, {col: date_str})
    for r in range(n_rows):
        temp = {}
        for c in range(n_cols):
            v = cell_val(r, c)
            if re.match(r'^\d+日\([月火水木金土日]\)$', v):
                temp[c] = v
        if len(temp) >= 3:
            blocks.append((r, temp))

    # 日付構造が見つからない場合は行ごと出力にフォールバック
    if not blocks:
        rows = []
        for i, row in df.iterrows():
            cells = [
                str(v).strip() if pd.notna(v) and str(v).strip() not in ("nan", "") else ""
                for v in row
            ]
            if any(cells):
                rows.append(f"行{i + 1}: " + " | ".join(cells))
        return "\n".join(rows)

    # 年月を抽出（例：「2026年09月」）
    year_month = ""
    month_num = ""
    for r in range(min(10, n_rows)):
        for c in range(min(10, n_cols)):
            v = cell_val(r, c)
            m = re.match(r'(\d{4})年(\d{2})月', v)
            if m:
                year_month = f"{m.group(1)}年{m.group(2)}月"
                month_num = str(int(m.group(2)))
                break
        if year_month:
            break

    skip_vals = {"[昼]", "[午後]", "献立名", "材料", "日付"}

    def is_valid_cell(v):
        if not v or v in skip_vals:
            return False
        if v.startswith("※"):  # 注記・免責文を除外
            return False
        return True

    lines = []
    if year_month:
        lines.append(f"# 献立データ {year_month}")
        lines.append("")

    for block_idx, (block_date_row, block_date_cols) in enumerate(blocks):
        # ブロック終端（次のブロック開始行 or ファイル末尾）
        block_end = blocks[block_idx + 1][0] if block_idx + 1 < len(blocks) else n_rows

        # このブロック内の材料セクション開始行
        block_mat_row = None
        for r in range(block_date_row + 1, block_end):
            for c in range(min(5, n_cols)):
                if cell_val(r, c) == "材料":
                    block_mat_row = r
                    break
            if block_mat_row is not None:
                break

        dish_end = block_mat_row if block_mat_row is not None else block_end

        for col_c in sorted(block_date_cols.keys()):
            raw_date = block_date_cols[col_c]
            dm = re.match(r'(\d+)日\(([月火水木金土日])\)', raw_date)
            if dm and month_num:
                date_label = f"{month_num}/{dm.group(1)}({dm.group(2)})"
            else:
                date_label = raw_date

            lines.append(f"【{date_label}】")

            # [午後]マーカーの行を検索（その日付列の1列前に出現する）
            afternoon_start = dish_end
            if col_c > 0:
                for r in range(block_date_row + 1, dish_end):
                    if cell_val(r, col_c - 1) == "[午後]":
                        afternoon_start = r
                        break

            # 昼食献立
            lunch = []
            for r in range(block_date_row + 1, afternoon_start):
                v = cell_val(r, col_c)
                if is_valid_cell(v):
                    lunch.append(v)

            # おやつ
            snack = []
            for r in range(afternoon_start, dish_end):
                v = cell_val(r, col_c)
                if is_valid_cell(v):
                    snack.append(v)

            # 材料（ブロック内のみ）
            mats = []
            if block_mat_row is not None:
                for r in range(block_mat_row, block_end):
                    v = cell_val(r, col_c)
                    if is_valid_cell(v):
                        mats.append(v)

            if lunch:
                lines.append(f"昼食: {' / '.join(lunch)}")
            if snack:
                lines.append(f"おやつ: {' / '.join(snack)}")
            if mats:
                lines.append(f"材料: {', '.join(mats)}")
            lines.append("")

    return "\n".join(lines)


def table_to_docx(markdown_text):
    from docx.oxml.ns import qn as _qn

    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    # 用紙サイズを A4縦 として明示
    pgSz = section._sectPr.find(_qn('w:pgSz'))
    if pgSz is not None:
        pgSz.set(_qn('w:orient'), 'portrait')
        pgSz.set(_qn('w:code'), '9')

    lines = markdown_text.strip().split('\n')
    table_lines = [
        l for l in lines
        if l.strip().startswith('|') and not re.match(r'\|[\s\-|]+\|', l.strip())
    ]

    if not table_lines:
        doc.add_paragraph(markdown_text)
        bio = BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio

    headers = [h.strip() for h in table_lines[0].split('|')[1:-1]]
    rows = [[c.strip() for c in line.split('|')[1:-1]] for line in table_lines[1:]]
    n_cols = len(headers)
    rows = [r[:n_cols] + [''] * max(0, n_cols - len(r)) for r in rows]

    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.style = 'Table Grid'
    table.allow_autofit = False

    # A4縦 本文幅 18cm を 4列に分配（日付 / 献立名 / おやつ / 結果）
    COL_WIDTHS = [Cm(2.0), Cm(5.5), Cm(3.5), Cm(7.0)]
    if n_cols != 4:
        each = Cm(18.0 / n_cols)
        COL_WIDTHS = [each] * n_cols

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.width = COL_WIDTHS[i]
        para = cell.paragraphs[0]
        if para.runs:
            run = para.runs[0]
            run.bold = True
            run.font.size = Pt(9)

    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = val
            cell.width = COL_WIDTHS[c_idx]
            para = cell.paragraphs[0]
            if para.runs:
                run = para.runs[0]
                run.font.size = Pt(8)
                if c_idx == n_cols - 1 and val.startswith('●'):
                    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def create_colored_excel(uploaded_file, sheet_name):
    """
    Excelを読み込み、食材カテゴリに応じてセルに色を付けた .xlsx を返す。
    .xlsx: load_workbook で全フォーマット（罫線・フォント・列幅等）を保持して色を追加。
    .xls:  xlrd でデータ＋マージセルを取得し openpyxl で再構築してから色を追加。
    """
    is_xls = uploaded_file.name.lower().endswith(".xls")
    file_bytes = uploaded_file.read()

    # ─── 色定義 ──────────────────────────────────────────────
    def _fill(hex6):
        return PatternFill(fill_type='solid', fgColor=hex6)

    FILL_YELLOW = _fill('FFFF99')  # 黄色
    FILL_GREEN  = _fill('CCFFCC')  # 緑色
    FILL_ORANGE = _fill('FFD9AD')  # オレンジ色
    FILL_PURPLE = _fill('E6CCFF')  # 紫色
    FILL_CYAN   = _fill('CCFFFF')  # 水色
    FILL_PINK   = _fill('FFB3C6')  # ピンク色

    ING_COLOR_RULES = [
        (['コーン', '人参', '黄パプリカ', '赤パプリカ', 'かぼちゃ'], FILL_YELLOW),
        (['ほうれん草', '小松菜', 'チンゲン菜', 'グリンピース',
          'いんげん', 'えだまめ', 'ブロッコリー', 'ピーマン'], FILL_GREEN),
        (['木綿豆腐', '焼き豆腐', '油揚げ', '厚揚げ', '大豆'], FILL_ORANGE),
        (['チーズ'], FILL_PURPLE),
        (['ちくわ', 'かにかま', 'ツナ', '赤かまぼこ'], FILL_CYAN),
        (['ロースハム', 'ベーコン', 'ウインナー'], FILL_PINK),
    ]

    # ─── ワークブックのロード ─────────────────────────────────
    if not is_xls:
        # .xlsx: openpyxl で直接ロード → 全フォーマット保持
        wb = load_workbook(BytesIO(file_bytes), data_only=True)
        ws = wb[sheet_name]
        n_rows = ws.max_row
        n_cols = ws.max_column

        def cell_val(r, c):
            if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
                return ""
            v = ws.cell(row=r + 1, column=c + 1).value
            if v is None:
                return ""
            s = str(v).strip()
            return "" if s in ("nan", "", "None") else s

        def apply_fill(r, c, fill):
            try:
                ws.cell(row=r + 1, column=c + 1).fill = fill
            except AttributeError:
                pass  # マージセルの非先頭セルは無視

    else:
        # .xls: xlrd でデータ＋マージセルを取得 → openpyxl で再構築
        import xlrd as _xlrd

        xls_wb = _xlrd.open_workbook(file_contents=file_bytes)
        xls_ws = xls_wb.sheet_by_name(sheet_name)
        n_rows = xls_ws.nrows
        n_cols = xls_ws.ncols

        def cell_val(r, c):
            if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
                return ""
            v = xls_ws.cell_value(r, c)
            if v is None:
                return ""
            s = str(v).strip()
            return "" if s in ("nan", "", "None") else s

        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name[:31]

        # 値をコピー
        for r_idx in range(n_rows):
            for c_idx in range(n_cols):
                v = cell_val(r_idx, c_idx)
                ws.cell(row=r_idx + 1, column=c_idx + 1, value=v or None)

        # マージセルをコピー（xlrd: row_hi/col_hi は exclusive）
        for row_lo, row_hi, col_lo, col_hi in xls_ws.merged_cells:
            try:
                ws.merge_cells(
                    start_row=row_lo + 1, start_column=col_lo + 1,
                    end_row=row_hi,       end_column=col_hi,
                )
            except Exception:
                pass

        # 列幅をコピー（xlrd 2.x では取得できない場合あり）
        try:
            for c_idx in range(n_cols):
                ci = xls_ws.colinfo_map.get(c_idx)
                if ci and ci.width > 0:
                    ws.column_dimensions[get_column_letter(c_idx + 1)].width = (
                        ci.width / 256
                    )
        except Exception:
            pass

        # 行高さをコピー（xlrd 2.x では取得できない場合あり）
        try:
            for r_idx in range(n_rows):
                ri = xls_ws.rowinfo_map.get(r_idx)
                if ri and ri.height > 0:
                    ws.row_dimensions[r_idx + 1].height = ri.height / 20
        except Exception:
            pass

        def apply_fill(r, c, fill):
            try:
                ws.cell(row=r + 1, column=c + 1).fill = fill
            except AttributeError:
                pass

    # ─── ブロック検出 ─────────────────────────────────────────
    blocks = []
    for r in range(n_rows):
        temp = {}
        for c in range(n_cols):
            v = cell_val(r, c)
            if re.match(r'^\d+日\([月火水木金土日]\)$', v):
                temp[c] = v
        if len(temp) >= 3:
            blocks.append((r, temp))

    # ─── 色付きセルの収集（材料欄のみ） ─────────────────────
    ing_cells = {}

    for block_idx, (block_date_row, block_date_cols) in enumerate(blocks):
        block_end = (
            blocks[block_idx + 1][0] if block_idx + 1 < len(blocks) else n_rows
        )

        block_mat_row = None
        for r in range(block_date_row + 1, block_end):
            for c in range(min(5, n_cols)):
                if cell_val(r, c) == "材料":
                    block_mat_row = r
                    break
            if block_mat_row is not None:
                break

        if block_mat_row is not None:
            for col_c in sorted(block_date_cols.keys()):
                for r in range(block_mat_row, block_end):
                    v = cell_val(r, col_c)
                    if not v:
                        continue
                    for kw_list, fill_color in ING_COLOR_RULES:
                        if any(kw in v for kw in kw_list):
                            ing_cells[(r, col_c)] = fill_color
                            break

    # ─── 色を適用 ────────────────────────────────────────────
    for (r_idx, c_idx), fc in ing_cells.items():
        apply_fill(r_idx, c_idx, fc)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def clean_result_column(text):
    """AI出力の結果列を後処理：OK説明文を強制除去してOK/●NGのみにする"""
    # これらのいずれかを含む●セルは「実NGなし」とみなしてOKに変換
    no_real_ng_patterns = [
        r'OK\s*$',            # 末尾がOK（→OK、/ OKなど含む）
        r'実NG[無な]し',      # 実NG無し / 実NGなし
        r'NG[無な]し',        # NGなし / NG無し
        r'OK扱い',            # OK扱い不可 / OK扱い等
        r'問題[無な]し',      # 問題なし / 問題無し
        r'許容$',             # 末尾が「許容」
        r'対象外\s*$',        # 末尾が「対象外」
        r'ではないが',        # 「月曜ではないがロールパン使用」等の非該当説明
        r'重複[無な]し',      # 重複なし / 重複無し（同日重複チェックでOKの場合）
    ]

    lines = text.split('\n')
    cleaned = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('|') and not re.match(r'\|[\s\-|]+\|', stripped):
            parts = stripped.split('|')
            if len(parts) >= 3:
                cell = parts[-2].strip()
                # 「●」で始まらない表現は「OK」に変換
                if cell and not cell.startswith('●'):
                    cell = 'OK'
                elif cell:
                    # 「●」で始まるが実NGなしを示すパターンが含まれる → OK
                    for pat in no_real_ng_patterns:
                        if re.search(pat, cell):
                            cell = 'OK'
                            break
                parts[-2] = f' {cell} '
                line = '|'.join(parts)
        cleaned.append(line)
    result = '\n'.join(cleaned)
    result = re.sub(r'<br\s*/?>', ' ／ ', result, flags=re.IGNORECASE)
    return result


def compute_all_python_ngs(excel_text):
    """
    全ルールベースNGをPythonで計算。
    Returns: (summary_text, day_ngs_dict)
      summary_text  … AIプロンプトに埋め込むNG一覧テキスト
      day_ngs_dict  … {date_str: [ng_str, ...]}
    """
    year, month_num, sorted_dates, entries = _parse_structured(excel_text)
    if not sorted_dates:
        return "", {}

    day_ngs = {ds: [] for ds in sorted_dates}

    def ing(ds):
        return entries[ds]['ingredients']

    def lunch(ds):
        return entries[ds]['lunch']

    def snack(ds):
        return entries[ds]['snack']

    def has(text, kw_list):
        return any(kw in text for kw in kw_list)

    # ── 禁止食材 ────────────────────────────────────────────────
    for ds in sorted_dates:
        for f in ['卵', 'マヨネーズ', '絹ごし豆腐']:
            if f in ing(ds):
                day_ngs[ds].append(f'● {f}（使用禁止食材）')

    # ── チーズ2日連続 ──────────────────────────────────────────
    for i in range(1, len(sorted_dates)):
        if 'チーズ' in ing(sorted_dates[i]) and 'チーズ' in ing(sorted_dates[i-1]):
            day_ngs[sorted_dates[i]].append('● チーズ類2日連続')

    # ── 肉類3日連続 ────────────────────────────────────────────
    for meat in MEAT_3DAY:
        for i in range(2, len(sorted_dates)):
            if all(meat in ing(sorted_dates[j]) for j in (i, i-1, i-2)):
                day_ngs[sorted_dates[i]].append(f'● {meat}3日連続')

    # ── 調味料2日連続（日祝を挟めばOK） ──────────────────────
    for seasoning in SEASONING_2DAY:
        for i in range(1, len(sorted_dates)):
            ds_c, ds_p = sorted_dates[i], sorted_dates[i-1]
            if seasoning in ing(ds_c) and seasoning in ing(ds_p):
                d_c, d_p = _parse_date(ds_c, year), _parse_date(ds_p, year)
                exempted = False
                if d_c and d_p:
                    d = d_p + datetime.timedelta(days=1)
                    while d < d_c:
                        if d.weekday() == 6 or _is_holiday(d):
                            exempted = True
                            break
                        d += datetime.timedelta(days=1)
                if not exempted:
                    day_ngs[ds_c].append(f'● {seasoning}2日連続')

    # ── 芋類3日連続 ────────────────────────────────────────────
    for i in range(2, len(sorted_dates)):
        if all(has(ing(sorted_dates[j]), IMO_KW) for j in (i, i-1, i-2)):
            day_ngs[sorted_dates[i]].append('● 芋類3日連続')

    # ── 汎用食材2日連続（免除リスト以外） ────────────────────
    for i in range(1, len(sorted_dates)):
        ds_c, ds_p = sorted_dates[i], sorted_dates[i-1]
        toks_c = {t for t in _split_ing(ing(ds_c)) if not _is_exempt(t) and len(t) >= 2}
        toks_p = {t for t in _split_ing(ing(ds_p)) if not _is_exempt(t) and len(t) >= 2}
        for token in sorted(toks_c & toks_p):
            day_ngs[ds_c].append(f'● {token}2日連続')

    # ── 同日チェック ───────────────────────────────────────────
    for ds in sorted_dates:
        i_text = ing(ds)
        toks = _split_ing(i_text)

        # 芋類2種類以上（同日）
        imo_cnt = sum(1 for kw in IMO_KW if kw in i_text)
        if imo_cnt >= 2:
            day_ngs[ds].append(f'● 同日芋類{imo_cnt}種類（1種類のみ可）')

        # 酢2回以上
        if toks.count('酢') >= 2:
            day_ngs[ds].append('● 同日「酢」2回使用')

        # みそ2回以上
        if toks.count('みそ') >= 2:
            day_ngs[ds].append('● 同日「みそ」2回使用')

        # ほうれん草＋小松菜
        if 'ほうれん草' in i_text and '小松菜' in i_text:
            day_ngs[ds].append('● 同日「ほうれん草」と「小松菜」重複')

        # 練り物2種類以上
        neri_cnt = sum(1 for kw in NERIMONO_KW if kw in i_text)
        if neri_cnt >= 2:
            day_ngs[ds].append(f'● 同日練り物{neri_cnt}種類（1種類のみ可）')

        # 玉ねぎ3回以上
        tama_cnt = toks.count('玉ねぎ')
        if tama_cnt >= 3:
            day_ngs[ds].append(f'● 同日「玉ねぎ」{tama_cnt}回使用')

        # 人参3回以上
        ninj_cnt = toks.count('人参')
        if ninj_cnt >= 3:
            day_ngs[ds].append(f'● 同日「人参」{ninj_cnt}回使用')

    # ── 月上限（4回目に到達した日に記録） ─────────────────────
    for item in ['バター', 'チーズ', 'マヨドレ']:
        found = [ds for ds in sorted_dates if item in ing(ds)]
        if len(found) >= 4:
            day_ngs[found[3]].append(f'● {item}月4回目（上限超過）')

    # ── 月曜・祝日翌日チェック ────────────────────────────────
    for ds in sorted_dates:
        dow_i = _dow(ds)
        d = _parse_date(ds, year)
        is_mon = (dow_i == 0)
        prev_holiday = _is_holiday(d - datetime.timedelta(days=1)) if d else False
        is_mon_or_post = is_mon or prev_holiday
        is_thu = (dow_i == 3)
        i_text = ing(ds)
        ls_text = lunch(ds) + ' ' + snack(ds)

        if is_mon_or_post:
            for item in MON_NG_ITEMS:
                if item in i_text:
                    day_ngs[ds].append(f'● 月曜/祝日翌日「{item}」NG')
            for m_kw in MUSHROOM_KW:
                if m_kw in i_text:
                    day_ngs[ds].append(f'● 月曜/祝日翌日「{m_kw}」NG')
                    break
            for fish_kw in FISH_KW:   # ツナはOKなので FISH_KW（ツナなし）を使う
                if fish_kw in i_text:
                    day_ngs[ds].append(f'● 月曜/祝日翌日「{fish_kw}」使用NG')
                    break
            # 月曜/祝日翌日に照り焼き・から揚げ・フライ
            for dish in ['照り焼き', 'から揚げ', 'フライ']:
                if dish in ls_text:
                    check = ls_text.replace('フライドポテト', '') if dish == 'フライ' else ls_text
                    if dish in check:
                        day_ngs[ds].append(f'● 月曜/祝日翌日「{dish}」NG')

        # にら・青ねぎ：月曜・祝日翌日・木曜がNG
        if is_mon_or_post or is_thu:
            for nk in ['にら', '青ねぎ']:
                if nk in i_text:
                    day_ngs[ds].append(f'● 月曜/祝日翌日/木曜「{nk}」NG')

    # ── 週次チェック（魚・豆腐・麺丼） ──────────────────────
    _dow_map = {'月': 0, '火': 1, '水': 2, '木': 3, '金': 4, '土': 5, '日': 6}

    def _week_key(ds):
        m2 = re.search(r'/(\d+)\(([月火水木金土日])\)', ds)
        return int(m2.group(1)) - _dow_map.get(m2.group(2), 0) if m2 else -1

    weeks = {}
    for ds in sorted_dates:
        weeks.setdefault(_week_key(ds), []).append(ds)

    for _, week_dates in sorted(weeks.items()):
        wd = sorted(week_dates, key=lambda d: _parse_date(d, year) or datetime.date.max)
        last = wd[-1]
        combined = ' '.join(lunch(ds) + ' ' + ing(ds) for ds in wd)

        if not has(combined, FISH_WITH_TUNA):
            day_ngs[last].append('● 今週魚なし（週1回以上必要）')

        tofu_ok = any(
            has(ing(ds), TOFU_KW) or has(lunch(ds), TOFU_KW)
            for ds in wd
        )
        if not tofu_ok:
            day_ngs[last].append('● 今週豆腐なし（週1回以上必要）')

        if not has(combined, NOODLE_KW):
            day_ngs[last].append('● 今週麺・丼なし（週1回以上必要）')

    # ── 仕込み食材7種類以上 ───────────────────────────────────
    for ds in sorted_dates:
        i_text = ing(ds)
        found = [kw for kw in PREP_KW if kw in i_text]
        if len(found) >= 7:
            day_ngs[ds].append(f'● 仕込み食材{len(found)}種類（7種類以上）: {", ".join(found)}')

    # ── 月の献立回数上限 ──────────────────────────────────────
    ingenge_days = [ds for ds in sorted_dates if 'いんげん' in (lunch(ds) + snack(ds))]
    fruit_days   = [ds for ds in sorted_dates if '果物' in (lunch(ds) + snack(ds))]
    if len(ingenge_days) >= 3:
        day_ngs[ingenge_days[2]].append('● 「いんげん」献立が月3回目（上限超過）')
    if len(fruit_days) >= 2:
        day_ngs[fruit_days[1]].append('● 「果物」献立が月2回目（上限超過）')

    # ── 土曜日に麺・丼なし ───────────────────────────────────
    for ds in sorted_dates:
        if _dow(ds) == 5 and not has(lunch(ds), NOODLE_KW):
            day_ngs[ds].append('● 土曜日に麺・丼なし')

    # ── ハンバーグ/フライ前日おやつゼリーなし ────────────────
    for i in range(1, len(sorted_dates)):
        ds_today, ds_next = sorted_dates[i-1], sorted_dates[i]
        l_next = lunch(ds_next)
        need_jelly = 'ハンバーグ' in l_next
        if not need_jelly and 'フライ' in l_next:
            if 'フライ' in l_next.replace('フライドポテト', ''):
                need_jelly = True
        if need_jelly and 'ゼリー' not in snack(ds_today):
            day_ngs[ds_today].append(
                f'● 翌日({ds_next})ハンバーグ/フライだが前日おやつにゼリーなし'
            )

    # ── 季節チェック（クリームシチュー） ─────────────────────
    if 5 <= month_num <= 9:
        for ds in sorted_dates:
            if 'クリームシチュー' in lunch(ds):
                day_ngs[ds].append('● 5〜9月クリームシチューNG')

    # ── 料理名から推定される必須材料チェック ─────────────────
    # (料理名に含む文字列, 材料欄に必要なキーワードのいずれか)
    RECIPE_RULES = [
        ('おかか',   ['かつお節', 'おかか']),
        ('ごま和え', ['白ごま', '黒ごま', 'すりごま', 'ごま']),
        ('あんかけ', ['片栗粉']),
        ('から揚げ', ['片栗粉']),
        ('唐揚げ',   ['片栗粉']),
        ('照り焼き', ['みりん', '醤油']),
        ('蒸しパン', ['ベーキングパウダー', 'BP', '重曹']),
    ]
    for ds in sorted_dates:
        ls_text = lunch(ds) + ' ' + snack(ds)
        i_text  = ing(ds)
        for kw, required_any in RECIPE_RULES:
            if kw in ls_text:
                if not any(req in i_text for req in required_any):
                    missing = '・'.join(required_any)
                    day_ngs[ds].append(f'● 「{kw}」があるが材料に{missing}なし')

    # ── 出力テキスト生成 ──────────────────────────────────────
    lines = [
        '【Python確定NGリスト】',
        '（AIはこの結果をそのまま採用。自分で再計算・再判断しないこと）',
        '',
    ]
    has_any = False
    for ds in sorted_dates:
        for ng in day_ngs[ds]:
            lines.append(f'{ds}: {ng}')
            has_any = True
    if not has_any:
        lines.append('（ルールベースのNG検出なし）')

    return '\n'.join(lines), day_ngs


def run_check(excel_text, rules_text, api_key, file_name, sheet_name):
    client = anthropic.Anthropic(api_key=api_key)
    python_ng_text, _ = compute_all_python_ngs(excel_text)

    prompt = f"""あなたは幼稚園給食の献立チェック専門家です。
担当は「ステップ４：献立名と材料の照合チェック」のみです。
その他の全ルール（連続使用・禁止食材・週次・月上限など）はPythonが計算済みです。

========================================
# エクセルデータ（ファイル：{file_name}　シート：{sheet_name}）
========================================
【月/日(曜日)】
昼食: 献立名1 / 献立名2 / ...
おやつ: おやつ名1 / ...
材料: 材料名1, 材料名2, ...

{excel_text}

========================================
# Python確定NGリスト（必ず結果欄に転記すること）
========================================
以下は全ルールをPythonで計算済みの確定NGです。
AIは再計算せず、そのまま結果欄に記入してください。

{python_ng_text}

========================================
# あなたが担当するチェック（ステップ４のみ）
========================================
【献立名と材料の照合チェック】
同じ日の【献立名】と【材料】を照らし合わせ、
献立名に入っている食材が材料欄にない場合はNG。

例外（NG対象外）：
・「ごはん」「おにぎり」→ 白米があればOK
・「チキン」「鶏」→ 鶏肉があればOK
・「すまし汁」「スープ」→ 食材なくてOK
・「ゼリー」→ 食材なくてOK
・「ごぼう」→ ささがきごぼうがあればOK

========================================
# 出力形式（絶対厳守）
========================================
マークダウンテーブルのみ。前後に文章・注釈禁止。

| 日付 | 献立名 | おやつ | 結果 |
|------|--------|--------|------|

【結果欄のルール】
・Python確定NGがある日 → そのNGをそのまま記入（複数なら改行で並べる）
・ステップ４でNGが見つかった日 → 「● ○○が材料欄にない」を追記
・どちらもなければ → 「OK」の2文字のみ

【絶対禁止】
× Python確定NGにないルールベースNGを自分で書く
× 「OK（確認済み）」「● ○○なし → OK」などOK説明文
× テーブル以外の文章

【その他】
・献立名はスラッシュ「/」でつなぐ
・おやつがない日は空欄
・給食のない日は出力しない
"""

    message = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )
    return clean_result_column(message.content[0].text)


with st.sidebar:
    st.title("🍱 献立チェックシステム")
    st.markdown("---")
    page = st.radio(
        "ページを選択",
        ["📋 献立チェック", "⚙️ ルール管理"],
        label_visibility="collapsed",
    )


if page == "📋 献立チェック":
    st.title("📋 献立チェック")

    api_key = get_api_key()
    if not api_key:
        st.error("⚠️ APIキーが設定されていません。管理者にお問い合わせください。")

    st.markdown("### 1. ファイルをアップロード")
    uploaded = st.file_uploader(
        "献立Excelファイル（.xls または .xlsx）",
        type=["xls", "xlsx"],
    )

    if uploaded:
        sheets = get_sheet_names(uploaded)
        if sheets:
            st.markdown("### 2. シートを選択")
            selected_sheet = st.selectbox("シート", sheets)

            st.markdown("### 3. 色付きExcel生成（目検用）")
            if st.button("🎨 色付きExcel生成"):
                with st.spinner("色付き処理中..."):
                    uploaded.seek(0)
                    colored = create_colored_excel(uploaded, selected_sheet)
                    st.session_state["colored_excel"] = colored.getvalue()
                    st.session_state["colored_fname"] = uploaded.name.rsplit(".", 1)[0]

            if st.session_state.get("colored_excel"):
                fname_c = st.session_state.get("colored_fname", "result")
                st.download_button(
                    "📥 色付きExcelをダウンロード",
                    data=st.session_state["colored_excel"],
                    file_name=f"色付き_{fname_c}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_colored",
                )

            st.markdown("### 4. チェック開始")
            if st.button("✅ チェックを開始する", type="primary", disabled=not api_key):
                rules = load_rules()
                if not rules.strip():
                    st.error("ルールファイルが空です。「ルール管理」ページでルールを設定してください。")
                else:
                    with st.spinner("Excelデータを読み込んでいます..."):
                        uploaded.seek(0)
                        excel_text = excel_to_text(uploaded, selected_sheet)

                    with st.spinner("チェック中です。1〜2分ほどお待ちください..."):
                        try:
                            uploaded.seek(0)
                            result = run_check(
                                excel_text, rules, api_key, uploaded.name, selected_sheet
                            )
                            st.session_state["last_result"] = result
                            st.session_state["last_filename"] = uploaded.name
                        except Exception as e:
                            st.error(f"エラーが発生しました: {e}")
                            result = None

            if st.session_state.get("last_result"):
                st.markdown("---")
                st.subheader("チェック結果")
                result_text = st.session_state["last_result"]
                st.markdown(result_text)
                fname = st.session_state.get("last_filename", "result").rsplit(".", 1)[0]
                docx_data = table_to_docx(result_text)
                st.download_button(
                    "📥 結果をWordファイルで保存",
                    data=docx_data,
                    file_name=f"チェック結果_{fname}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_word",
                )


elif page == "⚙️ ルール管理":
    st.title("⚙️ ルール管理")
    st.caption("ルールを編集して「保存」を押すと、アプリとGitHubの両方に反映されます。")

    current = load_rules()
    edited = st.text_area("チェックルール", value=current, height=600)

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("💾 保存", type="primary"):
            save_rules(edited)
            success, msg = push_rules_to_github(edited)
            if success:
                st.success(msg)
            else:
                st.warning(f"アプリには保存済みです。GitHub更新に失敗しました：{msg}")
    with col2:
        if st.button("🔄 元に戻す（最後の保存時点）"):
            st.rerun()
