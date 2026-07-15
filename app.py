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
from openpyxl.styles import PatternFill, Border, Side
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
RULES_JSON_FILE = BASE_DIR / "rules.json"

GITHUB_OWNER = "fukiyamiyazaki-blip"
GITHUB_REPO = "-"
GITHUB_BRANCH = "main"
GITHUB_RULES_PATH = "rules.txt"
GITHUB_RULES_JSON_PATH = "rules.json"

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


def load_rules_list():
    """複数ルールをリストで返す。GitHubを優先し、失敗時のみローカル→rules.txtにフォールバック。"""
    token = get_github_token()
    if token:
        api_url = (
            f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
            f"/contents/{GITHUB_RULES_JSON_PATH}?ref={GITHUB_BRANCH}"
        )
        headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
        }
        req = urllib.request.Request(api_url, headers=headers)
        try:
            with urllib.request.urlopen(req) as resp:
                data = json.loads(resp.read())
                content = base64.b64decode(data["content"]).decode("utf-8")
                return json.loads(content).get("rules", [])
        except Exception:
            pass  # GitHub取得失敗 → ローカルにフォールバック

    # ローカルファイルにフォールバック
    if RULES_JSON_FILE.exists():
        try:
            data = json.loads(RULES_JSON_FILE.read_text(encoding="utf-8"))
            return data.get("rules", [])
        except Exception:
            pass
    # 後方互換: rules.txt があれば移行して初期ルールとして返す
    legacy = load_rules()
    if legacy.strip():
        return [{"id": "default", "name": "共通ルール", "text": legacy}]
    return []


def save_rules_list(rules_list):
    """複数ルールをJSON保存（ローカル）。失敗してもエラーを出さない。"""
    try:
        RULES_JSON_FILE.write_text(
            json.dumps({"rules": rules_list}, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
    except Exception:
        pass  # Streamlit Cloud では書き込み不可の場合もあるが GitHub が主ストレージ


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
    _num = re.compile(r'^\d+\.?\d*$')  # 数値のみのトークン（数量・シリアル値）を除外
    for part in parts:
        # 半角・全角カッコも区切り文字として扱う
        for token in re.split(r'[\s　・／/()（）]+', part.strip()):
            t = token.strip()
            if t and t not in ('nan', '') and not _num.match(t):
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
    '昆布出し',  # 毎日使用OK（漢字表記のため'だし'部分一致に引っかからないため明示）
}
# 2日連続チェック免除（部分一致：これを含むトークンは免除）
_EXEMPT_SUB = ['醤油', '砂糖', 'みりん', '酒', '塩', '油',
               'だし', '出し', 'ごま', '水', '片栗粉', '小麦粉']
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
FRUIT_KW       = [
    '果物', 'みかん', 'りんご', 'バナナ', 'オレンジ', 'パイン', 'パイナップル',
    'もも', '桃', 'ぶどう', 'ブドウ', 'いちご', 'イチゴ', 'メロン', 'すいか',
    'スイカ', 'キウイ', 'なし', '梨', 'マンゴー', 'さくらんぼ',
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


def push_rules_list_to_github(rules_list):
    """複数ルール（rules.json）をGitHubに保存。"""
    token = get_github_token()
    if not token:
        return False, "GitHubトークンが未設定です"

    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
        f"/contents/{GITHUB_RULES_JSON_PATH}"
    )
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    }

    sha = None
    req = urllib.request.Request(
        api_url + f"?ref={GITHUB_BRANCH}", headers=headers
    )
    try:
        with urllib.request.urlopen(req) as resp:
            sha = json.loads(resp.read())["sha"]
    except Exception:
        pass  # 新規ファイルの場合はSHAなし

    content_text = json.dumps({"rules": rules_list}, ensure_ascii=False, indent=2)
    body = {"message": "ルール管理から更新", "branch": GITHUB_BRANCH,
            "content": base64.b64encode(content_text.encode("utf-8")).decode("utf-8")}
    if sha:
        body["sha"] = sha

    payload = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(api_url, data=payload, headers=headers, method="PUT")
    try:
        with urllib.request.urlopen(req):
            return True, "保存しました（GitHub反映済み）"
    except Exception as e:
        return False, f"GitHub更新エラー: {e}"


# ─────────────────────────────────────────────
# 多形式Excelパーサー（さかえ保育園・おおみや・ゆめのはな対応）
# ─────────────────────────────────────────────

def _detect_sheet_format(df):
    """シートのフォーマット種別を返す: 'sakae' / 'omiya' / 'mebaenomori' / 'yumehana' / 'yamazaki' / 'default'"""
    all_text = ' '.join(str(v) for v in df.values.flatten() if pd.notna(v))
    if '熱と力になるもの' in all_text:
        return 'sakae'
    if '◎は10時おやつ' in all_text or ('材料名' in all_text and '献立名' in all_text):
        return 'omiya'
    if 'つかみ食べ練習用野菜' in all_text:  # 美山保育園形式（月別シート・datetime日付・おやつcol10）
        return 'miyama'
    if 'うどんの日以外はおかゆ' in all_text:  # めばえの森（yumehanaより先に判定）
        return 'mebaenomori'
    if '初期には入りません' in all_text or 'おかゆが付きます' in all_text:
        return 'yumehana'
    if '初期・アレルギーには' in all_text:  # 歩学園バンビ形式（離乳食・datetime日付・おやつcol8-9）
        return 'ayumi'
    # 山崎幼稚園形式（横並び・3列/日）：セル単体が「材料表」と完全一致する場合のみ判定。
    # 部分一致にすると「献立材料表」のようなタイトル文言を含む他園（例：東久留米おひさま）を
    # 誤って山崎形式と判定してしまう（列オフセットが異なるため誤爆する）。
    cell_values = {str(v).strip() for v in df.values.flatten() if pd.notna(v)}
    if '材料表' in cell_values:
        return 'yamazaki'
    return 'default'


def _extract_year_month(df):
    """先頭10×10セルから年月を抽出。(year_int, month_int, label_str)"""
    n_rows, n_cols = df.shape
    for r in range(min(10, n_rows)):
        for c in range(min(10, n_cols)):
            v = str(df.iloc[r, c]).strip()
            m = re.match(r'(\d{4})年(\d{1,2})月', v)
            if m:
                y, mo = int(m.group(1)), int(m.group(2))
                return y, mo, f"{y}年{mo:02d}月"
    return 0, 0, ""


def _excel_to_text_sakae(df):
    """さかえ保育園形式（縦並び・4列材料）→ 構造化テキスト"""
    n_rows, n_cols = df.shape

    def cv(r, c):
        if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
            return ""
        v = str(df.iloc[r, c]).strip()
        return "" if v in ("nan", "", "None") else v

    year_num, month_num, year_month = _extract_year_month(df)

    days = []
    cur = None

    for r in range(n_rows):
        col0, col1, col2 = cv(r, 0), cv(r, 1), cv(r, 2)

        if re.match(r'^\d{1,2}(\.0)?$', col0) and col1 == '昼食':
            if cur is not None:
                days.append(cur)
            cur = {'day': int(float(col0)), 'dow': '?', 'lunch': [], 'snack': [], 'mats': [], 'in_snack': False}
            if col2:
                cur['lunch'].append(col2)
            for c in range(3, min(7, n_cols)):  # 列3-6が材料、列7はエネルギー値
                v = cv(r, c)
                if v:
                    cur['mats'].append(v)

        elif re.match(r'^[月火水木金土日]$', col0) and cur is not None:
            cur['dow'] = col0
            # 曜日行に献立名・材料が同居する場合（「土」行に鶏肉となすのみそ炒め等）
            if col2 and col1 == '':
                if cur['in_snack']:
                    cur['snack'].append(col2)
                else:
                    cur['lunch'].append(col2)
                for c in range(3, min(7, n_cols)):
                    v = cv(r, c)
                    if v:
                        cur['mats'].append(v)

        elif col1 == '午後おやつ' and cur is not None:
            cur['in_snack'] = True
            if col2:
                cur['snack'].append(col2)
            for c in range(3, min(7, n_cols)):
                v = cv(r, c)
                if v:
                    cur['mats'].append(v)

        elif cur is not None and col2 and col0 == '' and col1 == '':
            if cur['in_snack']:
                cur['snack'].append(col2)
            else:
                cur['lunch'].append(col2)
            for c in range(3, min(7, n_cols)):
                v = cv(r, c)
                if v:
                    cur['mats'].append(v)

    if cur is not None:
        days.append(cur)

    lines = []
    if year_month:
        lines += [f"# 献立データ {year_month}", ""]

    for d in days:
        label = f"{month_num}/{d['day']}({d['dow']})" if month_num else f"?/{d['day']}({d['dow']})"
        lines.append(f"【{label}】")
        if d['lunch']:
            lines.append(f"昼食: {' / '.join(d['lunch'])}")
        if d['snack']:
            lines.append(f"おやつ: {' / '.join(d['snack'])}")
        if d['mats']:
            lines.append(f"材料: {', '.join(d['mats'])}")
        lines.append("")

    return '\n'.join(lines)


def _excel_to_text_omiya(df):
    """おおみやこども園形式（縦並び・1セル全材料）→ 構造化テキスト"""
    n_rows, n_cols = df.shape

    def cv(r, c):
        if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
            return ""
        v = str(df.iloc[r, c]).strip()
        return "" if v in ("nan", "", "None") else v

    year_num, month_num, year_month = _extract_year_month(df)

    def clean_mat(text):
        text = re.sub(r'\(\d+g\)', '', text)         # 量(30g)を除去
        text = text.replace('*', '').replace('＊', '') # *印を除去
        text = re.sub(r'[／\n]', ',', text)           # ／と改行をカンマに
        return text

    def clean_oyatsu(text):
        text = text.replace('◎', '').replace('＊', '')
        text = re.sub(r'\(\d+g\)', '', text)
        return text.replace('\n', ' / ').strip()

    days = []

    for r in range(n_rows):
        col0 = cv(r, 0)
        col1 = cv(r, 1)
        col2 = cv(r, 2)

        if re.match(r'^\d{1,2}(\.0)?$', col0):
            day_int = int(float(col0))
            lunch_names = [x.strip() for x in re.split(r'[,\n]', col1) if x.strip()] if col1 else []
            mats_raw = clean_mat(col2) if col2 else ""
            mat_list = [x.strip() for x in re.split(r'[,、]', mats_raw) if x.strip()]

            oyatsu_names = []
            col5 = cv(r, 5) if n_cols > 5 else ""
            if col5:
                oyatsu_text = clean_oyatsu(col5)
                oyatsu_names = [x.strip() for x in re.split(r'[\s/／,、\n]+', oyatsu_text) if x.strip()]

            days.append({'day': day_int, 'dow': '?', 'lunch': lunch_names, 'snack': oyatsu_names, 'mats': mat_list})

        elif re.match(r'^[月火水木金土日]$', col0) and days:
            days[-1]['dow'] = col0

    lines = []
    if year_month:
        lines += [f"# 献立データ {year_month}", ""]

    for d in days:
        label = f"{month_num}/{d['day']}({d['dow']})" if month_num else f"?/{d['day']}({d['dow']})"
        lines.append(f"【{label}】")
        if d['lunch']:
            lines.append(f"昼食: {' / '.join(d['lunch'])}")
        if d['snack']:
            lines.append(f"おやつ: {' / '.join(d['snack'])}")
        if d['mats']:
            lines.append(f"材料: {', '.join(d['mats'])}")
        lines.append("")

    return '\n'.join(lines)


def _excel_to_text_yumehana(df):
    """ゆめのはなこども園形式（週別シート・Excelシリアル日付）→ 構造化テキスト"""
    n_rows, n_cols = df.shape

    def cv(r, c):
        if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
            return ""
        v = str(df.iloc[r, c]).strip()
        return "" if v in ("nan", "", "None") else v

    def clean_mat(v):
        return v.replace('＊', '').replace('*', '').strip()

    year_num, month_num, year_month = _extract_year_month(df)

    days = {}   # label → {'lunch', 'snack', 'mats'}
    day_order = []

    for r in range(n_rows):
        col0, col1, col2 = cv(r, 0), cv(r, 1), cv(r, 2)

        if re.match(r'^\d{5}(\.0)?$', col0):  # Excelシリアル番号（5桁）
            serial = int(float(col0))
            dow = col1 if re.match(r'^[月火水木金土日]$', col1) else '?'
            try:
                d = datetime.date(1899, 12, 30) + datetime.timedelta(days=serial)
                label = f"{d.month}/{d.day}({dow})"
                if not year_num:
                    year_num, month_num = d.year, d.month
            except Exception:
                label = f"?/?({dow})"

            if label not in days:
                days[label] = {'lunch': [], 'snack': [], 'mats': []}
                day_order.append(label)

            if col2 and '印は初期' not in col2 and 'おかゆが付きます' not in col2:
                days[label]['lunch'].append(col2)
            for c in range(3, n_cols):
                v = cv(r, c)
                if v:
                    days[label]['mats'].append(clean_mat(v))

        elif col0 == '' and col2 and day_order:
            if '印は初期' not in col2 and 'おかゆが付きます' not in col2:
                label = day_order[-1]
                days[label]['lunch'].append(col2)
                for c in range(3, n_cols):
                    v = cv(r, c)
                    if v:
                        days[label]['mats'].append(clean_mat(v))

    lines = []
    if year_num and month_num:
        lines += [f"# 献立データ {year_num}年{month_num:02d}月", ""]

    for label in day_order:
        d = days[label]
        lines.append(f"【{label}】")
        if d['lunch']:
            lines.append(f"昼食: {' / '.join(d['lunch'])}")
        if d['snack']:
            lines.append(f"おやつ: {' / '.join(d['snack'])}")
        if d['mats']:
            lines.append(f"材料: {', '.join(d['mats'])}")
        lines.append("")

    return '\n'.join(lines)


def _excel_to_text_mebaenomori(df):
    """めばえの森保育園形式（週別シート・datetime文字列日付・離乳食・2行/日）→ 構造化テキスト"""
    n_rows, n_cols = df.shape

    def cv(r, c):
        if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
            return ""
        v = str(df.iloc[r, c]).strip()
        return "" if v in ("nan", "", "None") else v

    def clean_mat(v):
        return v.replace('＊', '').replace('*', '').strip()

    _NOTE_PHRASES = ('初期には入りません', 'おかゆが付きます', '富喜屋', '株式会社')
    _HOLIDAY_KW   = ('山の日', '海の日', '振替休日', '祝日', '休園', '夏期休暇',
                     '冬期休暇', '春期休暇', 'こどもの日', '天皇誕生日')

    year_num, month_num = 0, 0

    days = {}     # label → {'lunch': [], 'mats': []}
    day_order = []

    for r in range(n_rows):
        col0 = cv(r, 0)
        col1 = cv(r, 1)
        col2 = cv(r, 2)

        # 日付行: pandasがdtype=strで読むと 'YYYY-MM-DD HH:MM:SS' 形式になる
        dm = re.match(r'^(\d{4})-(\d{2})-(\d{2})', col0)
        if dm:
            y, mo, day = int(dm.group(1)), int(dm.group(2)), int(dm.group(3))
            if not year_num:
                year_num, month_num = y, mo
            dow = col1 if re.match(r'^[月火水木金土日]$', col1) else '?'
            label = f"{mo}/{day}({dow})"

            if label not in days:
                days[label] = {'lunch': [], 'mats': []}
                day_order.append(label)

            # 祝日・休園は料理名として追加しない
            if col2 and not any(h in col2 for h in _HOLIDAY_KW):
                days[label]['lunch'].append(col2)
            for c in range(3, n_cols):
                v = clean_mat(cv(r, c))
                if v:
                    days[label]['mats'].append(v)

        # 2行目（col0・col1ともに空）: 同日の材料継続行または副料理行
        elif col0 == '' and col1 == '' and day_order:
            # 行全体のテキストに注記フレーズがあればスキップ
            row_text = ' '.join(cv(r, c) for c in range(n_cols))
            if any(p in row_text for p in _NOTE_PHRASES):
                continue
            mats_in_row = [clean_mat(cv(r, c)) for c in range(3, n_cols)
                           if clean_mat(cv(r, c))]
            if col2 or mats_in_row:
                label = day_order[-1]
                if col2:
                    days[label]['lunch'].append(col2)
                days[label]['mats'].extend(mats_in_row)

    lines = []
    if year_num and month_num:
        lines += [f"# 献立データ {year_num}年{month_num:02d}月", ""]

    for label in day_order:
        d = days[label]
        if not d['lunch'] and not d['mats']:
            continue  # 祝日・空日はスキップ
        lines.append(f"【{label}】")
        if d['lunch']:
            lines.append(f"昼食: {' / '.join(d['lunch'])}")
        if d['mats']:
            lines.append(f"材料: {', '.join(d['mats'])}")
        lines.append("")

    return '\n'.join(lines)


def _excel_to_text_yamazaki(df):
    """山崎幼稚園形式（横並び・2ブロック・3列/日・日付col_c/献立+材料col_c+1）→ 構造化テキスト"""
    n_rows, n_cols = df.shape

    def cv(r, c):
        if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
            return ""
        v = str(df.iloc[r, c]).strip()
        return "" if v in ("nan", "", "None") else v

    year_num, month_num, year_month = _extract_year_month(df)

    _SKIP = {"[昼]", "[午後]", "献立名", "材料", "材料表", "日付", "区分",
             "お箸もいります", "おはしもいります"}

    def is_valid(v):
        if not v or v in _SKIP:
            return False
        if v.startswith("※") or v.startswith("【"):
            return False
        try:
            float(v)   # 数量（g数等）は除外
            return False
        except ValueError:
            pass
        return True

    # 日付行を検索（「N日(曜)」パターンが3個以上ある行）
    blocks = []
    for r in range(n_rows):
        temp = {}
        for c in range(n_cols):
            v = cv(r, c)
            if re.match(r'^\d+日\([月火水木金土日]\)$', v):
                temp[c] = v
        if len(temp) >= 3:
            blocks.append((r, temp))

    lines = []
    if year_month:
        lines += [f"# 献立データ {year_month}", ""]

    for block_idx, (date_row, date_cols) in enumerate(blocks):
        block_end = blocks[block_idx + 1][0] if block_idx + 1 < len(blocks) else n_rows

        # 材料表開始行（col 1 が "材料表" または "材料"）
        mat_row = None
        for r in range(date_row + 1, block_end):
            if cv(r, 1) in ("材料表", "材料"):
                mat_row = r
                break

        dish_end = mat_row if mat_row is not None else block_end

        for col_c in sorted(date_cols.keys()):
            raw_date = date_cols[col_c]
            dm = re.match(r'(\d+)日\(([月火水木金土日])\)', raw_date)
            if dm and month_num:
                date_label = f"{month_num}/{dm.group(1)}({dm.group(2)})"
            else:
                date_label = raw_date

            # [午後]マーカーを col_c で検索（行は日によって異なる）
            afternoon_start = dish_end
            for r in range(date_row + 1, dish_end):
                if cv(r, col_c) == "[午後]":
                    afternoon_start = r
                    break

            # 昼食・おやつは col_c+1 から読む
            lunch, snack = [], []
            for r in range(date_row + 1, afternoon_start):
                v = cv(r, col_c + 1)
                if is_valid(v):
                    lunch.append(v)
            for r in range(afternoon_start, dish_end):
                v = cv(r, col_c + 1)
                if is_valid(v):
                    snack.append(v)

            # 材料は col_c+1 から読む（材料表セクション）
            mats = []
            if mat_row is not None:
                for r in range(mat_row, block_end):
                    v = cv(r, col_c + 1)
                    if is_valid(v):
                        mats.append(v)

            lines.append(f"【{date_label}】")
            if lunch:
                lines.append(f"昼食: {' / '.join(lunch)}")
            if snack:
                lines.append(f"おやつ: {' / '.join(snack)}")
            if mats:
                lines.append(f"材料: {', '.join(mats)}")
            lines.append("")

    return '\n'.join(lines)


def _excel_to_text_ayumi(df):
    """歩学園バンビ形式（週別シート・datetime日付・離乳食・おやつcol8-9）→ 構造化テキスト"""
    n_rows, n_cols = df.shape
    _DOW = '月火水木金土日'

    def cv(r, c):
        if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
            return ""
        v = str(df.iloc[r, c]).strip()
        return "" if v in ("nan", "", "None") else v

    def clean_mat(v):
        return v.replace('＊', '').replace('*', '').strip()

    _NOTE = ('初期・アレルギーには', '麺の日以外', '初期におやつ', 'お菓子は、',
             '完了期のアレルギー', '富喜屋', '株式会社')

    year_num, month_num = 0, 0
    days = {}
    day_order = []

    def _parse_snack(text):
        """'料理名\n（材料1、材料2）' → (name, [mats])"""
        if not text:
            return None, []
        text = re.sub(r'\s+', ' ', text.replace('\n', ' '))
        m = re.search(r'[（(]([^）)]+)[）)]', text)
        if m:
            name = text[:m.start()].strip()
            ings = [clean_mat(s.strip()) for s in re.split(r'[,、\s]+', m.group(1)) if s.strip()]
        else:
            name = text.strip()
            ings = []
        return name or None, ings

    for r in range(n_rows):
        col0 = cv(r, 0)
        col2 = cv(r, 2)

        # 日付行（pandasがdatetime→'YYYY-MM-DD HH:MM:SS'に変換）
        dm = re.match(r'^(\d{4})-(\d{2})-(\d{2})', col0)
        if dm:
            y, mo, day_i = int(dm.group(1)), int(dm.group(2)), int(dm.group(3))
            if not year_num:
                year_num, month_num = y, mo
            # 曜日を日付から計算（col1が空のケースを補完）
            try:
                dow = _DOW[datetime.date(y, mo, day_i).weekday()]
            except Exception:
                dow = '?'
            label = f"{mo}/{day_i}({dow})"
            if label not in days:
                days[label] = {'lunch': [], 'snack': [], 'mats': []}
                day_order.append(label)

            if col2 and not any(n in col2 for n in _NOTE):
                days[label]['lunch'].append(col2)
            for c in range(3, min(8, n_cols)):
                v = clean_mat(cv(r, c))
                if v and not any(n in v for n in _NOTE):
                    days[label]['mats'].append(v)

            # おやつ（col8: 中期～後期、col9: 完了期）
            for sc in (8, 9):
                sv = cv(r, sc)
                if not sv:
                    continue
                sname, sings = _parse_snack(sv)
                if sname and sname not in days[label]['snack']:
                    days[label]['snack'].append(sname)
                days[label]['mats'].extend(sings)

        # 継続行（col0が空）
        elif col0 == '' and day_order:
            row_text = ' '.join(cv(r, c) for c in range(n_cols))
            if any(n in row_text for n in _NOTE):
                continue
            label = day_order[-1]
            if col2:
                days[label]['lunch'].append(col2)
            for c in range(3, min(8, n_cols)):
                v = clean_mat(cv(r, c))
                if v and not any(n in v for n in _NOTE):
                    days[label]['mats'].append(v)

    lines = []
    if year_num and month_num:
        lines += [f"# 献立データ {year_num}年{month_num:02d}月", ""]

    for label in day_order:
        d = days[label]
        if not d['lunch'] and not d['mats']:
            continue
        lines.append(f"【{label}】")
        if d['lunch']:
            lines.append(f"昼食: {' / '.join(d['lunch'])}")
        if d['snack']:
            lines.append(f"おやつ: {' / '.join(d['snack'])}")
        if d['mats']:
            lines.append(f"材料: {', '.join(d['mats'])}")
        lines.append("")

    return '\n'.join(lines)


def _excel_to_text_miyama(df):
    """美山保育園形式（月別シート・datetime日付・離乳食・おやつcol10）→ 構造化テキスト"""
    n_rows, n_cols = df.shape
    _DOW = '月火水木金土日'

    def cv(r, c):
        if r < 0 or r >= n_rows or c < 0 or c >= n_cols:
            return ""
        v = str(df.iloc[r, c]).strip()
        return "" if v in ("nan", "", "None") else v

    def clean_mat(v):
        return v.replace('＊', '').replace('*', '').strip()

    _NOTE = ('*印は初期には入りません', '麺の日以外の主食', '初期におやつ',
             'お菓子は、', '富喜屋', '株式会社', 'ここに日、祝',
             'つかみ食べ練習用野菜', '給食提供が無い日')

    year_num, month_num = 0, 0
    days = {}
    day_order = []

    def _parse_snack(text):
        if not text:
            return None, []
        text = re.sub(r'\s+', ' ', text.replace('\n', ' '))
        m = re.search(r'[（(]([^）)]+)[）)]', text)
        if m:
            name = text[:m.start()].strip()
            ings = [clean_mat(s.strip()) for s in re.split(r'[,、\s・]+', m.group(1)) if s.strip()]
        else:
            name = text.strip()
            ings = []
        return name or None, ings

    for r in range(n_rows):
        col0 = cv(r, 0)
        col2 = cv(r, 2)

        dm = re.match(r'^(\d{4})-(\d{2})-(\d{2})', col0)
        if dm:
            y, mo, day_i = int(dm.group(1)), int(dm.group(2)), int(dm.group(3))
            if not year_num:
                year_num, month_num = y, mo
            if y != year_num or mo != month_num:
                continue  # 別月（祝日リスト等）はスキップ
            try:
                dow = _DOW[datetime.date(y, mo, day_i).weekday()]
            except Exception:
                dow = '?'
            label = f"{mo}/{day_i}({dow})"
            if label not in days:
                days[label] = {'lunch': [], 'snack': [], 'mats': []}
                day_order.append(label)

            if col2 and not any(n in col2 for n in _NOTE):
                days[label]['lunch'].append(col2)
            for c in range(3, min(10, n_cols)):
                v = clean_mat(cv(r, c))
                if v and not any(n in v for n in _NOTE):
                    days[label]['mats'].append(v)

            # おやつ（col10: 中期〜後期）
            sv = cv(r, 10)
            if sv and not any(n in sv for n in _NOTE):
                sname, sings = _parse_snack(sv)
                if sname and sname not in days[label]['snack']:
                    days[label]['snack'].append(sname)
                days[label]['mats'].extend(sings)

        elif col0 == '' and day_order:
            row_text = ' '.join(cv(r, c) for c in range(n_cols))
            if any(n in row_text for n in _NOTE):
                continue
            label = day_order[-1]
            if col2 and not any(n in col2 for n in _NOTE):
                days[label]['lunch'].append(col2)
            for c in range(3, min(10, n_cols)):
                v = clean_mat(cv(r, c))
                if v and not any(n in v for n in _NOTE):
                    days[label]['mats'].append(v)

    lines = []
    if year_num and month_num:
        lines += [f"# 献立データ {year_num}年{month_num:02d}月", ""]

    for label in day_order:
        d = days[label]
        if not d['lunch'] and not d['mats']:
            continue
        lines.append(f"【{label}】")
        if d['lunch']:
            lines.append(f"昼食: {' / '.join(d['lunch'])}")
        if d['snack']:
            lines.append(f"おやつ: {' / '.join(d['snack'])}")
        if d['mats']:
            lines.append(f"材料: {', '.join(d['mats'])}")
        lines.append("")

    return '\n'.join(lines)


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

    # フォーマット自動検出 → 専用パーサーに振り分け
    fmt = _detect_sheet_format(df)
    if fmt == 'sakae':
        return _excel_to_text_sakae(df)
    if fmt == 'omiya':
        return _excel_to_text_omiya(df)
    if fmt == 'mebaenomori':
        return _excel_to_text_mebaenomori(df)
    if fmt == 'yumehana':
        return _excel_to_text_yumehana(df)
    if fmt == 'yamazaki':
        return _excel_to_text_yamazaki(df)
    if fmt == 'ayumi':
        return _excel_to_text_ayumi(df)
    if fmt == 'miyama':
        return _excel_to_text_miyama(df)

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


def create_colored_excel(uploaded_file):
    """
    全シートを色付きにして .xlsx で返す。フォーマット自動検出により
    さかえ・おおみや・ゆめのはな・既存形式それぞれの材料列に色付け。
    .xls はデータ・マージセルを再構築（書式は一部失われる）。
    """
    is_xls = uploaded_file.name.lower().endswith(".xls")
    file_bytes = uploaded_file.read()

    # ─── 色定義 ──────────────────────────────────────────────
    def _fill(hex6):
        return PatternFill(fill_type='solid', fgColor=hex6)

    ING_COLOR_RULES = [
        (['コーン', '人参', '黄パプリカ', '赤パプリカ', 'かぼちゃ'], _fill('FFFF99')),
        (['ほうれん草', '小松菜', 'チンゲン菜', 'グリンピース',
          'いんげん', 'えだまめ', 'ブロッコリー', 'ピーマン'],      _fill('CCFFCC')),
        (['木綿豆腐', '焼き豆腐', '油揚げ', '厚揚げ', '大豆'],      _fill('FFD9AD')),
        (['チーズ'],                                                  _fill('E6CCFF')),
        (['ちくわ', 'かにかま', 'ツナ', '赤かまぼこ'],               _fill('CCFFFF')),
        (['ロースハム', 'ベーコン', 'ウインナー'],                    _fill('FFB3C6')),
    ]

    # ─── フォーマット別の色付けロジック ──────────────────────
    def _apply_colors(cv_fn, af_fn, n_rows, n_cols, fmt):
        """cv_fn(r,c)→str、af_fn(r,c,fill)→None を使って色付け"""
        if fmt == 'sakae':
            # 列3-6が材料
            for r in range(n_rows):
                for c in range(3, min(7, n_cols)):
                    v = cv_fn(r, c)
                    if not v:
                        continue
                    for kw_list, fc in ING_COLOR_RULES:
                        if any(kw in v for kw in kw_list):
                            af_fn(r, c, fc)
                            break
        elif fmt == 'omiya':
            # 列2が全材料（1セル）
            for r in range(n_rows):
                v = cv_fn(r, 2)
                if not v:
                    continue
                for kw_list, fc in ING_COLOR_RULES:
                    if any(kw in v for kw in kw_list):
                        af_fn(r, 2, fc)
                        break
        elif fmt in ('yumehana', 'mebaenomori'):
            # 列3-9が材料（ゆめのはな・めばえの森共通）
            for r in range(n_rows):
                for c in range(3, min(10, n_cols)):
                    v = cv_fn(r, c)
                    if not v:
                        continue
                    for kw_list, fc in ING_COLOR_RULES:
                        if any(kw in v for kw in kw_list):
                            af_fn(r, c, fc)
                            break
        elif fmt == 'yamazaki':
            # 「N日(曜)」のcol_c+1が材料列（日付ごとに3列ずつ）
            date_cols = set()
            for r in range(n_rows):
                for c in range(n_cols):
                    v = cv_fn(r, c)
                    if re.match(r'^\d+日\([月火水木金土日]\)$', v):
                        date_cols.add(c)
            for col_c in date_cols:
                for r in range(n_rows):
                    v = cv_fn(r, col_c + 1)
                    if not v:
                        continue
                    for kw_list, fc in ING_COLOR_RULES:
                        if any(kw in v for kw in kw_list):
                            af_fn(r, col_c + 1, fc)
                            break
        elif fmt in ('ayumi', 'miyama'):
            # 歩学園・美山保育園: col3〜col9が材料
            for r in range(n_rows):
                for c in range(3, min(10, n_cols)):
                    v = cv_fn(r, c)
                    if not v:
                        continue
                    for kw_list, fc in ING_COLOR_RULES:
                        if any(kw in v for kw in kw_list):
                            af_fn(r, c, fc)
                            break
        else:
            # 既存形式: 「N日(曜)」横並びブロック検出
            blocks = []
            for r in range(n_rows):
                temp = {}
                for c in range(n_cols):
                    v = cv_fn(r, c)
                    if re.match(r'^\d+日\([月火水木金土日]\)$', v):
                        temp[c] = v
                if len(temp) >= 3:
                    blocks.append((r, temp))
            for b_idx, (b_row, b_cols) in enumerate(blocks):
                b_end = blocks[b_idx + 1][0] if b_idx + 1 < len(blocks) else n_rows
                mat_row = None
                for r in range(b_row + 1, b_end):
                    for c in range(min(5, n_cols)):
                        if cv_fn(r, c) == "材料":
                            mat_row = r
                            break
                    if mat_row is not None:
                        break
                if mat_row is None:
                    continue
                for col_c in sorted(b_cols.keys()):
                    for r in range(mat_row, b_end):
                        v = cv_fn(r, col_c)
                        if not v:
                            continue
                        for kw_list, fc in ING_COLOR_RULES:
                            if any(kw in v for kw in kw_list):
                                af_fn(r, col_c, fc)
                                break

    # ─── .xlsx ───────────────────────────────────────────────
    if not is_xls:
        wb = load_workbook(BytesIO(file_bytes), data_only=True)
        for sh_name in wb.sheetnames:
            ws = wb[sh_name]
            n_rows, n_cols = ws.max_row, ws.max_column
            df_sh = pd.read_excel(BytesIO(file_bytes), sheet_name=sh_name,
                                  header=None, dtype=str, engine='openpyxl')
            fmt = _detect_sheet_format(df_sh)

            def _cv(r, c, _ws=ws):
                if r < 0 or r >= _ws.max_row or c < 0 or c >= _ws.max_column:
                    return ""
                v = _ws.cell(row=r + 1, column=c + 1).value
                return "" if v is None else ("" if str(v).strip() in ("nan", "", "None") else str(v).strip())

            def _af(r, c, fill, _ws=ws):
                try:
                    _ws.cell(row=r + 1, column=c + 1).fill = fill
                except AttributeError:
                    pass

            _apply_colors(_cv, _af, n_rows, n_cols, fmt)

    # ─── .xls ────────────────────────────────────────────────
    else:
        import xlrd as _xlrd
        # formatting_info=True で列幅・行高さを取得
        xls_wb = _xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
        wb = Workbook()
        first = True
        for sh_name in xls_wb.sheet_names():
            xls_ws = xls_wb.sheet_by_name(sh_name)
            n_cols = xls_ws.ncols

            # データがある実際の最終行を特定（空行の大量コピーを防ぐ）
            last_data_row = 0
            for r_i in range(xls_ws.nrows):
                if any(str(xls_ws.cell_value(r_i, c_i)).strip() not in ('', 'nan', 'None')
                       for c_i in range(n_cols)):
                    last_data_row = r_i
            n_rows = last_data_row + 1

            if first:
                ws = wb.active
                ws.title = sh_name[:31]
                first = False
            else:
                ws = wb.create_sheet(title=sh_name[:31])

            # 値をコピー
            for r_i in range(n_rows):
                for c_i in range(n_cols):
                    v = xls_ws.cell_value(r_i, c_i)
                    if v is not None and str(v).strip() not in ("nan", "", "None"):
                        ws.cell(row=r_i + 1, column=c_i + 1, value=str(v).strip() or None)

            # マージセルをコピー（実データ範囲内のみ）
            for row_lo, row_hi, col_lo, col_hi in xls_ws.merged_cells:
                if row_lo >= n_rows:
                    continue
                try:
                    ws.merge_cells(start_row=row_lo + 1, start_column=col_lo + 1,
                                   end_row=min(row_hi, n_rows), end_column=col_hi)
                except Exception:
                    pass

            # 列幅・行高さ（formatting_info=True で取得済み）
            try:
                for c_i in range(n_cols):
                    ci = xls_ws.colinfo_map.get(c_i)
                    if ci and ci.width > 0:
                        ws.column_dimensions[get_column_letter(c_i + 1)].width = ci.width / 256
            except Exception:
                pass
            try:
                for r_i in range(n_rows):
                    ri = xls_ws.rowinfo_map.get(r_i)
                    if ri and ri.height > 0:
                        ws.row_dimensions[r_i + 1].height = ri.height / 20
            except Exception:
                pass

            # 罫線をコピー
            _XLS_LINE = {
                1: 'thin', 2: 'medium', 3: 'dashed', 4: 'dotted',
                5: 'thick', 6: 'double', 7: 'hair', 8: 'mediumDashed',
                9: 'dashDot', 10: 'mediumDashDot', 11: 'dashDotDot',
                12: 'mediumDashDotDot', 13: 'slantDashDot',
            }

            def _side(line_type, colour_idx):
                style = _XLS_LINE.get(line_type)
                if not style:
                    return Side(border_style=None)
                rgb = xls_wb.colour_map.get(colour_idx)
                color = f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}' if rgb else '000000'
                return Side(border_style=style, color=color)

            for r_i in range(n_rows):
                for c_i in range(n_cols):
                    try:
                        xf = xls_wb.xf_list[xls_ws.cell_xf_index(r_i, c_i)]
                        b = xf.border
                        ws.cell(row=r_i + 1, column=c_i + 1).border = Border(
                            left=_side(b.left_line_style, b.left_colour_index),
                            right=_side(b.right_line_style, b.right_colour_index),
                            top=_side(b.top_line_style, b.top_colour_index),
                            bottom=_side(b.bottom_line_style, b.bottom_colour_index),
                        )
                    except Exception:
                        pass

            # フォーマット検出
            df_sh = pd.read_excel(BytesIO(file_bytes), sheet_name=sh_name,
                                  header=None, dtype=str, engine='xlrd')
            fmt = _detect_sheet_format(df_sh)

            def _cv(r, c, _xws=xls_ws):
                if r < 0 or r >= _xws.nrows or c < 0 or c >= _xws.ncols:
                    return ""
                v = _xws.cell_value(r, c)
                return "" if v is None else ("" if str(v).strip() in ("nan", "", "None") else str(v).strip())

            def _af(r, c, fill, _ws=ws):
                try:
                    _ws.cell(row=r + 1, column=c + 1).fill = fill
                except AttributeError:
                    pass

            _apply_colors(_cv, _af, n_rows, n_cols, fmt)

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


def compute_all_python_ngs(excel_text, rules_text=""):
    """
    全ルールベースNGをPythonで計算。
    選択中ルール本文（rules_text）にステップ見出しが含まれるかで各チェックのON/OFFを判定する
    （ルールごとにチェック内容が異なるため、ハードコード項目を無条件に全ルールへ適用しない）。
    Returns: (summary_text, day_ngs_dict)
      summary_text  … AIプロンプトに埋め込むNG一覧テキスト
      day_ngs_dict  … {date_str: [ng_str, ...]}
    """
    year, month_num, sorted_dates, entries = _parse_structured(excel_text)
    if not sorted_dates:
        return "", {}

    day_ngs = {ds: [] for ds in sorted_dates}

    def rule_has(*markers):
        return any(m in rules_text for m in markers)

    check_forbidden           = rule_has('使用禁止食材')
    check_consecutive_general = rule_has('食材の連続使用')
    check_cheese_consec       = rule_has('チーズ類の連続使用')
    check_meat3               = rule_has('肉類の3日連続使用', '肉類の３日連続使用')
    check_seasoning_consec    = rule_has('調味料・ベースの連続使用')
    check_imo                 = rule_has('芋類の連続・重複使用')
    check_same_day_dup        = rule_has('同じ日の材料内の重複')
    check_monthly_limit       = rule_has('月の使用上限')
    check_monday_holiday      = rule_has('月曜日・祝日翌日のチェック', '祝日翌日のチェック')
    check_weekly_tofu         = rule_has('週次 豆腐チェック', '週次豆腐チェック')
    check_weekly_menu         = rule_has('週次 献立チェック', '週次献立チェック')
    check_name_ingredient     = rule_has('献立名と材料の照合チェック')
    check_prep_ingredients    = rule_has('仕込み食材の集中チェック')
    check_monthly_count       = rule_has('献立名の出現回数チェック')
    check_saturday_noodle     = rule_has('土曜日」に麺', '土曜日に麺')
    check_jelly_prev          = rule_has('を含む文言がある日の前日')
    check_seasonal_stew       = rule_has('クリームシチュー')

    def ing(ds):
        return entries[ds]['ingredients']

    def lunch(ds):
        return entries[ds]['lunch']

    def snack(ds):
        return entries[ds]['snack']

    def has(text, kw_list):
        return any(kw in text for kw in kw_list)

    # ── 禁止食材 ────────────────────────────────────────────────
    if check_forbidden:
        for ds in sorted_dates:
            for f in ['卵', 'マヨネーズ', '絹ごし豆腐']:
                if f in ing(ds):
                    day_ngs[ds].append(f'● {f}（使用禁止食材）')

    # ── チーズ2日連続 ──────────────────────────────────────────
    if check_cheese_consec:
        for i in range(1, len(sorted_dates)):
            if 'チーズ' in ing(sorted_dates[i]) and 'チーズ' in ing(sorted_dates[i-1]):
                day_ngs[sorted_dates[i]].append('● チーズ類2日連続')

    # ── 肉類3日連続 ────────────────────────────────────────────
    if check_meat3:
        for meat in MEAT_3DAY:
            for i in range(2, len(sorted_dates)):
                if all(meat in ing(sorted_dates[j]) for j in (i, i-1, i-2)):
                    day_ngs[sorted_dates[i]].append(f'● {meat}3日連続')

    # ── 調味料2日連続（日祝を挟めばOK） ──────────────────────
    if check_seasoning_consec:
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
    if check_imo:
        for i in range(2, len(sorted_dates)):
            if all(has(ing(sorted_dates[j]), IMO_KW) for j in (i, i-1, i-2)):
                day_ngs[sorted_dates[i]].append('● 芋類3日連続')

    # ── 汎用食材2日連続（免除リスト以外） ────────────────────
    if check_consecutive_general:
        _NUM_ONLY = re.compile(r'^\d+\.?\d*$')  # 数値トークン（量・シリアル等）を除外
        for i in range(1, len(sorted_dates)):
            ds_c, ds_p = sorted_dates[i], sorted_dates[i-1]
            toks_c = {t for t in _split_ing(ing(ds_c)) if not _is_exempt(t) and len(t) >= 2 and not _NUM_ONLY.match(t)}
            toks_p = {t for t in _split_ing(ing(ds_p)) if not _is_exempt(t) and len(t) >= 2 and not _NUM_ONLY.match(t)}
            for token in sorted(toks_c & toks_p):
                day_ngs[ds_c].append(f'● {token}2日連続')

    # ── 同日チェック ───────────────────────────────────────────
    if check_imo or check_same_day_dup:
        for ds in sorted_dates:
            i_text = ing(ds)
            toks = _split_ing(i_text)

            # 芋類2種類以上（同日）
            if check_imo:
                imo_cnt = sum(1 for kw in IMO_KW if kw in i_text)
                if imo_cnt >= 2:
                    day_ngs[ds].append(f'● 同日芋類{imo_cnt}種類（1種類のみ可）')

            if check_same_day_dup:
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
    if check_monthly_limit:
        for item in ['バター', 'チーズ', 'マヨドレ']:
            found = [ds for ds in sorted_dates if item in ing(ds)]
            if len(found) >= 4:
                day_ngs[found[3]].append(f'● {item}月4回目（上限超過）')

    # ── 月曜・祝日翌日チェック ────────────────────────────────
    if check_monday_holiday:
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
    if check_weekly_tofu or check_weekly_menu:
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

            if check_weekly_menu and not has(combined, FISH_WITH_TUNA):
                day_ngs[last].append('● 今週魚なし（週1回以上必要）')

            if check_weekly_tofu:
                tofu_ok = any(
                    has(ing(ds), TOFU_KW) or has(lunch(ds), TOFU_KW)
                    for ds in wd
                )
                if not tofu_ok:
                    day_ngs[last].append('● 今週豆腐なし（週1回以上必要）')

            if check_weekly_menu and not has(combined, NOODLE_KW):
                day_ngs[last].append('● 今週麺・丼なし（週1回以上必要）')

    # ── 仕込み食材7種類以上 ───────────────────────────────────
    if check_prep_ingredients:
        for ds in sorted_dates:
            i_text = ing(ds)
            found = [kw for kw in PREP_KW if kw in i_text]
            if len(found) >= 7:
                day_ngs[ds].append(f'● 仕込み食材{len(found)}種類（7種類以上）: {", ".join(found)}')

    # ── 月の献立回数上限 ──────────────────────────────────────
    if check_monthly_count:
        ingenge_days = [ds for ds in sorted_dates if 'いんげん' in (lunch(ds) + snack(ds))]
        fruit_days   = [ds for ds in sorted_dates if '果物' in (lunch(ds) + snack(ds))]
        if len(ingenge_days) >= 3:
            day_ngs[ingenge_days[2]].append('● 「いんげん」献立が月3回目（上限超過）')
        if len(fruit_days) >= 2:
            day_ngs[fruit_days[1]].append('● 「果物」献立が月2回目（上限超過）')

    # ── 土曜日に麺・丼なし ───────────────────────────────────
    if check_saturday_noodle:
        for ds in sorted_dates:
            if _dow(ds) == 5 and not has(lunch(ds), NOODLE_KW):
                day_ngs[ds].append('● 土曜日に麺・丼なし')

    # ── ハンバーグ/フライ前日おやつゼリーなし ────────────────
    if check_jelly_prev:
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
    if check_seasonal_stew and 5 <= month_num <= 9:
        for ds in sorted_dates:
            if 'クリームシチュー' in lunch(ds):
                day_ngs[ds].append('● 5〜9月クリームシチューNG')

    # ── 料理名から推定される必須材料チェック ─────────────────
    # (料理名に含む文字列, 材料欄に必要なキーワードのいずれか)
    if check_name_ingredient:
        RECIPE_RULES = [
            ('おかか',   ['かつお節', 'おかか']),
            ('ごま和え', ['白ごま', '黒ごま', 'すりごま', 'ごま']),
            ('あんかけ', ['片栗粉']),
            ('から揚げ', ['片栗粉']),
            ('唐揚げ',   ['片栗粉']),
            ('照り焼き', ['みりん', '醤油']),
            ('蒸しパン', ['ベーキングパウダー', 'BP', '重曹']),
            ('味噌汁',           ['みそ']),
            ('みそ汁',           ['みそ']),
            ('果物',             FRUIT_KW),
            ('フルーツヨーグルト', FRUIT_KW),
            ('ヨーグルト',       ['ヨーグルト']),       # フルーツヨーグルト・桃ヨーグルト等すべて対象
            ('お菓子',           ['お菓子']),
        ]
        for ds in sorted_dates:
            ls_text = lunch(ds) + ' ' + snack(ds)
            i_text  = ing(ds)
            for kw, required_any in RECIPE_RULES:
                if kw in ls_text:
                    if not any(req in i_text for req in required_any):
                        missing = '・'.join(required_any)
                        day_ngs[ds].append(f'● 「{kw}」があるが材料に{missing}なし')

        # ── 献立名の果物名と材料の果物が一致しないチェック ────────
        # （例：「パインパンケーキ」なのに材料が「みかん缶」など、別の果物に
        #   すり替わっているケースを検出。表記ゆれ（パイン/パイナップル等）は同一視する）
        FRUIT_GROUPS = [
            ['パイン', 'パイナップル'],
            ['もも', '桃'],
            ['いちご', 'イチゴ'],
            ['ぶどう', 'ブドウ'],
            ['すいか', 'スイカ'],
            ['なし', '梨'],
            ['みかん'],
            ['りんご'],
            ['バナナ'],
            ['オレンジ', 'マーマレード'],  # 「オレンジ蒸しパン」等はマーマレードがあればOK
            ['メロン'],
            ['キウイ'],
            ['マンゴー'],
            ['さくらんぼ'],
        ]
        for ds in sorted_dates:
            ls_text = lunch(ds) + ' ' + snack(ds)
            i_text  = ing(ds)
            for group in FRUIT_GROUPS:
                if any(g in ls_text for g in group) and not any(g in i_text for g in group):
                    day_ngs[ds].append(f'● 献立名に「{group[0]}」があるが材料に見当たらない（別の果物に違っていないか要確認）')

        # ── おすまし・おすいものに「みそ」あり ─────────────────────
        for ds in sorted_dates:
            ls_text = lunch(ds) + ' ' + snack(ds)
            if ('おすまし' in ls_text or 'おすいもの' in ls_text) and 'みそ' in ing(ds):
                day_ngs[ds].append('● 「おすまし/おすいもの」があるが材料に「みそ」あり（不要）')

        # ── 料理名に含まれない魚が材料にあるチェック ────────────────
        # 「白身魚」は総称なので除外（白身魚のフライ→材料にタラ等があっても正常）
        _specific_fish = [f for f in FISH_KW if f != '白身魚']
        for ds in sorted_dates:
            ls_text = lunch(ds) + ' ' + snack(ds)
            i_text  = ing(ds)
            dish_fish = [f for f in _specific_fish if f in ls_text]
            ing_fish  = [f for f in _specific_fish if f in i_text]
            if dish_fish:
                for f in ing_fish:
                    if f not in dish_fish:
                        day_ngs[ds].append(f'● 材料に「{f}」があるが献立名に対応する料理なし（不要食材？）')

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
    python_ng_text, _ = compute_all_python_ngs(excel_text, rules_text)

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
・「オレンジ蒸しパン」などオレンジ風味の料理 → マーマレードがあればOK（オレンジ・みかんがなくても可）

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

            st.markdown("### 3. 使用するルールを選択")
            all_rules = load_rules_list()
            if not all_rules:
                st.warning("ルールが登録されていません。「ルール管理」ページでルールを追加してください。")
                chosen_name = ""
                selected_rule_text = ""
            else:
                rule_names = [r["name"] for r in all_rules]
                chosen_name = st.selectbox(
                    "ルール", rule_names,
                    key="rule_selector",
                    label_visibility="collapsed"
                )
                selected_rule_text = next(
                    (r["text"] for r in all_rules if r["name"] == chosen_name), ""
                )

            st.markdown("### 4. 色付きExcel生成（目検用）")
            if st.button("🎨 色付きExcel生成（全シート）"):
                with st.spinner("色付き処理中..."):
                    uploaded.seek(0)
                    colored = create_colored_excel(uploaded)
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

            st.markdown("### 5. チェック開始")
            if st.button("✅ チェックを開始する", type="primary", disabled=not api_key):
                rules = selected_rule_text
                if not rules.strip():
                    st.error("ルールが選択されていないか空です。「ルール管理」ページでルールを設定してください。")
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
                            st.session_state["last_sheet"] = selected_sheet
                            st.session_state["last_rule_name"] = chosen_name
                        except Exception as e:
                            st.error(f"エラーが発生しました: {e}")
                            result = None

            if st.session_state.get("last_result"):
                st.markdown("---")
                _fname_disp = st.session_state.get("last_filename", "")
                _sheet_disp = st.session_state.get("last_sheet", "")
                _rule_disp = st.session_state.get("last_rule_name", "")
                st.subheader(f"チェック結果 ― {_fname_disp}　シート: {_sheet_disp}")
                if _rule_disp:
                    st.caption(f"使用ルール：{_rule_disp}")
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
    st.caption("ルールを登録・編集すると、チェック画面のセレクトボックスに表示されます。保存はGitHubにも自動反映されます。")

    all_rules = load_rules_list()

    # ── 編集フォーム（新規追加 or 既存編集）────────────────────
    editing = st.session_state.get("rule_editing")  # {"id": str|None, "name": str, "text": str}

    if editing is not None:
        is_new = editing["id"] is None
        st.subheader("➕ 新しいルールを追加" if is_new else f"✏️ 編集：{editing['name']}")
        with st.form("rule_form"):
            new_name = st.text_input(
                "ルール名（例：さかえ保育園、おおみや共通 など）",
                value=editing["name"],
                max_chars=50,
            )
            new_text = st.text_area(
                "ルール本文",
                value=editing["text"],
                height=500,
                help="このルールをAIに渡してチェックします。",
            )
            col_s, col_c = st.columns([1, 5])
            with col_s:
                submitted = st.form_submit_button("💾 保存", type="primary")
            with col_c:
                cancelled = st.form_submit_button("✖ キャンセル")

        if submitted:
            if not new_name.strip():
                st.error("ルール名を入力してください。")
            else:
                if is_new:
                    new_id = str(int(datetime.datetime.now().timestamp() * 1000))
                    all_rules.append({"id": new_id, "name": new_name.strip(), "text": new_text})
                else:
                    for r in all_rules:
                        if r["id"] == editing["id"]:
                            r["name"] = new_name.strip()
                            r["text"] = new_text
                            break
                save_rules_list(all_rules)
                ok, msg = push_rules_list_to_github(all_rules)
                if ok:
                    st.success(msg)
                else:
                    st.warning(f"アプリには保存済みです。GitHub更新に失敗しました：{msg}")
                del st.session_state["rule_editing"]
                st.rerun()
        if cancelled:
            del st.session_state["rule_editing"]
            st.rerun()

    else:
        if st.button("➕ 新しいルールを追加", type="primary"):
            st.session_state["rule_editing"] = {"id": None, "name": "", "text": ""}
            st.rerun()

    # ── ルール一覧 ────────────────────────────────────────────
    if editing is None:
        st.markdown("---")
        if not all_rules:
            st.info("ルールがまだ登録されていません。上のボタンから追加してください。")
        else:
            st.markdown(f"**登録済みルール：{len(all_rules)} 件**")
            for rule in all_rules:
                with st.container(border=True):
                    col_n, col_e, col_d = st.columns([6, 1, 1])
                    with col_n:
                        st.markdown(f"**{rule['name']}**")
                        preview = rule["text"][:80].replace("\n", "  ") + ("…" if len(rule["text"]) > 80 else "")
                        st.caption(preview)
                    with col_e:
                        if st.button("✏ 編集", key=f"edit_{rule['id']}"):
                            st.session_state["rule_editing"] = {
                                "id": rule["id"],
                                "name": rule["name"],
                                "text": rule["text"],
                            }
                            st.rerun()
                    with col_d:
                        # 削除: 1回目クリックで確認フラグ、2回目で実行
                        confirm_key = f"confirm_del_{rule['id']}"
                        if st.session_state.get(confirm_key):
                            if st.button("本当に削除", key=f"yes_{rule['id']}",
                                         type="primary"):
                                all_rules = [r for r in all_rules if r["id"] != rule["id"]]
                                save_rules_list(all_rules)
                                push_rules_list_to_github(all_rules)
                                del st.session_state[confirm_key]
                                st.rerun()
                        else:
                            if st.button("🗑 削除", key=f"del_{rule['id']}"):
                                st.session_state[confirm_key] = True
                                st.rerun()
