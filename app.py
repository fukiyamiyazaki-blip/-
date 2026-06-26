import re
import json
import base64
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
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

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

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        para = cell.paragraphs[0]
        if para.runs:
            run = para.runs[0]
            run.bold = True
            run.font.size = Pt(10)

    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = val
            para = cell.paragraphs[0]
            if para.runs:
                run = para.runs[0]
                run.font.size = Pt(9)
                if c_idx == n_cols - 1 and val.startswith('●'):
                    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

    bio = BytesIO()
    doc.save(bio)
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
    return '\n'.join(cleaned)


def compute_weekly_summary(excel_text):
    """週次チェック項目（魚・豆腐・麺丼）をPythonで事前計算して返す"""
    fish_kw = ["サケ","サーモン","サバ","サワラ","タラ","アジ","イワシ","ブリ","カレイ",
               "メカジキ","タイ","マグロ","カツオ","シシャモ","ほっけ","白身魚","ツナ"]
    tofu_kw = ["木綿豆腐","焼き豆腐","厚揚げ","油揚げ","高野豆腐"]
    noodle_kw = ["スパゲティ","うどん","めん","麺","そば","丼"]
    dow = {"月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6}

    # 構造化テキストから日付ごとのテキストを抽出
    date_data = {}
    current = None
    for line in excel_text.split("\n"):
        m = re.match(r'【(\d+/\d+\([月火水木金土日]\))】', line)
        if m:
            current = m.group(1)
            date_data[current] = ""
        elif current and (line.startswith("昼食:") or line.startswith("おやつ:") or line.startswith("材料:")):
            date_data[current] += " " + line

    if not date_data:
        return ""

    def week_key(ds):
        m = re.search(r'/(\d+)\(([月火水木金土日])\)', ds)
        if not m:
            return -1
        return int(m.group(1)) - dow.get(m.group(2), 0)

    def day_num(ds):
        m = re.search(r'/(\d+)\(', ds)
        return int(m.group(1)) if m else 0

    # 週ごとにグループ化
    weeks = {}
    for ds in date_data:
        wk = week_key(ds)
        weeks.setdefault(wk, []).append(ds)

    lines = ["【週次チェック事前計算結果】（AIはこの結果をそのまま使用すること。自分で再計算しないこと）"]
    wn = 1
    for wk in sorted(weeks.keys()):
        dates = sorted(weeks[wk], key=day_num)
        start, end = dates[0], dates[-1]
        fish, tofu, noodle = [], [], []
        for ds in dates:
            text = date_data[ds]
            for kw in fish_kw:
                if kw in text:
                    fish.append(f"{kw}({ds})")
                    break
            for kw in tofu_kw:
                if kw in text:
                    tofu.append(f"{kw}({ds})")
                    break
            for kw in noodle_kw:
                if kw in text:
                    noodle.append(f"{kw}({ds})")
                    break
        fish_s = "あり: " + ", ".join(fish) if fish else "なし→NG"
        tofu_s = "あり: " + ", ".join(tofu) if tofu else "なし→NG"
        noodle_s = "あり: " + ", ".join(noodle) if noodle else "なし→NG"
        lines.append(f"第{wn}週({start}〜{end}): 魚={fish_s} / 豆腐={tofu_s} / 麺丼={noodle_s}")
        wn += 1

    return "\n".join(lines)


def run_check(excel_text, rules_text, api_key, file_name, sheet_name):
    client = anthropic.Anthropic(api_key=api_key)
    weekly_summary = compute_weekly_summary(excel_text)

    prompt = f"""あなたは幼稚園給食の献立チェック専門家です。
以下のチェックルールに従って、エクセルデータを厳密に確認してください。

========================================
# チェックルール
========================================
{rules_text}

========================================
# エクセルデータ（ファイル：{file_name}　シート：{sheet_name}）
========================================
各日付のデータは以下の形式で整理されています：
【月/日(曜日)】
昼食: 献立名1 / 献立名2 / ...（昼食の献立。汁物も含む）
おやつ: おやつ名1 / おやつ名2 / ...（午後のおやつ。「お菓子」は市販菓子類）
材料: 材料名1, 材料名2, ...（その日の全材料。みそ汁の具など汁物の中身も含む）

{excel_text}

========================================
# 週次チェック事前計算結果（絶対使用すること）
========================================
魚・豆腐・麺丼の週次チェックはPythonで正確に計算済みです。
AIは自分で再計算せず、必ず以下の結果をそのまま使用してください。

{weekly_summary}

========================================
# 出力形式（絶対厳守）
========================================
マークダウンテーブルのみを出力してください。テーブルの前後に文章・注釈・説明等は一切禁止です。

| 日付 | 献立名 | おやつ | 結果 |
|------|--------|--------|------|
| 9/1(火) | ごはん / 豚肉のポン酢炒め / もやしごま和え / みそ汁 | ツナトースト | ● 豚肉2日連続 |
| 9/2(水) | ごはん / サケのマーマレード焼き / 小松菜のごま和え / みそ汁 | ここア蒸しパン | OK |

【結果欄のルール（例外なし・絶対厳守）】

結果欄は「NG理由を書く列」であって「OK理由・確認内容を書く列」ではない。

書いてよいのは2種類のみ：
(1) 「OK」← 2文字のみ。何も追加しない
(2) 「● NG理由」← NG理由のみ

【絶対禁止パターン一覧】
× 「● ○○なし → OK」← 「なし」はNGがないということ。「OK」とだけ書け
× 「● ○○対象外 → OK」← 対象外なら「OK」とだけ書け
× 「● ソース・コンソメ2日連続なし → OK」← 禁止
× 「木曜だがしめじ等なし → OK」← 禁止
× 「太もやし対象外 → OK」← 禁止
× 「● NG理由 / ○○なし → OK」← NGとOK説明を混ぜるな
× 「OK（確認済み）」「OK（対象外）」「OK（問題なし）」← OKに何も付けるな
× 「● 魚なし（サバあり）」← 矛盾表現禁止

「なし」「対象外」「問題なし」「確認済み」「あり/OK」は結果欄に書かない。

【週次チェックの記入方法】
- 週次チェック（魚なし・豆腐なし・麺丼なし）のNGは、その週の土曜日の行に記入する
- 土曜が休日の場合は、その週最後の給食がある日の行に記入する
- 週次チェック事前計算結果で「あり」となっている週は、週次NGを絶対に書かない

【その他のルール】
- 献立名はスラッシュ「/」でつなぐ
- おやつがない日は空欄
- 給食のない日（土日・祝日）は出力しない
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

            st.markdown("### 3. チェック開始")
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
