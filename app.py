import re
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
    rows = []
    for i, row in df.iterrows():
        cells = [
            str(v).strip() if pd.notna(v) and str(v).strip() not in ("nan", "") else ""
            for v in row
        ]
        if any(cells):
            rows.append(f"行{i + 1}: " + " | ".join(cells))
    return "\n".join(rows)


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


def run_check(excel_text, rules_text, api_key, file_name, sheet_name):
    client = anthropic.Anthropic(api_key=api_key)

    prompt = f"""あなたは幼稚園給食の献立チェック専門家です。
以下のチェックルールに従って、エクセルデータを厳密に確認してください。

========================================
# チェックルール
========================================
{rules_text}

========================================
# エクセルデータ（ファイル：{file_name}　シート：{sheet_name}）
========================================
各行は「行N: セル1 | セル2 | ...」の形式です。空白セルは空欄として表示しています。

{excel_text}

========================================
# 出力形式（絶対厳守）
========================================
マークダウンテーブルのみを出力してください。テーブルの前後に文章・注釈・説明等は一切禁止です。

| 日付 | 献立名 | おやつ | 結果 |
|------|--------|--------|------|
| 9/1(火) | ごはん / 豚肉のポン酢炒め / もやしごま和え / みそ汁 | ツナトースト | ● 豚肉2日連続 |
| 9/2(水) | ごはん / サケのマーマレード焼き / 小松菜のごま和え / みそ汁 | ここア蒸しパン | OK |

========================================
# 結果欄のルール（例外なし・絶対厳守）
========================================

【記入できるのは以下の2パターンのみ】

パターン1：問題なし → 「OK」
  - 「OK」の2文字だけ。それ以外は何も書かない。
  - チェックした内容・理由・補足・括弧書きを一切付けない。
  - 例：「OK」← これだけ
  - 禁止例：「OK（対象外）」「●...あり/OK」「●...可」「● ゼリーあり/OK」
    「●...問題なし」「● 月曜にXXなし（該当外）」「バナナ（火曜は対象外）」

パターン2：確定NGあり → 「● NG理由を簡潔に」
  - 複数NGは「／」でつなぐ
  - 例：「● 豚肉3日連続／月曜に魚使用」

【中間表現は禁止】
「要確認」「確認待ち」「確認事項」「可能性あり」などは書かない。
OK か ● のどちらかのみ。

========================================
# 週次チェックの記入方法
========================================
- 週の区切りは月曜〜土曜（この範囲外の日は別の週として扱う）
- 週次チェック（魚なし・豆腐なし・麺丼なし）のNGは、その週の土曜日の行に記入する
- 土曜が休日の場合は、その週最後の給食がある日の行に記入する
- 例：9/7(月)〜9/13(土)の週に魚がなければ → 9/13(土)の行に「● 今週魚なし」

========================================
# その他のルール
========================================
- 献立名はスラッシュ「/」でつなぐ
- おやつがない日は空欄
- 給食のない日（土日・祝日）は出力しない
"""

    message = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


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
    st.caption("チェックに使うルールを確認・編集できます。変更後は「保存」を押してください。")

    current = load_rules()
    edited = st.text_area("チェックルール", value=current, height=600)

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("💾 保存", type="primary"):
            save_rules(edited)
            st.success("保存しました！")
    with col2:
        if st.button("🔄 元に戻す（最後の保存時点）"):
            st.rerun()
