import streamlit as st
import anthropic
import pandas as pd
from pathlib import Path

st.set_page_config(
    page_title="献立チェックシステム",
    layout="wide",
    page_icon="🍱"
)

BASE_DIR = Path(__file__).parent
RULES_FILE = BASE_DIR / "rules.txt"

# Ctrl+CでSteamlitのキャッシュクリアダイアログが出るのを防ぐ
st.markdown("""
<script>
document.addEventListener('keydown', function(e) {
    if (e.ctrlKey && (e.key === 'c' || e.key === 'C')) {
        e.stopPropagation();
    }
}, true);
</script>
""", unsafe_allow_html=True)


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


def run_check(excel_text, rules_text, api_key, file_name, sheet_name):
    client = anthropic.Anthropic(api_key=api_key)

    prompt = f"""あなたは幼稚園給食の献立チェック専門家です。
以下のチェックルールに従って、エクセルデータを厳密に確認し、NGの項目をすべて洗い出してください。

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
# 出力形式
========================================
以下のマークダウンテーブルのみを出力してください。
表の前後に説明文・注釈・ステップ結果・件数サマリー等は一切付けないでください。

| 日付 | 献立名 | おやつ | 結果 |
|------|--------|--------|------|
| 9/1(火) | ごはん / 豚肉のポン酢炒め / もやしごま和え / みそ汁 | ツナトースト | ● 【理由を簡潔に】 |
| 9/2(水) | ごはん / サケのマーマレード焼き / ... | ここア蒸しパン | OK |

ルール：
- 献立名はスラッシュ「/」でつなぐ
- おやつがない日は空欄
- NGがない日は「OK」とだけ記入する（理由・説明・注釈は絶対に付けない）
- NGがある日は「● 【内容】」（複数NGは「／」でつなぐ）
- 給食のない日（土日・祝日）は出力しない

テーブルの後に、必ず以下のセパレーターを1行で入れてください：
===NG_SUMMARY===

その後、確定NGと要確認事項のみを以下の形式で出力してください（余分なテキスト不要）：

1. **日付**：NG理由 → NG
2. **日付**：NG理由 → NG
（確定NGがない場合はこのセクションは省略）

**要確認事項**
- 内容
（要確認事項がない場合は「**要確認事項**」ごと省略）

最後に「確定NG件数：X件」の1行のみ
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
                raw = st.session_state["last_result"]
                if "===NG_SUMMARY===" in raw:
                    display_part, download_part = raw.split("===NG_SUMMARY===", 1)
                else:
                    display_part = raw
                    download_part = raw
                st.markdown(display_part)
                fname = st.session_state.get("last_filename", "result").rsplit(".", 1)[0]
                st.download_button(
                    "📥 結果をテキストファイルで保存",
                    data=download_part.strip().encode("utf-8-sig"),
                    file_name=f"チェック結果_{fname}.txt",
                    mime="text/plain",
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
