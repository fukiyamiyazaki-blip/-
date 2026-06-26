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
NGのみを以下の形式で出力してください。OKの項目は一切書かないでください。

## チェック結果
NG件数：X件

---

## NG一覧

（NGがある場合のみ、以下の形式で箇条書き）
- 【ステップN】日付（曜日）/ 料理名または食材名 → 理由を1行で

（NGが1件もない場合）
- NGなし

---

## 要確認（目視推奨）
（AIでは確定判断が難しく、担当者に目視確認をお願いしたい項目のみ記載。なければ省略）
"""

    message = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


# ==============================
# サイドバー
# ==============================
with st.sidebar:
    st.title("🍱 献立チェックシステム")
    st.markdown("---")
    page = st.radio(
        "ページを選択",
        ["📋 献立チェック", "⚙️ ルール管理"],
        label_visibility="collapsed",
    )


# ==============================
# ページ：献立チェック
# ==============================
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
                st.markdown(st.session_state["last_result"])
                fname = st.session_state.get("last_filename", "result").rsplit(".", 1)[0]
                st.download_button(
                    "📥 結果をテキストファイルで保存",
                    data=st.session_state["last_result"],
                    file_name=f"チェック結果_{fname}.txt",
                    mime="text/plain",
                )


# ==============================
# ページ：ルール管理
# ==============================
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
