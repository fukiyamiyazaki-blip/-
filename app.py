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
