import io
import json
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from google.oauth2 import service_account

# --- 1. パスワード認証機能 ---
def check_password():
    """パスワード認証を行い、認証成功時のみ True を返す"""
    def password_entered():
        correct_password = st.secrets.get("APP_PASSWORD", "sto0123")
        if st.session_state["password_input"] == correct_password:
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    st.text_input(
        "パスワードを入力してください",
        type="password",
        on_change=password_entered,
        key="password_input"
    )
    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("パスワードが違います")
    return False

# --- 2. メイン処理 ---
if check_password():
    st.title("請求書 単価推移・比較表作成 (Vertex AI)")
    st.write("GCP Vertex AI（エンタープライズ環境）で請求書PDFを解析し、月別単価推移表を生成します。")

    uploaded_files = st.file_uploader(
        "各月の請求書PDFを選択してください（複数可）",
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_files and st.button("単価推移表を生成"):
        if "gcp_service_account" not in st.secrets:
            st.error("Secrets に [gcp_service_account] が設定されていません。")
        else:
            with st.spinner("Vertex AI でPDF解析および名寄せ集計中..."):
                try:
                    # サービスアカウント認証オブジェクトの作成（OAuthスコープを明示）
                    creds_dict = dict(st.secrets["gcp_service_account"])
                    credentials = service_account.Credentials.from_service_account_info(
                        creds_dict,
                        scopes=["https://www.googleapis.com/auth/cloud-platform"]
                    )

                    # Vertex AI 用の GenAI クライアント初期化
                    client = genai.Client(
                        vertexai=True,
                        project=creds_dict.get("project_id"),
                        location=st.secrets.get("GCP_LOCATION", "us-central1"),
                        credentials=credentials
                    )

                    # アップロードされたPDFを並列データとしてセット
                    pdf_parts = []
                    for file in uploaded_files:
                        pdf_parts.append(
                            types.Part.from_bytes(
                                data=file.read(),
                                mime_type="application/pdf"
                            )
                        )

                    # 解析指示プロンプト
                    prompt = """
                    提供されたすべての請求書PDF（複数月分）を解析し、品目ごとの単価推移表（比較表）を作成してください。

                    【抽出・集計条件】
                    1. 品名、規格・サイズ、メーカー名を抽出してください。
                    2. 文字認識（OCR）の誤字・表記ゆれ（例: 「ベベル「ボード」と「ベベルボード」など）は同一品目として整理（名寄せ）してください。
                    3. 対象となる各月（例: 4月、5月、6月、7月）ごとの「単価」を抽出してください。
                    4. 単価に変更があった場合は「備考」に（例: 「7月に値上げ」）と記載してください。

                    【返却フォーマット】
                    以下の構造を持つ JSON 配列オブジェクトのみを出力してください。
                    [
                      {
                        "メーカー": "吉野石膏",
                        "品名": "ベベルボード",
                        "サイズ": "12.5mm 3x8",
                        "4月単価": 850,
                        "5月単価": 850,
                        "6月単価": 850,
                        "7月単価": 980,
                        "備考": "7月に値上げ"
                      }
                    ]
                    """

                    # Vertex AI モデル呼び出し
                    response = client.models.generate_content(
                        model="gemini-2.5-flash",
                        contents=[*pdf_parts, prompt],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json"
                        )
                    )

                    # 結果表示とデータフレーム化
                    data = json.loads(response.text)
                    df = pd.DataFrame(data)

                    st.success("解析および単価比較表の生成が完了しました。")
                    st.dataframe(df)

                    # Excelダウンロード機能
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="単価推移比較表")

                    st.download_button(
                        label="単価推移表（Excel）をダウンロード",
                        data=output.getvalue(),
                        file_name="月別品目単価推移表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"Vertex AI 処理中にエラーが発生しました: {e}")
