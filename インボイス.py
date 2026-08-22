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
    st.title("請求書 単価推移表 更新システム (Vertex AI)")
    st.write("既存の単価表エクセルと追加・更新したい請求書PDFを読み込み、データを照合・上書き更新します。")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. 既存の単価表エクセル（任意）")
        uploaded_excel = st.file_uploader(
            "更新対象のエクセルファイル (.xlsx) を選択",
            type=["xlsx"]
        )

    with col2:
        st.subheader("2. 追加・更新する請求書PDF（必須）")
        uploaded_pdfs = st.file_uploader(
            "追加・解析したいPDFを選択（複数可）",
            type=["pdf"],
            accept_multiple_files=True
        )

    if uploaded_pdfs and st.button("データ照合・更新を実行"):
        if "gcp_service_account" not in st.secrets:
            st.error("Secrets に [gcp_service_account] が設定されていません。")
        else:
            with st.spinner("Vertex AIでPDFを解析し、エクセルデータを照合・上書き中..."):
                try:
                    # 1. 既存エクセルの読み込み
                    existing_items_str = ""
                    df_existing = pd.DataFrame()
                    if uploaded_excel is not None:
                        df_existing = pd.read_excel(uploaded_excel)
                        if "品名" in df_existing.columns:
                            existing_items = df_existing["品名"].dropna().unique().tolist()
                            existing_items_str = f"\nなお、既存の単価表には以下の品名が存在します。表記ゆれがある場合は極力これらの既存品名に合わせて統一（名寄せ）してください:\n{json.dumps(existing_items, ensure_ascii=False)}"

                    # 2. Vertex AI クライアント初期化
                    creds_dict = dict(st.secrets["gcp_service_account"])
                    credentials = service_account.Credentials.from_service_account_info(
                        creds_dict,
                        scopes=["https://www.googleapis.com/auth/cloud-platform"]
                    )

                    client = genai.Client(
                        vertexai=True,
                        project=creds_dict.get("project_id"),
                        location=st.secrets.get("GCP_LOCATION", "us-central1"),
                        credentials=credentials
                    )

                    # 3. PDFバイナリのパッキング
                    pdf_parts = []
                    for file in uploaded_pdfs:
                        pdf_parts.append(
                            types.Part.from_bytes(
                                data=file.read(),
                                mime_type="application/pdf"
                            )
                        )

                    # 4. 解析指示プロンプトの構築（備考列を除外、月名を指定）
                    prompt = f"""
                    提供されたすべての請求書PDFを解析し、品目ごとの単価情報を抽出してください。
                    {existing_items_str}

                    【抽出・名寄せ条件】
                    1. メーカー名、品名、サイズ・規格を抽出してください。
                    2. 各月（12月単価、11月単価、... 1月単価）の単価を数値のみで抽出してください。単価がない月は null としてください。
                    3. 備考情報は抽出・出力しないでください。

                    【返却フォーマット】
                    以下の構造を持つ JSON 配列オブジェクトのみを出力してください。数値項目に空文字 "" や「備考」を含めないでください。
                    [
                      {{
                        "メーカー": "吉野石膏",
                        "品名": "ベベルボード",
                        "サイズ": "12.5mm 3x8",
                        "7月単価": 980,
                        "6月単価": 850,
                        "5月単価": 850,
                        "4月単価": 850
                      }}
                    ]
                    """

                    # 5. Gemini 呼び出し
                    response = client.models.generate_content(
                        model="gemini-2.5-flash",
                        contents=[*pdf_parts, prompt],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json"
                        )
                    )

                    # 6. PDF解析結果のデータフレーム化とクレンジング
                    new_data = json.loads(response.text)
                    df_new = pd.DataFrame(new_data)

                    # 数値列のデータ型クレンジング
                    for col in df_new.columns:
                        if "単価" in col or "金額" in col:
                            df_new[col] = pd.to_numeric(df_new[col].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce")

                    if not df_existing.empty:
                        for col in df_existing.columns:
                            if "単価" in col or "金額" in col:
                                df_existing[col] = pd.to_numeric(df_existing[col].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce")

                    # 7. 既存エクセルとの照合・重複排除・上書き・追加ロジック
                    if not df_existing.empty:
                        merge_key = "品名" if "品名" in df_existing.columns and "品名" in df_new.columns else None

                        if merge_key:
                            df_new_clean = df_new.drop_duplicates(subset=[merge_key], keep="last")
                            df_existing_clean = df_existing.drop_duplicates(subset=[merge_key], keep="last")

                            df_existing_indexed = df_existing_clean.set_index(merge_key)
                            df_new_indexed = df_new_clean.set_index(merge_key)

                            for col in df_new_indexed.columns:
                                if col not in df_existing_indexed.columns:
                                    df_existing_indexed[col] = None

                            df_existing_indexed.update(df_new_indexed)

                            new_rows = df_new_indexed.index.difference(df_existing_indexed.index)
                            if not new_rows.empty:
                                df_updated = pd.concat([df_existing_indexed, df_new_indexed.loc[new_rows]]).reset_index()
                            else:
                                df_updated = df_existing_indexed.reset_index()

                        else:
                            df_updated = pd.concat([df_existing, df_new], ignore_index=True)
                    else:
                        df_updated = df_new

                    # 8. 備考列の削除 ＆ 12月〜1月の降順並び替え整形
                    if "備考" in df_updated.columns:
                        df_updated = df_updated.drop(columns=["備考"])

                    base_cols = ["メーカー", "品名", "サイズ"]
                    month_cols = [f"{m}月単価" for m in range(12, 0, -1)]  # 12月〜1月の逆順
                    
                    # 存在する基本列 ＋ 12月〜1月順の存在する単価列の順に配置
                    ordered_cols = [c for c in base_cols if c in df_updated.columns] + \
                                   [c for c in month_cols if c in df_updated.columns]
                    
                    # その他もし予期せぬ列があれば末尾に配置
                    remaining_cols = [c for c in df_updated.columns if c not in ordered_cols]
                    df_updated = df_updated[ordered_cols + remaining_cols]

                    # 9. 画面表示とExcel出力（列幅自動調整）
                    st.success("データの照合・上書き・追加が完了しました。")
                    st.dataframe(df_updated)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        sheet_name = "単価推移比較表"
                        df_updated.to_excel(writer, index=False, sheet_name=sheet_name)

                        worksheet = writer.sheets[sheet_name]
                        for col in worksheet.columns:
                            max_length = 0
                            column_letter = col[0].column_letter
                            for cell in col:
                                if cell.value is not None:
                                    val_str = str(cell.value)
                                    length = sum(2 if ord(c) > 127 else 1 for c in val_str)
                                    if length > max_length:
                                        max_length = length
                            worksheet.column_dimensions[column_letter].width = max(max_length + 3, 10)

                    st.download_button(
                        label="更新後の単価推移表（Excel）をダウンロード",
                        data=output.getvalue(),
                        file_name="更新版_月別品目単価推移表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"データ処理中にエラーが発生しました: {e}")
