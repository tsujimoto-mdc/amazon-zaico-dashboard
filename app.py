# -*- coding: utf-8 -*-
"""
Amazon納品分 ZAICO在庫数 自動確認ダッシュボード
====================================================================
元の「zaico_tool.py（デスクトップ用GUIツール）」と、やっている事は
まったく同じです。ちがいは操作方法だけで、

  1. Amazon納品プランExcel(.xlsx) をアップロード
  2. ZAICOダウンロードCSV(.csv) をアップロード
  3. 「実行」ボタンを1回押す
  4. 「本社ZAICO / 3FZAICO」列が追記されたExcelをダウンロード

というブラウザ完結の流れになります。
アップロードしたファイルはサーバーに保存されず、その場のメモリ上だけで
処理されます（ブラウザを閉じる／再実行するとリセットされます）。
====================================================================
"""

import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from zaico_core import (
    LOCATION_3F,
    LOCATION_HONSHA,
    NEW_COL_3F,
    NEW_COL_HONSHA,
    run_transcription,
)

# ============================================================
# 基本設定
# ============================================================
st.set_page_config(
    page_title="Amazon納品分 ZAICO在庫数 自動確認",
    page_icon="📦",
    layout="wide",
)


# ============================================================
# パスワード保護（任意）
#   Secretsに APP_PASSWORD を設定した場合のみ、合言葉入力を求めます。
#   設定しなければ、そのまま誰でも使えます（社内限定URLでの運用を想定）。
# ============================================================
def check_password() -> bool:
    if "APP_PASSWORD" not in st.secrets:
        return True

    def password_entered():
        if st.session_state.get("password") == st.secrets.get("APP_PASSWORD"):
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct"):
        return True

    st.title("📦 Amazon納品分 ZAICO在庫数 自動確認")
    st.text_input(
        "合言葉を入力してください",
        type="password",
        on_change=password_entered,
        key="password",
    )
    if st.session_state.get("password_correct") is False:
        st.error("合言葉が正しくありません。")
    return False


# ============================================================
# ダウンロード用のZIP作成
# ============================================================
def make_zip(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, data in files.items():
            zf.writestr(filename, data)
    buf.seek(0)
    return buf.read()


# ============================================================
# メイン画面
# ============================================================
def main():
    if not check_password():
        return

    st.title("📦 Amazon納品分 ZAICO在庫数 自動確認")
    st.caption(
        "Amazonの納品プランExcelと、ZAICOのダウンロードCSVをアップロードして「実行」を押すと、"
        "「本社ZAICO / 3FZAICO」列を追記したExcelがダウンロードできます。"
    )

    with st.expander("このツールがやっていること（元のデスクトップ版と同じ）", expanded=False):
        st.markdown(
            f"""
- 納品プランExcelの **「SKU」列** と、ZAICO CSVの **「型番」列** を突き合わせます。
- ZAICO CSVで保管場所が **「{LOCATION_HONSHA}」** の数量を合計 → **「{NEW_COL_HONSHA}」列**
- ZAICO CSVで保管場所が **「{LOCATION_3F}」** の数量を合計 → **「{NEW_COL_3F}」列**
- 同じSKUがCSVに複数行ある場合は数量を合算します。
- ZAICO CSVに該当SKUが1件も無い場合は **0** として記載し、下に一覧で警告表示します（入力ミス確認用）。
- 「数量」列の右隣に2列追加して書き戻すだけで、既存の内容・書式は変更しません。
            """
        )

    # --- セッション初期化 ---
    if "results" not in st.session_state:
        st.session_state.results = None
        st.session_state.download_files = None
        st.session_state.zip_bytes = None

    # ------------------------------------------------------------
    st.subheader("① Amazon納品プランExcel をアップロード")
    excel_uploads = st.file_uploader(
        "納品プランExcel(.xlsx) を選択してください（複数まとめて処理できます）",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    st.subheader("② ZAICOダウンロードCSV をアップロード")
    csv_upload = st.file_uploader(
        "ZAICOからダウンロードしたCSV(.csv) を1つ選択してください",
        type=["csv"],
        accept_multiple_files=False,
    )

    st.subheader("③ 実行")
    ready = bool(excel_uploads) and csv_upload is not None
    run_button = st.button(
        "🚀 実行（ZAICO数量を転記してExcelを作成）",
        type="primary",
        disabled=not ready,
        use_container_width=True,
    )
    if not ready:
        st.caption("※ ①と②の両方をアップロードすると押せるようになります。")

    # ------------------------------------------------------------
    if run_button:
        excel_files = [(f.name, f.read()) for f in excel_uploads]
        csv_bytes = csv_upload.read()

        with st.status("処理を実行中です...", expanded=True) as status:
            status.write(f"ZAICO CSV: {csv_upload.name}")
            results, err = run_transcription(excel_files, csv_bytes)

            if err:
                status.update(label="処理を中止しました。", state="error")
                st.session_state.results = None
                st.session_state.download_files = None
                st.session_state.zip_bytes = None
                st.error("❌ " + err)
                return

            download_files = {}
            for r in results:
                if r.success:
                    status.write(
                        f"✅ {r.input_name} → {r.output_name}（SKU {r.total_sku}件"
                        + (f" / ZAICO未登録 {len(r.missing_skus)}件" if r.missing_skus else "")
                        + "）"
                    )
                    download_files[r.output_name] = r.output_bytes
                else:
                    status.write(f"❌ {r.input_name} の処理に失敗しました: {r.error}")

            st.session_state.results = results
            st.session_state.download_files = download_files
            st.session_state.zip_bytes = make_zip(download_files) if len(download_files) > 1 else None

            ok = sum(1 for r in results if r.success)
            ng = sum(1 for r in results if not r.success)
            if ng == 0:
                status.update(label=f"✅ 完了：成功 {ok}件", state="complete")
            else:
                status.update(label=f"⚠️ 完了：成功 {ok}件 / 失敗 {ng}件", state="error")

    # ------------------------------------------------------------
    results = st.session_state.get("results")
    if not results:
        st.info("①②をアップロードして「実行」を押すと、ここに結果とダウンロードボタンが表示されます。")
        return

    ok_results = [r for r in results if r.success]

    # --- サマリー ---
    st.subheader("結果サマリー")
    total_sku = sum(r.total_sku for r in ok_results)
    total_missing = sum(len(r.missing_skus) for r in ok_results)
    c1, c2, c3 = st.columns(3)
    c1.metric("処理ファイル数", f"{len(ok_results)} 件")
    c2.metric("合計SKU数", f"{total_sku} 件")
    c3.metric("ZAICO未登録SKU", f"{total_missing} 件")

    # --- ダウンロード ---
    st.subheader("④ ダウンロード")
    download_files = st.session_state.get("download_files") or {}
    if st.session_state.get("zip_bytes"):
        st.download_button(
            "📥 作成したExcelをまとめてダウンロード（ZIP）",
            data=st.session_state.zip_bytes,
            file_name=f"ZAICO数量記載_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )
    for fname, data in download_files.items():
        st.download_button(
            f"📥 {fname}",
            data=data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_{fname}",
        )

    # --- ファイルごとの詳細 ---
    st.subheader("転記内容の確認")
    st.caption(
        "※ この表は画面確認用です。ダウンロードするExcelは、元のデスクトップ版ツールと"
        "まったく同じ内容（「数量」列の右に「本社ZAICO / 3FZAICO」を追記しただけ）です。"
    )
    for r in results:
        with st.expander(
            f"{r.input_name} " + ("✅" if r.success else "❌"),
            expanded=(len(results) == 1),
        ):
            if not r.success:
                st.error(r.error)
                continue

            if r.missing_skus:
                st.warning(
                    "⚠️ 次のSKUはZAICO CSVに1件も見つからなかったため 0 として記載しました。"
                    "納品プラン側／ZAICO側の型番の表記ゆれがないかご確認ください。\n\n"
                    + "、".join(r.missing_skus)
                )

            df = pd.DataFrame(r.rows)
            if not df.empty:
                shortage = df["在庫判定"].astype(str).str.startswith("不足").sum()
                if shortage:
                    st.error(f"🔻 在庫合計が納品数量に満たないSKUが {int(shortage)} 件あります（下表「在庫判定」）。")

                def _highlight(row):
                    v = str(row.get("在庫判定", ""))
                    if v.startswith("不足"):
                        return ["background-color: #f8d7da"] * len(row)
                    if v == "ZAICO未登録":
                        return ["background-color: #fff3cd"] * len(row)
                    return [""] * len(row)

                st.dataframe(
                    df.style.apply(_highlight, axis=1),
                    use_container_width=True,
                    hide_index=True,
                )


if __name__ == "__main__":
    main()
