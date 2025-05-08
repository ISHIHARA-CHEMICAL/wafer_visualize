# streamlit_heatmap_app.py
"""Streamlit app: upload an Excel file (X, Y, Z columns),
    detect the first non‑empty cell automatically,
    draw a rainbow tricontour heatmap inside a radius n (inch),
    overlay measurement points & circle boundary,
    and list rows that were skipped with reasons.

Author: ChatGPT
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.tri as tri
import streamlit as st
from io import BytesIO
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------
# Streamlit page settings
# ------------------------------------------------------------
st.set_page_config(page_title="Excel Heatmap Viewer", layout="centered")

st.title("Excel Heatmap Viewer")
st.write(
    "テーブルの左上セルを **自動検出** して X, Y, Z を読み込み、"
    "半径 n inch 以内のヒートマップを表示します。"
)

# ------------------------------------------------------------
# Sidebar – user inputs
# ------------------------------------------------------------
sidebar = st.sidebar
sidebar.header("設定")

uploaded_file = sidebar.file_uploader("Excel ファイル (.xlsx) をアップロード", type=["xlsx"])

if uploaded_file:
    # read sheet names once to populate selectbox
    with BytesIO(uploaded_file.read()) as fh:
        xls = pd.ExcelFile(fh, engine="openpyxl")
        sheet_names = xls.sheet_names

    sheet_name = sidebar.selectbox("シートを選択", sheet_names, index=0)
    radius_inch: float = sidebar.number_input("半径 n (inch) (少し大きめに設定してください)", min_value=1.0, value=50.0)
    plot_btn = sidebar.button("ヒートマップを描画")
else:
    sidebar.info("まず Excel ファイルをアップロードしてください。")
    plot_btn = False

# ------------------------------------------------------------
# Helper – load & preprocess
# ------------------------------------------------------------

def load_and_prepare(df_raw: pd.DataFrame, radius_inch: float) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Return (data_ok, data_ng, meta) where
    * data_ok ― rows to plot (numeric, inside circle)
    * data_ng ― rows skipped with a '理由' column
    * meta ― dict with diagnostics (excel_cell, headers)
    """
    # ----- detect first non‑NA cell -----
    mask = df_raw.notna()
    if not mask.values.any():
        raise ValueError("シート内にデータが見つかりませんでした。")

    row0 = np.where(mask.any(axis=1))[0][0]
    col0 = np.where(mask.any(axis=0))[0][0]
    excel_cell = f"{get_column_letter(col0 + 1)}{row0 + 1}"

    # ----- assume three consecutive columns -----
    headers = df_raw.iloc[row0, col0 : col0 + 3].tolist()
    if len(headers) < 3:
        raise ValueError("連続した 3 列が見つからず、X, Y, Z の特定に失敗しました。")

    raw = df_raw.iloc[row0 + 1 :, col0 : col0 + 3].copy()
    raw.columns = headers
    raw.dropna(how="all", inplace=True)

    # ----- numeric conversion -----
    num = raw.apply(pd.to_numeric, errors="coerce")
    failed_cast = num.isna().any(axis=1)

    # ----- radius filter (only rows where cast succeeded) -----
    r = np.sqrt(num.loc[~failed_cast, headers[0]] ** 2 + num.loc[~failed_cast, headers[1]] ** 2)
    outside_circle = pd.Series(False, index=num.index)
    outside_circle.loc[~failed_cast] = r > radius_inch

    keep = (~failed_cast) & (~outside_circle)

    # ----- reason labels -----
    reason = np.select(
        [failed_cast, outside_circle],
        ["変換失敗 (非数値)", "円外"],
        default="描画対象",
    )
    num["理由"] = reason

    data_ok = num.loc[keep, headers]
    data_ng = num.loc[~keep, headers + ["理由"]]

    meta = {"excel_cell": excel_cell, "headers": headers}
    return data_ok, data_ng, meta

# ------------------------------------------------------------
# Main logic – triggered by button
# ------------------------------------------------------------
if plot_btn and uploaded_file:
    with st.spinner("読み込み & 描画中 ..."):
        # read selected sheet into DataFrame (header=None)
        with BytesIO(uploaded_file.getvalue()) as fh:
            df_raw = pd.read_excel(fh, sheet_name=sheet_name, header=None, engine="openpyxl")

        try:
            data_ok, data_ng, meta = load_and_prepare(df_raw, radius_inch)
        except Exception as e:
            st.error(str(e))
            st.stop()

        # unpack plotting data
        x = data_ok.iloc[:, 0].to_numpy()
        y = data_ok.iloc[:, 1].to_numpy()
        z = data_ok.iloc[:, 2].to_numpy()

        # ----- plotting -----
        fig, ax = plt.subplots(figsize=(6, 6))
        if len(x) >= 3:  # triangulation requires at least 3 points
            triang = tri.Triangulation(x, y)
            cont = ax.tricontourf(triang, z, levels=15, cmap="rainbow", antialiased=True)
        else:
            cont = ax.scatter(x, y, c=z, cmap="rainbow", s=40)

        # measurement points
        ax.plot(x, y, "k.", ms=4)
        # circle boundary
        circle = plt.Circle((0, 0), radius_inch, color="k", lw=2, fill=False)
        ax.add_patch(circle)

        # axis formatting
        ax.set_xlabel(meta["headers"][0], fontsize=14, fontweight="bold")
        ax.set_ylabel(meta["headers"][1], fontsize=14, fontweight="bold")
        ax.set_title(f"Heatmap (radius ≤ {radius_inch} inch)", fontsize=16, pad=12)
        ax.axis("equal")
        ax.set_xlim(-radius_inch, radius_inch)
        ax.set_ylim(-radius_inch, radius_inch)
        ticks = np.arange(-radius_inch, radius_inch + 1, radius_inch / 3)
        ax.set_xticks(ticks)
        ax.set_yticks(ticks)
        ax.grid(color="gray", linestyle="-", linewidth=1, alpha=0.5)

        # colorbar
        cbar = plt.colorbar(cont, ax=ax, pad=0.02)
        cbar.set_label(meta["headers"][2], fontsize=14, fontweight="bold")
        cbar.ax.tick_params(labelsize=12)

        st.pyplot(fig)

        # # diagnostics expander
        # with st.expander("内部情報 (デバッグ用)"):
        #     st.write(f"データ開始セル: **{meta['excel_cell']}**")
        #     st.write("検出した列ラベル:", meta["headers"])
        #     st.write(f"描画点数: {len(data_ok)} / 総行数: {len(df_raw)}")

        # skipped rows
        if not data_ng.empty:
            with st.expander("読み込めなかった点を表示"):
                st.write(f"合計 **{len(data_ng)}** 点が欠落しました。")
                st.dataframe(data_ng.reset_index(drop=True))
                csv = data_ng.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "CSV をダウンロード",
                    data=csv,
                    file_name="skipped_points.csv",
                    mime="text/csv",
                )
        else:
            st.success("すべての行を読み込めました 🎉")
else:
    st.info("左サイドバーでファイル・シート・半径を選んで **ヒートマップを描画** ボタンを押してください。")
