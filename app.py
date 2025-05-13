from __future__ import annotations

from io import BytesIO

import matplotlib.pyplot as plt
import matplotlib.tri as tri
import numpy as np
import pandas as pd
import streamlit as st
from matplotlib.ticker import FormatStrFormatter
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------
# Streamlit page settings
# ------------------------------------------------------------
st.set_page_config(page_title="Excel Heatmap Viewer", layout="centered")

# ------------------------------------------------------------
# グローバルフォント設定 (軸・カラーバーの文字を統一)
# ------------------------------------------------------------
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.size': 12,
})

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

uploaded_file = sidebar.file_uploader(
    "Excel ファイル (.xlsx) をアップロード", type=["xlsx"]
)

if uploaded_file:
    with BytesIO(uploaded_file.read()) as fh:
        xls = pd.ExcelFile(fh, engine="openpyxl")
        sheet_names = xls.sheet_names

    sheet_name = sidebar.selectbox("シートを選択", sheet_names, index=0)
    radius_inch: float = sidebar.number_input(
        "半径 n (inch)", min_value=1.0, value=7.0
    )
    margin_inch: float = sidebar.number_input(
        "余白 m (inch)", min_value=0.0, value=2.0
    )
    tick_step: float = sidebar.number_input(
        "目盛り間隔 d (inch, 0 で自動)", min_value=0.0, value=0.0
    )
    plot_btn = sidebar.button("ヒートマップを描画")
else:
    sidebar.info("まず Excel ファイルをアップロードしてください。")
    plot_btn = False

# ------------------------------------------------------------
# Helper – load & preprocess
# ------------------------------------------------------------
def load_and_prepare(
    df_raw: pd.DataFrame, radius_inch: float
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Return (data_ok, data_ng, meta)."""
    mask = df_raw.notna()
    if not mask.values.any():
        raise ValueError("シート内にデータが見つかりませんでした。")

    row0 = np.where(mask.any(axis=1))[0][0]
    col0 = np.where(mask.any(axis=0))[0][0]
    excel_cell = f"{get_column_letter(col0 + 1)}{row0 + 1}"

    headers = df_raw.iloc[row0, col0 : col0 + 3].tolist()
    if len(headers) < 3:
        raise ValueError("連続した 3 列が見つからず、X, Y, Z の特定に失敗しました。")

    raw = df_raw.iloc[row0 + 1 :, col0 : col0 + 3].copy()
    raw.columns = headers
    raw.dropna(how="all", inplace=True)

    num = raw.apply(pd.to_numeric, errors="coerce")
    failed_cast = num.isna().any(axis=1)

    r = np.sqrt(
        num.loc[~failed_cast, headers[0]] ** 2
        + num.loc[~failed_cast, headers[1]] ** 2
    )
    outside_circle = pd.Series(False, index=num.index)
    outside_circle.loc[~failed_cast] = r > radius_inch

    keep = (~failed_cast) & (~outside_circle)
    reason = np.select(
        [failed_cast, outside_circle],
        ["変換失敗 (非数値)", "円外"],
        default="描画対象",
    )
    num["理由"] = reason

    data_ok = num.loc[keep, headers]
    data_ng = num.loc[~keep, headers + ["理由"]]

    return data_ok, data_ng, {"excel_cell": excel_cell, "headers": headers}

# ------------------------------------------------------------
# Main logic – triggered by button
# ------------------------------------------------------------
if plot_btn and uploaded_file:
    with st.spinner("読み込み & 描画中 ..."):
        with BytesIO(uploaded_file.getvalue()) as fh:
            df_raw = pd.read_excel(
                fh, sheet_name=sheet_name, header=None, engine="openpyxl"
            )

        try:
            data_ok, data_ng, meta = load_and_prepare(df_raw, radius_inch)
        except Exception as e:
            st.error(str(e))
            st.stop()

        # ----- plotting data -----
        x = data_ok.iloc[:, 0].to_numpy()
        y = data_ok.iloc[:, 1].to_numpy()
        z = data_ok.iloc[:, 2].to_numpy()

        # ① 小数第3位で四捨五入 → 第2位までに
        z = np.round(z, 2)

        # ② z_min/z_max を取得し、等高線レベルを明示的に作成
        z_min, z_max = z.min(), z.max()
        num_levels = 15
        levels = np.linspace(z_min, z_max, num_levels)

        # ----- create figure -----
        fig, ax = plt.subplots(figsize=(6, 6))
        fig.subplots_adjust(right=0.85)

        # ----- draw heatmap or scatter -----
        ax.set_axisbelow(True)
        if len(x) >= 3:
            triang = tri.Triangulation(x, y)
            cont = ax.tricontourf(
                triang,
                z,
                levels=levels,
                cmap="gist_rainbow",
                antialiased=True,
                zorder=1
            )
        else:
            cont = ax.scatter(
                x,
                y,
                c=z,
                cmap="gist_rainbow",
                s=40,
                zorder=1,
                vmin=z_min,
                vmax=z_max
            )

        # グリッドを不透明で下に描画
        ax.grid(color="gray", linestyle="-", linewidth=1, alpha=1.0, zorder=0)

        # measurement points & circle
        ax.plot(x, y, "k.", ms=4, zorder=2)
        ax.add_patch(plt.Circle((0, 0), radius_inch, color="k", lw=2, fill=False, zorder=2))

        # ----- range & ticks -----
        plot_range = radius_inch + margin_inch
        ax.set_xlim(-plot_range, plot_range)
        ax.set_ylim(-plot_range, plot_range)
        ax.set_aspect("equal", adjustable="box")

        if tick_step > 0:
            ticks = np.arange(-plot_range, plot_range + tick_step, tick_step)
        else:
            ticks = np.linspace(-plot_range, plot_range, 7)
        ax.set_xticks(ticks)
        ax.set_yticks(ticks)
        ax.tick_params(
            axis='both',
            which='both',
            direction='in',
            length=6
        )

        # labels
        ax.set_xlabel(meta["headers"][0], fontsize=14, fontweight="bold")
        ax.set_ylabel(meta["headers"][1], fontsize=14, fontweight="bold")

        # ----- colorbar -----
        # ③ ２色ごとに(levels[::2])目盛りを表示
        cbar = fig.colorbar(
            cont,
            ax=ax,
            ticks=levels[::2],
            format='%.2f',
            fraction=0.05,
            pad=0.02
        )
        cbar.ax.tick_params(direction='in', length=6, pad=4)
        cont.set_clim(z_min, z_max)

        # ---- show in Streamlit ----
        st.pyplot(fig, bbox_inches="tight", use_container_width=True)

        # ----- skipped rows -----
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
    st.info(
        "左サイドバーでファイル・シート・半径・余白・目盛り間隔を選んで "
        "**ヒートマップを描画** ボタンを押してください。"
    )
