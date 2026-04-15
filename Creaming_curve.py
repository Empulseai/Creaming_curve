import os
import random
from io import BytesIO
from datetime import datetime, date

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

from openpyxl.styles import PatternFill, Font
from pptx import Presentation
from pptx.util import Inches


# --------------------------------------------------
# PAGE STYLE
# --------------------------------------------------
st.set_page_config(page_title="EI Matrix Tool", layout="wide")

st.markdown(
    """
    <style>
        body {
            background-color: #f4f4f4;
            font-family: 'Helvetica Neue', sans-serif;
        }
        .stApp {
            background-color: #f4f4f4;
        }
        h1 {
            color: #0078D7;
            text-align: center;
            font-size: 36px;
            font-family: 'Georgia', serif;
        }
        h4 {
            text-align: center;
            color: #444;
            font-family: 'Verdana', sans-serif;
        }
        .stDataFrame {
            background-color: white;
            border-radius: 10px;
            padding: 10px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def generate_random_color():
    return (
        random.randint(128, 255),
        random.randint(128, 255),
        random.randint(128, 255),
    )


def rgb_to_hex(rgb):
    return "#{:02x}{:02x}{:02x}".format(*rgb)


def prepare_ei_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    required_cols = ["Category", "Ideas", "Effort", "Impact"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    df = df.copy()
    df["Count of the category"] = df["Category"].map(df["Category"].value_counts())
    df = df.sort_values(by=["Count of the category", "Category"], ascending=False).reset_index(drop=True)
    df["Sl No"] = range(1, len(df) + 1)
    df["Effort1"] = np.where(df["Effort"].astype(str).str.strip().str.lower() == "high", 1.5, 0.5)
    df["Impact1"] = np.where(df["Impact"].astype(str).str.strip().str.lower() == "high", 1.5, 0.5)
    return df


def plot_cluster_cards(df: pd.DataFrame, title: str, color_mapping: dict):
    fig, ax = plt.subplots(figsize=(16, max(5, len(df) * 0.5)))
    ax.axis("off")
    ax.set_title(title, fontsize=14, fontweight="bold")

    if df.empty:
        ax.text(0.5, 0.5, "No data available", ha="center", va="center", fontsize=12)
        st.pyplot(fig)
        return

    y = len(df)
    for _, row in df.iterrows():
        category = str(row["Category"])
        idea = str(row["Ideas"])
        sl_no = row["Sl No"]
        box_color = rgb_to_hex(color_mapping[category])

        text = f"{sl_no}\n{category}\n\n{idea}"
        ax.text(
            0.02,
            y,
            text,
            fontsize=9,
            va="top",
            ha="left",
            bbox=dict(boxstyle="round,pad=0.4", facecolor=box_color, edgecolor="black"),
        )
        y -= 1

    ax.set_xlim(0, 1)
    ax.set_ylim(0, len(df) + 1)
    st.pyplot(fig)


def plot_ei_matrix(df: pd.DataFrame, color_mapping: dict):
    fig, ax = plt.subplots(figsize=(14, 10))

    ax.set_facecolor((242 / 255, 242 / 255, 242 / 255))
    ax.set_title("EI MATRIX", fontsize=16, fontweight="bold")
    ax.set_xlabel("EFFORT", fontsize=12, fontweight="bold")
    ax.set_ylabel("IMPACT", fontsize=12, fontweight="bold")

    ax.set_xlim(0, 2)
    ax.set_ylim(0, 2)

    ax.set_xticks(np.arange(0, 2.1, 0.5))
    ax.set_xticklabels([" ", "Low", " ", "High", " "])

    ax.set_yticks(np.arange(0, 2.1, 0.5))
    ax.set_yticklabels([" ", "Low", " ", "High", " "])

    ax.axvline(x=1, color="black", linewidth=1.5, linestyle="--")
    ax.axhline(y=1, color="black", linewidth=1.5, linestyle="--")

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_linewidth(2)
    ax.spines["bottom"].set_linewidth(2)

    quadrant_labels = [
        ("Quick Wins", 0.5, 1.5),
        ("Major Projects", 1.5, 1.5),
        ("Fill Ins", 0.5, 0.5),
        ("Time Wasters", 1.5, 0.5),
    ]

    for label, x, y in quadrant_labels:
        ax.text(x, y + 0.38, label, ha="center", va="center", fontsize=12, fontweight="bold")

    # Spread labels a bit inside each quadrant
    quadrant_offsets = {
        (0.5, 1.5): [],
        (0.5, 0.5): [],
        (1.5, 1.5): [],
        (1.5, 0.5): [],
    }

    for _, row in df.iterrows():
        x = row["Effort1"]
        y = row["Impact1"]
        quadrant_offsets[(x, y)].append(row)

    for (x, y), rows in quadrant_offsets.items():
        total = len(rows)
        if total == 0:
            continue

        for idx, row in enumerate(rows):
            dx = ((idx % 4) - 1.5) * 0.18
            dy = -((idx // 4) * 0.12)
            color = rgb_to_hex(color_mapping[str(row["Category"])])
            label = f'{row["Sl No"]}: {row["Category"]}'

            ax.text(
                x + dx,
                y + dy,
                label,
                fontsize=8,
                ha="center",
                va="center",
                bbox=dict(boxstyle="round,pad=0.25", facecolor=color, edgecolor="black"),
            )

    return fig


def create_ei_excel(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()

    quick_wins = df[(df["Effort1"] == 0.5) & (df["Impact1"] == 1.5)].copy()
    fill_ins = df[(df["Effort1"] == 0.5) & (df["Impact1"] == 0.5)].copy()
    major_projects = df[(df["Effort1"] == 1.5) & (df["Impact1"] == 1.5)].copy()
    time_wasters = df[(df["Effort1"] == 1.5) & (df["Impact1"] == 0.5)].copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="All Ideas")
        quick_wins.to_excel(writer, index=False, sheet_name="Quick Wins")
        fill_ins.to_excel(writer, index=False, sheet_name="Fill Ins")
        major_projects.to_excel(writer, index=False, sheet_name="Major Projects")
        time_wasters.to_excel(writer, index=False, sheet_name="Time Wasters")

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            header_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font

            for column_cells in ws.columns:
                max_length = 0
                col_letter = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

    output.seek(0)
    return output


def create_creaming_curve_excel(df: pd.DataFrame, budget: float) -> BytesIO:
    excel_stream = BytesIO()

    with pd.ExcelWriter(excel_stream, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DataFrame")
        worksheet = writer.sheets["DataFrame"]

        fill_blue = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
        font_white_bold = Font(color="FFFFFF", bold=True)

        font_green = Font(color="006400", bold=False)
        font_red = Font(color="8B0000", bold=False)

        for cell in worksheet[1]:
            cell.fill = fill_blue
            cell.font = font_white_bold

        for row_num in range(2, len(df) + 2):
            is_within_budget = df.iloc[row_num - 2]["Cumulative cost"] <= budget
            row_font = font_green if is_within_budget else font_red

            for cell in worksheet[row_num]:
                cell.font = row_font

        for column_cells in worksheet.columns:
            max_length = 0
            col_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            worksheet.column_dimensions[col_letter].width = min(max_length + 2, 35)

    excel_stream.seek(0)
    return excel_stream


def create_creaming_curve_ppt(fig, budget: float) -> BytesIO:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    if slide.shapes.title:
        slide.shapes.title.text = f"Creaming Curve with Budget ${budget:,.0f}"

    img_stream = BytesIO()
    fig.savefig(img_stream, format="png", bbox_inches="tight")
    img_stream.seek(0)

    slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.2), width=Inches(9))

    ppt_stream = BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream


# --------------------------------------------------
# EI MATRIX SECTION
# --------------------------------------------------
st.title("EI MATRIX TOOL")

file_upload1 = st.file_uploader(
    "📂 Upload the Excel file for EI matrix",
    type=["xlsx"],
    accept_multiple_files=False,
    key="ei_upload",
)

if file_upload1 is not None:
    try:
        df_ei = pd.read_excel(file_upload1, sheet_name="Consolidated Ideas")
        df_ei = prepare_ei_dataframe(df_ei)

        unique_categories = df_ei["Category"].unique()
        color_mapping = {category: generate_random_color() for category in unique_categories}

        st.subheader("Processed Data")
        st.data_editor(df_ei, use_container_width=True)

        st.subheader("Quadrant Summary")

        quick_wins = df_ei[(df_ei["Effort1"] == 0.5) & (df_ei["Impact1"] == 1.5)].copy()
        fill_ins = df_ei[(df_ei["Effort1"] == 0.5) & (df_ei["Impact1"] == 0.5)].copy()
        major_projects = df_ei[(df_ei["Effort1"] == 1.5) & (df_ei["Impact1"] == 1.5)].copy()
        time_wasters = df_ei[(df_ei["Effort1"] == 1.5) & (df_ei["Impact1"] == 0.5)].copy()

        col1, col2 = st.columns(2)
        with col1:
            plot_cluster_cards(quick_wins, "Quick Wins", color_mapping)
            plot_cluster_cards(fill_ins, "Fill Ins", color_mapping)
        with col2:
            plot_cluster_cards(major_projects, "Major Projects", color_mapping)
            plot_cluster_cards(time_wasters, "Time Wasters", color_mapping)

        st.subheader("EI Matrix")
        ei_fig = plot_ei_matrix(df_ei, color_mapping)
        st.pyplot(ei_fig)

        # Download EI matrix image
        ei_img = BytesIO()
        ei_fig.savefig(ei_img, format="png", bbox_inches="tight")
        ei_img.seek(0)

        st.download_button(
            "📥 Download EI Matrix Image",
            ei_img,
            "ei_matrix.png",
            "image/png",
        )

        # Download EI excel
        ei_excel = create_ei_excel(df_ei)
        st.download_button(
            "📥 Download EI Matrix Excel",
            ei_excel,
            "ei_matrix_output.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error while processing EI Matrix file: {e}")


# --------------------------------------------------
# CREAMING CURVE SECTION
# --------------------------------------------------
st.title("Creaming Curve Analyzer")

file = st.file_uploader(
    "📂 Upload Excel file for Creaming curve",
    type=["xlsx"],
    help="Ensure the file contains relevant cost and savings data",
    key="creaming_curve_upload",
)

if file is not None:
    try:
        df = pd.read_excel(file)

        required_cols = ["Cost $", "Annual Savings $ K", "Project Summary Name"]
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            st.error(f"⚠️ Missing required columns: {', '.join(missing_cols)}")
        else:
            df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

            # Avoid division by zero
            df["Cost $"] = pd.to_numeric(df["Cost $"], errors="coerce").fillna(0)
            df["Annual Savings $ K"] = pd.to_numeric(df["Annual Savings $ K"], errors="coerce").fillna(0)

            df["Savings ratio"] = np.where(
                df["Cost $"] != 0,
                df["Annual Savings $ K"] / df["Cost $"],
                0,
            )

            df = df.sort_values(by="Savings ratio", ascending=False).reset_index(drop=True)
            df["Cumulative cost"] = df["Cost $"].cumsum()
            df["Cumulative Savings"] = df["Annual Savings $ K"].cumsum()

            st.markdown("### Updated DataFrame")
            st.dataframe(
                df.style.format(
                    {"Cost $": "${:,.2f}", "Annual Savings $ K": "${:,.2f}"}
                ),
                use_container_width=True,
            )

            budget = st.number_input("💰 Enter your budget ($ K):", min_value=0.0, step=100.0)

            fig, ax = plt.subplots(figsize=(12, 6))
            fig.patch.set_facecolor("#f4f4f4")

            if budget > 0:
                within_budget = df["Cumulative cost"] <= budget

                ax.scatter(
                    df.loc[within_budget, "Cumulative cost"],
                    df.loc[within_budget, "Cumulative Savings"],
                    edgecolors="black",
                    s=100,
                    alpha=0.8,
                    label="Projects within Budget",
                )
                ax.scatter(
                    df.loc[~within_budget, "Cumulative cost"],
                    df.loc[~within_budget, "Cumulative Savings"],
                    edgecolors="black",
                    s=100,
                    alpha=0.8,
                    label="Projects outside Budget",
                )
                ax.axvline(
                    x=budget,
                    linestyle="--",
                    linewidth=2,
                    label=f"Budget: ${budget:,.0f} K",
                )
            else:
                ax.scatter(
                    df["Cumulative cost"],
                    df["Cumulative Savings"],
                    edgecolors="black",
                    s=100,
                    alpha=0.8,
                    label="Projects",
                )

            ax.set_title("Cumulative Cost vs. Cumulative Savings", fontsize=14, fontweight="bold")
            ax.set_xlabel("Cumulative Cost ($ K)", fontsize=10, fontweight="bold")
            ax.set_ylabel("Cumulative Savings ($ K)", fontsize=10, fontweight="bold")
            ax.grid(True, linestyle="--", alpha=0.5)

            x_labels = [
                f"{name} (${cost:,.0f})"
                for name, cost in zip(df["Project Summary Name"], df["Cumulative cost"])
            ]
            ax.set_xticks(df["Cumulative cost"])
            ax.set_xticklabels(x_labels, rotation=90, ha="center", fontsize=8)

            ax.legend(loc="upper left", frameon=True)

            st.pyplot(fig)

            ppt_stream = create_creaming_curve_ppt(fig, budget)
            st.download_button(
                "📥 Download PowerPoint",
                ppt_stream,
                "creaming_curve_presentation.pptx",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

            excel_stream = create_creaming_curve_excel(df, budget)
            st.download_button(
                "📥 Download Excel",
                excel_stream,
                "creaming_curve_data.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.success("✅ Analysis Complete!")

    except Exception as e:
        st.error(f"Error while processing Creaming Curve file: {e}")
