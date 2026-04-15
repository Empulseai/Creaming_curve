from io import BytesIO

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

from openpyxl.styles import PatternFill, Font
from pptx import Presentation
from pptx.util import Inches


# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Creaming Curve Analyzer", layout="wide")

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
# CREAMING CURVE SECTION
# --------------------------------------------------
st.title("Creaming Curve Analyzer")

file = st.file_uploader(
    "📂 Upload Excel file for Creaming curve",
    type=["xlsx"],
    help="Ensure the file contains relevant cost and savings data",
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

            budget = st.number_input(
                "💰 Enter your budget ($ K):",
                min_value=0.0,
                step=100.0
            )

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
