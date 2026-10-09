"""
app.py
======
MSI Surfaces - Planogram (POG) AI Audit Intelligence Platform
Streamlit Web Application with Power BI / Notion interactive dashboards.
"""

import os
import sys
import time
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image

# Import Core Audit Engine
from core_engine import AUDIT_MODELS, run_audit_pipeline

# Configure Streamlit page
st.set_page_config(
    page_title="POG AI Audit Platform | MSI Surfaces",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load custom MSI Surfaces CSS
CSS_PATH = Path(__file__).parent / "assets" / "style.css"
if CSS_PATH.exists():
    with open(CSS_PATH, "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Top Announcement Bar
st.markdown(
    """
    <div class="msi-announcement-bar">
      <span>MSI SURFACES ARCHITECTURAL QUALITY ASSURANCE & COMPLIANCE</span>
      <span>PLANOGRAM (POG) VISION AI v2.4</span>
    </div>
    """,
    unsafe_allow_html=True
)

# Hero Header
st.markdown(
    """
    <div class="msi-header-hero">
      <div>
        <h1 class="msi-header-title">Planogram AI Compliance Hub</h1>
        <div class="msi-header-subtitle">Automated Visual Audits for Retail Displays & Planogram Standards</div>
      </div>
      <div style="text-align: right; font-size: 12px; color: #6c757d; font-weight: 600; text-transform: uppercase;">
        Design System: MSI Surfaces Luxury 0px Architecture
      </div>
    </div>
    """,
    unsafe_allow_html=True
)

# Sidebar Controls
with st.sidebar:
    st.markdown("### 🏛️ AUDIT CONFIGURATION")
    st.markdown("---")

    audit_model_choice = st.selectbox(
        "SELECT AUDIT SPECIFICATION",
        options=list(AUDIT_MODELS.keys()),
        index=0,
        help="Select the trained vision model for this audit batch."
    )

    selected_cfg = AUDIT_MODELS[audit_model_choice]

    st.markdown(
        f"""
        <div style="background-color: #f8f9fa; border-left: 3px solid #8b6f4e; padding: 12px; margin-top: 10px; margin-bottom: 20px;">
          <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; color: #6c757d;">CALIBRATED STANDARD</div>
          <div style="font-size: 13px; font-weight: 600; color: #212529; margin-top: 4px;">{selected_cfg['name']}</div>
          <div style="font-size: 12px; color: #555; margin-top: 4px;">Target: <b>{selected_cfg['target_desc']}</b></div>
          <div style="font-size: 12px; color: #8b6f4e; font-weight: 700; margin-top: 6px;">Threshold: {selected_cfg['conf']:.3f} (Locked Standard)</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### 📋 WORKFLOW INSTRUCTIONS")
    st.markdown(
        """
        1. **Upload Audit Sheet**: Drag and drop your `.xlsx` or `.xlsm` report containing photo links.
        2. **Process Audit**: The engine downloads photos in parallel and runs YOLO deep vision inference.
        3. **Explore Dashboard**: Review Power BI analytics, store rankings, and side-by-side photo detections.
        4. **Export**: Download the audited Excel sheet and annotated bounding box package.
        """
    )
    st.markdown("---")
    st.caption("MSI Surfaces Quality & Merchandising Compliance")

# Main Container
if "audit_results" not in st.session_state:
    st.session_state.audit_results = None

tab_upload, tab_dashboard, tab_inspector = st.tabs([
    "📥 1. UPLOAD & RUN AUDIT",
    "📊 2. POWER BI / NOTION DASHBOARD",
    "🔍 3. VISUAL PHOTO INSPECTOR"
])

# TAB 1: UPLOAD & AUDIT LAUNCH
with tab_upload:
    st.markdown('<div class="notion-card-title">📤 PLANOGRAM AUDIT FILE UPLOAD</div>', unsafe_allow_html=True)
    st.markdown("Upload the retail execution Excel file containing bay photos or URLs.")

    uploaded_file = st.file_uploader(
        "Choose Planogram Audit Excel Spreadsheet",
        type=["xlsx", "xlsm"],
        help="Upload files such as Project - 1.xlsx"
    )

    if uploaded_file:
        file_details = {
            "Filename": uploaded_file.name,
            "File Size": f"{uploaded_file.size / (1024 * 1024):.2f} MB",
            "Selected Model": selected_cfg["name"],
            "Calibrated Threshold": f"{selected_cfg['conf']:.3f}"
        }
        st.write("📁 **File Loaded:**", file_details["Filename"], f"({file_details['File Size']})")

        start_btn = st.button("🚀 EXECUTE AI PLANOGRAM AUDIT", use_container_width=True)

        if start_btn:
            status_box = st.empty()
            progress_bar = st.progress(0.0)

            def update_status(text):
                status_box.info(f"⏳ **{text}**")

            def update_progress(val):
                progress_bar.progress(min(max(float(val), 0.0), 1.0))

            try:
                base_out = Path(__file__).parent / "audit_runs"
                base_out.mkdir(parents=True, exist_ok=True)

                t0 = time.time()
                file_bytes = uploaded_file.getvalue()

                results = run_audit_pipeline(
                    excel_file_bytes=file_bytes,
                    excel_filename=uploaded_file.name,
                    model_choice=audit_model_choice,
                    output_base_dir=base_out,
                    status_callback=update_status,
                    progress_callback=update_progress
                )
                duration = time.time() - t0

                st.session_state.audit_results = results
                status_box.success(f"✅ **Audit completed successfully in {duration:.1f} seconds!** Navigate to the **Power BI Dashboard** tab to view analytics.")
                progress_bar.progress(1.0)
            except Exception as e:
                status_box.error(f"❌ **Audit encountered an error:** {str(e)}")
                st.exception(e)

# TAB 2: POWER BI / NOTION DASHBOARD
with tab_dashboard:
    res = st.session_state.audit_results

    if not res:
        st.info("ℹ️ No audit has been executed yet. Please upload an Excel sheet in Tab 1 to generate analytics.")
    else:
        st.markdown(
            f"""
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
              <div>
                <h2 style="margin: 0; font-size: 24px;">EXECUTIVE AUDIT INTELLIGENCE</h2>
                <div style="font-size: 13px; color: #8b6f4e; font-weight: 600;">MODEL: {res['model_name']} | CALIBRATED CONFIDENCE: {res['confidence_threshold']:.3f}</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # Power BI Style KPI Metric Cards
        col1, col2, col3, col4, col5 = st.columns(5)

        with col1:
            st.markdown(
                f"""
                <div class="msi-kpi-card">
                  <div class="msi-kpi-label">TOTAL BAYS AUDITED</div>
                  <div class="msi-kpi-value">{res['total_rows']}</div>
                  <div class="msi-kpi-sub">100% Ingested</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col2:
            st.markdown(
                f"""
                <div class="msi-kpi-card" style="border-top-color: #2e7d32;">
                  <div class="msi-kpi-label">COMPLIANCE PASS RATE</div>
                  <div class="msi-kpi-value" style="color: #2e7d32;">{res['pass_rate']:.1f}%</div>
                  <div class="msi-kpi-sub">{res['approved_count']} Approved Bays</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col3:
            st.markdown(
                f"""
                <div class="msi-kpi-card" style="border-top-color: #c62828;">
                  <div class="msi-kpi-label">REJECTED BAYS</div>
                  <div class="msi-kpi-value" style="color: #c62828;">{res['rejected_count']}</div>
                  <div class="msi-kpi-sub">{(res['rejected_count'] / res['total_rows'] * 100) if res['total_rows'] else 0:.1f}% Non-Compliant</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col4:
            st.markdown(
                f"""
                <div class="msi-kpi-card" style="border-top-color: #7d6245;">
                  <div class="msi-kpi-label">STORE LOCATIONS</div>
                  <div class="msi-kpi-value">{len(res['store_summary'])}</div>
                  <div class="msi-kpi-sub">Distinct Locations</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col5:
            st.markdown(
                f"""
                <div class="msi-kpi-card" style="border-top-color: #616161;">
                  <div class="msi-kpi-label">ERRORS / UNREACHABLE</div>
                  <div class="msi-kpi-value" style="color: #616161;">{res['error_count']}</div>
                  <div class="msi-kpi-sub">Missing Image Links</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # Interactive Power BI Visualizations Row
        c_left, c_right = st.columns([1, 1.4])

        with c_left:
            st.markdown('<div class="notion-card-title">🍩 COMPLIANCE VERDICT DISTRIBUTION</div>', unsafe_allow_html=True)
            verdict_counts = {
                "Approved": res["approved_count"],
                "Rejected": res["rejected_count"],
                "Error": res["error_count"]
            }
            labels = [k for k, v in verdict_counts.items() if v > 0]
            values = [v for k, v in verdict_counts.items() if v > 0]
            colors_map = {
                "Approved": "#8b6f4e",
                "Rejected": "#212529",
                "Error": "#dee2e6"
            }

            fig_donut = go.Figure(data=[go.Pie(
                labels=labels,
                values=values,
                hole=0.55,
                marker=dict(colors=[colors_map.get(l, "#8b6f4e") for l in labels]),
                textinfo="percent+label",
                hoverinfo="label+value+percent",
                insidetextorientation="horizontal"
            )])
            fig_donut.update_layout(
                showlegend=True,
                margin=dict(t=10, b=10, l=10, r=10),
                height=320,
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="Open Sans", color="#212529")
            )
            st.plotly_chart(fig_donut, use_container_width=True)

        with c_right:
            st.markdown('<div class="notion-card-title">🏬 STORE COMPLIANCE RANKINGS</div>', unsafe_allow_html=True)

            # Build store dataframe
            st_data_list = []
            for s_id, s_info in res["store_summary"].items():
                pass_pct = (s_info["approved"] / s_info["total"] * 100) if s_info["total"] > 0 else 0
                st_data_list.append({
                    "Store": str(s_id),
                    "Total": s_info["total"],
                    "Approved": s_info["approved"],
                    "Rejected": s_info["rejected"],
                    "Pass Rate (%)": round(pass_pct, 1)
                })

            df_stores = pd.DataFrame(st_data_list).sort_values(by="Pass Rate (%)", ascending=True)

            fig_bar = px.bar(
                df_stores.head(15),
                x="Pass Rate (%)",
                y="Store",
                orientation="h",
                text="Pass Rate (%)",
                color="Pass Rate (%)",
                color_continuous_scale=[[0, "#212529"], [0.5, "#a88964"], [1, "#8b6f4e"]],
                labels={"Pass Rate (%)": "Compliance Rate (%)"}
            )
            fig_bar.update_layout(
                height=320,
                margin=dict(t=10, b=10, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                coloraxis_showscale=False,
                font=dict(family="Open Sans", color="#212529"),
                xaxis=dict(range=[0, 105], showgrid=True, gridcolor="#efefef"),
                yaxis=dict(type='category', showgrid=False)
            )
            fig_bar.update_traces(textposition="outside")
            st.plotly_chart(fig_bar, use_container_width=True)

        # Notion-Style Interactive Database Explorer
        st.markdown('<div class="notion-card-title">🗂️ NOTION-STYLE AUDIT DATABASE</div>', unsafe_allow_html=True)

        df_rows = pd.DataFrame(res["row_results"])

        # Filters row
        f1, f2, f3 = st.columns([1, 1, 2])
        with f1:
            filter_verdict = st.multiselect(
                "Filter Verdict",
                options=["All"] + sorted(list(df_rows["verdict"].unique())),
                default=["All"]
            )
        with f2:
            filter_store = st.multiselect(
                "Filter Store",
                options=["All"] + sorted(list(df_rows["store"].unique())),
                default=["All"]
            )
        with f3:
            search_query = st.text_input("Search Bay ID or Detected Class", placeholder="e.g. Bay 2, open...")

        filtered_df = df_rows.copy()
        if "All" not in filter_verdict:
            filtered_df = filtered_df[filtered_df["verdict"].isin(filter_verdict)]
        if "All" not in filter_store:
            filtered_df = filtered_df[filtered_df["store"].isin(filter_store)]
        if search_query:
            q = search_query.lower()
            filtered_df = filtered_df[
                filtered_df["bay"].str.lower().str.contains(q) |
                filtered_df["format"].str.lower().str.contains(q)
            ]

        # Display table
        display_cols = ["row_idx", "store", "bay", "verdict", "confidence", "format", "boxes", "latency_ms"]
        st.dataframe(
            filtered_df[display_cols].rename(columns={
                "row_idx": "Row",
                "store": "Store / Location",
                "bay": "Bay ID",
                "verdict": "AI Verdict",
                "confidence": "Max Conf",
                "format": "Detected Class",
                "boxes": "Detections",
                "latency_ms": "Latency (ms)"
            }),
            use_container_width=True,
            height=340
        )

        st.markdown("<br>", unsafe_allow_html=True)

        # Export & Download Center
        st.markdown('<div class="notion-card-title">📦 AUDIT EXPORT & DELIVERABLES</div>', unsafe_allow_html=True)
        d_col1, d_col2 = st.columns(2)

        with d_col1:
            if Path(res["output_excel_path"]).exists():
                with open(res["output_excel_path"], "rb") as f_excel:
                    st.download_button(
                        label="📥 DOWNLOAD AUDITED EXCEL REPORT (.XLSX)",
                        data=f_excel.read(),
                        file_name=Path(res["output_excel_path"]).name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        with d_col2:
            if res.get("zip_buffer"):
                st.download_button(
                    label="📦 DOWNLOAD ANNOTATED IMAGES (.ZIP)",
                    data=res["zip_buffer"],
                    file_name=f"{Path(res['output_excel_path']).stem}_Annotated_Images.zip",
                    mime="application/zip",
                    use_container_width=True
                )

# TAB 3: VISUAL PHOTO INSPECTOR
with tab_inspector:
    res = st.session_state.audit_results
    if not res:
        st.info("ℹ️ Execute an audit first in Tab 1 to inspect photos.")
    else:
        st.markdown('<div class="notion-card-title">🔍 SIDE-BY-SIDE VISUAL PHOTO INSPECTOR</div>', unsafe_allow_html=True)
        st.caption("Select any audited row from the dropdown below to compare the original photo against the AI bounding box detection.")

        row_options = [
            f"Row {r['row_idx']} | Store {r['store']} | Bay {r['bay']} | [{r['verdict']}]"
            for r in res["row_results"]
        ]

        selected_row_label = st.selectbox("SELECT ITEM FOR VISUAL INSPECTION", options=row_options, index=0)
        selected_idx = int(selected_row_label.split(" ")[1])

        item = next((r for r in res["row_results"] if r["row_idx"] == selected_idx), None)

        if item:
            # Metadata banner
            st.markdown(
                f"""
                <div style="background-color: #ffffff; border: 1px solid #dee2e6; border-left: 4px solid #8b6f4e; padding: 14px; margin-bottom: 20px;">
                  <span style="font-weight: 700; color: #212529;">STORE:</span> {item['store']} &nbsp;|&nbsp;
                  <span style="font-weight: 700; color: #212529;">BAY:</span> {item['bay']} &nbsp;|&nbsp;
                  <span style="font-weight: 700; color: #212529;">VERDICT:</span> <b>{item['verdict']}</b> &nbsp;|&nbsp;
                  <span style="font-weight: 700; color: #212529;">CONFIDENCE:</span> {item['confidence']:.3f} &nbsp;|&nbsp;
                  <span style="font-weight: 700; color: #212529;">DETECTED:</span> {item['format']} ({item['boxes']} boxes)
                </div>
                """,
                unsafe_allow_html=True
            )

            p_col1, p_col2 = st.columns(2)

            with p_col1:
                st.markdown("**ORIGINAL SUBMITTED PHOTO**")
                if item["original_path"] and Path(item["original_path"]).exists():
                    st.image(item["original_path"], use_container_width=True)
                else:
                    st.warning("Original photo not available or download failed.")

            with p_col2:
                st.markdown("**AI BOUNDING BOX PREDICTION**")
                if item["annotated_path"] and Path(item["annotated_path"]).exists():
                    st.image(item["annotated_path"], use_container_width=True)
                else:
                    st.info("No detections or photo unavailable.")
