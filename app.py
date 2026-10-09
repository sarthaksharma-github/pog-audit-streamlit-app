"""
app.py
======
Planogram (POG) AI Audit Intelligence Platform
Streamlit Web Application featuring Anthropic typography styling,
clean architectural palette, and interactive dashboards.
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
    page_title="Planogram AI Compliance Hub",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load custom Anthropic Typography CSS
CSS_PATH = Path(__file__).parent / "assets" / "style.css"
if CSS_PATH.exists():
    with open(CSS_PATH, "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Top Announcement Bar
st.markdown(
    """
    <div class="msi-announcement-bar">
      <span>Quality Assurance & Planogram Compliance</span>
      <span>Planogram Vision AI</span>
    </div>
    """,
    unsafe_allow_html=True
)

# Hero Header (Removed "Design System : MSI Surfaces...")
st.markdown(
    """
    <div class="msi-header-hero">
      <div>
        <h1 class="msi-header-title">Planogram AI Compliance Hub</h1>
        <div class="msi-header-subtitle">Automated visual compliance audits for retail displays & planogram standards</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True
)

# Sidebar Controls
with st.sidebar:
    st.markdown("### Audit Configuration")
    st.markdown("---")

    audit_model_choice = st.selectbox(
        "Select Audit Specification",
        options=list(AUDIT_MODELS.keys()),
        index=0,
        help="Select the trained vision model for this audit batch."
    )

    selected_cfg = AUDIT_MODELS[audit_model_choice]

    st.markdown(
        f"""
        <div style="background-color: #faf9f6; border-left: 2px solid #8b6f4e; padding: 14px; margin-top: 12px; margin-bottom: 24px;">
          <div style="font-size: 11px; font-weight: 600; text-transform: uppercase; color: #737373; letter-spacing: 0.04em;">Calibrated Standard</div>
          <div style="font-size: 14px; font-weight: 600; color: #191919; margin-top: 4px;">{selected_cfg['name']}</div>
          <div style="font-size: 12px; color: #4b4b4b; margin-top: 4px;">Target: <b>{selected_cfg['target_desc']}</b></div>
          <div style="font-size: 12px; color: #8b6f4e; font-weight: 600; margin-top: 6px;">Threshold: {selected_cfg['conf']:.3f} (Locked Standard)</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.expander("Workflow Instructions", expanded=False):
        st.markdown(
            """
            **1. Ingest Data**  
            Upload an `.xlsx` or `.xlsm` report with photo links.

            **2. Automated Audit**  
            Downloads photos in parallel and runs YOLO AI inference.

            **3. Analytics & Inspection**  
            Review compliance rates, store rankings, and side-by-side photo detections.

            **4. Export Results**  
            Download audited Excel files and annotated image archives.
            """
        )
    st.markdown("---")
    st.caption("Retail Operations Compliance Hub")

# Main Container
if "audit_results" not in st.session_state:
    st.session_state.audit_results = None

tab_upload, tab_dashboard, tab_inspector = st.tabs([
    "Upload & Run Audit",
    "Analytics & Database",
    "Visual Photo Inspector"
])

# TAB 1: UPLOAD & AUDIT LAUNCH
with tab_upload:
    st.markdown('<div class="section-title">Audit File Ingestion</div>', unsafe_allow_html=True)
    st.markdown("Upload the retail execution Excel file containing bay photos or remote image links.")

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
        st.write("File Loaded:", f"**{file_details['Filename']}**", f"({file_details['File Size']})")

        start_btn = st.button("Execute Planogram AI Audit", use_container_width=True)

        if start_btn:
            status_box = st.empty()
            progress_bar = st.progress(0.0)

            def update_status(text):
                status_box.info(text)

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
                status_box.success(f"Audit completed successfully in {duration:.1f} seconds. Switch to the Analytics tab to view results.")
                progress_bar.progress(1.0)
            except Exception as e:
                status_box.error(f"Audit encountered an error: {str(e)}")
                st.exception(e)

# TAB 2: POWER BI / INTERACTIVE ANALYTICS
with tab_dashboard:
    res = st.session_state.audit_results

    if not res:
        st.info("No audit has been executed yet. Please upload an Excel sheet in Tab 1 to generate analytics.")
    else:
        st.markdown(
            f"""
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
              <div>
                <h2 style="margin: 0; font-size: 24px;">Executive Audit Summary</h2>
                <div style="font-size: 13px; color: #8b6f4e; font-weight: 500;">Specification: {res['model_name']} | Calibrated Threshold: {res['confidence_threshold']:.3f}</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # KPI Metric Summary Cards
        col1, col2, col3, col4, col5 = st.columns(5)

        with col1:
            st.markdown(
                f"""
                <div class="msi-kpi-card">
                  <div class="msi-kpi-label">Total Bays Audited</div>
                  <div class="msi-kpi-value">{res['total_rows']}</div>
                  <div class="msi-kpi-sub">100% Ingested</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col2:
            st.markdown(
                f"""
                <div class="msi-kpi-card" style="border-top-color: #276738;">
                  <div class="msi-kpi-label">Compliance Pass Rate</div>
                  <div class="msi-kpi-value" style="color: #276738;">{res['pass_rate']:.1f}%</div>
                  <div class="msi-kpi-sub">{res['approved_count']} Approved Bays</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col3:
            st.markdown(
                f"""
                <div class="msi-kpi-card" style="border-top-color: #b32626;">
                  <div class="msi-kpi-label">Rejected Bays</div>
                  <div class="msi-kpi-value" style="color: #b32626;">{res['rejected_count']}</div>
                  <div class="msi-kpi-sub">{(res['rejected_count'] / res['total_rows'] * 100) if res['total_rows'] else 0:.1f}% Non-Compliant</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col4:
            st.markdown(
                f"""
                <div class="msi-kpi-card" style="border-top-color: #7d6245;">
                  <div class="msi-kpi-label">Store Locations</div>
                  <div class="msi-kpi-value">{len(res['store_summary'])}</div>
                  <div class="msi-kpi-sub">Distinct Stores</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col5:
            st.markdown(
                f"""
                <div class="msi-kpi-card" style="border-top-color: #63615a;">
                  <div class="msi-kpi-label">Errors / Unreachable</div>
                  <div class="msi-kpi-value" style="color: #63615a;">{res['error_count']}</div>
                  <div class="msi-kpi-sub">Missing Image Links</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # Interactive Visualizations
        c_left, c_right = st.columns([1, 1.4])

        with c_left:
            st.markdown('<div class="section-title">Compliance Verdict Breakdown</div>', unsafe_allow_html=True)
            verdict_counts = {
                "Approved": res["approved_count"],
                "Rejected": res["rejected_count"],
                "Error": res["error_count"]
            }
            labels = [k for k, v in verdict_counts.items() if v > 0]
            values = [v for k, v in verdict_counts.items() if v > 0]
            colors_map = {
                "Approved": "#8b6f4e",
                "Rejected": "#191919",
                "Error": "#e2dfd7"
            }

            fig_donut = go.Figure(data=[go.Pie(
                labels=labels,
                values=values,
                hole=0.6,
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
                font=dict(family="Anthropic Sans, Inter, sans-serif", color="#191919")
            )
            st.plotly_chart(fig_donut, use_container_width=True)

        with c_right:
            st.markdown('<div class="section-title">Store Location Compliance Rankings</div>', unsafe_allow_html=True)

            # Build store dataframe
            st_data_list = []
            for s_id, s_info in res["store_summary"].items():
                pass_pct = (s_info["approved"] / s_info["total"] * 100) if s_info["total"] > 0 else 0
                st_data_list.append({
                    "Store": f"Store {s_id}",
                    "Total": s_info["total"],
                    "Approved": s_info["approved"],
                    "Rejected": s_info["rejected"],
                    "Pass Rate (%)": round(pass_pct, 1)
                })

            df_stores = pd.DataFrame(st_data_list)

            # Separate multi-bay stores (richest insights) and single-bay stores
            multi_bay_stores = df_stores[df_stores["Total"] > 1].sort_values(by=["Pass Rate (%)", "Total"], ascending=[True, False])
            
            if len(multi_bay_stores) >= 5:
                chart_df = multi_bay_stores.head(15)
                st.caption("Displaying stores with multiple audited bays ranked by compliance rate.")
            else:
                chart_df = df_stores.sort_values(by="Pass Rate (%)", ascending=True).head(15)
                st.caption("Displaying stores ranked by compliance rate.")

            fig_bar = go.Figure()

            # Approved bars
            fig_bar.add_trace(go.Bar(
                y=chart_df["Store"],
                x=chart_df["Approved"],
                name="Approved",
                orientation='h',
                marker=dict(color='#8b6f4e'),
                text=chart_df["Approved"],
                textposition='auto'
            ))

            # Rejected bars
            fig_bar.add_trace(go.Bar(
                y=chart_df["Store"],
                x=chart_df["Rejected"],
                name="Rejected",
                orientation='h',
                marker=dict(color='#191919'),
                text=chart_df["Rejected"],
                textposition='auto'
            ))

            fig_bar.update_layout(
                barmode='stack',
                height=320,
                margin=dict(t=10, b=10, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                font=dict(family="Anthropic Sans, Inter, sans-serif", color="#191919"),
                xaxis=dict(title="Number of Bays Audited", showgrid=True, gridcolor="#f2f0ea"),
                yaxis=dict(type='category', showgrid=False)
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        # Database Explorer (Renamed from Notion-Style)
        st.markdown('<div class="section-title">Audit Record Database</div>', unsafe_allow_html=True)

        df_rows = pd.DataFrame(res["row_results"])

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

        display_cols = ["row_idx", "store", "bay", "verdict", "confidence", "format", "boxes", "latency_ms"]
        st.dataframe(
            filtered_df[display_cols].rename(columns={
                "row_idx": "Row",
                "store": "Store / Location",
                "bay": "Bay ID",
                "verdict": "Verdict",
                "confidence": "Confidence",
                "format": "Detected Class",
                "boxes": "Detections",
                "latency_ms": "Latency (ms)"
            }),
            use_container_width=True,
            height=340
        )

        st.markdown("<br>", unsafe_allow_html=True)

        # Export & Download Center
        st.markdown('<div class="section-title">Audit Export & Deliverables</div>', unsafe_allow_html=True)
        d_col1, d_col2 = st.columns(2)

        with d_col1:
            if Path(res["output_excel_path"]).exists():
                with open(res["output_excel_path"], "rb") as f_excel:
                    st.download_button(
                        label="Download Audited Excel Report (.xlsx)",
                        data=f_excel.read(),
                        file_name=Path(res["output_excel_path"]).name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        with d_col2:
            if res.get("zip_buffer"):
                st.download_button(
                    label="Download Annotated Images (.zip)",
                    data=res["zip_buffer"],
                    file_name=f"{Path(res['output_excel_path']).stem}_Annotated_Images.zip",
                    mime="application/zip",
                    use_container_width=True
                )

# TAB 3: VISUAL PHOTO INSPECTOR
with tab_inspector:
    res = st.session_state.audit_results
    if not res:
        st.info("Execute an audit first in Tab 1 to inspect photos.")
    else:
        st.markdown('<div class="section-title">Side-by-Side Visual Photo Inspector</div>', unsafe_allow_html=True)
        st.caption("Select any audited row from the dropdown to compare the original photo against the AI bounding box detection.")

        row_options = [
            f"Row {r['row_idx']} | Store {r['store']} | Bay {r['bay']} | [{r['verdict']}]"
            for r in res["row_results"]
        ]

        selected_row_label = st.selectbox("Select Item for Inspection", options=row_options, index=0)
        selected_idx = int(selected_row_label.split(" ")[1])

        item = next((r for r in res["row_results"] if r["row_idx"] == selected_idx), None)

        if item:
            st.markdown(
                f"""
                <div style="background-color: #faf9f6; border: 1px solid #e2dfd7; border-left: 3px solid #8b6f4e; padding: 14px; margin-bottom: 20px;">
                  <span style="font-weight: 600; color: #191919;">Store:</span> {item['store']} &nbsp;|&nbsp;
                  <span style="font-weight: 600; color: #191919;">Bay:</span> {item['bay']} &nbsp;|&nbsp;
                  <span style="font-weight: 600; color: #191919;">Verdict:</span> <b>{item['verdict']}</b> &nbsp;|&nbsp;
                  <span style="font-weight: 600; color: #191919;">Confidence:</span> {item['confidence']:.3f} &nbsp;|&nbsp;
                  <span style="font-weight: 600; color: #191919;">Detected:</span> {item['format']} ({item['boxes']} detections)
                </div>
                """,
                unsafe_allow_html=True
            )

            p_col1, p_col2 = st.columns(2)

            with p_col1:
                st.markdown("**Original Submitted Photo**")
                if item["original_path"] and Path(item["original_path"]).exists():
                    st.image(item["original_path"], use_container_width=True)
                else:
                    st.warning("Original photo not available or download failed.")

            with p_col2:
                st.markdown("**AI Vision Bounding Box Prediction**")
                if item["annotated_path"] and Path(item["annotated_path"]).exists():
                    st.image(item["annotated_path"], use_container_width=True)
                else:
                    st.info("No detections or photo unavailable.")
