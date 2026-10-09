# Planogram (POG) AI Audit Platform | MSI Surfaces Design

An enterprise-grade, automated computer vision platform for Planogram (POG) retail photo compliance audits. Built with **Streamlit**, **Ultralytics YOLOv8**, and **Plotly**, styled according to the **MSI Surfaces** architectural luxury design system (`#8b6f4e` bronze palette, 0px border-radius).

---

## 🏛️ Features

- **MSI Surfaces Architectural UI**:
  - Warm Bronze (`#8b6f4e`), dark charcoal (`#212529`), and stone neutral canvas.
  - Crisp zero border-radius (`0px`) on all tiles, cards, dropzones, and buttons.
- **Calibrated Vision AI Models**:
  - **Rubber Mat Audit**: Detects open stacks & boxed cartons (`conf = 0.25`).
  - **Threshold Fixture Audit**: Detects 2-tier upright threshold fixtures (`conf = 0.710`).
- **Power BI / Notion-Style Interactive Analytics**:
  - Live KPI metric cards with compliance pass rates, rejection counts, and error tracking.
  - Interactive Plotly Donut chart for verdict breakdowns.
  - Store/Location performance leaderboard ranking compliance by store.
  - Notion-style searchable and filterable database table.
- **Side-by-Side Visual Photo Inspector**:
  - Compares original submitted store photos against AI bounding box detections.
- **1-Click Export & Deliverables**:
  - Formatted Excel report (`.xlsx`) with an executive dashboard and color-coded status cells.
  - Annotated bounding-box images bundle (`.zip`).

---

## 🚀 Quickstart

### 1. Installation
Clone the repository and install dependencies:
```bash
git clone https://github.com/sarthaksharma-github/pog-audit-streamlit-app.git
cd pog-audit-streamlit-app
pip install -r requirements.txt
```

### 2. Run the App
Launch the Streamlit web application:
```bash
streamlit run app.py
```
Open your browser at `http://localhost:8501`.

---

## 📁 Repository Structure
```text
pog-audit-streamlit-app/
├── .streamlit/
│   └── config.toml          # MSI Surfaces theme configuration
├── assets/
│   └── style.css            # Custom CSS for zero-radius luxury design
├── models/
│   ├── yolov8n_mat.pt       # Trained weights for Rubber Mat model
│   └── yolov8n_fixture.pt   # Trained weights for Fixture model
├── app.py                   # Main Streamlit web application
├── core_engine.py           # Parallel downloader, YOLO inference & report engine
├── requirements.txt         # Production dependencies
└── README.md                # Documentation
```

---

## 🔒 Security & Privacy
Original uploaded spreadsheets are kept untouched; all audit computations and reports are generated into dedicated output run directories.
