# ✦ AI Dashboard Studio

A Streamlit app that uses **Google Gemini** to analyse multi-sheet Excel files and generate interactive dashboards — exportable to PowerPoint.

---

## Features

| Feature | Details |
|--------|---------|
| 📊 Multi-sheet Excel | Reads all tabs, understands inter-sheet relationships |
| 🤖 Gemini Analysis | Auto-generates KPIs, chart configs, and insights |
| 🎨 Interactive Dashboard | Metric tiles, Plotly charts, AI insight cards |
| 💬 Chat Modification | Live chat to reshape layout and data presentation |
| 📤 PPT Export | Polished PowerPoint with title, KPIs, charts, insights |

---

## Quick Start

### 1. Clone / copy the files
```
dashboard_app/
├── app.py                   # Main Streamlit application
├── requirements.txt         # Python dependencies
├── generate_sample_data.py  # Test data generator
└── README.md
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Get a Gemini API Key
1. Go to https://aistudio.google.com/app/apikey
2. Create a free API key
3. Copy it — you'll paste it into the sidebar

### 4. (Optional) Generate sample data
```bash
python generate_sample_data.py
# Creates: sample_business_data.xlsx  (4 sheets, ~500 rows)
```

### 5. Run the app
```bash
streamlit run app.py
```
Opens at: http://localhost:8501

---

## Usage Flow

```
1. Paste Gemini API key in sidebar
2. Upload Excel file (.xlsx / .xls)
3. Click "Generate Dashboard"
4. Explore tiles, charts, and insights
5. Chat to modify: "Change bar chart to line" / "Add avg order value KPI"
6. Click "Export to PowerPoint" → download .pptx
```

---

## Chat Commands (examples)

- *"Change the sales chart to a line chart"*
- *"Add a KPI for total marketing spend"*
- *"Show customer NPS trend over time"*
- *"Remove the histogram and add a scatter plot of spend vs conversions"*
- *"Reorganise KPIs to show revenue metrics first"*
- *"Add an insight about the best performing region"*

---

## Architecture

```
┌─────────────────────────────────────────────────────┐
│                  Streamlit Frontend                  │
│  Sidebar: API key, upload, generate, export         │
│  Main: KPI tiles, Plotly charts, insights, chat     │
└──────────────┬──────────────────────┬───────────────┘
               │                      │
        ┌──────▼──────┐        ┌──────▼──────┐
        │  Gemini API  │        │   Plotly    │
        │ analyse_and_ │        │  Charts     │
        │    plan()    │        │             │
        │ handle_chat()│        └─────────────┘
        └──────────────┘
               │
        ┌──────▼──────┐
        │  python-pptx │
        │  export_to_  │
        │   pptx()     │
        └─────────────┘
```

### Key Functions

| Function | Purpose |
|----------|---------|
| `load_excel()` | Reads all sheets into DataFrames |
| `build_data_summary()` | Creates text summary for Gemini prompt |
| `analyse_and_plan()` | Gemini → JSON dashboard config |
| `compute_kpi()` | Calculates KPI values from DataFrames |
| `build_figure()` | Converts chart config → Plotly figure |
| `handle_chat()` | Gemini chat → updated dashboard config |
| `export_to_pptx()` | Dashboard config + charts → .pptx file |

---

## Customisation

### Change Gemini model
In `app.py`, find:
```python
return genai.GenerativeModel("gemini-1.5-flash")
```
Change to `gemini-1.5-pro` for better analysis on complex datasets.

### Add more chart types
In `build_figure()`, add new cases to the `ctype` if/elif chain using Plotly Express.

### Modify PPT theme
Colors are defined at the top of `export_to_pptx()`:
```python
PURPLE = RGBColor(0x7C, 0x3A, 0xED)
DARK_BG = RGBColor(0x0A, 0x0A, 0x0F)
# etc.
```

---

## Requirements

- Python 3.9+
- Gemini API key (free at aistudio.google.com)
- Excel file with 1+ sheets
