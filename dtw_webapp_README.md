# NeoAsia Sales Order DTW Web Application

A Streamlit web app for generating SAP DTW import files from Excel data.

## Quick Start (Streamlit Cloud - FREE)

### Option 1: Deploy to Streamlit Community Cloud (Recommended)

1. Create a GitHub account if you don't have one
2. Create a new repository
3. Upload these files to the repository:
   - `app.py`
   - `requirements.txt`
4. Go to [share.streamlit.io](https://share.streamlit.io)
5. Sign in with GitHub
6. Click "New app"
7. Select your repository
8. Set main file path: `app.py`
9. Click "Deploy"

**Done!** You'll get a URL like `https://your-app.streamlit.app`

### Option 2: Run Locally

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

Open http://localhost:8501 in your browser.

### Option 3: Deploy with Docker

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY app.py .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

Build and run:
```bash
docker build -t dtw-app .
docker run -p 8501:8501 dtw-app
```

## User Guide

### Step 1: Download Template
- Click "Download Excel Template" in the app
- Save the template to your computer

### Step 2: Fill in Data
- Open the template in Excel
- Fill in **Sales Order Entry** sheet (one row per order)
- Fill in **Line Items Entry** sheet (lines for each order)
- Look up Sales Codes in **SalesEmployees** sheet

### Step 3: Upload and Generate
- Go to the web app
- Upload your filled Excel file
- Fix any validation errors
- Click "Generate DTW Files"
- Download the ZIP file

### Step 4: Import to SAP
- Extract ORDR_*.txt and RDR1_*.txt from the ZIP
- Copy to Terminal Server
- Open SAP DTW > Import > Transactional Data > Sales Orders
- Select both files and run

## Data Format Rules

| Field | Format | Example |
|-------|--------|---------|
| Document Date | YYYYMMDD | 20260121 |
| Due Date | YYYYMMDD | 20260131 |
| Order # | Unique number | 5001 |
| Line # | Starts at 1 per order | 1, 2, 3 |
| Sales Code | From SalesEmployees sheet | 143 |

## Files

| File | Description |
|------|-------------|
| `app.py` | Main Streamlit application |
| `requirements.txt` | Python dependencies |

## Support

This is an internal tool for NeoAsia Business Analytics.
Contact the BA team for support.
