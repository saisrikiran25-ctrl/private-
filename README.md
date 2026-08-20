# Listing Factory — Multi-Marketplace Packaging Studio

Automated packaging, validation, and delivery engine for Indian e-commerce cataloging agencies.
Takes AI-generated JSON copy (Claude/Gemini) and loose image files (ChatGPT/Canva), validates them against Amazon.in, Flipkart, and Meesho upload constraints, organizes images into SKU subfolders, auto-populates official marketplace multi-tab Excel flat files, and outputs a ready-to-deliver client ZIP archive.

## 🚀 Live Web App (GitHub Pages)

Access the live studio directly in your browser with zero installation:
**[https://saisrikiran25-ctrl.github.io/private-/](https://saisrikiran25-ctrl.github.io/private-/)**

---

## ✨ Features

- **Dual-Mode AI JSON Ingestion**: Upload `.json` files or paste directly from LLMs with real-time validation badges.
- **Pre-Flight Validation Engine**:
  - Amazon backend keywords capped at $\le 240$ bytes
  - Amazon title length warnings ($\le 180$ characters)
  - Flipkart & Meesho mandatory attribute completion checks
- **Interactive Multi-Image Drop Zone**:
  - Auto-matches images to SKUs by filename prefix (e.g. `SKU_01_MAIN.jpg`)
  - Live visual checklist for core image slots (`_MAIN`, `_PT01_Size`, `_PT02_Fabric`, `_PT03_Care`)
  - Isolates unmatched images into `Unassigned_Assets/`
- **Master Multi-Tab Excel Workbook**:
  1. `Master_Summary`: Executive overview with upload readiness status
  2. `01_Amazon_Bulk_Import`: Official 18-column Amazon.in flat file schema
  3. `02_Flipkart_Bulk_Import`: Official 17-column Flipkart Seller Hub schema
  4. `03_Meesho_Bulk_Import`: Official 8-column Meesho Supplier Panel schema
- **Complete Client ZIP Delivery Package**:
  - Styled Master Excel workbook
  - `README_Upload_Instructions.txt` (3-step marketplace guide)
  - Organized SKU image folders

---

## 💻 Local Setup & Development

### 1. Clone the repository
```bash
git clone https://github.com/saisrikiran25-ctrl/private-.git
cd private-
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the local server
```bash
python app.py
```
Open **[http://127.0.0.1:8000](http://127.0.0.1:8000)** in your browser.
