# Listing Factory — Multi-Marketplace Packaging Studio (v2.0)

**Listing Factory** is a fully client-side, zero-install web application for Indian e-commerce cataloging agencies managing multi-channel uploads across **Amazon.in**, **Flipkart**, and **Meesho**.

It takes raw AI-generated listing copy and loose product images, validates them against strict marketplace rules without silent mutation, organizes assets into canonical SKU folders, generates structured multi-tab Excel mapping workbooks, and packages everything into a ready-to-deliver client handover ZIP.

---

## 🌐 Canonical Web Application (GitHub Pages)

The primary, canonical version of Listing Factory runs 100% in your browser with zero installation, zero data telemetry, and ultra-fast client-side packaging:

👉 **[https://saisrikiran25-ctrl.github.io/private-/](https://saisrikiran25-ctrl.github.io/private-/)**

---

## 🏗️ The Two-Stage Production Workflow

```
┌─────────────────────────────────────────────────────────────┐
│ STAGE 1: AI Listing Generation (Claude / Gemini / GPT-4)    │
│ Generates strictly formatted JSON copy with 5x A/B angles   │
└──────────────────────────────┬──────────────────────────────┘
                               │ JSON Payload
                               ▼
┌─────────────────────────────────────────────────────────────┐
│ STAGE 2: Listing Factory Studio (GitHub Pages Web App)      │
│  • Client-Side Strict Validation (No Silent Truncation)     │
│  • Dynamic Tax & Commercial Ingestion (GST, HSN, Quantity)  │
│  • Color Field Ingestion across Amazon, Flipkart & Meesho   │
│  • Canonical Image Routing (SKU_XX_MAIN.jpg ... PT05)       │
│  • Multi-Tab Marketplace Mapping Workbooks (.xlsx)          │
│  • 5x Alternate Marketing Copies Workbook (.xlsx)           │
│  • Client Handover Delivery Archive (.zip)                  │
└─────────────────────────────────────────────────────────────┘
```

---

## 📊 Marketplace Mapping Workbooks & Color Field Mapping

Listing Factory generates two structured mapping workbooks inside the client ZIP:

### 1. `[Client]_Master_Marketplace_Upload.xlsx`
Structured data sheets designed for transferring copy into official marketplace upload templates:
- **`Master_Summary`**: Executive catalog overview including `Color`, `Fabric`, `Sizes Available`, core coverage (`4/4 Core`), validation status (`✅ Pass`), and package readiness (`✅ Ready for Review`).
- **`01_Amazon_Bulk_Import`**: Amazon.in 23-column mapping schema with dynamic `quantity`, generic keywords ($\le 240$ bytes), 5 bullet points, `size`, `color` (`sku.color`), and canonical image filenames.
- **`02_Flipkart_Bulk_Import`**: Flipkart Seller Hub 23-column schema with controlled attributes (fabric, kurta type, neck, sleeve, length, pattern, occasion), `Color` (`sku.color`), seller GST/HSN, and Flipkart-specific description.
- **`03_Meesho_Bulk_Import`**: Meesho 17-column schema with dual Hinglish + English hook descriptions, `Color` (`sku.color`), 4 highlight badges, seller GST/HSN, dynamic Meesho price, and 5-slot image mappings.

### 2. `[Client]_Alternate_Listing_Copies.xlsx`
Contains 5 distinct marketing angle variations per SKU for A/B testing and seasonal refreshes:
- **V1**: Daily Office & Workwear
- **V2**: Festive & Wedding Celebrations
- **V3**: Summer Heat & Breathable Comfort
- **V4**: High-Value Everyday Essential
- **V5**: Gifting & Modern Fusion

---

## 📸 Canonical Image Naming Scheme

All image assets adhere to this standardized naming convention:

| Role | Canonical Filename Pattern | Purpose / Content |
|---|---|---|
| **Primary Hero** | `SKU_XX_MAIN.jpg` | Pure white background (`#FFFFFF`), primary product cutout |
| **Other Image 1** | `SKU_XX_PT01.jpg` | Size chart, measurement specifications, & fit guide |
| **Other Image 2** | `SKU_XX_PT02.jpg` | Fabric texture, weave detail, & material spec |
| **Other Image 3** | `SKU_XX_PT03.jpg` | Wash care instructions & styling recommendations |
| **Other Image 4** | `SKU_XX_PT04.jpg` | Back view / alternate product angle |
| **Other Image 5** | `SKU_XX_PT05.jpg` | Detail close-up / lifestyle photo |

*Where `XX` represents the zero-padded SKU identifier (e.g., `SKU_01_MAIN.jpg`).*

---

## 🔒 Strict Validation & Schema Rules

- **Zero Silent Mutation**: Over-length titles and search terms are rejected with clear, actionable validation errors rather than silently truncated.
- **Amazon Constraints**: Title $\le 180$ chars; Backend search terms $\le 240$ UTF-8 bytes; exactly 5 bullet points.
- **Flipkart Taxonomy**: Fabric, kurta type, neck, sleeve, length type, pattern, and occasion validated against allowed marketplace literals.
- **Meesho Rules**: Title $\le 60$ chars; exactly 4 highlights; mandatory Hinglish and English hook descriptions.
- **Dynamic Commercials**: GST (%), HSN code, Amazon stock quantity, and color are dynamically ingested from the JSON payload.

---

## 📦 Client ZIP Structure

```
[Client]_[Batch]_Handover_Package/
├── [Client]_Master_Marketplace_Upload.xlsx
├── [Client]_Alternate_Listing_Copies.xlsx  (when alternates provided)
├── README_Upload_Instructions.txt
└── Organized_SKU_Images/
    ├── SKU_01/
    │   ├── SKU_01_MAIN.jpg
    │   ├── SKU_01_PT01.jpg
    │   ├── SKU_01_PT02.jpg
    │   ├── SKU_01_PT03.jpg
    │   └── SKU_01_PT04.jpg
    ├── SKU_02/
    │   └── ...
    └── Unassigned_Assets/
```

---

## 🛠️ Reference Backend Implementation (FastAPI)

For headless automation or programmatic batch generation, a Python FastAPI reference backend is provided in `app.py`:

```bash
# Optional local reference server
pip install -r requirements.txt
python app.py
```

---

## 📝 Important Notice & Disclaimer
The Excel workbooks produced by Listing Factory are structured **Mapping Workbooks** formatted to organize catalog data for seamless copy-pasting into official Seller Central, Seller Hub, and Supplier Panel upload templates. Sellers must verify category-specific rules, tax rates, and portal guidelines prior to final submission.
