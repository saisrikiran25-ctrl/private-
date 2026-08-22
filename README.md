# Listing Factory — Multi-Marketplace Packaging Studio (v2.0)

**Listing Factory** is a fully client-side, zero-install web application for Indian e-commerce cataloging agencies managing multi-channel uploads across **Amazon.in**, **Flipkart**, and **Meesho**.

It takes raw AI-generated listing copy and loose product images, validates them against strict marketplace rules without silent mutation, organizes assets into canonical SKU folders, generates structured multi-tab Excel mapping workbooks, creates a cryptographic audit trail (`package_metadata.json`), and packages everything into a ready-to-deliver client handover ZIP.

See [CHANGELOG.md](file:///c:/Users/Sai%20Kiran/Downloads/Listing_Factory/CHANGELOG.md) for detailed version history and known technical limitations.

---

## 🌐 Canonical Web Application (GitHub Pages)

The primary, canonical version of Listing Factory runs 100% in your browser with zero installation, zero data telemetry, and ultra-fast client-side packaging:

👉 **[https://saisrikiran25-ctrl.github.io/private-/](https://saisrikiran25-ctrl.github.io/private-/)**

---

## 🏗️ The Two-Stage Production Workflow

```
┌─────────────────────────────────────────────────────────────┐
│ STAGE 1: AI Listing Generation (Claude / Gemini / GPT-4)    │
│  • Truth Boundary: Forbids hallucination of fabric/specs    │
│  • Separation of Facts vs Copy (verified.* product input)   │
│  • Enforces schema_version: "v2.0" & batch_config defaults  │
│  • Generates strictly formatted JSON with 5x A/B angles     │
└──────────────────────────────┬──────────────────────────────┘
                               │ JSON Payload (schema_version: "v2.0")
                               ▼
┌─────────────────────────────────────────────────────────────┐
│ STAGE 2: Listing Factory Studio (GitHub Pages Web App)      │
│  • Schema Version Contract Enforcement ("v2.0")             │
│  • Structured Per-SKU Validation (No Silent Truncation)     │
│  • "To be confirmed" Sentinel Handling & Review Flags       │
│  • Clear "Structurally Complete – Seller Review Required"   │
│  • Dynamic Tax & Commercial Ingestion (GST, HSN, Quantity)  │
│  • Color Field Ingestion across Amazon, Flipkart & Meesho   │
│  • Cryptographic Package Audit Trail (SHA-256 Metadata)     │
│  • Canonical Image Routing (SKU_XX_MAIN.jpg ... PT05)       │
│  • Multi-Tab Marketplace Mapping Workbooks (.xlsx)          │
│  • 5x Alternate Marketing Copies Workbook (.xlsx)           │
│  • Client Handover Delivery Archive (.zip)                  │
└─────────────────────────────────────────────────────────────┘
```

---

## 🤖 Stage 1: AI Prompt Layer & Truth Boundary Contracts

When generating JSON listing payloads with LLMs, the master generation prompt must enforce these foundational contracts:

1. **Truth Boundary / No Hallucination**:
   - The LLM is strictly forbidden from inventing fabric composition, wash care rules, stitching styles, or technical claims not present in the verified product record (`verified.*`).
   - If an attribute is unknown or unverified, the LLM must mark it as `"To be confirmed"` or use neutral generic phrasing (e.g., *"as per product label"*).
2. **Separation of Facts vs. Copy**:
   - **Stage A (Factual Basis)**: Verified physical product records (`sku_id`, `brand`, `product_type`, `color`, `sizes`, `mrp`, `meesho_price`, `seller_config`).
   - **Stage B (Creative Copywriting)**: Marketplace-specific copy, bullet points, Hinglish hooks, and 5x marketing variations derived strictly from verified facts.
3. **Formal JSON Schema Contract**:
   - Every payload must declare top-level `"schema_version": "v2.0"`.
   - Exact array lengths: Amazon bullet points ($= 5$), Meesho highlights ($= 4$), Alternates ($= 5$).
   - Strict length/byte ceilings: Amazon Title $\le 180$ chars, Amazon BST $\le 240$ UTF-8 bytes, Meesho Title $\le 60$ chars.

---

## ✨ Features & Capabilities

### 1. 🛡️ Cryptographic Package Audit Trail (`package_metadata.json`)
Every generated package contains a top-level `package_metadata.json` documenting:
- **`tool_version`**: Tool build (`Listing Factory v2.0`).
- **`json_prompt_version`**: Master generation prompt release (`JSON Prompt v1.0 – 2026-08-22`).
- **`schema_version`**: Schema compatibility contract (`v2.0`).
- **`generated_at`**: Exact ISO 8601 timestamp with local timezone offset.
- **Per-SKU Cryptographic Hashes**:
  - `input_hash`: SHA-256 digest of verified product input records.
  - `output_hash`: SHA-256 digest of the complete generated listing record.

### 2. 📜 Schema Version Contract (`schema_version: "v2.0"`)
- Requires `"schema_version": "v2.0"` in input JSON payloads.
- Validates prompt output compatibility against the application engine to prevent silent schema drift.

### 3. ⚠️ Explicit "To be confirmed" Sentinel Handling
- Fields marked with `"To be confirmed"` or `"TBC"` are accepted without failing hard schema validation.
- Surfaced as advisory review warnings in the UI and flagged in the `Master_Summary` Excel tab under **`Review Flags: ⚠️ Has unconfirmed fields`**.

### 4. 🏷️ Clear "Ready" Phrasing & Seller Checklist
To prevent misunderstanding "Ready" as automatic "marketplace acceptance", status labels are explicitly worded:
- `✅ Structurally Complete – Seller Review Required`
- `⚠️ Warnings – Seller Review Required`
- `❌ Not Ready – Fix Errors First`

### 5. 🔍 Structured Per-SKU Validation Breakdown
- Replaces generic error lists with structured per-SKU diagnostic cards separating hard schema errors from non-blocking advisory warnings.

### 6. ⚙️ Batch-Level Configuration (`batch_config`)
Allows specifying shared brand, category, and seller configuration at the batch level while allowing individual SKU overrides:
```json
{
  "schema_version": "v2.0",
  "batch_config": {
    "brand": "Janasya",
    "category": "Women Ethnic Wear",
    "seller_config": {
      "amazon_quantity": 50,
      "gst_percent": 5,
      "hsn_code": "62114200"
    }
  },
  "skus": [
    { "sku_id": "SKU_01", "product_type": "Kurti", "color": "Navy Blue", ... }
  ]
}
```

### 7. 🧪 Sample / Dry-Run Mode
- **Client Studio:** One-click `✨ Load Sample (v2.0)` button in the top navigation loads a complete, valid sample payload for immediate testing.
- **Backend API:** Dedicated `POST /api/generate-sample` endpoint generates a full mock client handover ZIP for testing pipelines.

### 8. 🗄️ Versioned Output Archival & Rollback
- The backend server automatically archives generated packages into `output/` as `<client>_<batch>_v1.zip`, `v2.zip`, etc., for audit and rollback.
- Query historical archives via `GET /api/history`.

---

## ⚖️ Scope & Limitations Boundary

- **Listing Copy & Mapping Preparation:** Prepares structured listing copy and multi-tab Excel mapping workbooks based on seller-provided catalog data.
- **No Acceptance Guarantee:** Does not guarantee automatic marketplace approval, indexation, or exemption from category ungating requirements.
- **Asset Slotting:** Assigns image slots based on canonical filename patterns; visual content, white background quality, and typography are not evaluated by computer vision.
- **Tax & Classification Disclaimer:** GST percentages and HSN codes are seller-provided configurations and do not constitute official tax advice.
- **Mandatory Final Review:** All data must be reviewed, verified, and confirmed by the brand prior to live portal submission.

---

## 📊 Marketplace Mapping Workbooks

Listing Factory generates two structured mapping workbooks inside the client ZIP:

### 1. `[Client]_Master_Marketplace_Upload.xlsx`
Structured data sheets designed for transferring copy into official marketplace upload templates:
- **`Master_Summary`**: Catalog overview including `Color`, `Fabric`, `Sizes Available`, core coverage (`4/4 Core`), `Validation Status` (`✅ Pass`), `Review Flags` (`⚠️ Has unconfirmed fields` / `—`), and `Package Readiness` (`✅ Structurally Complete – Seller Review Required`).
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

| Role | Canonical Filename Pattern | Purpose / Content |
|---|---|---|
| **Primary Hero** | `SKU_XX_MAIN.jpg` | Pure white background (`#FFFFFF`), primary product cutout |
| **Other Image 1** | `SKU_XX_PT01.jpg` | Size chart, measurement specifications, & fit guide |
| **Other Image 2** | `SKU_XX_PT02.jpg` | Fabric texture, weave detail, & material spec |
| **Other Image 3** | `SKU_XX_PT03.jpg` | Wash care instructions & styling recommendations |
| **Other Image 4** | `SKU_XX_PT04.jpg` | Back view / alternate product angle |
| **Other Image 5** | `SKU_XX_PT05.jpg` | Detail close-up / lifestyle photo |

---

## 📦 Client ZIP Structure

```
[Client]_[Batch]_Handover_Package/
├── [Client]_Master_Marketplace_Upload.xlsx
├── [Client]_Alternate_Listing_Copies.xlsx  (when alternates provided)
├── package_metadata.json                   (cryptographic audit trail)
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

For headless automation, sample generation, or batch pipelines, a Python FastAPI reference backend is provided:

```bash
pip install -r requirements.txt
python app.py
```

---

## 📞 Support & SLA Expectations

Support is provided on a best-effort basis (Monday–Friday, 10:00–18:00 IST). Critical bugs (app crashes, broken ZIP archives) are prioritized. For technical inquiries, contact `support@listingfactory.internal`.
