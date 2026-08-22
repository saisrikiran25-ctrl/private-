# Listing Factory — Changelog & Known Issues

All notable changes to the Listing Factory packaging studio will be documented in this file.

---

## [v2.0] – 2026-08-22

### 🌟 New Features & Enhancements
- **Amazon Title Quality & Order Pattern**: Added Amazon title quality validation enforcing the standard pattern `[Brand] [Fabric] [Pattern] [Product Type] with [Verified Detail] ([Color])`. Blocks duplicate brand names, repeated adjacent tokens, and redundant phrase stacks as hard errors; flags excessive delimiters and stacked product types as advisory warnings.
- **Truth Boundary & Claim Safety Validator**: Implemented comprehensive truthfulness validation across primary copy and all 5 alternate variants. Enforces hard errors on unverified technical/guarantee claims (`breathable`, `cooling`, `sweat-absorbent`, `quick-dry`, `lightweight`, `zero fading`, `superior quality`, `guaranteed`), advisory warnings on subjective fit/comfort phrases (`comfortable fit`, `perfect fit`, `soft feel`), and permits subjective styling suggestions.
- **Structural Readiness Scope & Disclaimer**: Standardized `STRUCTURAL_READINESS_DISCLAIMER` across generated README, `Master_Summary` 15th column (`Status Scope / Meaning`), UI Section 4 notice, and README.md. Preserves exact 3 readiness status labels without collapsing to a simple "Ready".
- **Declared Image Roles & Verification Notice**: Standardized `IMAGE_ROLES` to explicit declared slots (`Declared Hero Image`, `Declared Size Chart Slot`, etc.), added `IMAGE_ROLE_DISCLAIMER`, and labeled all folder hierarchy assets in handover instructions with `(manual visual check required)`.
- **Cryptographic Package Metadata (`package_metadata.json`)**: Every generated package now contains an audit trail recording `tool_version`, `json_prompt_version`, `schema_version`, local ISO 8601 generation timestamp, and deterministic SHA-256 digests (`input_hash` & `output_hash`) for each SKU.
- **Explicit Schema Version Contract**: Mandates top-level `"schema_version": "v2.0"` in input JSON payloads with strict compatibility checking between prompt output and packaging engines.
- **"To be confirmed" Sentinel Handling**: Fields marked with `"To be confirmed"` or `"TBC"` are ingested safely as advisory review flags rather than throwing hard schema errors.
- **Clear "Ready" Messaging & Seller Checklist**: Renamed status labels across Master Summary workbooks and UI to explicit actionable statements:
  - `✅ Structurally Complete – Seller Review Required`
  - `⚠️ Warnings – Seller Review Required`
  - `❌ Not Ready – Fix Errors First`
- **Structured Per-SKU Validation**: Replaced generic aggregate error outputs with structured per-SKU diagnostic cards separating hard schema errors from non-blocking advisory warnings.
- **Batch-Level Configuration (`batch_config`)**: Enables batch-level inheritance for shared attributes (`brand`, `category`, `seller_config`), with per-SKU overrides.
- **Sample / Dry-Run Mode**: Built-in sample generator endpoint (`/api/generate-sample`) and one-click demo data in the client UI for instant testing.
- **Rollback & Versioned Output Archive**: Automatically maintains sequential package versions (`output/<client>_<batch>_v1.zip`, `v2.zip`, etc.) in the backend repository for audit and rollback.
- **Performance & Soft Limits**: Surfaces non-blocking advisory warnings when batch size exceeds 50 SKUs or images exceed 200 MB.
- **Scope & Limitations & Support Documentation**: Formalized operational boundaries, seller responsibilities, and SLA support expectations in README and in-app documentation.

---

## [v1.0] – 2026-08-10

### 🚀 Initial Release
- Initial multi-marketplace mapping engine for Amazon.in, Flipkart Seller Hub, and Meesho Supplier Panel.
- Client-side and FastAPI packaging with JSZip and ExcelJS.
- Automated image routing by canonical SKU prefix.
- Multi-tab Master Excel workbook builder.
- Handover instructions generation (`README_Upload_Instructions.txt`).

---

## 📌 Known Issues & Technical Boundaries

1. **Filename-Based Image Role Assignment**: Image roles (Hero, Size Chart, Fabric Spec, Care Guide, Back View) are assigned strictly via filename tokens (`_MAIN`, `_PT01` to `_PT05`). Image visual content, background colors, and typography are not evaluated by computer vision.
2. **Category Template Alignment**: The Excel workbooks produced by Listing Factory are structured mapping files. Because Indian e-commerce portals frequently update their vertical schemas, catalog teams must copy-paste data into the latest official portal upload flat files.
3. **Marketplace Processing Times**: Processing windows stated in documentation (~15–30 min for Amazon, ~24–48 hr for Flipkart, ~2–4 hr for Meesho) are industry averages and vary depending on portal queue load and account standing.
