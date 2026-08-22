# Listing Factory — Changelog & Known Issues

All notable changes to the Listing Factory packaging studio will be documented in this file.

---

## [v2.1] – 2026-08-22

### 🧬 Category-Profile-Driven Architecture
- **13 Supported Product Families**: Refactored the core catalog pipeline from hardcoded Women Ethnic Wear attributes into a dynamic, extensible Category Profile Registry covering:
  1. `women_ethnic_kurta` (Women Ethnic Wear — Kurta / Kurti / Tunic / Set)
  2. `saree` (Sarees — Daily / Festive / Traditional)
  3. `coord_set` (Co-ord Sets — 2-Piece / 3-Piece Sets)
  4. `women_dress` (Women's Dresses — A-Line / Maxi / Midi)
  5. `women_top` (Women's Western Tops & Tunics)
  6. `men_shirt` (Men's Casual & Formal Shirts)
  7. `men_tshirt` (Men's T-Shirts & Polos)
  8. `men_bottomwear` (Men's Bottomwear — Jeans / Trousers / Chinos)
  9. `women_bottomwear` (Women's Bottomwear — Jeans / Palazzo / Pants)
  10. `men_ethnic` (Men's Ethnic Wear — Kurta / Pyjama / Set)
  11. `kidswear` (Kidswear — Boys & Girls Garments / Sets)
  12. `footwear` (Footwear — Casual / Formal / Sports)
  13. `home_textiles` (Home Textiles & Furnishing — Bedsheets / Curtains)
- **Profile-Driven Required Fact Enforcement**: Each category profile enforces mandatory verified product facts (e.g. Saree length, Co-ord set package contents, Footwear sole material, Home textile dimensions, Kidswear age group) before passing validation.
- **Category-Specific Prohibited Claims**: Block unverified high-risk claims per vertical:
  - *Footwear*: Blocks unverified `anti-slip`, `cushioning`, `arch support`, `memory foam`.
  - *Sarees*: Blocks unverified `handloom`, `Banarasi`, `Kanjivaram`, `pure silk`.
  - *Home Textiles*: Blocks unverified thread count (`400 TC`), `GSM`, `hypoallergenic`.
  - *Kidswear*: Blocks unverified `skin-safe`, `gentle on skin`, `organic`.
  - *Apparel*: Flags unverified `slim fit`, `tailored fit`, `comfort fit` with advisory warnings.
- **Strict Exclusion of Non-Supported Profiles**: Explicitly rejects `blouse`, `lingerie`, `innerwear`, `shapewear`, `bra`, `underwear`, and intimate apparel at the validation layer with the message: `"This category profile is intentionally not supported in Listing Factory v2.0."`
- **Dynamic Excel Mapping Workbooks**:
  - `Master_Summary`: 17 columns including `Category Profile` and `Status Scope / Meaning`.
  - `01_Amazon_Bulk_Import`: Ingests profile-specific `item_type_keyword` and `feed_product_type`.
  - `02_Flipkart_Bulk_Import`: Dynamic column headers built from active profile controlled attributes (no kurta columns on sarees, footwear, or shirts).
  - `03_Meesho_Bulk_Import`: Dynamic material and dimension ingestion.
- **Profile-Driven UI Studio (index.html)**:
  - Category Profile selector dropdown with optgroups.
  - Interactive Profile Inspector card displaying required facts, controlled attributes, and marketplace hints.
  - One-click sample JSON loader for all 13 supported profiles.
  - Excluded categories policy notice banner.
- **Full Backward Compatibility**: Legacy v2.0 payloads lacking `category_profile` infer `women_ethnic_kurta` with an advisory warning when `category == "Women Ethnic Wear"` or attributes resemble kurta data.

---

## [v2.0] – 2026-08-22

### 🌟 New Features & Enhancements
- **Amazon Title Quality & Order Pattern**: Added Amazon title quality validation enforcing standard pattern `[Brand] [Fabric] [Pattern] [Product Type] with [Verified Detail] ([Color])`. Blocks duplicate brand names, repeated adjacent tokens, and redundant phrase stacks.
- **Truth Boundary & Claim Safety Validator**: Implemented truthfulness validation across primary copy and 5 alternate variants.
- **Structural Readiness Scope & Disclaimer**: Standardized `STRUCTURAL_READINESS_DISCLAIMER` across generated README, `Master_Summary`, UI, and documentation.
- **Declared Image Roles & Verification Notice**: Standardized `IMAGE_ROLES` to explicit declared slots (`Declared Hero Image`, `Declared Size Chart Slot`, etc.), added `IMAGE_ROLE_DISCLAIMER`.
- **Cryptographic Package Metadata (`package_metadata.json`)**: Every package contains an audit trail recording `tool_version`, `json_prompt_version`, `schema_version`, local ISO 8601 timestamp, and SHA-256 digests (`input_hash` & `output_hash`).
- **Explicit Schema Version Contract**: Mandates top-level `"schema_version": "v2.0"`.
- **"To be confirmed" Sentinel Handling**: Fields marked with `"To be confirmed"` or `"TBC"` are ingested safely as advisory review flags.

---

## [v1.0] – 2026-08-10

### 🚀 Initial Release
- Initial multi-marketplace mapping engine for Amazon.in, Flipkart Seller Hub, and Meesho Supplier Panel.
- Client-side and FastAPI packaging with JSZip and ExcelJS.
