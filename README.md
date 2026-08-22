# Listing Factory — Category-Profile-Driven Packaging Studio (v2.1)

**Listing Factory** is a robust, category-profile-driven web application for Indian e-commerce cataloging agencies managing multi-channel uploads across **Amazon.in**, **Flipkart**, and **Meesho** for 13 distinct product families.

It takes AI-generated listing copy and loose product images, validates them against strict category profile contracts and marketplace rules without silent mutation, organizes assets into canonical SKU folders, generates structured multi-tab Excel mapping workbooks with profile-specific attribute columns, creates a cryptographic audit trail (`package_metadata.json`), and packages everything into a ready-to-deliver client handover ZIP.

See [CHANGELOG.md](file:///c:/Users/Sai%20Kiran/Downloads/Listing_Factory/CHANGELOG.md) for detailed version history.

---

## 🌐 Canonical Web Application (GitHub Pages)

The primary, canonical version of Listing Factory runs 100% in your browser with zero installation, zero data telemetry, and client-side packaging:

👉 **[https://saisrikiran25-ctrl.github.io/private-/](https://saisrikiran25-ctrl.github.io/private-/)**

---

## 🧬 Category Profile Architecture (13 Product Families)

Listing Factory v2.1 replaces single-garment hardcoding with a central **Category Profile Registry**. The validation engine, prompt guidance, and Excel output dynamically adapt based on the selected `category_profile`:

| # | Profile ID | Product Family | Category Group | Amazon Item Keyword | Flipkart Vertical | Meesho Vertical |
|---|---|---|---|---|---|---|
| 1 | `women_ethnic_kurta` | Women Ethnic Wear | Women Ethnic Wear | `kurtas-and-ethnic-tops` | `Ethnic Wear / Kurta` | `Women Ethnic Wear / Kurtis` |
| 2 | `saree` | Sarees (Daily / Festive) | Sarees | `sarees` | `Saree` | `Women Ethnic Wear / Sarees` |
| 3 | `coord_set` | Co-ord Sets (2/3-Piece) | Women Western & Sets | `co-ord-sets` | `Co-ords` | `Women Western / Co-ord Sets` |
| 4 | `women_dress` | Women's Dresses | Women Western Wear | `dresses` | `Dress` | `Women Western / Dresses` |
| 5 | `women_top` | Women's Western Tops | Women Western Wear | `tunics-and-western-tops` | `Top` | `Women Western / Tops` |
| 6 | `men_shirt` | Men's Shirts | Men Western Wear | `mens-casual-shirts` | `Shirt` | `Men Western / Shirts` |
| 7 | `men_tshirt` | Men's T-Shirts & Polos | Men Western Wear | `mens-t-shirts` | `T-Shirt` | `Men Western / T-Shirts` |
| 8 | `men_bottomwear` | Men's Bottomwear | Men Western Wear | `mens-trousers-and-jeans` | `Trouser / Jeans` | `Men Western / Bottomwear` |
| 9 | `women_bottomwear` | Women's Bottomwear | Women Western & Ethnic | `womens-bottomwear` | `Women Bottomwear` | `Women Western / Bottomwear` |
| 10 | `men_ethnic` | Men's Ethnic Wear | Men Ethnic Wear | `mens-ethnic-wear` | `Men Ethnic Wear` | `Men Ethnic / Kurtas` |
| 11 | `kidswear` | Kidswear (Boys & Girls) | Kidswear | `kids-apparel` | `Kids Apparel` | `Kids / Clothing Sets` |
| 12 | `footwear` | Footwear (Casual / Formal) | Footwear | `casual-shoes` | `Footwear` | `Footwear / Casual Shoes` |
| 13 | `home_textiles` | Home Furnishing | Home & Furnishing | `home-furnishing` | `Home Furnishing` | `Home & Kitchen / Bedding` |

### 🚫 Explicitly Excluded Profiles (Strict Rejection)
The following category profiles are **intentionally not supported** in Listing Factory and will be rejected at the validation layer with an explicit notice:
- `blouse`, `lingerie`, `innerwear`, `shapewear`, `bra`, `underwear`, or any sexualized / intimate apparel.
- *Rejection Message*: `"This category profile is intentionally not supported in Listing Factory v2.0."`

---

## 🤖 AI Prompt Guidance for JSON Generation

When prompting an LLM (Claude, Gemini, GPT-4) to generate listing records, provide the verified product facts and select the exact `category_profile`. The LLM must adhere to these category-specific contracts:

### 1. Required Verified Facts by Profile
- **`women_ethnic_kurta`**: `fabric`, `sleeve`, `neckline`, `length`, `pattern`, `care_label`, `product_type`, `color`, `sizes`.
- **`saree`**: `fabric`, `pattern`, `saree_length`, `blouse_piece_included`, `care_label`, `color`, `occasion`.
- **`coord_set`**: `fabric`, `pattern`, `top_type`, `bottom_type`, `sleeve`, `neckline`, `top_length`, `bottom_length`, `care_label`, `package_contents`, `sizes`, `color`.
- **`women_dress`**: `fabric`, `dress_type`, `neckline`, `sleeve`, `length`, `pattern`, `care_label`, `sizes`, `color`.
- **`women_top`**: `fabric`, `top_type`, `neckline`, `sleeve`, `length`, `pattern`, `care_label`, `sizes`, `color`.
- **`men_shirt`**: `fabric`, `shirt_type`, `collar_type`, `sleeve`, `pattern`, `care_label`, `sizes`, `color`.
- **`men_tshirt`**: `fabric`, `tshirt_type`, `neckline`, `sleeve`, `pattern`, `care_label`, `sizes`, `color`.
- **`men_bottomwear`**: `bottom_type`, `fabric`, `pattern`, `waist_type`, `length_or_inseam`, `closure_type`, `care_label`, `sizes`, `color`.
- **`women_bottomwear`**: `bottom_type`, `fabric`, `pattern`, `waist_type`, `length_or_inseam`, `closure_type`, `care_label`, `sizes`, `color`.
- **`men_ethnic`**: `garment_type`, `fabric`, `neckline_or_collar`, `sleeve`, `length`, `pattern`, `care_label`, `sizes`, `color`.
- **`kidswear`**: `target_gender`, `age_group`, `garment_type`, `fabric`, `pattern`, `care_label`, `sizes`, `color`.
- **`footwear`**: `footwear_type`, `material`, `closure_type`, `sole_material`, `toe_shape`, `heel_type_or_flat`, `care_label`, `sizes`, `color`.
- **`home_textiles`**: `product_type`, `material`, `pattern`, `dimensions`, `package_contents`, `care_label`, `color`.

### 2. Category-Specific Prohibited Claims
- **Footwear**: Forbids unverified `anti-slip`, `cushioning`, `arch support`, `memory foam`, `waterproof` without lab certification.
- **Sarees**: Forbids unverified `handloom`, `Banarasi`, `Kanjivaram`, `pure silk`, `artisanal` without Silk Mark / GI certification.
- **Home Textiles**: Forbids unverified thread count (e.g. `400 TC`), `GSM`, `hypoallergenic`, `stain-resistant` without test documentation.
- **Kidswear**: Forbids unverified `skin-safe`, `gentle on skin`, `100% organic` without dermatological / GOTS certification.
- **Apparel (Shirts & Bottomwear)**: Flags unverified `slim fit`, `tailored fit`, `comfort fit` unless verified in source specifications.

### 3. Title Template Formulas by Profile
- **Apparel**: `[Brand] [Gender]'s [Fabric] [Pattern] [Product Type] with [Verified Detail] ([Color])`
- **Sarees**: `[Brand] Women's [Fabric] [Pattern] [Saree Type] with [Blouse Piece Info] ([Color], [Saree Length])`
- **Co-ord Sets**: `[Brand] Women's [Fabric] [Pattern] 2-Piece Co-ord Set with [Top Type] and [Bottom Type] ([Color])`
- **Kidswear**: `[Brand] [Target Gender] [Age Group] [Fabric] [Pattern] [Garment Type] ([Color])`
- **Footwear**: `[Brand] [Gender] [Material] [Footwear Type] with [Closure Type] and [Sole Material] ([Color])`
- **Home Textiles**: `[Brand] [Material] [Pattern] [Product Type] with [Package Contents] ([Dimensions], [Color])`

---

## 📊 Marketplace Mapping Workbooks

Listing Factory v2.1 dynamically formats Excel workbooks per profile:

1. **`Master_Summary` (17 Columns)**:
   - `SKU ID`, `Brand`, `Category Profile`, `Category`, `Product Type`, `Color`, `Key Attributes / Fabric`, `Sizes / Dimensions`, `Amazon Title Preview`, `Flipkart Title Preview`, `Meesho Hinglish Hook Preview`, `Core Images Found`, `Core Coverage`, `Validation Status`, `Review Flags`, `Package Readiness`, **`Status Scope / Meaning`**.
2. **`01_Amazon_Bulk_Import`**:
   - Ingests profile-specific `item_type_keyword` and `feed_product_type`, `standard_price`, `quantity`, 5 bullet points, backend search terms ($\le 240$ bytes), and 6 canonical image URLs (`main_image_url`, `other_image_url1` through `other_image_url5`).
3. **`02_Flipkart_Bulk_Import` (Dynamic Columns)**:
   - Replaces hardcoded kurta columns with dynamic attribute columns derived from active profiles (e.g. Saree gets `Saree Length`, `Blouse Piece`, `Border Type`; Footwear gets `Footwear Type`, `Material`, `Sole Material`, `Closure Type`, `Toe Shape`, `Heel Type`).
4. **`03_Meesho_Bulk_Import`**:
   - Ingests Hinglish and English hook descriptions, primary material, dimensions, 4 highlight badges, dynamic seller pricing, and 6 image slots.
5. **`[Client]_Alternate_Listing_Copies.xlsx`**:
   - Contains 5 distinct marketing angle variations (V1 to V5) across Amazon, Flipkart, and Meesho.

---

## 📸 Canonical Image Naming Scheme (6 Declared Slots)

| Role | Canonical Filename Pattern | Declared Slot / Purpose | Verification Note |
|---|---|---|---|
| **Primary Hero** | `SKU_XX_MAIN.jpg` | Pure white background (`#FFFFFF`), primary product cutout | Declared Hero Image — manual visual check required |
| **Other Image 1** | `SKU_XX_PT01.jpg` | Size chart, measurement specifications, & fit guide | Declared Size Chart Slot — manual visual check required |
| **Other Image 2** | `SKU_XX_PT02.jpg` | Fabric texture, weave detail, & material spec | Declared Fabric Specification Slot — manual visual check required |
| **Other Image 3** | `SKU_XX_PT03.jpg` | Wash care instructions & styling recommendations | Declared Care Guide Slot — manual visual check required |
| **Other Image 4** | `SKU_XX_PT04.jpg` | Back view / alternate product angle | Declared Back View Slot — manual visual check required |
| **Other Image 5** | `SKU_XX_PT05.jpg` | Detail close-up / alternate lifestyle shot | Declared Other Image 5 Slot — manual visual check required |

---

## ⚖️ Scope & Operational Boundary

- **Listing Preparation Tool**: Prepares structured listing copy and multi-tab Excel mapping workbooks based on seller-provided catalog data.
- **No Acceptance Guarantee**: Does not guarantee automatic marketplace approval or exemption from category ungating requirements.
- **Declared Image Roles vs. Visual Verification**: Assigns image slots based on declared filename patterns (`_MAIN` through `_PT05`).
  > **Image Role Disclaimer**: *"Image roles are assigned from filenames only. Listing Factory does not visually verify image content. The seller must confirm that each declared slot contains the intended image before marketplace upload."*
- **Structural Readiness Disclaimer**:
  > *"Structural completeness confirms schema and package checks only. It does not verify product facts, image content, tax classification, marketplace-policy compliance, or marketplace acceptance. Seller review is required before upload."*

---

## 🛠️ Local Development & Reference Backend

```bash
pip install -r requirements.txt
python app.py
```

Runs at `http://127.0.0.1:8000`.
