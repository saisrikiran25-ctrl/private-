"""
Listing Factory: Multi-Marketplace Packaging Studio
====================================================
A FastAPI-powered local web application for e-commerce cataloging agencies.
Takes AI-generated JSON copy + loose image files, validates against
Amazon.in / Flipkart / Meesho constraints, organizes images into SKU
subfolders, auto-populates marketplace Excel flat files, and outputs
a ready-to-deliver client ZIP archive.

Run:  python app.py
Open: http://127.0.0.1:8000
"""

from __future__ import annotations

import io
import json
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, List, Optional

import pandas as pd
from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    NamedStyle,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field, field_validator
import uvicorn

# ──────────────────────────────────────────────
# Pydantic v2 Models
# ──────────────────────────────────────────────

class AmazonData(BaseModel):
    title: str = ""
    bullet_points: list[str] = Field(default_factory=list)
    backend_search_terms: str = ""
    description: str = ""


class FlipkartData(BaseModel):
    title: str = ""
    fabric: str = ""
    kurta_type: str = ""
    neck: str = ""
    sleeve: str = ""
    length_type: str = ""
    pattern: str = ""
    occasion: str = ""
    search_keywords: str = ""


class MeeshoData(BaseModel):
    title: str = ""
    hinglish_hook_description: str = ""
    highlights: list[str] = Field(default_factory=list)


class AlternateAmazonData(BaseModel):
    title: str = ""
    bullet_points: list[str] = Field(default_factory=list)
    backend_search_terms: str = ""
    description: str = ""


class AlternateFlipkartData(BaseModel):
    title: str = ""
    search_keywords: str = ""
    description: str = ""


class AlternateMeeshoData(BaseModel):
    title: str = ""
    hinglish_hook_description: str = ""
    english_hook_description: str = ""
    highlights: list[str] = Field(default_factory=list)


class AlternateVariant(BaseModel):
    variant_id: str = ""
    angle_theme: str = ""
    amazon: AlternateAmazonData = Field(default_factory=AlternateAmazonData)
    flipkart: AlternateFlipkartData = Field(default_factory=AlternateFlipkartData)
    meesho: AlternateMeeshoData = Field(default_factory=AlternateMeeshoData)


class SKUItem(BaseModel):
    sku_id: str
    brand: str = ""
    product_type: str = ""
    sizes: str = ""
    mrp: float = 0
    meesho_price: float = 0
    amazon: AmazonData = Field(default_factory=AmazonData)
    flipkart: FlipkartData = Field(default_factory=FlipkartData)
    meesho: MeeshoData = Field(default_factory=MeeshoData)
    alternates: list[AlternateVariant] = Field(default_factory=list)


# ──────────────────────────────────────────────
# Category mapping
# ──────────────────────────────────────────────
CATEGORY_MAP = {
    "Women Ethnic Wear": "kurtas-and-ethnic-tops",
    "Men Western Wear": "mens-casual-shirts",
    "Sarees": "sarees",
    "Footwear": "casual-shoes",
    "Home & Kitchen": "home-furnishing",
}

# Image role suffixes
IMAGE_ROLES = {
    "_MAIN": "Main Hero",
    "_PT01": "Size Chart",
    "_PT02": "Fabric Spec",
    "_PT03": "Care Guide",
    "_PT04": "Back View",
    "_PT05": "Lifestyle 1",
    "_PT06": "Lifestyle 2",
    "_PT07": "Detail Shot",
    "_PT08": "Packaging",
}

CORE_SUFFIXES = ["_MAIN", "_PT01", "_PT02", "_PT03"]


# ──────────────────────────────────────────────
# Excel Builder
# ──────────────────────────────────────────────

def _style_header_row(ws, num_cols: int):
    """Apply premium styling to the header row of a worksheet."""
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", bold=True, size=11, color="10B981")
    thin_border = Border(
        left=Side(style="thin", color="334155"),
        right=Side(style="thin", color="334155"),
        top=Side(style="thin", color="334155"),
        bottom=Side(style="thin", color="334155"),
    )
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _auto_width(ws, min_width=12, max_width=45):
    """Auto-fit column widths based on content."""
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            try:
                cell_len = len(str(cell.value)) if cell.value else 0
                max_len = max(max_len, cell_len)
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 4, max_width))


def _build_master_summary(wb: Workbook, skus: list[SKUItem], image_map: dict[str, list[str]]):
    ws = wb.active
    ws.title = "Master_Summary"
    headers = [
        "SKU ID", "Brand", "Product Type", "Fabric", "Sizes Available",
        "Amazon Title Preview", "Flipkart Title Preview", "Meesho Hook Preview",
        "Image Assets Count", "Upload Status",
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        img_count = len(image_map.get(sku.sku_id, []))
        has_core = all(
            any(suffix.lower() in fn.lower() for fn in image_map.get(sku.sku_id, []))
            for suffix in CORE_SUFFIXES
        )
        status = "✅ Ready" if has_core else "⚠️ Missing Images"
        ws.append([
            sku.sku_id,
            sku.brand,
            sku.product_type,
            sku.flipkart.fabric,
            sku.sizes,
            (sku.amazon.title[:80] + "…") if len(sku.amazon.title) > 80 else sku.amazon.title,
            (sku.flipkart.title[:80] + "…") if len(sku.flipkart.title) > 80 else sku.flipkart.title,
            (sku.meesho.hinglish_hook_description[:80] + "…") if len(sku.meesho.hinglish_hook_description) > 80 else sku.meesho.hinglish_hook_description,
            img_count,
            status,
        ])
    _auto_width(ws)


def _build_amazon_tab(wb: Workbook, skus: list[SKUItem], category: str):
    ws = wb.create_sheet("01_Amazon_Bulk_Import")
    item_type = CATEGORY_MAP.get(category, "generic-item")
    headers = [
        "Seller SKU", "Product Name (Title)", "Brand Name", "Item Type Keyword",
        "Standard Price (INR)", "Quantity",
        "Main Image URL / Name",
        "Other Image URL 1 (Size Chart)",
        "Other Image URL 2 (Fabric Spec)",
        "Other Image URL 3 (Care Guide)",
        "Other Image URL 4",
        "Key Product Features (Bullet 1)",
        "Key Product Features (Bullet 2)",
        "Key Product Features (Bullet 3)",
        "Key Product Features (Bullet 4)",
        "Key Product Features (Bullet 5)",
        "Generic Keywords (Backend Search)",
        "Product Description",
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        bullets = sku.amazon.bullet_points + [""] * 5  # pad
        # Enforce 240 byte cap on backend search terms
        bst = sku.amazon.backend_search_terms
        while len(bst.encode("utf-8")) > 240:
            bst = bst[:-1]
        ws.append([
            sku.sku_id,
            sku.amazon.title[:180],
            sku.brand,
            item_type,
            sku.mrp,
            50,
            f"{sku.sku_id}_MAIN.jpg",
            f"{sku.sku_id}_PT01_Size.jpg",
            f"{sku.sku_id}_PT02_Fabric.jpg",
            f"{sku.sku_id}_PT03_Care.jpg",
            f"{sku.sku_id}_PT04_Back.jpg",
            bullets[0],
            bullets[1],
            bullets[2],
            bullets[3],
            bullets[4],
            bst,
            sku.amazon.description,
        ])
    _auto_width(ws)


def _build_flipkart_tab(wb: Workbook, skus: list[SKUItem]):
    ws = wb.create_sheet("02_Flipkart_Bulk_Import")
    headers = [
        "Seller SKU ID", "Product Title", "Brand", "Style Code", "Size",
        "Pattern", "Type / Kurta Type", "Fabric", "Neck", "Sleeve",
        "Length Type", "Occasion", "Search Keywords",
        "Main Image Name", "Angle 1 Image", "Angle 2 Image",
        "Description",
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        ws.append([
            sku.sku_id,
            sku.flipkart.title,
            sku.brand,
            sku.sku_id,
            sku.sizes,
            sku.flipkart.pattern,
            sku.flipkart.kurta_type,
            sku.flipkart.fabric,
            sku.flipkart.neck,
            sku.flipkart.sleeve,
            sku.flipkart.length_type,
            sku.flipkart.occasion,
            sku.flipkart.search_keywords,
            f"{sku.sku_id}_MAIN.jpg",
            f"{sku.sku_id}_PT01_Size.jpg",
            f"{sku.sku_id}_PT02_Fabric.jpg",
            sku.amazon.description,
        ])
    _auto_width(ws)


def _build_meesho_tab(wb: Workbook, skus: list[SKUItem]):
    ws = wb.create_sheet("03_Meesho_Bulk_Import")
    headers = [
        "Product ID / SKU", "Product Name",
        "Product Description (Hinglish/Hindi Hook)",
        "Fabric", "Available Sizes", "Color", "Occasion",
        "Key Highlights", "GST (%)", "HSN Code",
        "Recommended Meesho Price (INR)",
        "Primary Image (Hero)",
        "Other Image 1 (Size Chart)",
        "Other Image 2 (Fabric Spec)",
        "Other Image 3 (Care Guide)",
        "Other Image 4 (Back View)",
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        highlights = " \u2022 ".join(sku.meesho.highlights)
        ws.append([
            sku.sku_id,
            sku.meesho.title,
            sku.meesho.hinglish_hook_description,
            sku.flipkart.fabric,
            sku.sizes,
            "",
            sku.flipkart.occasion or "",
            highlights,
            5,
            "62114200",
            sku.meesho_price,
            f"{sku.sku_id}_MAIN.jpg",
            f"{sku.sku_id}_PT01_Size.jpg",
            f"{sku.sku_id}_PT02_Fabric.jpg",
            f"{sku.sku_id}_PT03_Care.jpg",
            f"{sku.sku_id}_PT04_Back.jpg",
        ])
    _auto_width(ws)


def build_workbook(skus: list[SKUItem], category: str, image_map: dict[str, list[str]]) -> bytes:
    """Build the full multi-tab Excel workbook and return raw bytes."""
    wb = Workbook()
    _build_master_summary(wb, skus, image_map)
    _build_amazon_tab(wb, skus, category)
    _build_flipkart_tab(wb, skus)
    _build_meesho_tab(wb, skus)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_alternates_workbook(skus: list[SKUItem]) -> bytes:
    """Build the 5x Alternate Listing Copies workbook with 3 marketplace tabs."""
    wb = Workbook()

    # -- Tab 1: Amazon Alternates --
    ws_amz = wb.active
    ws_amz.title = "Amazon_Alternate_Copies"
    ws_amz.append([
        "SKU ID", "Brand", "Variant", "Marketing Angle / Theme",
        "Alternative Title (Amazon)",
        "Bullet 1 (Fabric & Comfort)", "Bullet 2 (Design & Utility)",
        "Bullet 3 (Colorfast & Durable)", "Bullet 4 (Versatile Styling)",
        "Bullet 5 (Sizing & Care)",
        "Alternative Backend Search Terms (<240 bytes)",
        "Alternative Description",
    ])
    _style_header_row(ws_amz, 12)

    # -- Tab 2: Flipkart Alternates --
    ws_fk = wb.create_sheet("Flipkart_Alternate_Copies")
    ws_fk.append([
        "SKU ID", "Brand", "Variant", "Marketing Angle / Theme",
        "Alternative Flipkart Title",
        "Alternative Search Keywords",
        "Alternative Product Description",
    ])
    _style_header_row(ws_fk, 7)

    # -- Tab 3: Meesho Alternates (Hinglish + English) --
    ws_me = wb.create_sheet("Meesho_Alternate_Copies")
    ws_me.append([
        "SKU ID", "Brand", "Variant", "Marketing Angle / Theme",
        "Alternative Product Title",
        "Hinglish Hook Description (Conversational)",
        "English Hook Description (Formal/Metro)",
        "Key Highlights (Badges)",
    ])
    _style_header_row(ws_me, 8)

    for sku in skus:
        for idx, alt in enumerate(sku.alternates):
            v_tag = alt.variant_id or f"V{idx+1}"
            theme = alt.angle_theme or f"Angle {idx+1}"

            # Amazon row
            bp = (alt.amazon.bullet_points + [""] * 5)[:5]
            ws_amz.append([
                sku.sku_id, sku.brand, v_tag, theme,
                alt.amazon.title,
                bp[0], bp[1], bp[2], bp[3], bp[4],
                alt.amazon.backend_search_terms,
                alt.amazon.description,
            ])

            # Flipkart row
            ws_fk.append([
                sku.sku_id, sku.brand, v_tag, theme,
                alt.flipkart.title,
                alt.flipkart.search_keywords,
                alt.flipkart.description,
            ])

            # Meesho row
            ws_me.append([
                sku.sku_id, sku.brand, v_tag, theme,
                alt.meesho.title,
                alt.meesho.hinglish_hook_description,
                alt.meesho.english_hook_description,
                " \u2022 ".join(alt.meesho.highlights),
            ])

    _auto_width(ws_amz)
    _auto_width(ws_fk)
    _auto_width(ws_me)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────
# README generator
# ──────────────────────────────────────────────

def generate_readme(client: str, batch: str) -> str:
    return f"""
================================================================================
  LISTING FACTORY HANDOVER PACKAGE: {client} -- {batch}
================================================================================

  Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
  Tool: Listing Factory v1.0 — Multi-Marketplace Packaging Studio

================================================================================
  HOW TO UPLOAD — 3-STEP GUIDE
================================================================================

  STEP 1 ▸  AMAZON SELLER CENTRAL
  ─────────────────────────────────
  1. Log in to sellercentral.amazon.in
  2. Navigate to ▸ Catalog ▸ Add Products via Upload
  3. Choose "Flat File" upload and select your category template
  4. Open the "01_Amazon_Bulk_Import" tab in the Master Excel file
  5. Copy-paste the rows into Amazon's template (match columns carefully)
  6. Upload images from each SKU folder to your product listing
  7. Submit and wait for processing (usually 15-30 minutes)

  STEP 2 ▸  FLIPKART SELLER HUB
  ─────────────────────────────────
  1. Log in to seller.flipkart.com
  2. Navigate to ▸ Listings ▸ Add in Bulk
  3. Download the Flipkart bulk listing template for your category
  4. Open the "02_Flipkart_Bulk_Import" tab in the Master Excel file
  5. Transfer data into Flipkart's template
  6. Upload the filled template + images
  7. Submit for QC review (usually 24-48 hours)

  STEP 3 ▸  MEESHO SUPPLIER PANEL
  ─────────────────────────────────
  1. Log in to supplier.meesho.com
  2. Navigate to ▸ Catalog ▸ Add Catalog
  3. Select the product category
  4. Open the "03_Meesho_Bulk_Import" tab in the Master Excel file
  5. Fill in the Meesho catalog form or bulk upload using the data
  6. Upload primary images from each SKU folder
  7. Submit and wait for approval (usually 2-4 hours)

================================================================================
  FOLDER STRUCTURE
================================================================================

  Organized_SKU_Images/
    ├── [SKU_ID]/            ← One folder per SKU
    │   ├── [SKU]_MAIN.jpg   ← Primary product image (hero cutout)
    │   ├── [SKU]_PT01*.jpg  ← Size & fit chart
    │   ├── [SKU]_PT02*.jpg  ← Fabric & feature spec
    │   ├── [SKU]_PT03*.jpg  ← Wash-care & styling guide
    │   └── [SKU]_PT04+.jpg  ← Additional angles / lifestyle
    └── Unassigned_Assets/   ← Images that didn't match any SKU

================================================================================
  NOTES
================================================================================

  • All titles, descriptions, and keywords are pre-optimized for search.
  • Amazon backend search terms are capped at ≤240 bytes.
  • Amazon titles are capped at ≤180 characters.
  • Image filenames follow marketplace naming conventions.
  • If any images are in "Unassigned_Assets", manually assign them.

  For support, contact your cataloging agency team.
================================================================================
"""


# ──────────────────────────────────────────────
# Image Router
# ──────────────────────────────────────────────

def route_images(
    sku_ids: list[str],
    image_files: list[tuple[str, bytes]],
) -> tuple[dict[str, list[tuple[str, bytes]]], list[tuple[str, bytes]]]:
    """
    Routes image files to their SKU folders based on prefix matching.
    Returns (matched_map, unassigned_list).
    """
    matched: dict[str, list[tuple[str, bytes]]] = {sid: [] for sid in sku_ids}
    unassigned: list[tuple[str, bytes]] = []

    # Sort SKU IDs longest-first so more specific IDs match first
    sorted_ids = sorted(sku_ids, key=len, reverse=True)

    for fname, data in image_files:
        stem = Path(fname).stem.upper()
        found = False
        for sid in sorted_ids:
            if stem.startswith(sid.upper()):
                matched[sid].append((fname, data))
                found = True
                break
        if not found:
            unassigned.append((fname, data))

    return matched, unassigned


# ──────────────────────────────────────────────
# ZIP Builder
# ──────────────────────────────────────────────

def build_zip(
    client: str,
    batch: str,
    category: str,
    skus: list[SKUItem],
    image_files: list[tuple[str, bytes]],
) -> bytes:
    """Build the full delivery ZIP archive in-memory."""
    sku_ids = [s.sku_id for s in skus]
    matched_images, unassigned_images = route_images(sku_ids, image_files)

    # Build image map for workbook (just filenames per SKU)
    image_map: dict[str, list[str]] = {}
    for sid, files in matched_images.items():
        image_map[sid] = [f[0] for f in files]

    xlsx_bytes = build_workbook(skus, category, image_map)
    readme_text = generate_readme(client, batch)

    # Build alternates workbook if any SKU has alternates
    has_alternates = any(len(s.alternates) > 0 for s in skus)
    alt_xlsx_bytes = build_alternates_workbook(skus) if has_alternates else None

    prefix = f"{client}_{batch}_Handover_Package"
    xlsx_name = f"{client}_Master_Marketplace_Upload.xlsx"

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Excel workbook
        zf.writestr(f"{prefix}/{xlsx_name}", xlsx_bytes)
        # Alternate copies workbook
        if alt_xlsx_bytes:
            alt_xlsx_name = f"{client}_Alternate_Listing_Copies.xlsx"
            zf.writestr(f"{prefix}/{alt_xlsx_name}", alt_xlsx_bytes)
        # README
        zf.writestr(f"{prefix}/README_Upload_Instructions.txt", readme_text)
        # Organized images
        for sid, files in matched_images.items():
            for fname, data in files:
                zf.writestr(f"{prefix}/Organized_SKU_Images/{sid}/{fname}", data)
        # Unassigned
        for fname, data in unassigned_images:
            zf.writestr(f"{prefix}/Organized_SKU_Images/Unassigned_Assets/{fname}", data)

    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────
# FastAPI Application
# ──────────────────────────────────────────────

app = FastAPI(title="Listing Factory", version="1.0.0")


# ── Validation endpoint ──

@app.post("/api/validate-json")
async def validate_json(request: Request):
    """Validate pasted/uploaded JSON and return parsed SKU info."""
    try:
        body = await request.json()
        raw = body.get("json_text", "")
        data = json.loads(raw)
        if not isinstance(data, list):
            data = [data]
        skus = [SKUItem(**item) for item in data]

        # Validation warnings
        warnings = []
        for sku in skus:
            if len(sku.amazon.title) > 180:
                warnings.append(f"{sku.sku_id}: Amazon title exceeds 180 chars ({len(sku.amazon.title)})")
            bst_bytes = len(sku.amazon.backend_search_terms.encode("utf-8"))
            if bst_bytes > 240:
                warnings.append(f"{sku.sku_id}: Backend search terms exceed 240 bytes ({bst_bytes})")
            # Flipkart attribute checks
            for attr in ["fabric", "kurta_type", "neck", "sleeve", "pattern", "occasion"]:
                if not getattr(sku.flipkart, attr, ""):
                    warnings.append(f"{sku.sku_id}: Flipkart '{attr}' is empty")
            # Meesho checks
            if not sku.meesho.hinglish_hook_description:
                warnings.append(f"{sku.sku_id}: Meesho Hinglish hook is empty")

        sku_summaries = []
        for sku in skus:
            sku_summaries.append({
                "sku_id": sku.sku_id,
                "brand": sku.brand,
                "product_type": sku.product_type,
                "amazon_title_len": len(sku.amazon.title),
                "bst_bytes": len(sku.amazon.backend_search_terms.encode("utf-8")),
                "bullet_count": len(sku.amazon.bullet_points),
            })

        return JSONResponse({
            "valid": True,
            "sku_count": len(skus),
            "skus": sku_summaries,
            "warnings": warnings,
        })

    except json.JSONDecodeError as e:
        return JSONResponse({"valid": False, "error": f"Invalid JSON syntax: {e}"}, status_code=400)
    except Exception as e:
        return JSONResponse({"valid": False, "error": str(e)}, status_code=400)


# ── Generation endpoint ──

@app.post("/api/generate")
async def generate_package(
    client_name: str = Form(...),
    batch_id: str = Form(...),
    category: str = Form(...),
    json_data: str = Form(...),
    images: list[UploadFile] = File(default=[]),
):
    """Generate the full delivery ZIP package."""
    try:
        raw = json.loads(json_data)
        if not isinstance(raw, list):
            raw = [raw]
        skus = [SKUItem(**item) for item in raw]
    except Exception as e:
        return JSONResponse({"error": f"JSON parse error: {e}"}, status_code=400)

    # Read all image files into memory
    image_files: list[tuple[str, bytes]] = []
    for img in images:
        data = await img.read()
        image_files.append((img.filename, data))

    # Sanitize names
    safe_client = re.sub(r"[^\w\-]", "_", client_name)
    safe_batch = re.sub(r"[^\w\-]", "_", batch_id)

    zip_bytes = build_zip(safe_client, safe_batch, category, skus, image_files)
    zip_name = f"{safe_client}_{safe_batch}_Handover_Package.zip"

    return StreamingResponse(
        io.BytesIO(zip_bytes),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{zip_name}"'},
    )


# ── Frontend — reads index.html from disk (single source of truth with GitHub Pages) ──

@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = Path(__file__).parent / "index.html"
    if html_path.exists():
        return html_path.read_text(encoding="utf-8")
    return HTMLResponse("<h1>index.html not found. Run from the project directory.</h1>", status_code=500)




# ──────────────────────────────────────────────
# Entry Point
# ──────────────────────────────────────────────

if __name__ == "__main__":
    print("\n  [*] Listing Factory - Multi-Marketplace Packaging Studio")
    print("  " + "=" * 57)
    print("  > Server running at: http://127.0.0.1:8000")
    print("  > Press Ctrl+C to stop\n")
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
