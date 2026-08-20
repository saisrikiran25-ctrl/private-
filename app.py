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
        "Fabric", "Available Sizes", "Key Highlights",
        "Recommended Meesho Price (INR)", "Primary Image",
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        highlights = " • ".join(sku.meesho.highlights)
        ws.append([
            sku.sku_id,
            sku.meesho.title,
            sku.meesho.hinglish_hook_description,
            sku.flipkart.fabric,
            sku.sizes,
            highlights,
            sku.meesho_price,
            f"{sku.sku_id}_MAIN.jpg",
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


# ──────────────────────────────────────────────
# README generator
# ──────────────────────────────────────────────

def generate_readme(client: str, batch: str) -> str:
    return f"""
================================================================================
  📦  {client} — {batch}  |  LISTING FACTORY HANDOVER PACKAGE
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

    prefix = f"{client}_{batch}_Handover_Package"
    xlsx_name = f"{client}_Master_Marketplace_Upload.xlsx"

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Excel workbook
        zf.writestr(f"{prefix}/{xlsx_name}", xlsx_bytes)
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


# ── Frontend ──

@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_TEMPLATE


# ──────────────────────────────────────────────
# Inline Frontend Template
# ──────────────────────────────────────────────

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en" class="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Listing Factory — Multi-Marketplace Packaging Studio</title>
  <meta name="description" content="Automated packaging, validation, and delivery engine for Indian e-commerce cataloging. Supports Amazon.in, Flipkart, and Meesho bulk uploads.">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwindcss.config = {
      darkMode: 'class',
      theme: {
        extend: {
          fontFamily: {
            sans: ['"Plus Jakarta Sans"', 'system-ui', 'sans-serif'],
            mono: ['"JetBrains Mono"', 'monospace'],
          },
          colors: {
            brand: {
              50: '#ecfdf5', 100: '#d1fae5', 200: '#a7f3d0', 300: '#6ee7b7',
              400: '#34d399', 500: '#10b981', 600: '#059669', 700: '#047857',
            }
          },
          animation: {
            'pulse-slow': 'pulse 3s cubic-bezier(0.4, 0, 0.6, 1) infinite',
            'float': 'float 6s ease-in-out infinite',
            'shimmer': 'shimmer 2s linear infinite',
            'slide-up': 'slideUp 0.5s ease-out',
            'fade-in': 'fadeIn 0.3s ease-out',
          },
          keyframes: {
            float: {
              '0%, 100%': { transform: 'translateY(0px)' },
              '50%': { transform: 'translateY(-10px)' },
            },
            shimmer: {
              '0%': { backgroundPosition: '-200% 0' },
              '100%': { backgroundPosition: '200% 0' },
            },
            slideUp: {
              '0%': { opacity: '0', transform: 'translateY(20px)' },
              '100%': { opacity: '1', transform: 'translateY(0)' },
            },
            fadeIn: {
              '0%': { opacity: '0' },
              '100%': { opacity: '1' },
            }
          }
        }
      }
    };
  </script>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
      background: #020617;
      color: #e2e8f0;
      min-height: 100vh;
    }

    /* Glassmorphism panels */
    .glass-panel {
      background: rgba(15, 23, 42, 0.6);
      backdrop-filter: blur(20px);
      -webkit-backdrop-filter: blur(20px);
      border: 1px solid rgba(51, 65, 85, 0.5);
      border-radius: 16px;
    }
    .glass-panel-lighter {
      background: rgba(30, 41, 59, 0.4);
      backdrop-filter: blur(12px);
      border: 1px solid rgba(71, 85, 105, 0.3);
      border-radius: 12px;
    }

    /* Gradient mesh background */
    .bg-mesh {
      position: fixed;
      inset: 0;
      z-index: 0;
      overflow: hidden;
      pointer-events: none;
    }
    .bg-mesh::before {
      content: '';
      position: absolute;
      width: 600px; height: 600px;
      background: radial-gradient(circle, rgba(16,185,129,0.08) 0%, transparent 70%);
      top: -200px; right: -100px;
      animation: float 8s ease-in-out infinite;
    }
    .bg-mesh::after {
      content: '';
      position: absolute;
      width: 500px; height: 500px;
      background: radial-gradient(circle, rgba(6,182,212,0.06) 0%, transparent 70%);
      bottom: -100px; left: -100px;
      animation: float 10s ease-in-out infinite reverse;
    }

    /* Custom scrollbar */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #0f172a; }
    ::-webkit-scrollbar-thumb { background: #334155; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #475569; }

    /* Textarea code styling */
    .json-editor {
      font-family: 'JetBrains Mono', monospace;
      font-size: 13px;
      line-height: 1.6;
      tab-size: 2;
      background: rgba(2, 6, 23, 0.8);
      border: 1px solid rgba(51, 65, 85, 0.6);
      border-radius: 12px;
      color: #a5f3fc;
      resize: vertical;
      transition: border-color 0.3s ease, box-shadow 0.3s ease;
    }
    .json-editor:focus {
      outline: none;
      border-color: #10b981;
      box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.15), 0 0 20px rgba(16, 185, 129, 0.05);
    }
    .json-editor.invalid {
      border-color: #ef4444;
      box-shadow: 0 0 0 3px rgba(239, 68, 68, 0.15);
    }
    .json-editor.valid {
      border-color: #10b981;
      box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.15);
    }

    /* Dropzone */
    .dropzone {
      border: 2px dashed rgba(71, 85, 105, 0.5);
      border-radius: 16px;
      transition: all 0.3s ease;
      position: relative;
      overflow: hidden;
    }
    .dropzone::before {
      content: '';
      position: absolute;
      inset: 0;
      background: linear-gradient(135deg, rgba(16,185,129,0.02) 0%, transparent 50%);
      opacity: 0;
      transition: opacity 0.3s ease;
    }
    .dropzone.drag-over {
      border-color: #10b981;
      background: rgba(16, 185, 129, 0.05);
      transform: scale(1.005);
    }
    .dropzone.drag-over::before { opacity: 1; }

    /* SKU card */
    .sku-card {
      background: rgba(15, 23, 42, 0.5);
      border: 1px solid rgba(51, 65, 85, 0.4);
      border-radius: 12px;
      padding: 14px;
      transition: all 0.3s ease;
    }
    .sku-card:hover {
      border-color: rgba(16, 185, 129, 0.3);
      transform: translateY(-2px);
      box-shadow: 0 8px 25px rgba(0,0,0,0.3);
    }
    .sku-card.complete { border-color: rgba(16, 185, 129, 0.4); }
    .sku-card.incomplete { border-color: rgba(234, 179, 8, 0.4); }

    /* Badge styles */
    .badge {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 3px 8px;
      border-radius: 6px;
      font-size: 11px;
      font-weight: 600;
      letter-spacing: 0.02em;
    }
    .badge-present {
      background: rgba(16, 185, 129, 0.15);
      color: #34d399;
      border: 1px solid rgba(16, 185, 129, 0.3);
    }
    .badge-missing {
      background: rgba(100, 116, 139, 0.15);
      color: #94a3b8;
      border: 1px solid rgba(100, 116, 139, 0.3);
    }

    /* Generate button */
    .btn-generate {
      position: relative;
      background: linear-gradient(135deg, #059669 0%, #10b981 50%, #34d399 100%);
      color: white;
      font-weight: 700;
      font-size: 16px;
      padding: 16px 40px;
      border-radius: 14px;
      border: none;
      cursor: pointer;
      overflow: hidden;
      transition: all 0.3s ease;
      text-transform: none;
      letter-spacing: 0.01em;
    }
    .btn-generate::before {
      content: '';
      position: absolute;
      inset: 0;
      background: linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent);
      background-size: 200% 100%;
      opacity: 0;
      transition: opacity 0.3s ease;
    }
    .btn-generate:hover {
      transform: translateY(-2px);
      box-shadow: 0 12px 35px rgba(16, 185, 129, 0.3);
    }
    .btn-generate:hover::before {
      opacity: 1;
      animation: shimmer 1.5s linear infinite;
    }
    .btn-generate:active { transform: translateY(0); }
    .btn-generate:disabled {
      opacity: 0.5;
      cursor: not-allowed;
      transform: none !important;
      box-shadow: none !important;
    }

    /* Loading spinner */
    .spinner {
      width: 20px; height: 20px;
      border: 2.5px solid rgba(255,255,255,0.3);
      border-top-color: white;
      border-radius: 50%;
      animation: spin 0.7s linear infinite;
      display: inline-block;
      vertical-align: middle;
      margin-right: 8px;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    /* Warning badge */
    .warning-item {
      background: rgba(234, 179, 8, 0.1);
      border: 1px solid rgba(234, 179, 8, 0.2);
      border-radius: 8px;
      padding: 8px 12px;
      font-size: 13px;
      color: #fbbf24;
    }

    /* Section header */
    .section-number {
      width: 32px; height: 32px;
      background: linear-gradient(135deg, #059669, #10b981);
      border-radius: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: 800;
      font-size: 14px;
      color: white;
      flex-shrink: 0;
    }

    /* Input styling */
    .input-field {
      background: rgba(2, 6, 23, 0.6);
      border: 1px solid rgba(51, 65, 85, 0.5);
      border-radius: 10px;
      padding: 10px 14px;
      color: #e2e8f0;
      font-size: 14px;
      transition: all 0.3s ease;
      width: 100%;
    }
    .input-field:focus {
      outline: none;
      border-color: #10b981;
      box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.12);
    }
    .input-field::placeholder { color: #475569; }

    select.input-field {
      appearance: none;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 12px center;
      padding-right: 36px;
    }

    /* Tooltip */
    .tooltip { position: relative; }
    .tooltip::after {
      content: attr(data-tip);
      position: absolute;
      bottom: 100%;
      left: 50%;
      transform: translateX(-50%);
      background: #1e293b;
      color: #e2e8f0;
      padding: 6px 10px;
      border-radius: 6px;
      font-size: 11px;
      white-space: nowrap;
      opacity: 0;
      pointer-events: none;
      transition: opacity 0.2s ease;
      border: 1px solid #334155;
    }
    .tooltip:hover::after { opacity: 1; }

    /* Stats counter */
    .stat-value {
      font-size: 28px;
      font-weight: 800;
      background: linear-gradient(135deg, #10b981, #34d399);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }
  </style>
</head>
<body class="antialiased">
  <!-- Background mesh -->
  <div class="bg-mesh"></div>

  <!-- Main container -->
  <div class="relative z-10 min-h-screen">
    <!-- Header -->
    <header class="border-b border-slate-800/60 bg-slate-950/80 backdrop-blur-xl sticky top-0 z-50">
      <div class="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-500 to-teal-400 flex items-center justify-center shadow-lg shadow-emerald-500/20">
            <svg class="w-5 h-5 text-white" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="m20.25 7.5-.625 10.632a2.25 2.25 0 0 1-2.247 2.118H6.622a2.25 2.25 0 0 1-2.247-2.118L3.75 7.5m8.25 3v6.75m0 0-3-3m3 3 3-3M3.375 7.5h17.25c.621 0 1.125-.504 1.125-1.125v-1.5c0-.621-.504-1.125-1.125-1.125H3.375c-.621 0-1.125.504-1.125 1.125v1.5c0 .621.504 1.125 1.125 1.125Z"/></svg>
          </div>
          <div>
            <h1 class="text-lg font-bold text-white tracking-tight">Listing Factory</h1>
            <p class="text-xs text-slate-400 font-medium">Multi-Marketplace Packaging Studio</p>
          </div>
        </div>
        <div class="flex items-center gap-3">
          <div class="hidden md:flex items-center gap-2 text-xs text-slate-500">
            <span class="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
            Local Engine Ready
          </div>
          <div class="px-3 py-1.5 rounded-lg bg-slate-800/60 border border-slate-700/50 text-xs font-mono text-slate-400">
            v1.0
          </div>
        </div>
      </div>
    </header>

    <!-- Content -->
    <main class="max-w-7xl mx-auto px-6 py-8 space-y-8">

      <!-- ═══════════════ SECTION 1: Configuration ═══════════════ -->
      <section class="glass-panel p-6 animate-slide-up">
        <div class="flex items-center gap-3 mb-6">
          <div class="section-number">1</div>
          <div>
            <h2 class="text-base font-bold text-white">Client & Batch Configuration</h2>
            <p class="text-xs text-slate-400 mt-0.5">Set up the delivery metadata for this packaging run</p>
          </div>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label class="block text-xs font-semibold text-slate-300 mb-2 uppercase tracking-wider">Client / Brand Name</label>
            <input type="text" id="clientName" class="input-field" placeholder="e.g., Anvi_Fabrics" value="">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-300 mb-2 uppercase tracking-wider">Batch Identifier</label>
            <input type="text" id="batchId" class="input-field" placeholder="e.g., Batch_01_50SKU" value="">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-300 mb-2 uppercase tracking-wider">Product Category</label>
            <select id="category" class="input-field">
              <option value="">Select category…</option>
              <option value="Women Ethnic Wear">Women Ethnic Wear</option>
              <option value="Men Western Wear">Men Western Wear</option>
              <option value="Sarees">Sarees</option>
              <option value="Footwear">Footwear</option>
              <option value="Home & Kitchen">Home &amp; Kitchen</option>
            </select>
          </div>
        </div>
      </section>

      <!-- ═══════════════ SECTION 2: JSON Input ═══════════════ -->
      <section class="glass-panel p-6 animate-slide-up" style="animation-delay: 0.1s">
        <div class="flex items-center justify-between mb-6">
          <div class="flex items-center gap-3">
            <div class="section-number">2</div>
            <div>
              <h2 class="text-base font-bold text-white">AI-Generated JSON Ingestion</h2>
              <p class="text-xs text-slate-400 mt-0.5">Paste or upload your Claude / Gemini JSON output</p>
            </div>
          </div>
          <div class="flex items-center gap-2">
            <!-- JSON status badge -->
            <div id="jsonBadge" class="badge badge-missing">
              <svg class="w-3 h-3" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/></svg>
              <span id="jsonBadgeText">No JSON</span>
            </div>
          </div>
        </div>

        <!-- Mode tabs -->
        <div class="flex gap-1 mb-4 p-1 bg-slate-900/50 rounded-lg w-fit">
          <button onclick="setJsonMode('paste')" id="tabPaste" class="px-4 py-2 text-xs font-semibold rounded-md transition-all bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
            ✏️ Paste JSON
          </button>
          <button onclick="setJsonMode('file')" id="tabFile" class="px-4 py-2 text-xs font-semibold rounded-md transition-all text-slate-400 hover:text-slate-300">
            📁 Upload File
          </button>
        </div>

        <!-- Paste mode -->
        <div id="pasteMode">
          <textarea id="jsonTextarea" class="json-editor w-full h-64 p-4" spellcheck="false"
            placeholder='Paste your JSON array here...

[
  {
    "sku_id": "SKU_01",
    "brand": "Anvi Fabrics",
    "product_type": "Anarkali Kurti",
    ...
  }
]'></textarea>
        </div>

        <!-- File mode -->
        <div id="fileMode" class="hidden">
          <div class="border-2 border-dashed border-slate-700/50 rounded-xl p-8 text-center cursor-pointer hover:border-emerald-500/40 transition-all"
               onclick="document.getElementById('jsonFileInput').click()">
            <svg class="w-10 h-10 mx-auto mb-3 text-slate-500" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m6.75 12-3-3m0 0-3 3m3-3v6m-1.5-15H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z"/>
            </svg>
            <p class="text-sm text-slate-400 font-medium">Click to upload <span class="text-emerald-400">.json</span> file</p>
            <p class="text-xs text-slate-500 mt-1">Or drag and drop your JSON file here</p>
            <input type="file" id="jsonFileInput" accept=".json" class="hidden" onchange="handleJsonFile(event)">
          </div>
          <div id="jsonFileName" class="hidden mt-3 px-4 py-2 bg-emerald-500/10 border border-emerald-500/20 rounded-lg text-sm text-emerald-400 flex items-center gap-2">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="m4.5 12.75 6 6 9-13.5"/></svg>
            <span id="jsonFileLabel"></span>
          </div>
        </div>

        <!-- Warnings -->
        <div id="warningsPanel" class="hidden mt-4 space-y-2">
          <div class="flex items-center gap-2 text-xs font-semibold text-amber-400 mb-2">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126ZM12 15.75h.007v.008H12v-.008Z"/></svg>
            Validation Warnings
          </div>
          <div id="warningsList" class="space-y-1.5"></div>
        </div>

        <!-- SKU summary stats -->
        <div id="skuStats" class="hidden mt-5 grid grid-cols-2 sm:grid-cols-4 gap-3">
        </div>
      </section>

      <!-- ═══════════════ SECTION 3: Image Upload ═══════════════ -->
      <section class="glass-panel p-6 animate-slide-up" style="animation-delay: 0.2s">
        <div class="flex items-center justify-between mb-6">
          <div class="flex items-center gap-3">
            <div class="section-number">3</div>
            <div>
              <h2 class="text-base font-bold text-white">Image Assets Drop Zone</h2>
              <p class="text-xs text-slate-400 mt-0.5">Drag & drop your product images — auto-matched to SKUs by filename prefix</p>
            </div>
          </div>
          <div id="imageCounter" class="badge badge-missing">
            <span>0 images</span>
          </div>
        </div>

        <div id="imageDropzone" class="dropzone p-10 text-center cursor-pointer transition-all"
             onclick="document.getElementById('imageInput').click()">
          <input type="file" id="imageInput" multiple accept=".jpg,.jpeg,.png,.webp" class="hidden" onchange="handleImages(event)">
          <div id="dropzoneContent">
            <div class="w-16 h-16 mx-auto mb-4 rounded-2xl bg-gradient-to-br from-slate-800 to-slate-700 flex items-center justify-center">
              <svg class="w-8 h-8 text-slate-400" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="m2.25 15.75 5.159-5.159a2.25 2.25 0 0 1 3.182 0l5.159 5.159m-1.5-1.5 1.409-1.409a2.25 2.25 0 0 1 3.182 0l2.909 2.909M3.75 21h16.5A2.25 2.25 0 0 0 22.5 18.75V5.25A2.25 2.25 0 0 0 20.25 3H3.75A2.25 2.25 0 0 0 1.5 5.25v13.5A2.25 2.25 0 0 0 3.75 21Z"/>
              </svg>
            </div>
            <p class="text-sm font-semibold text-slate-300">Drop images here or click to browse</p>
            <p class="text-xs text-slate-500 mt-2">Supports <span class="text-slate-400">.jpg .jpeg .png .webp</span> — Name format: <span class="font-mono text-emerald-400/80">SKU_01_MAIN.jpg</span></p>
          </div>
        </div>

        <!-- Live SKU image matching grid -->
        <div id="skuImageGrid" class="hidden mt-6 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
        </div>

        <!-- Unassigned images -->
        <div id="unassignedPanel" class="hidden mt-4 glass-panel-lighter p-4">
          <div class="flex items-center gap-2 text-xs font-semibold text-amber-400 mb-2">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126ZM12 15.75h.007v.008H12v-.008Z"/></svg>
            Unassigned Images
          </div>
          <div id="unassignedList" class="flex flex-wrap gap-2 text-xs text-slate-400"></div>
        </div>
      </section>

      <!-- ═══════════════ SECTION 4: Pre-Flight & Generate ═══════════════ -->
      <section class="glass-panel p-6 animate-slide-up" style="animation-delay: 0.3s">
        <div class="flex items-center gap-3 mb-6">
          <div class="section-number">4</div>
          <div>
            <h2 class="text-base font-bold text-white">Pre-Flight Check & Generate</h2>
            <p class="text-xs text-slate-400 mt-0.5">Review validation status and generate your delivery package</p>
          </div>
        </div>

        <!-- Pre-flight checklist -->
        <div id="preflightChecks" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mb-6">
          <div class="glass-panel-lighter p-4 flex items-center gap-3">
            <div id="checkConfig" class="w-8 h-8 rounded-lg bg-slate-800 flex items-center justify-center text-slate-500 transition-all">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/></svg>
            </div>
            <div>
              <p class="text-xs font-semibold text-slate-300">Configuration</p>
              <p class="text-xs text-slate-500" id="checkConfigText">Not set</p>
            </div>
          </div>
          <div class="glass-panel-lighter p-4 flex items-center gap-3">
            <div id="checkJson" class="w-8 h-8 rounded-lg bg-slate-800 flex items-center justify-center text-slate-500 transition-all">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/></svg>
            </div>
            <div>
              <p class="text-xs font-semibold text-slate-300">JSON Data</p>
              <p class="text-xs text-slate-500" id="checkJsonText">No data</p>
            </div>
          </div>
          <div class="glass-panel-lighter p-4 flex items-center gap-3">
            <div id="checkImages" class="w-8 h-8 rounded-lg bg-slate-800 flex items-center justify-center text-slate-500 transition-all">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/></svg>
            </div>
            <div>
              <p class="text-xs font-semibold text-slate-300">Images</p>
              <p class="text-xs text-slate-500" id="checkImagesText">None uploaded</p>
            </div>
          </div>
          <div class="glass-panel-lighter p-4 flex items-center gap-3">
            <div id="checkWarnings" class="w-8 h-8 rounded-lg bg-slate-800 flex items-center justify-center text-slate-500 transition-all">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/></svg>
            </div>
            <div>
              <p class="text-xs font-semibold text-slate-300">Warnings</p>
              <p class="text-xs text-slate-500" id="checkWarningsText">—</p>
            </div>
          </div>
        </div>

        <!-- Generate button -->
        <div class="flex flex-col items-center gap-4 pt-2">
          <button id="generateBtn" class="btn-generate" onclick="generatePackage()" disabled>
            <span id="btnContent">
              <svg class="w-5 h-5 inline-block mr-2 -mt-0.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="m20.25 7.5-.625 10.632a2.25 2.25 0 0 1-2.247 2.118H6.622a2.25 2.25 0 0 1-2.247-2.118L3.75 7.5m8.25 3v6.75m0 0-3-3m3 3 3-3M3.375 7.5h17.25c.621 0 1.125-.504 1.125-1.125v-1.5c0-.621-.504-1.125-1.125-1.125H3.375c-.621 0-1.125.504-1.125 1.125v1.5c0 .621.504 1.125 1.125 1.125Z"/></svg>
              Generate &amp; Download Delivery Package (ZIP)
            </span>
            <span id="btnLoading" class="hidden">
              <span class="spinner"></span>
              Generating Package…
            </span>
          </button>
          <p class="text-xs text-slate-500">Builds Excel workbook + organizes images + creates ZIP — all in-memory</p>
        </div>
      </section>

      <!-- Footer -->
      <footer class="text-center py-8 text-xs text-slate-600">
        <p>Listing Factory v1.0 — Built for speed. Designed for cataloging agencies.</p>
        <p class="mt-1">Amazon.in · Flipkart · Meesho — All marketplace schemas supported</p>
      </footer>
    </main>
  </div>

  <!-- ═══════════════ JavaScript Engine ═══════════════ -->
  <script>
    // ── State ──
    let currentJsonText = '';
    let parsedSkus = [];
    let validationWarnings = [];
    let imageFiles = [];  // File objects
    let imageMap = {};    // { sku_id: [filename, ...] }
    let unassignedImages = [];
    let jsonValid = false;
    let jsonMode = 'paste';

    // ── JSON Mode Toggle ──
    function setJsonMode(mode) {
      jsonMode = mode;
      const tabPaste = document.getElementById('tabPaste');
      const tabFile = document.getElementById('tabFile');
      const pasteDiv = document.getElementById('pasteMode');
      const fileDiv = document.getElementById('fileMode');

      if (mode === 'paste') {
        tabPaste.className = 'px-4 py-2 text-xs font-semibold rounded-md transition-all bg-emerald-500/20 text-emerald-400 border border-emerald-500/30';
        tabFile.className = 'px-4 py-2 text-xs font-semibold rounded-md transition-all text-slate-400 hover:text-slate-300';
        pasteDiv.classList.remove('hidden');
        fileDiv.classList.add('hidden');
      } else {
        tabFile.className = 'px-4 py-2 text-xs font-semibold rounded-md transition-all bg-emerald-500/20 text-emerald-400 border border-emerald-500/30';
        tabPaste.className = 'px-4 py-2 text-xs font-semibold rounded-md transition-all text-slate-400 hover:text-slate-300';
        fileDiv.classList.remove('hidden');
        pasteDiv.classList.add('hidden');
      }
    }

    // ── JSON File Upload ──
    function handleJsonFile(e) {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        currentJsonText = ev.target.result;
        document.getElementById('jsonFileName').classList.remove('hidden');
        document.getElementById('jsonFileLabel').textContent = file.name;
        validateJson();
      };
      reader.readAsText(file);
    }

    // ── JSON Textarea Live Validation ──
    const jsonTextarea = document.getElementById('jsonTextarea');
    let debounceTimer;
    jsonTextarea.addEventListener('input', () => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        currentJsonText = jsonTextarea.value;
        validateJson();
      }, 400);
    });

    async function validateJson() {
      if (!currentJsonText.trim()) {
        setJsonBadge(false, 'No JSON');
        jsonValid = false;
        parsedSkus = [];
        validationWarnings = [];
        hideSkuStats();
        hideWarnings();
        updatePreflightChecks();
        refreshImageGrid();
        return;
      }

      try {
        const resp = await fetch('/api/validate-json', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ json_text: currentJsonText }),
        });
        const data = await resp.json();

        if (data.valid) {
          jsonValid = true;
          parsedSkus = data.skus;
          validationWarnings = data.warnings || [];
          setJsonBadge(true, `${data.sku_count} SKU${data.sku_count > 1 ? 's' : ''} parsed`);
          if (jsonMode === 'paste') {
            jsonTextarea.classList.remove('invalid');
            jsonTextarea.classList.add('valid');
          }
          showSkuStats(data);
          showWarnings(validationWarnings);
          refreshImageGrid();
        } else {
          jsonValid = false;
          parsedSkus = [];
          setJsonBadge(false, 'Invalid JSON');
          if (jsonMode === 'paste') {
            jsonTextarea.classList.remove('valid');
            jsonTextarea.classList.add('invalid');
          }
          hideSkuStats();
          hideWarnings();
          refreshImageGrid();
        }
      } catch (err) {
        jsonValid = false;
        setJsonBadge(false, 'Parse Error');
        if (jsonMode === 'paste') {
          jsonTextarea.classList.remove('valid');
          jsonTextarea.classList.add('invalid');
        }
      }
      updatePreflightChecks();
    }

    function setJsonBadge(valid, text) {
      const badge = document.getElementById('jsonBadge');
      const badgeText = document.getElementById('jsonBadgeText');
      badgeText.textContent = text;
      if (valid) {
        badge.className = 'badge badge-present';
        badge.querySelector('svg').innerHTML = '<path stroke-linecap="round" stroke-linejoin="round" d="m4.5 12.75 6 6 9-13.5"/>';
      } else {
        badge.className = 'badge badge-missing';
        badge.querySelector('svg').innerHTML = '<circle cx="12" cy="12" r="10"/>';
      }
    }

    function showSkuStats(data) {
      const container = document.getElementById('skuStats');
      container.classList.remove('hidden');
      container.innerHTML = '';

      const stats = [
        { label: 'Total SKUs', value: data.sku_count, icon: '📦' },
        { label: 'Brands', value: [...new Set(data.skus.map(s => s.brand))].length, icon: '🏷️' },
        { label: 'Avg Bullets', value: data.skus.length ? (data.skus.reduce((a, s) => a + s.bullet_count, 0) / data.skus.length).toFixed(1) : 0, icon: '📝' },
        { label: 'Warnings', value: validationWarnings.length, icon: validationWarnings.length > 0 ? '⚠️' : '✅' },
      ];

      stats.forEach(s => {
        const div = document.createElement('div');
        div.className = 'glass-panel-lighter p-3 text-center animate-fade-in';
        div.innerHTML = `
          <div class="text-lg mb-0.5">${s.icon}</div>
          <div class="stat-value text-xl">${s.value}</div>
          <div class="text-xs text-slate-400 mt-0.5">${s.label}</div>
        `;
        container.appendChild(div);
      });
    }

    function hideSkuStats() {
      document.getElementById('skuStats').classList.add('hidden');
    }

    function showWarnings(warnings) {
      const panel = document.getElementById('warningsPanel');
      const list = document.getElementById('warningsList');
      if (warnings.length === 0) {
        panel.classList.add('hidden');
        return;
      }
      panel.classList.remove('hidden');
      list.innerHTML = '';
      warnings.forEach(w => {
        const div = document.createElement('div');
        div.className = 'warning-item animate-fade-in';
        div.textContent = w;
        list.appendChild(div);
      });
    }

    function hideWarnings() {
      document.getElementById('warningsPanel').classList.add('hidden');
    }

    // ── Image Handling ──
    const dropzone = document.getElementById('imageDropzone');

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('drag-over');
    });
    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('drag-over');
    });
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('drag-over');
      const files = Array.from(e.dataTransfer.files).filter(f =>
        /\.(jpg|jpeg|png|webp)$/i.test(f.name)
      );
      addImages(files);
    });

    function handleImages(e) {
      addImages(Array.from(e.target.files));
    }

    function addImages(files) {
      imageFiles = [...imageFiles, ...files];
      // Remove duplicates by name
      const seen = new Set();
      imageFiles = imageFiles.filter(f => {
        if (seen.has(f.name)) return false;
        seen.add(f.name);
        return true;
      });

      document.getElementById('imageCounter').innerHTML = `<span>${imageFiles.length} image${imageFiles.length !== 1 ? 's' : ''}</span>`;
      if (imageFiles.length > 0) {
        document.getElementById('imageCounter').className = 'badge badge-present';
      }
      routeImages();
      refreshImageGrid();
      updatePreflightChecks();
    }

    function routeImages() {
      imageMap = {};
      unassignedImages = [];
      const skuIds = parsedSkus.map(s => s.sku_id);

      // Sort longest first for matching
      const sortedIds = [...skuIds].sort((a, b) => b.length - a.length);

      imageFiles.forEach(file => {
        const stem = file.name.replace(/\.[^.]+$/, '').toUpperCase();
        let matched = false;
        for (const sid of sortedIds) {
          if (stem.startsWith(sid.toUpperCase())) {
            if (!imageMap[sid]) imageMap[sid] = [];
            imageMap[sid].push(file.name);
            matched = true;
            break;
          }
        }
        if (!matched) {
          unassignedImages.push(file.name);
        }
      });
    }

    function refreshImageGrid() {
      const grid = document.getElementById('skuImageGrid');
      const unPanel = document.getElementById('unassignedPanel');

      if (parsedSkus.length === 0 && imageFiles.length === 0) {
        grid.classList.add('hidden');
        unPanel.classList.add('hidden');
        return;
      }

      if (parsedSkus.length > 0 && imageFiles.length > 0) {
        grid.classList.remove('hidden');
        grid.innerHTML = '';

        const roles = ['_MAIN', '_PT01', '_PT02', '_PT03', '_PT04', '_PT05', '_PT06', '_PT07', '_PT08'];
        const roleLabels = {
          '_MAIN': 'Hero', '_PT01': 'Size', '_PT02': 'Fabric', '_PT03': 'Care',
          '_PT04': 'Back', '_PT05': 'Life1', '_PT06': 'Life2', '_PT07': 'Detail', '_PT08': 'Pack'
        };

        parsedSkus.forEach(sku => {
          const files = imageMap[sku.sku_id] || [];
          const corePresent = ['_MAIN', '_PT01', '_PT02', '_PT03'].every(suffix =>
            files.some(f => f.toUpperCase().includes(suffix))
          );

          const card = document.createElement('div');
          card.className = `sku-card ${corePresent ? 'complete' : (files.length > 0 ? 'incomplete' : '')} animate-fade-in`;

          let badgesHtml = '';
          roles.forEach(role => {
            const present = files.some(f => f.toUpperCase().includes(role));
            if (present || roles.indexOf(role) < 4) {
              badgesHtml += `<span class="badge ${present ? 'badge-present' : 'badge-missing'}">${roleLabels[role]}</span>`;
            }
          });

          card.innerHTML = `
            <div class="flex items-center justify-between mb-2.5">
              <span class="font-mono font-bold text-sm text-white">${sku.sku_id}</span>
              <span class="text-xs ${corePresent ? 'text-emerald-400' : (files.length > 0 ? 'text-amber-400' : 'text-slate-500')}">
                ${corePresent ? '✅ Complete' : (files.length > 0 ? '⚠️ Partial' : '—')}
              </span>
            </div>
            <div class="flex flex-wrap gap-1.5">${badgesHtml}</div>
            <div class="mt-2 text-xs text-slate-500">${files.length} file${files.length !== 1 ? 's' : ''} matched</div>
          `;
          grid.appendChild(card);
        });
      } else {
        grid.classList.add('hidden');
      }

      // Unassigned
      if (unassignedImages.length > 0) {
        unPanel.classList.remove('hidden');
        document.getElementById('unassignedList').innerHTML = unassignedImages.map(f =>
          `<span class="px-2 py-1 bg-slate-800/60 rounded text-slate-400 font-mono">${f}</span>`
        ).join('');
      } else {
        unPanel.classList.add('hidden');
      }
    }

    // ── Pre-Flight Checks ──
    function updatePreflightChecks() {
      const client = document.getElementById('clientName').value.trim();
      const batch = document.getElementById('batchId').value.trim();
      const cat = document.getElementById('category').value;
      const configOk = client && batch && cat;

      setCheck('checkConfig', 'checkConfigText', configOk, configOk ? 'All set' : 'Not set');
      setCheck('checkJson', 'checkJsonText', jsonValid, jsonValid ? `${parsedSkus.length} SKUs` : 'No data');
      setCheck('checkImages', 'checkImagesText', imageFiles.length > 0, imageFiles.length > 0 ? `${imageFiles.length} files` : 'None uploaded');

      const warnCount = validationWarnings.length;
      setCheck('checkWarnings', 'checkWarningsText',
        warnCount === 0 && jsonValid,
        warnCount > 0 ? `${warnCount} warning${warnCount > 1 ? 's' : ''}` : (jsonValid ? 'None' : '—'),
        warnCount > 0 ? 'warn' : null
      );

      // Enable generate button if config + json are valid (images optional)
      const canGenerate = configOk && jsonValid;
      document.getElementById('generateBtn').disabled = !canGenerate;
    }

    function setCheck(iconId, textId, ok, text, type) {
      const icon = document.getElementById(iconId);
      const textEl = document.getElementById(textId);
      textEl.textContent = text;

      if (type === 'warn') {
        icon.className = 'w-8 h-8 rounded-lg bg-amber-500/15 flex items-center justify-center text-amber-400 transition-all';
        icon.innerHTML = '<svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126ZM12 15.75h.007v.008H12v-.008Z"/></svg>';
      } else if (ok) {
        icon.className = 'w-8 h-8 rounded-lg bg-emerald-500/15 flex items-center justify-center text-emerald-400 transition-all';
        icon.innerHTML = '<svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="m4.5 12.75 6 6 9-13.5"/></svg>';
      } else {
        icon.className = 'w-8 h-8 rounded-lg bg-slate-800 flex items-center justify-center text-slate-500 transition-all';
        icon.innerHTML = '<svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/></svg>';
      }
    }

    // Listen for config changes
    ['clientName', 'batchId', 'category'].forEach(id => {
      document.getElementById(id).addEventListener('input', updatePreflightChecks);
      document.getElementById(id).addEventListener('change', updatePreflightChecks);
    });

    // ── Package Generation ──
    async function generatePackage() {
      const btn = document.getElementById('generateBtn');
      const btnContent = document.getElementById('btnContent');
      const btnLoading = document.getElementById('btnLoading');

      btn.disabled = true;
      btnContent.classList.add('hidden');
      btnLoading.classList.remove('hidden');

      try {
        const formData = new FormData();
        formData.append('client_name', document.getElementById('clientName').value.trim());
        formData.append('batch_id', document.getElementById('batchId').value.trim());
        formData.append('category', document.getElementById('category').value);
        formData.append('json_data', currentJsonText);

        // Append images
        imageFiles.forEach(file => {
          formData.append('images', file);
        });

        const resp = await fetch('/api/generate', {
          method: 'POST',
          body: formData,
        });

        if (!resp.ok) {
          const err = await resp.json();
          throw new Error(err.error || 'Generation failed');
        }

        // Download the ZIP
        const blob = await resp.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const disposition = resp.headers.get('Content-Disposition');
        let filename = 'package.zip';
        if (disposition) {
          const match = disposition.match(/filename="?(.+?)"?$/);
          if (match) filename = match[1];
        }
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        // Success flash
        btn.style.background = 'linear-gradient(135deg, #059669, #10b981)';
        btnLoading.classList.add('hidden');
        btnContent.classList.remove('hidden');
        btnContent.innerHTML = `
          <svg class="w-5 h-5 inline-block mr-2 -mt-0.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="m4.5 12.75 6 6 9-13.5"/></svg>
          Package Downloaded!
        `;
        setTimeout(() => {
          btnContent.innerHTML = `
            <svg class="w-5 h-5 inline-block mr-2 -mt-0.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="m20.25 7.5-.625 10.632a2.25 2.25 0 0 1-2.247 2.118H6.622a2.25 2.25 0 0 1-2.247-2.118L3.75 7.5m8.25 3v6.75m0 0-3-3m3 3 3-3M3.375 7.5h17.25c.621 0 1.125-.504 1.125-1.125v-1.5c0-.621-.504-1.125-1.125-1.125H3.375c-.621 0-1.125.504-1.125 1.125v1.5c0 .621.504 1.125 1.125 1.125Z"/></svg>
            Generate &amp; Download Delivery Package (ZIP)
          `;
          btn.disabled = false;
        }, 2500);

      } catch (err) {
        alert('Error: ' + err.message);
        btnLoading.classList.add('hidden');
        btnContent.classList.remove('hidden');
        btn.disabled = false;
      }
    }

    // ── Init ──
    updatePreflightChecks();
  </script>
</body>
</html>"""


# ──────────────────────────────────────────────
# Entry Point
# ──────────────────────────────────────────────

if __name__ == "__main__":
    print("\n  [*] Listing Factory - Multi-Marketplace Packaging Studio")
    print("  " + "=" * 57)
    print("  > Server running at: http://127.0.0.1:8000")
    print("  > Press Ctrl+C to stop\n")
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
