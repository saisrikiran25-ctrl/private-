"""
Listing Factory: Multi-Marketplace Packaging Studio (v2.0)
===========================================================
A FastAPI-powered local web application for e-commerce cataloging agencies.
Takes AI-generated JSON copy + loose image files, validates against
Amazon.in / Flipkart / Meesho constraints, organizes images into SKU
subfolders, populates marketplace mapping Excel workbooks, and outputs
a ready-to-deliver client ZIP archive with full package audit metadata.

Version: Listing Factory v2.0
Run:     python app.py
Open:    http://127.0.0.1:8000
"""

from __future__ import annotations

import glob
import hashlib
import io
import json
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, List, Literal, Optional, Union

import pandas as pd
from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field, field_validator
import uvicorn


# ──────────────────────────────────────────────
# Global Constants & Versioning
# ──────────────────────────────────────────────

TOOL_VERSION = "Listing Factory v2.0"
JSON_PROMPT_VERSION = "JSON Prompt v1.0 – 2026-08-22"
EXPECTED_SCHEMA_VERSION = "v2.0"
TO_BE_CONFIRMED = "To be confirmed"

MAX_SKUS_SOFT_LIMIT = 50
MAX_TOTAL_IMAGE_SIZE_MB = 200

OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


# ──────────────────────────────────────────────
# Flipkart Controlled Attributes Literal Types
# ──────────────────────────────────────────────

FabricLiteral = Literal[
    "Pure Cotton", "Rayon", "Georgette", "Silk Blend", "Crepe", "Chanderi Cotton", "Poly Cotton",
    "To be confirmed"
]
KurtaTypeLiteral = Literal[
    "Anarkali", "Straight", "A-line", "Flared", "Kaftan", "Frontslit", "Pathani",
    "To be confirmed"
]
NeckLiteral = Literal[
    "Mandarin Neck", "Round Neck", "V-Neck", "Boat Neck", "Sweetheart Neck", "Collar Neck",
    "To be confirmed"
]
SleeveLiteral = Literal[
    "3/4 Sleeve", "Full Sleeve", "Half Sleeve", "Sleeveless", "Short Sleeve",
    "To be confirmed"
]
LengthTypeLiteral = Literal[
    "Calf Length", "Ankle Length", "Knee Length", "Above Knee",
    "To be confirmed"
]
PatternLiteral = Literal[
    "Floral Print", "Solid", "Printed", "Embroidered", "Geometric Print", "Self Design", "Bandhani",
    "To be confirmed"
]
OccasionLiteral = Literal[
    "Casual", "Festive", "Casual & Festive", "Party", "Formal",
    "To be confirmed"
]


# ──────────────────────────────────────────────
# Pydantic v2 Models (Strict Validation Schema)
# ──────────────────────────────────────────────

class SellerConfig(BaseModel):
    amazon_quantity: int = 50
    gst_percent: int = 5
    hsn_code: str = "62114200"


class BatchConfig(BaseModel):
    brand: Optional[str] = None
    category: Optional[str] = None
    seller_config: Optional[SellerConfig] = None


class AmazonData(BaseModel):
    title: str
    bullet_points: list[str] = Field(..., min_length=5, max_length=5)
    backend_search_terms: str
    description: str

    @field_validator("title")
    @classmethod
    def validate_title(cls, v: str) -> str:
        if v != TO_BE_CONFIRMED and len(v) > 180:
            raise ValueError(f"Amazon title must be ≤ 180 characters (got {len(v)} chars)")
        if not v.strip():
            raise ValueError("Amazon title cannot be empty")
        return v

    @field_validator("backend_search_terms")
    @classmethod
    def validate_bst(cls, v: str) -> str:
        if v != TO_BE_CONFIRMED:
            b = len(v.encode("utf-8"))
            if b > 240:
                raise ValueError(f"Amazon backend search terms must be ≤ 240 bytes (got {b} bytes)")
        if not v.strip():
            raise ValueError("Amazon backend search terms cannot be empty")
        return v


class FlipkartData(BaseModel):
    title: str
    fabric: FabricLiteral
    kurta_type: KurtaTypeLiteral
    neck: NeckLiteral
    sleeve: SleeveLiteral
    length_type: LengthTypeLiteral
    pattern: PatternLiteral
    occasion: OccasionLiteral
    search_keywords: str
    description: str


class MeeshoData(BaseModel):
    title: str
    hinglish_hook_description: str
    english_hook_description: str
    highlights: list[str] = Field(..., min_length=4, max_length=4)

    @field_validator("title")
    @classmethod
    def validate_title(cls, v: str) -> str:
        if v != TO_BE_CONFIRMED and len(v) > 60:
            raise ValueError(f"Meesho title must be ≤ 60 characters (got {len(v)} chars)")
        if not v.strip():
            raise ValueError("Meesho title cannot be empty")
        return v


class AlternateAmazonData(BaseModel):
    title: str
    bullet_points: list[str] = Field(..., min_length=5, max_length=5)
    backend_search_terms: str
    description: str

    @field_validator("title")
    @classmethod
    def validate_title(cls, v: str) -> str:
        if v != TO_BE_CONFIRMED and len(v) > 180:
            raise ValueError(f"Alternate Amazon title must be ≤ 180 characters (got {len(v)} chars)")
        return v

    @field_validator("backend_search_terms")
    @classmethod
    def validate_bst(cls, v: str) -> str:
        if v != TO_BE_CONFIRMED:
            b = len(v.encode("utf-8"))
            if b > 240:
                raise ValueError(f"Alternate Amazon backend search terms must be ≤ 240 bytes (got {b} bytes)")
        return v


class AlternateFlipkartData(BaseModel):
    title: str
    search_keywords: str
    description: str


class AlternateMeeshoData(BaseModel):
    title: str
    hinglish_hook_description: str
    english_hook_description: str
    highlights: list[str] = Field(..., min_length=4, max_length=4)

    @field_validator("title")
    @classmethod
    def validate_title(cls, v: str) -> str:
        if v != TO_BE_CONFIRMED and len(v) > 60:
            raise ValueError(f"Alternate Meesho title must be ≤ 60 characters (got {len(v)} chars)")
        return v


class AlternateVariant(BaseModel):
    variant_id: str
    angle_theme: str
    amazon: AlternateAmazonData
    flipkart: AlternateFlipkartData
    meesho: AlternateMeeshoData


class SKUItem(BaseModel):
    sku_id: str
    brand: str
    product_type: str
    color: str
    category: str = "Women Ethnic Wear"
    sizes: str
    mrp: float
    meesho_price: float
    seller_config: SellerConfig
    amazon: AmazonData
    flipkart: FlipkartData
    meesho: MeeshoData
    alternates: list[AlternateVariant] = Field(default_factory=list, max_length=5)

    @field_validator("alternates")
    @classmethod
    def validate_alternates(cls, v: list[AlternateVariant]) -> list[AlternateVariant]:
        if v and len(v) != 5:
            raise ValueError(f"When alternates are provided, exactly 5 variants (V1-V5) are required (got {len(v)})")
        return v


# ──────────────────────────────────────────────
# Category mapping & Image Roles
# ──────────────────────────────────────────────

CATEGORY_MAP = {
    "Women Ethnic Wear": "kurtas-and-ethnic-tops",
    "Men Western Wear": "mens-casual-shirts",
    "Sarees": "sarees",
    "Footwear": "casual-shoes",
    "Home & Kitchen": "home-furnishing",
}

# Declared Image Roles (Assigned by filename suffix; does not verify visual content)
IMAGE_ROLES = {
    "_MAIN": "Primary Image (Hero)",
    "_PT01": "Other Image 1 (Size Chart)",
    "_PT02": "Other Image 2 (Fabric Spec)",
    "_PT03": "Other Image 3 (Care Guide)",
    "_PT04": "Other Image 4 (Back View)",
    "_PT05": "Other Image 5",
}

CORE_SUFFIXES = ["_MAIN", "_PT01", "_PT02", "_PT03"]


# ──────────────────────────────────────────────
# Search Term & Bullet 3 Safety Helpers
# ──────────────────────────────────────────────

UNVERIFIED_DURABILITY_TERMS = [
    "colorfast", "colourfast", "fade resistant", "zero fade", "no fading",
    "zero bleeding", "no bleeding", "shrink resistant", "anti-shrink", "pre-shrunk",
    "pilling", "anti-pilling", "durable", "reinforced", "tested", "technology",
    "vat dye", "reactive dye", "color lock", "colour lock"
]


def validate_amazon_backend_search_terms(title: str, bst: str, brand: str) -> tuple[list[str], list[str]]:
    """
    Validates Amazon Backend Search Terms hygiene:
    - Must be lowercase ASCII only (no uppercase, no non-ASCII characters).
    - Must not contain punctuation, commas, or repeated whitespace.
    - Must not contain the brand name.
    - Warns on exact title duplicates or close singular/plural title word matches.
    """
    errors: list[str] = []
    warnings: list[str] = []

    if not bst.strip() or bst == TO_BE_CONFIRMED:
        return errors, warnings

    # 1. Non-ASCII check
    try:
        bst.encode("ascii")
    except UnicodeEncodeError:
        errors.append("Amazon backend search terms must contain ASCII characters only.")

    # 2. Uppercase check
    if re.search(r"[A-Z]", bst):
        errors.append("Amazon backend search terms must be lowercase only (uppercase characters found).")

    # 3. Punctuation / commas check
    invalid_chars = set(re.findall(r"[^a-z0-9\s]", bst))
    if invalid_chars:
        errors.append(f"Amazon backend search terms must not contain punctuation or commas (found: {', '.join(sorted(invalid_chars))}). Use single spaces only.")

    # 4. Repeated whitespace
    if re.search(r"\s{2,}", bst):
        errors.append("Amazon backend search terms must not contain repeated whitespace. Use single spaces between terms.")

    # 5. Brand occurrence check
    bst_lower = bst.lower()
    bst_tokens = bst_lower.split()
    brand_tokens = [t.lower() for t in re.findall(r"[a-z0-9]+", brand) if len(t) >= 2]
    for bt in brand_tokens:
        if bt in bst_tokens or bt in bst_lower:
            errors.append(f"Amazon backend search terms must not contain the brand name ('{brand}' / '{bt}').")
            break

    # 6. Title duplicate checks (Warnings)
    stopwords = {"and", "or", "with", "for", "in", "of", "a", "an", "the", "to", "by", "on", "at", "from", "is", "it", "as"}
    title_tokens = set(re.findall(r"[a-z0-9]+", title.lower())) - stopwords

    exact_dupes = []
    close_dupes = []

    for word in bst_tokens:
        clean_w = re.sub(r"[^a-z0-9]", "", word)
        if not clean_w or clean_w in stopwords:
            continue
        if clean_w in title_tokens:
            exact_dupes.append(clean_w)
        else:
            if clean_w.endswith("s") and clean_w[:-1] in title_tokens and len(clean_w) > 3:
                close_dupes.append(f"'{clean_w}' (plural of '{clean_w[:-1]}' in title)")
            elif (clean_w + "s") in title_tokens:
                close_dupes.append(f"'{clean_w}' (singular of '{clean_w}s' in title)")

    if exact_dupes:
        warnings.append(f"Amazon backend search terms contain word(s) already in the title: {', '.join(sorted(set(exact_dupes)))}. Consider replacing with non-title search keywords.")

    if close_dupes:
        warnings.append(f"Amazon backend search terms contain close duplicate(s) of title words: {', '.join(sorted(set(close_dupes)))}.")

    return errors, warnings


def validate_bullet_3_safety(bullet_3: str, is_verified_durability: bool = False) -> tuple[list[str], list[str]]:
    """
    Validates that Bullet 3 begins with 'COLORFAST & DURABLE:' and does not contain
    unsupported durability/guarantee/shrinkage/colorfastness claims without explicit verified support.
    """
    errors: list[str] = []
    warnings: list[str] = []

    if not bullet_3.strip() or bullet_3 == TO_BE_CONFIRMED:
        return errors, warnings

    b3_upper = bullet_3.upper().strip()
    if not b3_upper.startswith("COLORFAST & DURABLE:"):
        errors.append("Bullet 3 must start with the mandated heading 'COLORFAST & DURABLE:'.")

    body = bullet_3.split(":", 1)[1] if ":" in bullet_3 else bullet_3
    body_lower = body.lower()

    if not is_verified_durability:
        offending = []
        for term in UNVERIFIED_DURABILITY_TERMS:
            pattern = r"\b" + re.escape(term) + r"\b"
            if re.search(pattern, body_lower):
                offending.append(term)

        if offending:
            errors.append(
                f"Bullet 3 ('COLORFAST & DURABLE:') contains unverified technical/guarantee claim(s): {', '.join(sorted(set(offending)))}. "
                "Use neutral care-led language (e.g. 'COLORFAST & DURABLE: Follow the provided care label to help maintain the fabric\\'s appearance and color.')"
            )

    return errors, warnings


# ──────────────────────────────────────────────
# Audit Trail & Package Metadata Builder
# ──────────────────────────────────────────────

def build_package_metadata(
    client: str,
    batch: str,
    category: str,
    skus: list[SKUItem],
    tool_version: str = TOOL_VERSION,
    json_prompt_version: str = JSON_PROMPT_VERSION,
    schema_version: str = EXPECTED_SCHEMA_VERSION,
) -> dict:
    """
    Build package audit metadata including SHA-256 input and output hashes per SKU.
    """
    sku_entries = []
    for sku in skus:
        # Input hash over the verified product record subset
        input_record = {
            "sku_id": sku.sku_id,
            "brand": sku.brand,
            "product_type": sku.product_type,
            "color": sku.color,
            "category": sku.category or category,
            "sizes": sku.sizes,
            "mrp": sku.mrp,
            "meesho_price": sku.meesho_price,
            "seller_config": sku.seller_config.model_dump(),
        }
        input_json = json.dumps(input_record, sort_keys=True)
        input_hash = hashlib.sha256(input_json.encode("utf-8")).hexdigest()

        # Output hash over the complete internal SKU object
        output_json = json.dumps(sku.model_dump(), sort_keys=True)
        output_hash = hashlib.sha256(output_json.encode("utf-8")).hexdigest()

        sku_entries.append({
            "sku_id": sku.sku_id,
            "input_hash": input_hash,
            "output_hash": output_hash,
        })

    now_iso = datetime.now().astimezone().isoformat()

    return {
        "tool_version": tool_version,
        "json_prompt_version": json_prompt_version,
        "schema_version": schema_version,
        "generated_at": now_iso,
        "client": client,
        "batch": batch,
        "category": category,
        "skus": sku_entries,
    }


# ──────────────────────────────────────────────
# Helper to Detect "To be confirmed" Fields
# ──────────────────────────────────────────────

def _sku_has_unconfirmed_fields(sku: SKUItem) -> bool:
    """Check if any attribute in the SKU is set to TO_BE_CONFIRMED."""
    def check_val(v: Any) -> bool:
        if isinstance(v, str) and (v.strip().lower() == TO_BE_CONFIRMED.lower() or v.strip().upper() == "TBC"):
            return True
        if isinstance(v, list):
            return any(check_val(item) for item in v)
        if isinstance(v, dict):
            return any(check_val(val) for val in v.values())
        return False

    return check_val(sku.model_dump())


# ──────────────────────────────────────────────
# Excel Workbook Builders (Complete PT05 Support)
# ──────────────────────────────────────────────

def _style_header_row(ws, num_cols: int, header_color: str = "10B981"):
    """Apply styling to the header row of a worksheet."""
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", bold=True, size=11, color=header_color)
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
    ws.row_dimensions[1].height = 30


def _auto_width(ws, min_width=12, max_width=50):
    """Auto-fit column widths based on content."""
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            try:
                cell_len = len(str(cell.value)) if cell.value is not None else 0
                max_len = max(max_len, cell_len)
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 4, max_width))


def _build_master_summary(
    wb: Workbook,
    skus: list[SKUItem],
    image_map: dict[str, list[str]],
    validation_status: str = "✅ Pass"
):
    ws = wb.active
    ws.title = "Master_Summary"
    headers = [
        "SKU ID", "Brand", "Product Type", "Color", "Fabric", "Sizes Available",
        "Amazon Title Preview", "Flipkart Title Preview", "Meesho Hinglish Hook Preview",
        "Core Images Found", "Core Coverage", "Validation Status", "Review Flags", "Package Readiness"
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        files = image_map.get(sku.sku_id, [])
        core_found = sum(
            1 for suffix in CORE_SUFFIXES
            if any(suffix.lower() in fn.lower() for fn in files)
        )
        coverage_pct = f"{int((core_found / 4) * 100)}% ({core_found}/4 Core)"

        has_tbc = _sku_has_unconfirmed_fields(sku)
        review_flags = "⚠️ Has unconfirmed fields" if has_tbc else "—"
        
        if validation_status == "✅ Pass" and core_found == 4 and not has_tbc:
            readiness = "✅ Structurally Complete – Seller Review Required"
        elif validation_status == "✅ Pass":
            readiness = "⚠️ Warnings – Seller Review Required"
        elif "Warning" in validation_status:
            readiness = "⚠️ Warnings – Seller Review Required"
        else:
            readiness = "❌ Not Ready – Fix Errors First"

        ws.append([
            sku.sku_id,
            sku.brand,
            sku.product_type,
            sku.color,
            sku.flipkart.fabric,
            sku.sizes,
            (sku.amazon.title[:75] + "…") if len(sku.amazon.title) > 75 else sku.amazon.title,
            (sku.flipkart.title[:75] + "…") if len(sku.flipkart.title) > 75 else sku.flipkart.title,
            (sku.meesho.hinglish_hook_description[:75] + "…") if len(sku.meesho.hinglish_hook_description) > 75 else sku.meesho.hinglish_hook_description,
            f"{core_found}/4 Slots",
            coverage_pct,
            validation_status,
            review_flags,
            readiness,
        ])
    _auto_width(ws)


def _build_amazon_tab(wb: Workbook, skus: list[SKUItem], category: str):
    ws = wb.create_sheet("01_Amazon_Bulk_Import")
    item_type = CATEGORY_MAP.get(category, "kurtas-and-ethnic-tops")
    headers = [
        "item_sku", "item_name", "brand_name", "feed_product_type", "item_type_keyword",
        "standard_price", "currency", "quantity", "condition_type",
        "main_image_url", "other_image_url1", "other_image_url2", "other_image_url3", "other_image_url4", "other_image_url5",
        "bullet_point1", "bullet_point2", "bullet_point3", "bullet_point4", "bullet_point5",
        "generic_keyword", "item_description", "size", "color"
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        bp = sku.amazon.bullet_points
        ws.append([
            sku.sku_id,
            sku.amazon.title,
            sku.brand,
            "Kurta",
            item_type,
            sku.mrp,
            "INR",
            sku.seller_config.amazon_quantity,
            "New",
            f"{sku.sku_id}_MAIN.jpg",
            f"{sku.sku_id}_PT01.jpg",
            f"{sku.sku_id}_PT02.jpg",
            f"{sku.sku_id}_PT03.jpg",
            f"{sku.sku_id}_PT04.jpg",
            f"{sku.sku_id}_PT05.jpg",
            bp[0],
            bp[1],
            bp[2],
            bp[3],
            bp[4],
            sku.amazon.backend_search_terms,
            sku.amazon.description,
            sku.sizes,
            sku.color,
        ])
    _auto_width(ws)


def _build_flipkart_tab(wb: Workbook, skus: list[SKUItem]):
    ws = wb.create_sheet("02_Flipkart_Bulk_Import")
    headers = [
        "Seller SKU ID", "Product Title", "Brand", "Style Code", "Ideal For", "Size", "Color",
        "Pattern", "Type / Kurta Type", "Fabric", "Neck", "Sleeve", "Length Type", "Occasion",
        "Net Quantity", "GST (%)", "HSN Code", "Search Keywords",
        "Main Image Name", "Angle 1 Image", "Angle 2 Image", "Angle 3 Image", "Angle 4 Image", "Angle 5 Image",
        "Description"
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        ws.append([
            sku.sku_id,
            sku.flipkart.title,
            sku.brand,
            sku.sku_id,
            "Women",
            sku.sizes,
            sku.color,
            sku.flipkart.pattern,
            sku.flipkart.kurta_type,
            sku.flipkart.fabric,
            sku.flipkart.neck,
            sku.flipkart.sleeve,
            sku.flipkart.length_type,
            sku.flipkart.occasion,
            1,
            sku.seller_config.gst_percent,
            sku.seller_config.hsn_code,
            sku.flipkart.search_keywords,
            f"{sku.sku_id}_MAIN.jpg",
            f"{sku.sku_id}_PT01.jpg",
            f"{sku.sku_id}_PT02.jpg",
            f"{sku.sku_id}_PT03.jpg",
            f"{sku.sku_id}_PT04.jpg",
            f"{sku.sku_id}_PT05.jpg",
            sku.flipkart.description,
        ])
    _auto_width(ws)


def _build_meesho_tab(wb: Workbook, skus: list[SKUItem]):
    ws = wb.create_sheet("03_Meesho_Bulk_Import")
    headers = [
        "Product ID / SKU", "Product Name",
        "Product Description (Hinglish Hook)",
        "English Hook Description",
        "Fabric", "Available Sizes", "Color", "Occasion",
        "Key Highlights", "GST (%)", "HSN Code",
        "Recommended Price (INR)",
        "Primary Image (Hero)",
        "Other Image 1 (Size Chart)",
        "Other Image 2 (Fabric Spec)",
        "Other Image 3 (Care Guide)",
        "Other Image 4 (Back View)",
        "Other Image 5"
    ]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    for sku in skus:
        highlights_str = " • ".join(sku.meesho.highlights)
        ws.append([
            sku.sku_id,
            sku.meesho.title,
            sku.meesho.hinglish_hook_description,
            sku.meesho.english_hook_description,
            sku.flipkart.fabric,
            sku.sizes,
            sku.color,
            sku.flipkart.occasion,
            highlights_str,
            sku.seller_config.gst_percent,
            sku.seller_config.hsn_code,
            sku.meesho_price,
            f"{sku.sku_id}_MAIN.jpg",
            f"{sku.sku_id}_PT01.jpg",
            f"{sku.sku_id}_PT02.jpg",
            f"{sku.sku_id}_PT03.jpg",
            f"{sku.sku_id}_PT04.jpg",
            f"{sku.sku_id}_PT05.jpg",
        ])
    _auto_width(ws)


def build_workbook(
    skus: list[SKUItem],
    category: str,
    image_map: dict[str, list[str]],
    validation_status: str = "✅ Pass"
) -> bytes:
    """Build the full multi-tab Master Excel mapping workbook and return raw bytes."""
    wb = Workbook()
    _build_master_summary(wb, skus, image_map, validation_status)
    _build_amazon_tab(wb, skus, category)
    _build_flipkart_tab(wb, skus)
    _build_meesho_tab(wb, skus)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_alternates_workbook(skus: list[SKUItem]) -> bytes:
    """Build the 5x Alternate Listing Copies mapping workbook with 3 marketplace tabs."""
    wb = Workbook()

    # -- Tab 1: Amazon Alternates --
    ws_amz = wb.active
    ws_amz.title = "Amazon_Alternate_Copies"
    headers_amz = [
        "SKU ID", "Brand", "Variant", "Marketing Angle / Theme",
        "Alternative Title (Amazon)",
        "Bullet 1 (Fabric & Comfort)", "Bullet 2 (Design & Utility)",
        "Bullet 3 (Colorfast & Durable)", "Bullet 4 (Versatile Styling)",
        "Bullet 5 (Sizing & Care)",
        "Alternative Backend Search Terms (<240 bytes)",
        "Alternative Description"
    ]
    ws_amz.append(headers_amz)
    _style_header_row(ws_amz, len(headers_amz), header_color="38BDF8")

    # -- Tab 2: Flipkart Alternates --
    ws_fk = wb.create_sheet("Flipkart_Alternate_Copies")
    headers_fk = [
        "SKU ID", "Brand", "Variant", "Marketing Angle / Theme",
        "Alternative Flipkart Title",
        "Alternative Search Keywords",
        "Alternative Product Description"
    ]
    ws_fk.append(headers_fk)
    _style_header_row(ws_fk, len(headers_fk), header_color="38BDF8")

    # -- Tab 3: Meesho Alternates (Hinglish + English) --
    ws_me = wb.create_sheet("Meesho_Alternate_Copies")
    headers_me = [
        "SKU ID", "Brand", "Variant", "Marketing Angle / Theme",
        "Alternative Product Title",
        "Hinglish Hook Description (Conversational)",
        "English Hook Description (Formal/Metro)",
        "Key Highlights (Badges)"
    ]
    ws_me.append(headers_me)
    _style_header_row(ws_me, len(headers_me), header_color="38BDF8")

    for sku in skus:
        for idx, alt in enumerate(sku.alternates):
            v_tag = alt.variant_id or f"V{idx+1}"
            theme = alt.angle_theme or f"Angle {idx+1}"

            # Amazon row
            bp = alt.amazon.bullet_points
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
                " • ".join(alt.meesho.highlights),
            ])

    _auto_width(ws_amz)
    _auto_width(ws_fk)
    _auto_width(ws_me)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────
# README generator (Handover Documentation)
# ──────────────────────────────────────────────

def generate_readme(client: str, batch: str, category: str = "Women Ethnic Wear") -> str:
    category_keyword = CATEGORY_MAP.get(category, "kurtas-and-ethnic-tops")
    return f"""================================================================================
  LISTING FACTORY HANDOVER PACKAGE: {client} -- {batch}
================================================================================

  Generated           : {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
  Tool Version        : {TOOL_VERSION}
  JSON Prompt Version : {JSON_PROMPT_VERSION}
  Schema Version      : {EXPECTED_SCHEMA_VERSION}
  Target Category     : {category}

================================================================================
  IMPORTANT NOTICE & DISCLAIMER
================================================================================

  This package is structurally complete but must be reviewed by the seller for
  marketplace policy compliance, category template alignment, and final attribute
  verification before upload.

  The Excel files in this package are MARKETPLACE MAPPING WORKBOOKS designed to
  organize, validate, and prepare listing data. They are not direct official portal
  flat files. Sellers and catalog teams must transfer this structured data into
  their respective Seller Central, Seller Hub, or Supplier Panel category upload
  templates.

  Processing and approval times are approximate examples only. They may vary by
  marketplace, category, seller account status, validation results, review queue,
  and current platform workload. Listing Factory does not guarantee approval or
  processing timelines.

================================================================================
  SELLER PRE-UPLOAD CHECKLIST
================================================================================

  Before uploading to marketplaces, please confirm:
    [ ] Fabric, care, size, and color details match your physical product.
    [ ] GST (%), HSN code, and pricing values are correct for your account.
    [ ] Data is mapped into the latest official marketplace templates for your category.
    [ ] Images meet each marketplace's current image policy (size, white background, count).
    [ ] All fields marked 'To be confirmed' have been finalized with confirmed values.
    [ ] All visual image content has been manually verified against declared filename roles.

================================================================================
  SCOPE & LIMITATIONS
================================================================================

  • This tool prepares listing copy and mapping workbooks based on provided data.
  • It does not guarantee marketplace acceptance or policy compliance.
  • It does not verify image visual content beyond filename-based slot assignment.
  • It does not provide legal, tax, or official classification advice (GST/HSN).
  • All data must be reviewed and confirmed by the seller before portal upload.

================================================================================
  TEMPLATE REFERENCE NOTES
================================================================================

  This package mapping is tailored for:
    • Amazon.in  : '{category}' (Keyword: {category_keyword})
    • Flipkart   : {category} Vertical
    • Meesho     : {category} / Kurtis Vertical

  Always download the latest official flat file or category template from each
  respective marketplace before uploading.

================================================================================
  FIELDS MARKED 'TO BE CONFIRMED'
================================================================================

  Some fields in this package may be marked as '{TO_BE_CONFIRMED}'.
  These indicate placeholder or unverified data points that must be finalized
  and confirmed by the brand/seller prior to live marketplace submission.

================================================================================
  HOW TO UPLOAD -- 3-STEP GUIDE
================================================================================

  STEP 1 ▸ AMAZON SELLER CENTRAL (sellercentral.amazon.in)
  ────────────────────────────────────────────────────────
  1. Navigate to: Catalog > Add Products via Upload > Upload your Inventory File.
  2. Download the category-specific inventory flat file template for your product.
  3. Open the "01_Amazon_Bulk_Import" tab in the Master Excel mapping file.
  4. Copy-paste rows into Amazon's template columns matching headers carefully.
  5. Link image assets from each SKU folder (MAIN, PT01, PT02, PT03, PT04, PT05).
  6. Submit file. (Processing and review timelines vary by platform queue).

  STEP 2 ▸ FLIPKART SELLER HUB (seller.flipkart.com)
  ──────────────────────────────────────────────────
  1. Navigate to: Listings > Add in Bulk.
  2. Download Flipkart's bulk listing template for the selected vertical.
  3. Open the "02_Flipkart_Bulk_Import" tab in the Master Excel mapping file.
  4. Map Flipkart-specific copy, style code, fabric, and controlled attributes.
  5. Upload the completed template and corresponding SKU image assets (MAIN, PT01-PT05).
  6. Submit for QC review. (QC review timelines vary by portal queue and category).

  STEP 3 ▸ MEESHO SUPPLIER PANEL (supplier.meesho.com)
  ────────────────────────────────────────────────────
  1. Navigate to: Catalog > Add Catalog (Single / Bulk).
  2. Select your exact product vertical.
  3. Open the "03_Meesho_Bulk_Import" tab in the Master Excel mapping file.
  4. Use the Hinglish hook copy, English hook copy, and bullet highlight badges.
  5. Upload primary hero cutouts and angle shots from SKU folders (MAIN, PT01-PT05).
  6. Submit for catalog approval. (Catalog approval timelines vary by vertical queue).

================================================================================
  IMAGE NAMING & ASSET STRUCTURE
================================================================================

  Canonical Image Naming Scheme (Declared Filename Slots):
    • Primary Image (Hero)        : SKU_XX_MAIN.jpg (Declared Hero cutout)
    • Other Image 1 (Size Chart)  : SKU_XX_PT01.jpg (Declared Size Guide)
    • Other Image 2 (Fabric Spec) : SKU_XX_PT02.jpg (Declared Fabric Highlight)
    • Other Image 3 (Care Guide)  : SKU_XX_PT03.jpg (Declared Care / Styling)
    • Other Image 4 (Back View)   : SKU_XX_PT04.jpg (Declared Angle / Back View)
    • Other Image 5               : SKU_XX_PT05.jpg (Declared Other Image 5)

  Note on Image Verification:
    All image roles are assigned based on filename suffixes. Filename assignment
    does NOT verify visual image content, white background quality, or resolution.
    Catalog managers must manually visually verify image assets before upload.

  Folder Hierarchy in this ZIP:
    Organized_SKU_Images/
      ├── [SKU_ID]/              ← One subfolder per SKU
      │   ├── [SKU]_MAIN.jpg     ← Declared Primary hero
      │   ├── [SKU]_PT01.jpg     ← Declared Size & measurement guide
      │   ├── [SKU]_PT02.jpg     ← Declared Fabric texture & spec highlight
      │   ├── [SKU]_PT03.jpg     ← Declared Care & styling recommendations
      │   ├── [SKU]_PT04.jpg     ← Declared Additional angle / back view
      │   └── [SKU]_PT05.jpg     ← Declared Other Image 5
      └── Unassigned_Assets/     ← Assets requiring manual prefix assignment

================================================================================
  SUPPORT & SLA EXPECTATION
================================================================================

  Support is provided on a best-effort basis (Monday–Friday, 10:00–18:00 IST).
  Critical packaging issues and broken ZIP archives are prioritized.
  Contact catalog support: support@listingfactory.internal

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
    Routes image files to their SKU folders based on prefix matching (e.g. SKU_01_MAIN.jpg).
    Returns (matched_map, unassigned_list).
    """
    matched: dict[str, list[tuple[str, bytes]]] = {sid: [] for sid in sku_ids}
    unassigned: list[tuple[str, bytes]] = []

    sorted_ids = sorted(sku_ids, key=len, reverse=True)

    for fname, data in image_files:
        stem = Path(fname).stem.upper()
        found = False
        for sid in sorted_ids:
            if stem.startswith(sid.upper() + "_") or stem == sid.upper():
                matched[sid].append((fname, data))
                found = True
                break
        if not found:
            unassigned.append((fname, data))

    return matched, unassigned


# ──────────────────────────────────────────────
# JSON Payload Normalization & Batch Config Merge
# ──────────────────────────────────────────────

def normalize_and_merge_json(raw_data: Any) -> tuple[Optional[str], Optional[BatchConfig], list[dict]]:
    """
    Normalizes JSON payload which can be:
    1. An object with schema_version, batch_config, and skus:
       { "schema_version": "v2.0", "batch_config": {...}, "skus": [...] }
    2. A raw list of SKU dicts: [ {...}, {...} ]
    3. A single SKU dict: { "sku_id": "...", ... }
    Returns (schema_version, batch_config, merged_skus).
    """
    schema_ver: Optional[str] = None
    batch_cfg: Optional[BatchConfig] = None
    sku_raw_list: list[dict] = []

    if isinstance(raw_data, dict):
        schema_ver = raw_data.get("schema_version")
        if "batch_config" in raw_data or "skus" in raw_data:
            if "batch_config" in raw_data and isinstance(raw_data["batch_config"], dict):
                batch_cfg = BatchConfig(**raw_data["batch_config"])
            skus_val = raw_data.get("skus", [])
            sku_raw_list = skus_val if isinstance(skus_val, list) else [skus_val]
        else:
            sku_raw_list = [raw_data]
    elif isinstance(raw_data, list):
        sku_raw_list = raw_data
    else:
        raise ValueError("Invalid JSON structure: Expected a JSON object or array of SKUs.")

    merged_skus = []
    for item in sku_raw_list:
        sku_dict = dict(item)
        if batch_cfg:
            if not sku_dict.get("brand") and batch_cfg.brand:
                sku_dict["brand"] = batch_cfg.brand
            if not sku_dict.get("category") and batch_cfg.category:
                sku_dict["category"] = batch_cfg.category
            if batch_cfg.seller_config:
                existing_sc = sku_dict.get("seller_config", {})
                if not isinstance(existing_sc, dict):
                    existing_sc = {}
                default_sc = batch_cfg.seller_config.model_dump()
                default_sc.update(existing_sc)
                sku_dict["seller_config"] = default_sc
        merged_skus.append(sku_dict)

    return schema_ver, batch_cfg, merged_skus


# ──────────────────────────────────────────────
# ZIP Builder
# ──────────────────────────────────────────────

def build_zip(
    client: str,
    batch: str,
    category: str,
    skus: list[SKUItem],
    image_files: list[tuple[str, bytes]],
    validation_status: str = "✅ Pass",
    schema_version: str = EXPECTED_SCHEMA_VERSION
) -> bytes:
    """Build the full delivery ZIP archive in-memory."""
    sku_ids = [s.sku_id for s in skus]
    matched_images, unassigned_images = route_images(sku_ids, image_files)

    image_map: dict[str, list[str]] = {}
    for sid, files in matched_images.items():
        image_map[sid] = [f[0] for f in files]

    xlsx_bytes = build_workbook(skus, category, image_map, validation_status)
    readme_text = generate_readme(client, batch, category)
    metadata_dict = build_package_metadata(
        client, batch, category, skus, TOOL_VERSION, JSON_PROMPT_VERSION, schema_version
    )
    metadata_bytes = json.dumps(metadata_dict, indent=2, ensure_ascii=False).encode("utf-8")

    has_alternates = any(len(s.alternates) > 0 for s in skus)
    alt_xlsx_bytes = build_alternates_workbook(skus) if has_alternates else None

    prefix = f"{client}_{batch}_Handover_Package"
    xlsx_name = f"{client}_Master_Marketplace_Upload.xlsx"

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Master Excel mapping workbook
        zf.writestr(f"{prefix}/{xlsx_name}", xlsx_bytes)
        # Alternate copies workbook
        if alt_xlsx_bytes:
            alt_xlsx_name = f"{client}_Alternate_Listing_Copies.xlsx"
            zf.writestr(f"{prefix}/{alt_xlsx_name}", alt_xlsx_bytes)
        # Package audit metadata (Feature #1)
        zf.writestr(f"{prefix}/package_metadata.json", metadata_bytes)
        # Handover README
        zf.writestr(f"{prefix}/README_Upload_Instructions.txt", readme_text)
        # Organized images
        for sid, files in matched_images.items():
            for fname, data in files:
                zf.writestr(f"{prefix}/Organized_SKU_Images/{sid}/{fname}", data)
        # Unassigned images
        for fname, data in unassigned_images:
            zf.writestr(f"{prefix}/Organized_SKU_Images/Unassigned_Assets/{fname}", data)

    buf.seek(0)
    return buf.getvalue()


# ──────────────────────────────────────────────
# Sample Data Generator for Dry-Run (Section 3 - #3)
# ──────────────────────────────────────────────

def get_sample_skus() -> list[dict]:
    return [
        {
            "sku_id": "SKU_SAMPLE_01",
            "brand": "SampleBrand",
            "product_type": "Pure Cotton Kurti",
            "color": "Indigo Blue",
            "category": "Women Ethnic Wear",
            "sizes": "S, M, L, XL, XXL",
            "mrp": 1299.0,
            "meesho_price": 349.0,
            "seller_config": {
                "amazon_quantity": 50,
                "gst_percent": 5,
                "hsn_code": "62114200"
            },
            "amazon": {
                "title": "SampleBrand Women's Pure Cotton Straight Kurti with Pocket | Daily Office Wear",
                "bullet_points": [
                    "100% PURE COTTON: Soft, breathable 60s count fabric for all-day comfort.",
                    "POCKET UTILITY: Includes convenient deep side pocket for phone and essentials.",
                    "COLORFAST & DURABLE: Follow the provided care label to help maintain the fabric's appearance and color.",
                    "VERSATILE STYLING: Pairs easily with denim, palazzo pants, or ethnic trousers.",
                    "PRECISE SIZING: Standard regular Indian fit from S(36) to XXL(44)."
                ],
                "backend_search_terms": "cotton kurti straight kurta daily wear office ethnic top",
                "description": "Experience daily comfort and elegance with SampleBrand Pure Cotton Straight Kurti."
            },
            "flipkart": {
                "title": "SampleBrand Women Solid Pure Cotton Straight Kurta (Blue)",
                "fabric": "Pure Cotton",
                "kurta_type": "Straight",
                "neck": "Round Neck",
                "sleeve": "3/4 Sleeve",
                "length_type": "Calf Length",
                "pattern": "Solid",
                "occasion": "Casual",
                "search_keywords": "cotton kurti, straight kurta women, summer kurti",
                "description": "Crafted from breathable pure cotton, this straight kurta combines understated elegance with everyday practicality."
            },
            "meesho": {
                "title": "Pure Cotton Straight Kurti with Pocket",
                "hinglish_hook_description": "100% Pure Cotton fabric jo garmi mein bhi de pura comfort! Daily wear aur office ke liye best choice.",
                "english_hook_description": "Premium pure cotton straight kurti tailored with a functional pocket for smart everyday wear.",
                "highlights": [
                    "Fabric: 100% Pure Cotton",
                    "Pocket: 1 Side Pocket Included",
                    "Fit: Regular Straight Fit",
                    "Care: Hand or Machine Wash"
                ]
            },
            "alternates": [
                {
                    "variant_id": f"V{i+1}",
                    "angle_theme": theme,
                    "amazon": {
                        "title": f"SampleBrand Women's Cotton Kurti - {theme} Edition",
                        "bullet_points": [
                            f"COMFORT {i+1}: Ultra-breathable weave for active daily ventilation.",
                            "DESIGN: Tailored silhouette designed for versatile ethnic styling.",
                            "COLORFAST & DURABLE: Follow the provided care label to help maintain the fabric's appearance.",
                            "VERSATILITY: Complements office trousers and ethnic skirts seamlessly.",
                            "CARE: Easy machine wash with mild detergent."
                        ],
                        "backend_search_terms": f"cotton kurti {theme.lower().replace('&', '').replace('  ', ' ')} women ethnic kurta",
                        "description": f"Designed specifically for {theme.lower()}. A premium wardrobe essential."
                    },
                    "flipkart": {
                        "title": f"SampleBrand Women Cotton Kurta - {theme}",
                        "search_keywords": f"kurti {theme.lower()}, cotton kurta",
                        "description": f"Elegant pure cotton kurta tailored for {theme.lower()}."
                    },
                    "meesho": {
                        "title": f"Cotton Kurti - {theme}",
                        "hinglish_hook_description": f"Har din ke liye badhiya cotton kurti! {theme} ke liye ekdum perfect.",
                        "english_hook_description": f"Premium daily cotton kurti crafted for {theme.lower()}.",
                        "highlights": ["Fabric: Cotton", f"Theme: {theme}", "Fit: Regular", "Wash: Normal"]
                    }
                }
                for i, theme in enumerate([
                    "Daily Office Workwear",
                    "Festive Celebrations",
                    "Summer Cooling Comfort",
                    "Everyday Essential",
                    "Modern Fusion Styling"
                ])
            ]
        }
    ]


# ──────────────────────────────────────────────
# FastAPI Application
# ──────────────────────────────────────────────

app = FastAPI(title="Listing Factory", version="2.0.0")


# ── Validation endpoint ──

@app.post("/api/validate-json")
async def validate_json(request: Request):
    """
    Strictly validate AI-generated JSON listing data against the schema.
    Enforces schema_version contract, structured per-SKU errors, backend search terms hygiene,
    Bullet 3 truth-boundary safety, and advisory warnings.
    """
    try:
        body = await request.json()
        raw = body.get("json_text", "")
        if not raw or not raw.strip():
            return JSONResponse({
                "valid": False,
                "global_errors": ["JSON payload is empty."],
                "global_warnings": [],
                "sku_results": []
            }, status_code=400)
        
        data = json.loads(raw)
        schema_ver, batch_cfg, merged_skus = normalize_and_merge_json(data)

        global_warnings = []
        global_errors = []

        # Quaternary Section 2 (#2): Schema Version Contract
        if schema_ver is not None and schema_ver != EXPECTED_SCHEMA_VERSION:
            global_errors.append(
                f"Expected schema_version '{EXPECTED_SCHEMA_VERSION}', got '{schema_ver}'. "
                "Please regenerate JSON using the current prompt."
            )
            return JSONResponse({
                "valid": False,
                "schema_version": schema_ver,
                "expected_schema_version": EXPECTED_SCHEMA_VERSION,
                "global_errors": global_errors,
                "global_warnings": [],
                "sku_results": []
            }, status_code=400)

        # Performance Safeguards (Soft Limit on SKU count)
        if len(merged_skus) > MAX_SKUS_SOFT_LIMIT:
            global_warnings.append(
                f"Batch has {len(merged_skus)} SKUs (recommended soft limit: {MAX_SKUS_SOFT_LIMIT}). "
                "Consider splitting into smaller batches for optimal performance."
            )

        sku_results = []
        parsed_skus = []
        overall_valid = True

        for idx, item in enumerate(merged_skus):
            sk_id = item.get("sku_id") or f"SKU_{idx+1:02d}"
            sku_errors = []
            sku_warnings = []

            # Check for "To be confirmed" fields
            for key, val in item.items():
                if isinstance(val, str) and val.strip().lower() == TO_BE_CONFIRMED.lower():
                    sku_warnings.append(f"Field '{key}' is marked '{TO_BE_CONFIRMED}'")
                elif isinstance(val, dict):
                    for subk, subv in val.items():
                        if isinstance(subv, str) and subv.strip().lower() == TO_BE_CONFIRMED.lower():
                            sku_warnings.append(f"{key}.{subk} is marked '{TO_BE_CONFIRMED}'")

            try:
                sku_obj = SKUItem(**item)
                parsed_skus.append(sku_obj)

                # Near-limit advisory checks
                t_len = len(sku_obj.amazon.title)
                if sku_obj.amazon.title != TO_BE_CONFIRMED and t_len > 165:
                    sku_warnings.append(f"Amazon title length is near limit ({t_len}/180 chars)")
                
                b_len = len(sku_obj.amazon.backend_search_terms.encode("utf-8"))
                if sku_obj.amazon.backend_search_terms != TO_BE_CONFIRMED and b_len > 225:
                    sku_warnings.append(f"Amazon search terms bytes near limit ({b_len}/240 bytes)")
                
                m_len = len(sku_obj.meesho.title)
                if sku_obj.meesho.title != TO_BE_CONFIRMED and m_len > 55:
                    sku_warnings.append(f"Meesho title length is near limit ({m_len}/60 chars)")

                # Section 2: Amazon Backend Search-Term Hygiene Check
                bst_errs, bst_warns = validate_amazon_backend_search_terms(
                    sku_obj.amazon.title, sku_obj.amazon.backend_search_terms, sku_obj.brand
                )
                if bst_errs:
                    sku_errors.extend(bst_errs)
                if bst_warns:
                    sku_warnings.extend(bst_warns)

                # Section 3: Bullet 3 Safety Check (Neutralize COLORFAST & DURABLE)
                if len(sku_obj.amazon.bullet_points) >= 3:
                    b3_errs, b3_warns = validate_bullet_3_safety(sku_obj.amazon.bullet_points[2])
                    if b3_errs:
                        sku_errors.extend(b3_errs)
                    if b3_warns:
                        sku_warnings.extend(b3_warns)

                # Validate alternates backend search terms & Bullet 3 as well
                for alt in sku_obj.alternates:
                    alt_bst_errs, alt_bst_warns = validate_amazon_backend_search_terms(
                        alt.amazon.title, alt.amazon.backend_search_terms, sku_obj.brand
                    )
                    if alt_bst_errs:
                        sku_errors.extend([f"Alternate {alt.variant_id}: {e}" for e in alt_bst_errs])
                    if alt_bst_warns:
                        sku_warnings.extend([f"Alternate {alt.variant_id}: {w}" for w in alt_bst_warns])

                    if len(alt.amazon.bullet_points) >= 3:
                        alt_b3_errs, alt_b3_warns = validate_bullet_3_safety(alt.amazon.bullet_points[2])
                        if alt_b3_errs:
                            sku_errors.extend([f"Alternate {alt.variant_id}: {e}" for e in alt_b3_errs])

            except Exception as pe:
                sku_errors.append(str(pe))

            if sku_errors:
                overall_valid = False

            sku_results.append({
                "sku_id": sk_id,
                "valid": len(sku_errors) == 0,
                "errors": sku_errors,
                "warnings": sku_warnings,
            })

        if not overall_valid:
            return JSONResponse({
                "valid": False,
                "schema_version": schema_ver or EXPECTED_SCHEMA_VERSION,
                "expected_schema_version": EXPECTED_SCHEMA_VERSION,
                "sku_count": len(merged_skus),
                "global_errors": global_errors,
                "global_warnings": global_warnings,
                "sku_results": sku_results,
            }, status_code=400)

        return JSONResponse({
            "valid": True,
            "schema_version": schema_ver or EXPECTED_SCHEMA_VERSION,
            "expected_schema_version": EXPECTED_SCHEMA_VERSION,
            "sku_count": len(parsed_skus),
            "global_errors": [],
            "global_warnings": global_warnings,
            "sku_results": sku_results,
        })

    except json.JSONDecodeError as e:
        return JSONResponse({
            "valid": False,
            "global_errors": [f"Invalid JSON syntax at line {e.lineno}, column {e.colno}: {e.msg}"],
            "global_warnings": [],
            "sku_results": []
        }, status_code=400)
    except Exception as e:
        return JSONResponse({
            "valid": False,
            "global_errors": [str(e)],
            "global_warnings": [],
            "sku_results": []
        }, status_code=400)


# ── Generation endpoint ──

@app.post("/api/generate")
async def generate_package(
    client_name: str = Form(...),
    batch_id: str = Form(...),
    category: str = Form(...),
    json_data: str = Form(...),
    images: list[UploadFile] = File(default=[]),
):
    """Generate the full delivery ZIP package with mapping workbooks and versioned disk archive."""
    try:
        raw = json.loads(json_data)
        schema_ver, batch_cfg, merged_skus = normalize_and_merge_json(raw)
        skus = [SKUItem(**item) for item in merged_skus]
    except Exception as e:
        return JSONResponse({"error": f"Schema validation error: {e}"}, status_code=400)

    image_files: list[tuple[str, bytes]] = []
    total_bytes = 0
    for img in images:
        data = await img.read()
        total_bytes += len(data)
        image_files.append((img.filename, data))

    safe_client = re.sub(r"[^\w\-]", "_", client_name.strip() or "Client")
    safe_batch = re.sub(r"[^\w\-]", "_", batch_id.strip() or "Batch_01")
    cat = category.strip() or "Women Ethnic Wear"
    ver_str = schema_ver or EXPECTED_SCHEMA_VERSION

    zip_bytes = build_zip(
        safe_client, safe_batch, cat, skus, image_files,
        validation_status="✅ Pass", schema_version=ver_str
    )

    # Versioned Output Archival
    pattern = str(OUTPUT_DIR / f"{safe_client}_{safe_batch}_v*.zip")
    existing_files = glob.glob(pattern)
    existing_versions = []
    for f in existing_files:
        m = re.search(r"_v(\d+)\.zip$", f)
        if m:
            existing_versions.append(int(m.group(1)))
    next_ver = (max(existing_versions) + 1) if existing_versions else 1
    versioned_file = OUTPUT_DIR / f"{safe_client}_{safe_batch}_v{next_ver}.zip"
    try:
        versioned_file.write_bytes(zip_bytes)
    except Exception as io_err:
        print(f"Warning: could not write versioned archive {versioned_file}: {io_err}")

    zip_name = f"{safe_client}_{safe_batch}_Handover_Package.zip"

    return StreamingResponse(
        io.BytesIO(zip_bytes),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{zip_name}"'},
    )


# ── Sample / Dry-Run Endpoint ──

@app.post("/api/generate-sample")
async def generate_sample():
    """Generate and return a sample handover ZIP archive for testing and demos."""
    sample_skus_raw = get_sample_skus()
    skus = [SKUItem(**item) for item in sample_skus_raw]
    zip_bytes = build_zip(
        client="SampleBrand",
        batch="Sample_Batch_01",
        category="Women Ethnic Wear",
        skus=skus,
        image_files=[],
        validation_status="✅ Pass",
        schema_version=EXPECTED_SCHEMA_VERSION
    )
    return StreamingResponse(
        io.BytesIO(zip_bytes),
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="Sample_Handover_Package.zip"'},
    )


# ── History & Rollback Endpoint ──

@app.get("/api/history")
async def get_history(client: Optional[str] = None, batch: Optional[str] = None):
    """Retrieve archived package versions from the output directory."""
    files = []
    for p in OUTPUT_DIR.glob("*.zip"):
        st = p.stat()
        files.append({
            "filename": p.name,
            "size_bytes": st.st_size,
            "created_at": datetime.fromtimestamp(st.st_mtime).isoformat(),
        })
    files.sort(key=lambda x: x["created_at"], reverse=True)
    return JSONResponse({"packages": files})


# ── Frontend route (serves index.html from disk) ──

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
    print(f"\n  [*] {TOOL_VERSION} - Multi-Marketplace Packaging Studio")
    print("  " + "=" * 63)
    print("  > Server running at: http://127.0.0.1:8000")
    print("  > Press Ctrl+C to stop\n")
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
