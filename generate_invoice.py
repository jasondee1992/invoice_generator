from __future__ import annotations

import argparse
import math
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

import pandas as pd
from PIL import Image, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 40
ACCENT_COLOR = colors.HexColor("#1F4E79")
LIGHT_BORDER = colors.HexColor("#D9E2F3")
TEXT_COLOR = colors.HexColor("#222222")
SUPPORTED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
BUILDER_NAME = "JasonD"
APP_TITLE = f"Invoice Generator Build by: {BUILDER_NAME}"


def get_base_dir() -> Path:
    """Return the folder where the script or packaged executable lives."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolve_input_path(path_value: str) -> Path:
    """Resolve input paths relative to the script/exe folder."""
    path = Path(path_value)
    if path.is_absolute():
        return path
    return get_base_dir() / path


def notify_user(title: str, message: str, is_error: bool = False) -> None:
    """Show a popup when packaged as an exe, otherwise just print to the console."""
    if not getattr(sys, "frozen", False):
        return

    try:
        import tkinter
        from tkinter import messagebox

        root = tkinter.Tk()
        root.withdraw()
        if is_error:
            messagebox.showerror(title, message)
        else:
            messagebox.showinfo(title, message)
        root.destroy()
    except Exception:
        pass


def parse_arguments() -> argparse.Namespace:
    """Read configurable file paths from the command line."""
    parser = argparse.ArgumentParser(
        description="Generate a PDF invoice from an Excel file and evidence images."
    )
    parser.add_argument(
        "--input-excel-path",
        default="data/invoice_data.xlsx",
        help="Path to the Excel workbook containing invoice_info and line_items sheets.",
    )
    parser.add_argument(
        "--evidence-folder-path",
        default="evidences",
        help="Folder containing evidence images to append after the invoice page.",
    )
    parser.add_argument(
        "--output-pdf-path",
        default="output/invoice_output.pdf",
        help="Path of the PDF file that will be generated.",
    )
    return parser.parse_args()


def require_file(path: Path, label: str) -> None:
    """Raise a clear error if a required file is missing."""
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")
    if not path.is_file():
        raise FileNotFoundError(f"{label} is not a file: {path}")


def require_folder(path: Path, label: str) -> None:
    """Raise a clear error if a required folder is missing."""
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")
    if not path.is_dir():
        raise NotADirectoryError(f"{label} is not a folder: {path}")


def normalize_text(value: object) -> str:
    """Convert Excel values into safe display text."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def parse_decimal(value: object, default: Decimal = Decimal("0")) -> Decimal:
    """Convert numeric-looking values to Decimal for safer money formatting."""
    text = normalize_text(value)
    if not text:
        return default
    cleaned = text.replace(",", "").replace("$", "").strip()
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return default


def format_currency(amount: Decimal, currency_symbol: str) -> str:
    """Format money with a currency symbol and two decimal places."""
    return f"{currency_symbol} {amount:,.2f}"


def validate_columns(df: pd.DataFrame, required_columns: Iterable[str], sheet_name: str) -> None:
    """Ensure the Excel sheet contains all required columns."""
    missing = [column for column in required_columns if column not in df.columns]
    if missing:
        raise ValueError(
            f"Missing columns in sheet '{sheet_name}': {', '.join(missing)}"
        )


def build_multiline_block(*parts: object) -> str:
    """Join non-empty text values into a clean multi-line block."""
    lines = [normalize_text(part) for part in parts if normalize_text(part)]
    return "\n".join(lines)


def load_invoice_data(excel_path: Path) -> tuple[dict, list[dict]]:
    """
    Load one invoice header row and multiple line items from Excel.

    Expected sheets:
    - invoice_info: one row only
    - line_items: one or more rows
    """
    require_file(excel_path, "Excel file")

    workbook = pd.ExcelFile(excel_path)
    required_sheets = {"invoice_info", "line_items"}
    missing_sheets = required_sheets.difference(workbook.sheet_names)
    if missing_sheets:
        raise ValueError(
            f"Missing worksheet(s): {', '.join(sorted(missing_sheets))}. "
            "Expected sheets: invoice_info and line_items."
        )

    invoice_df = pd.read_excel(excel_path, sheet_name="invoice_info")
    line_items_df = pd.read_excel(excel_path, sheet_name="line_items")

    validate_columns(
        invoice_df,
        [
            "invoice_number",
            "invoice_date",
            "due_date",
            "sender_name",
            "sender_address",
            "client_name",
            "client_address",
            "payment_details",
            "contact_email",
        ],
        "invoice_info",
    )
    validate_columns(
        line_items_df,
        ["description", "service_date", "rate"],
        "line_items",
    )

    if invoice_df.empty:
        raise ValueError("Sheet 'invoice_info' is empty.")
    if len(invoice_df) > 1:
        raise ValueError(
            "Sheet 'invoice_info' should contain only one invoice row per run."
        )
    if line_items_df.empty:
        raise ValueError("Sheet 'line_items' is empty.")

    invoice_row = invoice_df.iloc[0].to_dict()

    currency_symbol = normalize_text(invoice_row.get("currency_symbol")) or "$"
    items: list[dict] = []
    calculated_subtotal = Decimal("0")

    for _, row in line_items_df.iterrows():
        description = normalize_text(row.get("description"))
        service_date = normalize_text(row.get("service_date"))
        rate = parse_decimal(row.get("rate"))
        quantity = parse_decimal(row.get("quantity"), default=Decimal("1"))
        explicit_amount = normalize_text(row.get("amount"))
        amount = parse_decimal(explicit_amount) if explicit_amount else rate * quantity

        if not description:
            continue

        item = {
            "description": description,
            "service_date": service_date,
            "quantity": quantity,
            "rate": rate,
            "amount": amount,
        }
        items.append(item)
        calculated_subtotal += amount

    if not items:
        raise ValueError("No valid line items were found in sheet 'line_items'.")

    subtotal = parse_decimal(invoice_row.get("subtotal"), default=calculated_subtotal)
    total = parse_decimal(invoice_row.get("total"), default=subtotal)

    invoice_data = {
        "invoice_number": normalize_text(invoice_row.get("invoice_number")),
        "invoice_date": normalize_text(invoice_row.get("invoice_date")),
        "due_date": normalize_text(invoice_row.get("due_date")),
        "sender_block": build_multiline_block(
            invoice_row.get("sender_name"),
            invoice_row.get("sender_address"),
            invoice_row.get("sender_phone"),
            invoice_row.get("sender_email"),
        ),
        "client_block": build_multiline_block(
            invoice_row.get("client_name"),
            invoice_row.get("client_company"),
            invoice_row.get("client_address"),
            invoice_row.get("client_email"),
        ),
        "payment_details": normalize_text(invoice_row.get("payment_details")),
        "contact_email": normalize_text(invoice_row.get("contact_email")),
        "notes": normalize_text(invoice_row.get("notes")),
        "currency_symbol": currency_symbol,
        "subtotal": subtotal,
        "total": total,
    }

    return invoice_data, items


def get_image_files(evidence_folder: Path) -> list[Path]:
    """Collect supported image files from the evidence folder."""
    require_folder(evidence_folder, "Evidence folder")

    image_files = sorted(
        path
        for path in evidence_folder.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
    )
    if not image_files:
        raise ValueError(
            "No image files found in the evidence folder. "
            "Supported formats: .jpg, .jpeg, .png, .webp, .bmp"
        )
    return image_files


def draw_text_block(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    width: float,
    font_name: str = "Helvetica",
    font_size: int = 10,
    leading: int = 14,
) -> float:
    """Draw multi-line text and return the ending Y position."""
    text_object = pdf.beginText()
    text_object.setTextOrigin(x, y)
    text_object.setFont(font_name, font_size)
    text_object.setLeading(leading)

    for line in text.splitlines() or [""]:
        safe_line = line.strip()
        if safe_line:
            text_object.textLine(safe_line[: max(1, int(width / (font_size * 0.52)))])
        else:
            text_object.textLine("")

    pdf.drawText(text_object)
    line_count = max(1, len(text.splitlines()))
    return y - (line_count * leading)


def draw_wrapped_line(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    font_name: str = "Helvetica",
    font_size: int = 10,
) -> None:
    """Draw a single line trimmed to fit inside a table cell."""
    safe_text = normalize_text(text)
    while safe_text and pdf.stringWidth(safe_text, font_name, font_size) > max_width:
        safe_text = safe_text[:-1]
    if safe_text != normalize_text(text):
        safe_text = safe_text[:-3].rstrip() + "..."
    pdf.setFont(font_name, font_size)
    pdf.drawString(x, y, safe_text)


def wrap_text_lines(
    pdf: canvas.Canvas,
    text: str,
    max_width: float,
    font_name: str = "Helvetica",
    font_size: int = 10,
) -> list[str]:
    """Wrap text to a cell width while preserving manual line breaks from Excel."""
    raw_text = normalize_text(text)
    if not raw_text:
        return [""]

    wrapped_lines: list[str] = []
    for paragraph in raw_text.splitlines() or [""]:
        paragraph = paragraph.strip()
        if not paragraph:
            wrapped_lines.append("")
            continue

        current_line = ""
        for word in paragraph.split():
            trial_line = word if not current_line else f"{current_line} {word}"
            if pdf.stringWidth(trial_line, font_name, font_size) <= max_width:
                current_line = trial_line
            else:
                if current_line:
                    wrapped_lines.append(current_line)
                current_line = word

        if current_line:
            wrapped_lines.append(current_line)

    return wrapped_lines or [""]


def draw_cell_lines(
    pdf: canvas.Canvas,
    lines: list[str],
    x: float,
    top_y: float,
    font_name: str = "Helvetica",
    font_size: int = 10,
    leading: int = 12,
) -> None:
    """Draw multiple lines inside a table cell from top to bottom."""
    pdf.setFont(font_name, font_size)
    y = top_y
    for line in lines:
        pdf.drawString(x, y, line)
        y -= leading


def calculate_block_height(line_count: int, padding_top: float, padding_bottom: float, leading: float) -> float:
    """Calculate the required box height for a multi-line text block."""
    safe_line_count = max(1, line_count)
    return padding_top + padding_bottom + ((safe_line_count - 1) * leading)


def prepare_line_item_layout(
    pdf: canvas.Canvas,
    item: dict,
    col_widths: list[float],
    min_row_height: float,
    cell_padding_top: float,
    cell_padding_bottom: float,
    cell_leading: float,
) -> dict:
    """Precompute wrapped cell content and the row height for one line item."""
    description_lines = wrap_text_lines(
        pdf,
        item["description"],
        col_widths[0] - 16,
    )
    date_lines = wrap_text_lines(
        pdf,
        item["service_date"],
        col_widths[1] - 14,
    )
    line_count = max(len(description_lines), len(date_lines), 1)
    row_height = max(
        min_row_height,
        cell_padding_top + cell_padding_bottom + ((line_count - 1) * cell_leading),
    )
    return {
        "description_lines": description_lines,
        "date_lines": date_lines,
        "row_height": row_height,
        "item": item,
    }


def draw_invoice_footer(pdf: canvas.Canvas, invoice: dict) -> None:
    """Draw the footer used on invoice pages."""
    if invoice["notes"]:
        pdf.setFont("Helvetica", 9)
        pdf.setFillColor(colors.HexColor("#555555"))
        pdf.drawString(MARGIN, 68, f"Notes: {invoice['notes']}")

    pdf.setStrokeColor(ACCENT_COLOR)
    pdf.line(MARGIN, 48, PAGE_WIDTH - MARGIN, 48)
    pdf.setFillColor(colors.HexColor("#555555"))
    pdf.setFont("Helvetica", 9)
    pdf.drawCentredString(
        PAGE_WIDTH / 2,
        34,
        f"Contact: {invoice['contact_email']}",
    )
    pdf.setFillColor(TEXT_COLOR)


def get_footer_top_y(invoice: dict) -> float:
    """Return the highest Y position occupied by the footer area."""
    return 82 if invoice["notes"] else 52


def draw_invoice_top_sections(pdf: canvas.Canvas, invoice: dict, page_number: int) -> float:
    """Draw the page header and return the Y position where the table header should start."""
    top_y = PAGE_HEIGHT - MARGIN
    pdf.setStrokeColor(ACCENT_COLOR)
    pdf.setFillColor(TEXT_COLOR)

    if page_number == 1:
        pdf.setTitle(f"Invoice {invoice['invoice_number']}")
        pdf.setAuthor(BUILDER_NAME)
        pdf.setSubject(f"Generated by {BUILDER_NAME}")
        pdf.setFont("Helvetica-Bold", 22)
        pdf.setFillColor(ACCENT_COLOR)
        pdf.drawString(MARGIN, top_y, "INVOICE")
        pdf.setFillColor(TEXT_COLOR)

        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(MARGIN, top_y - 28, "Billed By")
        draw_text_block(
            pdf,
            invoice["sender_block"],
            MARGIN,
            top_y - 44,
            width=230,
            font_size=10,
            leading=13,
        )

        right_x = PAGE_WIDTH - 210
        pdf.roundRect(right_x, top_y - 92, 170, 82, 8, stroke=1, fill=0)
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(right_x + 12, top_y - 28, "Invoice Number")
        pdf.drawRightString(PAGE_WIDTH - 52, top_y - 28, invoice["invoice_number"])
        pdf.drawString(right_x + 12, top_y - 48, "Invoice Date")
        pdf.drawRightString(PAGE_WIDTH - 52, top_y - 48, invoice["invoice_date"])
        pdf.drawString(right_x + 12, top_y - 68, "Due Date")
        pdf.drawRightString(PAGE_WIDTH - 52, top_y - 68, invoice["due_date"])

        bill_to_y = top_y - 148
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(MARGIN, bill_to_y, "Bill To")
        pdf.roundRect(MARGIN, bill_to_y - 84, 250, 72, 8, stroke=1, fill=0)
        draw_text_block(
            pdf,
            invoice["client_block"],
            MARGIN + 12,
            bill_to_y - 28,
            width=226,
            font_size=10,
            leading=13,
        )
        return bill_to_y - 130

    pdf.setFont("Helvetica-Bold", 20)
    pdf.setFillColor(ACCENT_COLOR)
    pdf.drawString(MARGIN, top_y, "INVOICE")
    pdf.setFillColor(TEXT_COLOR)
    pdf.setFont("Helvetica", 10)
    pdf.drawString(MARGIN, top_y - 22, f"Invoice Number: {invoice['invoice_number']}")
    pdf.drawString(MARGIN + 220, top_y - 22, f"Invoice Date: {invoice['invoice_date']}")
    pdf.drawRightString(PAGE_WIDTH - MARGIN, top_y - 22, f"Page {page_number}")
    return top_y - 52


def draw_table_header(
    pdf: canvas.Canvas,
    table_x: float,
    table_y: float,
    table_width: float,
    headers: list[str],
    col_widths: list[float],
    header_height: float,
) -> None:
    """Draw the table header row."""
    pdf.setFillColor(ACCENT_COLOR)
    pdf.rect(table_x, table_y, table_width, header_height, stroke=0, fill=1)
    pdf.setFillColor(colors.white)
    pdf.setFont("Helvetica-Bold", 10)

    cursor_x = table_x + 8
    for header, width in zip(headers, col_widths):
        pdf.drawString(cursor_x, table_y + 8, header)
        cursor_x += width

    pdf.setFillColor(TEXT_COLOR)
    pdf.setStrokeColor(LIGHT_BORDER)


def draw_line_item_row(
    pdf: canvas.Canvas,
    layout: dict,
    table_x: float,
    current_y: float,
    table_width: float,
    col_widths: list[float],
    currency_symbol: str,
    cell_padding_top: float,
    cell_leading: float,
) -> None:
    """Draw one prepared line item row."""
    row_height = layout["row_height"]
    item = layout["item"]

    pdf.rect(table_x, current_y, table_width, row_height, stroke=1, fill=0)

    inner_x = table_x + 8
    text_top_y = current_y + row_height - cell_padding_top
    draw_cell_lines(
        pdf,
        layout["description_lines"],
        inner_x,
        text_top_y,
        leading=cell_leading,
    )
    inner_x += col_widths[0]
    draw_cell_lines(
        pdf,
        layout["date_lines"],
        inner_x - 16,
        text_top_y,
        leading=cell_leading,
    )
    inner_x += col_widths[1]
    pdf.drawString(inner_x + 2, text_top_y, f"{item['quantity']:g}")
    inner_x += col_widths[2]
    pdf.drawString(
        inner_x + 2,
        text_top_y,
        format_currency(item["rate"], currency_symbol),
    )
    inner_x += col_widths[3]
    pdf.drawString(
        inner_x + 2,
        text_top_y,
        format_currency(item["amount"], currency_symbol),
    )


def draw_summary_and_payment(pdf: canvas.Canvas, invoice: dict, current_y: float, currency_symbol: str) -> None:
    """Draw subtotal/total and payment details below the final table row."""
    summary_box_width = 170
    summary_box_height = 54
    summary_gap_from_table = 14
    payment_gap_from_summary = 16
    payment_padding_top = 20
    payment_padding_bottom = 16
    payment_leading = 13
    payment_title_gap = 18
    payment_text_width = PAGE_WIDTH - (MARGIN * 2) - 24
    payment_lines = wrap_text_lines(
        pdf,
        invoice["payment_details"],
        payment_text_width,
        font_size=10,
    )
    payment_box_height = max(
        52,
        calculate_block_height(
            len(payment_lines),
            payment_padding_top + payment_title_gap,
            payment_padding_bottom,
            payment_leading,
        ),
    )

    footer_top_y = get_footer_top_y(invoice)
    payment_box_y = footer_top_y + 14
    summary_x = PAGE_WIDTH - MARGIN - summary_box_width
    summary_y = payment_box_y + payment_box_height + payment_gap_from_summary

    minimum_summary_y = current_y - summary_box_height - summary_gap_from_table
    if summary_y < minimum_summary_y:
        summary_y = minimum_summary_y

    pdf.roundRect(summary_x, summary_y, summary_box_width, summary_box_height, 8, stroke=1, fill=0)
    pdf.setFont("Helvetica", 10)
    pdf.drawString(summary_x + 12, summary_y + 34, "Subtotal")
    pdf.drawRightString(
        summary_x + summary_box_width - 12,
        summary_y + 34,
        format_currency(invoice["subtotal"], currency_symbol),
    )
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(summary_x + 12, summary_y + 14, "Total")
    pdf.drawRightString(
        summary_x + summary_box_width - 12,
        summary_y + 14,
        format_currency(invoice["total"], currency_symbol),
    )

    pdf.roundRect(MARGIN, payment_box_y, PAGE_WIDTH - (MARGIN * 2), payment_box_height, 8, stroke=1, fill=0)
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(MARGIN + 12, payment_box_y + payment_box_height - 20, "Payment Details")
    draw_cell_lines(
        pdf,
        payment_lines,
        MARGIN + 12,
        payment_box_y + payment_box_height - (payment_padding_top + payment_title_gap),
        font_size=10,
        leading=payment_leading,
    )


def draw_invoice_pages(pdf: canvas.Canvas, invoice: dict, items: list[dict]) -> int:
    """Draw one or more invoice pages with repeated table headers and final totals on the last page."""
    pdf.setTitle(f"Invoice {invoice['invoice_number']}")
    table_x = MARGIN
    table_width = PAGE_WIDTH - (MARGIN * 2)
    min_row_height = 24
    header_height = 26
    col_widths = [220, 90, 60, 75, 75]
    headers = ["Description", "Date", "Qty", "Rate", "Amount"]
    currency_symbol = invoice["currency_symbol"]
    cell_padding_top = 16
    cell_padding_bottom = 8
    cell_leading = 12
    footer_top_y = get_footer_top_y(invoice)
    continuation_bottom_y = footer_top_y + 12

    summary_box_height = 54
    summary_gap_from_table = 14
    payment_gap_from_summary = 16
    payment_padding_top = 20
    payment_padding_bottom = 16
    payment_leading = 13
    payment_title_gap = 18
    payment_text_width = PAGE_WIDTH - (MARGIN * 2) - 24
    payment_lines = wrap_text_lines(
        pdf,
        invoice["payment_details"],
        payment_text_width,
        font_size=10,
    )
    payment_box_height = max(
        52,
        calculate_block_height(
            len(payment_lines),
            payment_padding_top + payment_title_gap,
            payment_padding_bottom,
            payment_leading,
        ),
    )
    final_bottom_y = (
        (footer_top_y + 14)
        + payment_box_height
        + payment_gap_from_summary
        + summary_box_height
        + summary_gap_from_table
        + 12
    )

    prepared_items = [
        prepare_line_item_layout(
            pdf,
            item,
            col_widths,
            min_row_height,
            cell_padding_top,
            cell_padding_bottom,
            cell_leading,
        )
        for item in items
    ]

    page_number = 1
    table_y = draw_invoice_top_sections(pdf, invoice, page_number)
    draw_table_header(pdf, table_x, table_y, table_width, headers, col_widths, header_height)
    current_y = table_y

    for index, layout in enumerate(prepared_items):
        is_last_item = index == len(prepared_items) - 1
        min_bottom_y = final_bottom_y if is_last_item else continuation_bottom_y

        if current_y - layout["row_height"] < min_bottom_y:
            draw_invoice_footer(pdf, invoice)
            pdf.showPage()
            page_number += 1
            table_y = draw_invoice_top_sections(pdf, invoice, page_number)
            draw_table_header(pdf, table_x, table_y, table_width, headers, col_widths, header_height)
            current_y = table_y

        current_y -= layout["row_height"]
        draw_line_item_row(
            pdf,
            layout,
            table_x,
            current_y,
            table_width,
            col_widths,
            currency_symbol,
            cell_padding_top,
            cell_leading,
        )

    draw_summary_and_payment(pdf, invoice, current_y, currency_symbol)
    draw_invoice_footer(pdf, invoice)
    pdf.showPage()
    return page_number


def draw_image_in_box(
    pdf: canvas.Canvas,
    image_path: Path,
    x: float,
    y: float,
    box_width: float,
    box_height: float,
) -> None:
    """Draw one image inside a bounding box without distorting it."""
    with Image.open(image_path) as image:
        prepared_image = ImageOps.exif_transpose(image)
        if prepared_image.mode not in ("RGB", "RGBA"):
            prepared_image = prepared_image.convert("RGB")

        # Resize large source images before embedding them to keep the PDF file size manageable.
        target_dpi = 150
        max_pixel_width = max(1, int((box_width / 72) * target_dpi))
        max_pixel_height = max(1, int((box_height / 72) * target_dpi))
        prepared_image = prepared_image.copy()
        prepared_image.thumbnail((max_pixel_width, max_pixel_height), Image.Resampling.LANCZOS)

        image_width, image_height = prepared_image.size
        scale = min(box_width / image_width, box_height / image_height)
        draw_width = image_width * scale
        draw_height = image_height * scale

        draw_x = x + (box_width - draw_width) / 2
        draw_y = y + (box_height - draw_height) / 2

        pdf.drawImage(
            ImageReader(prepared_image),
            draw_x,
            draw_y,
            width=draw_width,
            height=draw_height,
            preserveAspectRatio=True,
            anchor="c",
        )


def draw_evidence_pages(pdf: canvas.Canvas, image_files: list[Path], starting_page_number: int) -> None:
    """Draw evidence images in a 2x2 grid, four images per page."""
    images_per_page = 4
    caption_height = 14
    gap = 16
    grid_top = PAGE_HEIGHT - MARGIN - 40
    grid_bottom = MARGIN + 30
    grid_height = grid_top - grid_bottom
    cell_width = (PAGE_WIDTH - (MARGIN * 2) - gap) / 2
    cell_height = (grid_height - gap) / 2

    for page_index in range(math.ceil(len(image_files) / images_per_page)):
        page_images = image_files[page_index * images_per_page : (page_index + 1) * images_per_page]

        pdf.setFont("Helvetica-Bold", 16)
        pdf.setFillColor(ACCENT_COLOR)
        pdf.drawString(MARGIN, PAGE_HEIGHT - MARGIN, "Evidence Images")
        pdf.setFillColor(TEXT_COLOR)
        pdf.setFont("Helvetica", 9)
        pdf.drawRightString(
            PAGE_WIDTH - MARGIN,
            PAGE_HEIGHT - MARGIN + 2,
            f"Page {starting_page_number + page_index}",
        )

        for index, image_path in enumerate(page_images):
            row = index // 2
            col = index % 2

            x = MARGIN + col * (cell_width + gap)
            y = grid_top - ((row + 1) * cell_height) - (row * gap)

            pdf.setStrokeColor(LIGHT_BORDER)
            pdf.roundRect(x, y, cell_width, cell_height, 8, stroke=1, fill=0)

            usable_height = cell_height - caption_height - 10
            draw_image_in_box(pdf, image_path, x + 8, y + caption_height + 6, cell_width - 16, usable_height - 8)

            pdf.setFillColor(colors.HexColor("#666666"))
            pdf.setFont("Helvetica", 8)
            caption = image_path.name
            max_caption_width = cell_width - 16
            while caption and pdf.stringWidth(caption, "Helvetica", 8) > max_caption_width:
                caption = caption[:-1]
            if caption != image_path.name:
                caption = caption[:-3].rstrip() + "..."
            pdf.drawString(x + 8, y + 6, caption)
            pdf.setFillColor(TEXT_COLOR)

        pdf.showPage()


def generate_invoice_pdf(excel_path: Path, evidence_folder: Path, output_pdf_path: Path) -> Path:
    """Main workflow for reading data and writing the final PDF."""
    invoice_data, line_items = load_invoice_data(excel_path)
    image_files = get_image_files(evidence_folder)

    output_pdf_path.parent.mkdir(parents=True, exist_ok=True)

    pdf = canvas.Canvas(str(output_pdf_path), pagesize=A4)
    invoice_page_count = draw_invoice_pages(pdf, invoice_data, line_items)
    draw_evidence_pages(pdf, image_files, starting_page_number=invoice_page_count + 1)
    pdf.save()

    return output_pdf_path


def main() -> None:
    args = parse_arguments()
    excel_path = resolve_input_path(args.input_excel_path)
    evidence_folder = resolve_input_path(args.evidence_folder_path)
    output_pdf_path = resolve_input_path(args.output_pdf_path)

    try:
        generated_pdf = generate_invoice_pdf(excel_path, evidence_folder, output_pdf_path)
        message = (
            f"PDF generated successfully:\n{generated_pdf}\n\n"
        )
        print(message)
        notify_user(APP_TITLE, message)
    except Exception as error:
        message = f"Error: {error}"
        print(message)
        notify_user(APP_TITLE, message, is_error=True)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
