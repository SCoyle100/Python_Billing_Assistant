try:
    import win32com.client
except ImportError:
    win32com = None
import datetime
import itertools
import re
import sys
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

# Word constants
wdActiveEndPageNumber = 3
wdReplaceOne = 1
wdFindStop = 0
wdCollapseEnd = 0  # Collapse to end of range
wdCharacter = 1    # Unit for character movement
PENSACOLA_MARGIN_MULTIPLIER = 1108.0 / 950.0
SERVICE_PERIOD_PATTERN = re.compile(
    r"(?P<start_month>\d{1,3})/(?P<start_day>\d{1,3})/(?P<start_year>\d{2}|\d{4})"
    r"\s*[-\u2013\u2014]\s*"
    r"(?P<end_month>\d{1,3})/(?P<end_day>\d{1,3})/(?P<end_year>\d{2}|\d{4})"
)
SINGLE_SERVICE_DATE_PATTERN = re.compile(
    r"\d{1,2}/\d{1,2}/(?:\d{2}|\d{4})"
)
# Matrix placements are normally four-week/month-length periods. Date
# subtraction is 27 days for a 28-day inclusive service period.
TARGET_MATRIX_SERVICE_PERIOD_DAYS = 27
MIN_MATRIX_SERVICE_PERIOD_DAYS = 20
MAX_MATRIX_SERVICE_PERIOD_DAYS = 35


def _component_candidates(value, maximum, max_digits=2):
    """Return valid values obtainable without reordering the source digits."""
    digits = str(value)
    candidates = []

    # Only remove overflow digits. A normal-width but invalid component such as
    # month 99 is not safe to guess at.
    length = min(len(digits), max_digits)
    for positions in itertools.combinations(range(len(digits)), length):
        candidate_text = "".join(digits[position] for position in positions)
        candidate = int(candidate_text)
        if 1 <= candidate <= maximum and candidate not in candidates:
            candidates.append(candidate)

    return candidates


def _year_value(value):
    year = int(value)
    return 2000 + year if len(str(value)) == 2 else year


def _date_candidates(month_text, day_text, year_text):
    candidates = []
    for month in _component_candidates(month_text, 12):
        for day in _component_candidates(day_text, 31):
            try:
                candidates.append(datetime.date(_year_value(year_text), month, day))
            except ValueError:
                continue
    return candidates


def _is_valid_date_parts(month_text, day_text, year_text):
    try:
        datetime.date(_year_value(year_text), int(month_text), int(day_text))
        return len(month_text) <= 2 and len(day_text) <= 2
    except ValueError:
        return False


def correct_matrix_service_period(value):
    """Repair malformed Matrix date ranges by preferring a roughly monthly span.

    For example, ``8/313/26-9/27/26`` has two plausible start days (31 and
    13).  A start on August 31 produces the normal 27-day billing period, so it
    is preferred over the unusual 45-day alternative.
    """
    text = str(value or "")

    def replace_match(match):
        parts = match.groupdict()
        start_valid = _is_valid_date_parts(
            parts["start_month"], parts["start_day"], parts["start_year"]
        )
        end_valid = _is_valid_date_parts(
            parts["end_month"], parts["end_day"], parts["end_year"]
        )

        # Valid ranges are left untouched, including their original spacing.
        if start_valid and end_valid:
            return match.group(0)

        start_candidates = _date_candidates(
            parts["start_month"], parts["start_day"], parts["start_year"]
        )
        end_candidates = _date_candidates(
            parts["end_month"], parts["end_day"], parts["end_year"]
        )
        plausible_ranges = []
        for start_date in start_candidates:
            for end_date in end_candidates:
                duration = (end_date - start_date).days
                if (
                    MIN_MATRIX_SERVICE_PERIOD_DAYS
                    <= duration
                    <= MAX_MATRIX_SERVICE_PERIOD_DAYS
                ):
                    plausible_ranges.append((start_date, end_date, duration))

        if not plausible_ranges:
            return match.group(0)

        start_date, end_date, _ = min(
            plausible_ranges,
            key=lambda item: (
                abs(item[2] - TARGET_MATRIX_SERVICE_PERIOD_DAYS),
                -item[2],
            ),
        )
        start_year = (
            str(start_date.year)
            if len(parts["start_year"]) == 4
            else f"{start_date.year % 100:02d}"
        )
        end_year = (
            str(end_date.year)
            if len(parts["end_year"]) == 4
            else f"{end_date.year % 100:02d}"
        )
        return (
            f"{start_date.month}/{start_date.day}/{start_year} - "
            f"{end_date.month}/{end_date.day}/{end_year}"
        )

    return SERVICE_PERIOD_PATTERN.sub(replace_match, text)


def first_matrix_service_period(value):
    """Return only the first date or date range found in a Matrix service cell."""
    corrected = correct_matrix_service_period(value)
    range_match = SERVICE_PERIOD_PATTERN.search(corrected)
    if range_match:
        return range_match.group(0).strip()

    date_match = SINGLE_SERVICE_DATE_PATTERN.search(corrected)
    if date_match:
        return date_match.group(0).strip()

    return str(corrected or "").strip()


def normalize_page_service_periods(page_market_mapping):
    """Keep the pre-read page mapping aligned with corrections saved to Word."""
    if not page_market_mapping:
        return

    for page_num, page_data in list(page_market_mapping.items()):
        if isinstance(page_data, tuple) and len(page_data) >= 2:
            corrected = first_matrix_service_period(page_data[1])
            page_market_mapping[page_num] = (page_data[0], corrected, *page_data[2:])

def parse_dollar_amount(dollar_str):
    """
    Converts a string like '$1,234.56' to a float (e.g. 1234.56).
    """
    cleaned = re.sub(r'[^\d\.]', '', dollar_str)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

def format_dollar_amount(value):
    """
    Formats a float 1234.56 into '$1,234.56' format.
    Ensures comma separators for values over 1000.
    """
    formatted = f"${value:,.2f}"
    # Extra check to ensure comma is present for values over 1000
    if value >= 1000 and ',' not in formatted:
        # Alternative formatting method if f-string doesn't work
        whole_part = int(value)
        formatted = '${:,}.{:02d}'.format(whole_part, int((value - whole_part) * 100))
    return formatted


def amount_to_cents(value):
    """Convert a currency-like value to integer cents without float rounding."""
    cleaned = re.sub(r"[^\d.\-]", "", str(value or ""))
    if not cleaned:
        return None
    try:
        decimal_value = Decimal(cleaned).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return None
    return int(decimal_value * 100)


def format_cents(cents):
    return f"${Decimal(cents) / Decimal(100):,.2f}"


def prompt_matrix_line_item_amounts(line_items, email_total):
    """Show one dialog for correcting every Matrix line item on a page."""
    try:
        import tkinter as tk
        from tkinter import messagebox
    except ImportError as exc:
        raise RuntimeError(
            "Matrix line-item amounts need manual review, but tkinter is unavailable."
        ) from exc

    root = tk.Tk()
    root.withdraw()
    dialog = tk.Toplevel(root)
    dialog.title("Review Matrix line items")
    dialog.resizable(False, False)
    dialog.attributes("-topmost", True)

    target_cents = amount_to_cents(email_total)
    result = {"values": None}

    tk.Label(
        dialog,
        text=(
            f"The Matrix backup has {len(line_items)} line items that do not add up "
            f"to the email total of {format_cents(target_cents)}.\n"
            "Enter the correct amount for each line item. The entries must add up "
            "to the email total."
        ),
        justify="left",
        padx=12,
        pady=10,
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    entries = []
    for index, item in enumerate(line_items, start=1):
        label = str(item.get("label") or f"Line item {index}").strip()
        tk.Label(dialog, text=label, anchor="w", padx=12).grid(
            row=index, column=0, sticky="w"
        )
        entry = tk.Entry(dialog, width=16, justify="right")
        entry.insert(0, format_cents(amount_to_cents(item.get("amount"))))
        entry.grid(row=index, column=1, padx=(8, 12), pady=3)
        entries.append(entry)

    def submit():
        entered_cents = [amount_to_cents(entry.get()) for entry in entries]
        if any(value is None or value < 0 for value in entered_cents):
            messagebox.showerror(
                "Invalid amount",
                "Enter a valid non-negative dollar amount for every line item.",
                parent=dialog,
            )
            return
        if sum(entered_cents) != target_cents:
            messagebox.showerror(
                "Amounts do not match",
                (
                    f"The line items add up to {format_cents(sum(entered_cents))}, "
                    f"but the email total is {format_cents(target_cents)}."
                ),
                parent=dialog,
            )
            return
        result["values"] = entered_cents
        dialog.destroy()

    def cancel():
        dialog.destroy()

    button_row = len(line_items) + 1
    tk.Button(dialog, text="Cancel", width=10, command=cancel).grid(
        row=button_row, column=0, padx=12, pady=12, sticky="e"
    )
    tk.Button(dialog, text="Use amounts", width=12, command=submit).grid(
        row=button_row, column=1, padx=12, pady=12, sticky="e"
    )
    dialog.protocol("WM_DELETE_WINDOW", cancel)
    entries[0].focus_set()
    dialog.grab_set()
    root.wait_window(dialog)
    root.destroy()

    if result["values"] is None:
        raise RuntimeError("Matrix line-item review was cancelled; no amounts were changed.")
    return result["values"]


def reconcile_matrix_line_item_amounts(line_items, email_total, prompt_callback=None):
    """Preserve matching multi-line totals; otherwise request exact line amounts."""
    if len(line_items) <= 1:
        return [item.get("amount", "") for item in line_items]

    target_cents = amount_to_cents(email_total)
    original_cents = [amount_to_cents(item.get("amount")) for item in line_items]
    if target_cents is None or any(value is None for value in original_cents):
        raise ValueError("Could not compare the Matrix line items with the email total.")

    if sum(original_cents) == target_cents:
        return [format_cents(value) for value in original_cents]

    prompt = prompt_callback or prompt_matrix_line_item_amounts
    entered_values = prompt(line_items, email_total)
    entered_cents = [
        value if isinstance(value, int) else amount_to_cents(value)
        for value in entered_values
    ]
    if len(entered_cents) != len(line_items):
        raise ValueError("A corrected amount is required for every Matrix line item.")
    if any(value is None or value < 0 for value in entered_cents):
        raise ValueError("Every corrected Matrix line-item amount must be valid.")
    if sum(entered_cents) != target_cents:
        raise ValueError("Corrected Matrix line items must add up to the email total.")
    return [format_cents(value) for value in entered_cents]


def normalize_amount_override_value(value):
    cleaned = str(value or "").replace("$", "").replace(",", "").strip()
    if not cleaned:
        return ""

    try:
        return format_dollar_amount(float(cleaned))
    except ValueError:
        return str(value or "").strip()


def normalize_override_match_text(value):
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def find_amount_override(amount_overrides, page_num, market_text=""):
    if not amount_overrides:
        return ""

    page_key = str(page_num)
    page_overrides = amount_overrides.get(page_num) or amount_overrides.get(page_key) or []
    searchable_market = normalize_override_match_text(market_text)

    for override in page_overrides:
        if not isinstance(override, dict):
            continue

        override_amount = normalize_amount_override_value(override.get("amount"))
        if not override_amount:
            continue

        override_market = normalize_override_match_text(override.get("market"))
        override_description = normalize_override_match_text(override.get("description"))

        if not override_market and not override_description:
            return override_amount

        if override_market and override_market in searchable_market:
            return override_amount
        if override_description and override_description in searchable_market:
            return override_amount
        if searchable_market and override_market and searchable_market in override_market:
            return override_amount

    return ""


def calculate_updated_amount(original_amount, market_text=""):
    parsed_value = parse_dollar_amount(original_amount)
    normalized_market = str(market_text or "").upper()

    if "ONEONTA" in normalized_market:
        multiplied_value = parsed_value * 1.3177
    elif "PENSACOLA" in normalized_market:
        multiplied_value = parsed_value * PENSACOLA_MARGIN_MULTIPLIER
    else:
        multiplied_value = parsed_value / 0.85

    if multiplied_value != int(multiplied_value):
        multiplied_value = int(multiplied_value)

    return format_dollar_amount(multiplied_value)


def get_page_market_text(page_market_mapping, page_num):
    if not page_market_mapping:
        return ""

    page_data = page_market_mapping.get(page_num)
    if not page_data:
        return ""

    if isinstance(page_data, tuple):
        return " ".join(str(part or "") for part in page_data)

    return str(page_data or "")


def analyze_word_document(file_path, page_market_mapping=None, amount_overrides=None):
    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Change to True for debugging

    # Regex pattern to match dollar amounts, e.g., $999.00 up to $99,999.00
    dollar_amount_pattern = re.compile(r"\$(\d{1,3}(?:,\d{3})*\.\d{2})")

    # Open the document
    doc = word.Documents.Open(file_path)
    page_to_market = {}
    normalize_page_service_periods(page_market_mapping)


    try:
        # 1. Build a mapping of page_number -> table object
        page_tables = {}
        for table in doc.Tables:
            page_num = table.Range.Information(wdActiveEndPageNumber)
            page_tables[page_num] = table

        # 2. Build a mapping of page_number -> list of shapes
        page_shapes = {}
        for shape in doc.Shapes:
            if not shape.Anchor:
                continue
            page_num = shape.Anchor.Information(wdActiveEndPageNumber)
            if page_num not in page_shapes:
                page_shapes[page_num] = []
            page_shapes[page_num].append(shape)

        # 3. Process each page: update the table and the text boxes
        for page_num, table in page_tables.items():
            # Find the relevant column indices in the header row.
            amount_col_index = None
            service_period_col_index = None
            market_col_index = None
            description_col_index = None
            num_cols = table.Columns.Count

            for col_idx in range(1, num_cols + 1):
                header_text = table.Cell(1, col_idx).Range.Text.strip()
                if "Amount" in header_text:
                    amount_col_index = col_idx
                elif "Service Period" in header_text:
                    service_period_col_index = col_idx
                elif "Market" in header_text:
                    market_col_index = col_idx
                elif "Description" in header_text:
                    description_col_index = col_idx

            # If no "Amount" column found, skip this table
            if amount_col_index is None:
                continue

            # A Matrix page can contain multiple component charges for one email
            # total (for example, production plus shipping). Never replace every
            # component row with the full email total. Preserve the source rows
            # when their sum already matches; otherwise ask the user for the
            # corrected allocation before changing the document.
            page_market_text = get_page_market_text(page_market_mapping, page_num)
            line_items = []
            for row_idx in range(2, table.Rows.Count + 1):
                amount_text = (
                    table.Cell(row_idx, amount_col_index)
                    .Range.Text.replace("\r", "")
                    .replace("\a", "")
                    .strip()
                )
                row_matches = list(dollar_amount_pattern.finditer(amount_text))
                if market_col_index is not None:
                    row_market = table.Cell(row_idx, market_col_index).Range.Text.strip()
                    page_market_text = f"{page_market_text} {row_market}".strip()
                else:
                    row_market = ""
                if description_col_index is not None:
                    row_description = (
                        table.Cell(row_idx, description_col_index)
                        .Range.Text.replace("\r", "")
                        .replace("\a", "")
                        .strip()
                    )
                else:
                    row_description = ""
                for match_index, match in enumerate(row_matches):
                    line_items.append(
                        {
                            "row_idx": row_idx,
                            "match_index": match_index,
                            "amount": match.group(0),
                            "label": row_description or row_market or f"Row {row_idx}",
                        }
                    )

            page_email_total = find_amount_override(
                amount_overrides,
                page_num,
                page_market_text,
            )
            line_item_overrides = {}
            if page_email_total and len(line_items) > 1:
                reconciled_amounts = reconcile_matrix_line_item_amounts(
                    line_items,
                    page_email_total,
                )
                line_item_overrides = {
                    (item["row_idx"], item["match_index"]): amount
                    for item, amount in zip(line_items, reconciled_amounts)
                }

            # Update amounts in the "Amount" column for each data row
            num_rows = table.Rows.Count
            for row_idx in range(2, num_rows + 1):  # Start from second row
                if service_period_col_index is not None:
                    service_cell = table.Cell(row_idx, service_period_col_index)
                    service_range = service_cell.Range
                    original_period = (
                        service_range.Text.replace("\r", "").replace("\a", "").strip()
                    )
                    corrected_period = first_matrix_service_period(original_period)
                    if corrected_period != original_period:
                        # Assign the cell text directly. A PDF-to-Word conversion
                        # may separate repeated dates with hidden paragraph marks,
                        # so searching for the cleaned/crammed string is unreliable.
                        service_cell.Range.Text = corrected_period
                        print(
                            f"Corrected Matrix service period '{original_period}' "
                            f"to '{corrected_period}'"
                        )

                cell = table.Cell(row_idx, amount_col_index)
                cell_range = cell.Range
                cell_text = cell_range.Text.replace("\r", "").replace("\a", "").strip()

                print(f"Page {page_num}, Row {row_idx}, Column {amount_col_index}: '{cell_text}'")

                matches = list(dollar_amount_pattern.finditer(cell_text))
                if not matches:
                    continue

                for match_index, match in enumerate(matches):
                    original_amount = match.group(0)
                    # Get market name from current row
                    market_text = ""
                    if market_col_index:
                        market_text = table.Cell(row_idx, market_col_index).Range.Text.strip()
                    page_market_text = get_page_market_text(page_market_mapping, page_num)
                    market_text = f"{market_text} {page_market_text}".strip()
                    updated_amount = line_item_overrides.get((row_idx, match_index))
                    if not updated_amount:
                        updated_amount = find_amount_override(
                            amount_overrides,
                            page_num,
                            market_text,
                        ) or calculate_updated_amount(original_amount, market_text)

                    # Use Word's Find/Replace with wildcard matching
                    find = cell_range.Find
                    find.ClearFormatting()
                    find.Replacement.ClearFormatting()
                    
                    find.Text = original_amount
                    find.Replacement.Text = updated_amount
                    find.Forward = True
                    find.Wrap = wdFindStop
                    find.MatchCase = True  
                    find.MatchWholeWord = False
                    find.MatchWildcards = False  

                    # **Execute replacement and reset range after each match**
                    result = find.Execute(
                        FindText=original_amount,
                        MatchCase=True, 
                        MatchWholeWord=False,
                        MatchWildcards=False,
                        MatchSoundsLike=False,
                        MatchAllWordForms=False,
                        Forward=True,
                        Wrap=wdFindStop,
                        Format=False,
                        ReplaceWith=updated_amount,
                        Replace=wdReplaceOne
                    )

                    print(f"Page {page_num}, Row {row_idx}, Find/Replace result: {result}")
                    print(f"Attempted to replace '{original_amount}' with '{updated_amount}'")
                    
                    # Debug print to verify formatted amount has commas where needed
                    print(f"Formatted amount: {updated_amount}, Has comma: {',' in updated_amount}")

                    if result:  
                        # Reset the range to continue searching in the same cell
                        cell_range.Collapse(wdCollapseEnd)

                print(f"Cell text after: '{cell.Range.Text.strip()}'")

            # Optionally auto-fit the table to tidy up columns
            table.AutoFitBehavior(2)  # wdAutoFitContent = 2

            # Update text boxes (Shapes) on the same page, if any
            if page_num in page_shapes:
                for shape in page_shapes[page_num]:
                    if shape.TextFrame.HasText:
                        text_range = shape.TextFrame.TextRange
                        shape_text = text_range.Text.replace("\r", "").replace("\a", "").strip()

                        matches = list(dollar_amount_pattern.finditer(shape_text))
                        if not matches:
                            continue

                        for match in matches:
                            original_amount = match.group(0)

                            # Check this page's table market rows for market-specific pricing.
                            market_col_index = None
                            page_market_text = get_page_market_text(page_market_mapping, page_num)
                            table = page_tables.get(page_num)
                            if table:
                                for col_idx in range(1, table.Columns.Count + 1):
                                    header_text = table.Cell(1, col_idx).Range.Text.strip()
                                    if "Market" in header_text:
                                        market_col_index = col_idx
                                        break
                                
                                if market_col_index:
                                    for row_idx in range(2, table.Rows.Count + 1):
                                        market_text = table.Cell(row_idx, market_col_index).Range.Text.strip()
                                        page_market_text = f"{page_market_text} {market_text}".strip()

                            updated_amount = find_amount_override(
                                amount_overrides,
                                page_num,
                                page_market_text,
                            ) or calculate_updated_amount(original_amount, page_market_text)

                            # Find and replace in shape text
                            find = text_range.Find
                            find.ClearFormatting()
                            find.Replacement.ClearFormatting()

                            # Execute with proper parameters
                            result = find.Execute(
                                FindText=original_amount,
                                MatchCase=True, 
                                MatchWholeWord=False,
                                MatchWildcards=False,  
                                MatchSoundsLike=False,
                                MatchAllWordForms=False,
                                Forward=True,
                                Wrap=wdFindStop,
                                Format=False,
                                ReplaceWith=updated_amount,
                                Replace=wdReplaceOne
                            )

                            print(f"Page {page_num}, Shape Text, Find/Replace result: {result}")
                            print(f"Attempted to replace '{original_amount}' with '{updated_amount}'")
                            
                            # Debug print to verify formatted amount has commas where needed
                            print(f"Formatted amount: {updated_amount}, Has comma: {',' in updated_amount}")

                            if result:  
                                # Reset the range to continue searching
                                text_range.Collapse(wdCollapseEnd)

        print("Amounts updated successfully while preserving formatting.")

    finally:
        # Save and close document
        doc.Save()
        doc.Close(True)
        word.Quit()

        




   

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python matrix_media_logic.py <path_to_docx>")
        sys.exit(1)
    file_path = sys.argv[1]
    analyze_word_document(file_path)







