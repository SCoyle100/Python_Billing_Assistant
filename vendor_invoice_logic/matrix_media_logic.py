try:
    import win32com.client
except ImportError:
    win32com = None
import datetime
import itertools
import re
import sys

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


def normalize_page_service_periods(page_market_mapping):
    """Keep the pre-read page mapping aligned with corrections saved to Word."""
    if not page_market_mapping:
        return

    for page_num, page_data in list(page_market_mapping.items()):
        if isinstance(page_data, tuple) and len(page_data) >= 2:
            corrected = correct_matrix_service_period(page_data[1])
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
            num_cols = table.Columns.Count

            for col_idx in range(1, num_cols + 1):
                header_text = table.Cell(1, col_idx).Range.Text.strip()
                if "Amount" in header_text:
                    amount_col_index = col_idx
                elif "Service Period" in header_text:
                    service_period_col_index = col_idx

            # If no "Amount" column found, skip this table
            if amount_col_index is None:
                continue

            # Update amounts in the "Amount" column for each data row
            num_rows = table.Rows.Count
            for row_idx in range(2, num_rows + 1):  # Start from second row
                if service_period_col_index is not None:
                    service_cell = table.Cell(row_idx, service_period_col_index)
                    service_range = service_cell.Range
                    original_period = (
                        service_range.Text.replace("\r", "").replace("\a", "").strip()
                    )
                    corrected_period = correct_matrix_service_period(original_period)
                    if corrected_period != original_period:
                        service_find = service_range.Find
                        service_find.ClearFormatting()
                        service_find.Replacement.ClearFormatting()
                        service_find.Execute(
                            FindText=original_period,
                            MatchCase=True,
                            MatchWholeWord=False,
                            MatchWildcards=False,
                            MatchSoundsLike=False,
                            MatchAllWordForms=False,
                            Forward=True,
                            Wrap=wdFindStop,
                            Format=False,
                            ReplaceWith=corrected_period,
                            Replace=wdReplaceOne,
                        )
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

                for match in matches:
                    original_amount = match.group(0)
                    # Get market name from current row
                    market_cell_index = None
                    market_text = ""
                    for col_idx in range(1, num_cols + 1):
                        header_text = table.Cell(1, col_idx).Range.Text.strip()
                        if "Market" in header_text:
                            market_cell_index = col_idx
                            break

                    if market_cell_index:
                        market_text = table.Cell(row_idx, market_cell_index).Range.Text.strip()
                    page_market_text = get_page_market_text(page_market_mapping, page_num)
                    market_text = f"{market_text} {page_market_text}".strip()
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







