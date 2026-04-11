import argparse
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from vendor_invoice_logic.capitol_media_dataframe_1 import build_dataframe_from_capitol_media
from vendor_invoice_logic.capitol_media_rebuild import rebuild_capitol_media_table


def build_output_path(input_path, output_path=None):
    if output_path:
        return Path(output_path)
    return input_path.with_name(f"{input_path.stem}_capitol_rebuilt{input_path.suffix}")


def main():
    parser = argparse.ArgumentParser(
        description="Test the Capitol Media extraction + rebuild flow on a DOCX copy."
    )
    parser.add_argument("input_docx", help="Path to the original Capitol Media DOCX")
    parser.add_argument(
        "--output",
        help="Optional output DOCX path. Defaults to <input>_capitol_rebuilt.docx",
    )
    args = parser.parse_args()

    input_path = Path(args.input_docx).resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input DOCX not found: {input_path}")

    output_path = build_output_path(input_path, args.output).resolve()
    shutil.copy2(input_path, output_path)

    df_invoices = build_dataframe_from_capitol_media(str(output_path))
    print("Extracted dataframe:")
    print(df_invoices)

    if df_invoices.empty:
        print("No Capitol invoice rows were extracted. Output copy left unchanged:")
        print(output_path)
        return

    if "Market" not in df_invoices.columns or "Amount" not in df_invoices.columns:
        raise ValueError(
            f"Expected columns ['Market', 'Amount'], found {df_invoices.columns.tolist()}"
        )

    invoices_list = list(df_invoices[["Market", "Amount"]].itertuples(index=False, name=None))
    rebuild_capitol_media_table(str(output_path), invoices_list)

    print("\nRebuilt DOCX written to:")
    print(output_path)


if __name__ == "__main__":
    main()
