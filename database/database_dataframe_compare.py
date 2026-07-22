import logging
import sqlite3
from datetime import datetime

from vendor_invoice_logic import matrix_media_dataframe
from utils.openai_json import chat_completion_json


today_date = datetime.now().strftime("%Y%m%d")
file_path = r"D:\Programming\Billing_PDF_Automation\output\Matrix Media Services Invoice.docx"


def compare_dataframe_to_database(dataframe_records, database_records):
    return chat_completion_json(
        system_prompt=(
            "Compare transformed invoice dataframe records to database invoice records. "
            "Return a JSON object with keys: matches and discrepancies. "
            "Both values must be arrays of objects. "
            "For close matches, prefer the database amount and database market naming."
        ),
        user_prompt=(
            f"DataFrame records:\n{dataframe_records}\n\n"
            f"Database records:\n{database_records}"
        ),
        max_completion_tokens=2500,
    )


def compare_dataframes_for_margin(df_transformed, df_original):
    return chat_completion_json(
        system_prompt=(
            "Compare transformed invoice records against original invoice records. "
            "Return a JSON object with one key, unchanged_amounts, containing records where the amount "
            "did not change after margin was expected to be applied. Ignore the Dothan $1,003.00 exception."
        ),
        user_prompt=(
            f"Transformed records:\n{df_transformed}\n\n"
            f"Original records:\n{df_original}"
        ),
        max_completion_tokens=2000,
    )


if __name__ == "__main__":
    df_transformed, df_original = matrix_media_dataframe.build_dataframe_from_word_document(file_path)

    conn = sqlite3.connect("invoices.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM invoices WHERE batch_id LIKE ?", (f"{today_date}%",))
    columns = [desc[0] for desc in cursor.description]
    database_records = [dict(zip(columns, row)) for row in cursor.fetchall()]
    conn.close()

    print(f"Filtered database records for batch ID {today_date}: {database_records}")

    try:
        response_db = compare_dataframe_to_database(
            df_transformed.to_dict(orient="records"),
            database_records,
        )
        print("OpenAI Database Comparison Response:", response_db)
        matches = response_db.get("matches", [])
        discrepancies_db = response_db.get("discrepancies", [])
        logging.info(
            "Found %s matches and %s discrepancies in DB comparison.",
            len(matches),
            len(discrepancies_db),
        )
    except Exception as exc:
        logging.error("Error during OpenAI database comparison: %s", exc)

    try:
        response_margin = compare_dataframes_for_margin(
            df_transformed.to_dict(orient="records"),
            df_original.to_dict(orient="records"),
        )
        print("OpenAI Margin Comparison Response:", response_margin)
        unchanged_amounts = response_margin.get("unchanged_amounts", [])
        if unchanged_amounts:
            logging.warning(
                "Found %s records where margin was not applied:",
                len(unchanged_amounts),
            )
            for record in unchanged_amounts:
                print(record)
        else:
            logging.info("All amounts have margin applied correctly.")
    except Exception as exc:
        logging.error("Error during OpenAI margin comparison: %s", exc)
