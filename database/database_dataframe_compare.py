'''
import logging
import dspy
import sqlite3
from vendor_invoice_logic import matrix_media_dataframe

from datetime import datetime


# Get today's date in the desired format (e.g., 'YYYY-MM-DD')
today_date = datetime.now().strftime('%Y%m%d')

dspy.configure(lm=dspy.LM('openai/gpt-4o'))

file_path = r"D:\Programming\Billing_PDF_Automation\output\Matrix Media Services Invoice.docx"




class MatchDataFrameToDatabase(dspy.Signature):

    """
    Match records from a DataFrame to a database, considering market name similarity and amount proximity.
    If a match is found, the database amount overrides the DataFrame amount, and the database market overrides 
    the dataframe market.
    """

    dataframe_records: list[dict[str, str]] = dspy.InputField(
        desc="All records from the DataFrame"
    )
    database_records: list[dict[str, str]] = dspy.InputField(
        desc="All database records"
    )
    matches: list[dict[str, str]] = dspy.OutputField(
        desc="Best matches for each DataFrame record"
    )



compare_data = dspy.Predict(MatchDataFrameToDatabase)


# Example DataFrame records
dataframe_records = matrix_media_dataframe.build_dataframe_from_word_document(file_path)


# Example database query to fetch all records
conn = sqlite3.connect('invoices.db')
cursor = conn.cursor()
cursor.execute("SELECT * FROM invoices WHERE batch_id LIKE ?", (f"{today_date}%",))
columns = [desc[0] for desc in cursor.description]  # Get column names
database_records = [dict(zip(columns, row)) for row in cursor.fetchall()]
conn.close()

print(f"Filtered database records for batch ID {today_date}: {database_records}")

try:
    response = compare_data(
        dataframe_records=dataframe_records,
        database_records=database_records
    )



# Debug: Print the response structure
    print("DSPy response:", response)


    
    # Extract matches and discrepancies
    matches = response.results.get('matches', [])
    discrepancies = response.results.get('discrepancies', [])

    logging.info(f"Found {len(matches)} matches and {len(discrepancies)} discrepancies.")
except Exception as e:
    logging.error(f"Error during DSPy comparison: {e}")

    '''


import logging
import dspy
import sqlite3
from vendor_invoice_logic import matrix_media_dataframe
from datetime import datetime

# Get today's date in the desired format (e.g., 'YYYY-MM-DD')
today_date = datetime.now().strftime('%Y%m%d')

dspy.configure(lm=dspy.LM('openai/gpt-4o'))

file_path = r"D:\Programming\Billing_PDF_Automation\output\Matrix Media Services Invoice.docx"


class MatchDataFrameToDatabase(dspy.Signature):
    """
    Match records from a DataFrame to a database, considering market name similarity and amount proximity.
    If a match is found, the database amount overrides the DataFrame amount, and the database market overrides 
    the DataFrame market.
    """
    dataframe_records: list[dict[str, str]] = dspy.InputField(
        desc="All records from the DataFrame"
    )
    database_records: list[dict[str, str]] = dspy.InputField(
        desc="All database records"
    )
    matches: list[dict[str, str]] = dspy.OutputField(
        desc="Best matches for each DataFrame record"
    )


class CompareDataFramesForMargin(dspy.Signature):
    """
    Compare transformed and original DataFrames to identify records where margin
    was not applied (amounts are identical) not including:  Market: Dothan, Amount: $1,003.00
    """
    df_transformed: list[dict[str, str]] = dspy.InputField(
        desc="Transformed DataFrame records"
    )
    df_original: list[dict[str, str]] = dspy.InputField(
        desc="Original DataFrame records"
    )
    unchanged_amounts: list[dict[str, str]] = dspy.OutputField(
        desc="Records where transformed and original amounts are the same"
    )





# Predict functions
compare_data_to_db = dspy.Predict(MatchDataFrameToDatabase)
compare_dataframes = dspy.Predict(CompareDataFramesForMargin)

# Example DataFrame records
df_transformed, df_original = matrix_media_dataframe.build_dataframe_from_word_document(file_path)

# Example database query to fetch all records
conn = sqlite3.connect('invoices.db')
cursor = conn.cursor()
cursor.execute("SELECT * FROM invoices WHERE batch_id LIKE ?", (f"{today_date}%",))
columns = [desc[0] for desc in cursor.description]  # Get column names
database_records = [dict(zip(columns, row)) for row in cursor.fetchall()]
conn.close()

print(f"Filtered database records for batch ID {today_date}: {database_records}")

# Compare DataFrame to Database
try:
    response_db = compare_data_to_db(
        dataframe_records=df_transformed.to_dict(orient='records'),
        database_records=database_records
    )

    print("DSPy Database Comparison Response:", response_db)

    # Extract matches and discrepancies from DB comparison
    matches = response_db.results.get('matches', [])
    discrepancies_db = response_db.results.get('discrepancies', [])

    logging.info(f"Found {len(matches)} matches and {len(discrepancies_db)} discrepancies in DB comparison.")

except Exception as e:
    logging.error(f"Error during DSPy database comparison: {e}")

# Compare Transformed vs Original DataFrames
try:
    response_margin = compare_dataframes(
        df_transformed=df_transformed.to_dict(orient='records'),
        df_original=df_original.to_dict(orient='records')
    )

    print("DSPy Margin Comparison Response:", response_margin)

    # Extract records with unchanged amounts
    unchanged_amounts = response_margin.results.get('unchanged_amounts', [])

    if unchanged_amounts:
        logging.warning(f"Found {len(unchanged_amounts)} records where margin was not applied:")
        for record in unchanged_amounts:
            print(record)
    else:
        logging.info("All amounts have margin applied correctly.")

except Exception as e:
    logging.error(f"Error during DSPy margin comparison: {e}")