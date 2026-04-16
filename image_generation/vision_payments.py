
import os
import base64
import json
import re
from PIL import Image
import pandas as pd
from openai import OpenAI
from invoice_processor import process_invoice  # Import the invoice processing function

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def analyze_image_with_openai(image_path):
    client = OpenAI()
    base64_image = encode_image(image_path)

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "Please provide the following columns from the table: invoice # and net invoice amount. Please ensure it is formatted like this example:- 112401: $1,300.00",
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{base64_image}",
                            "detail": "high",
                        },
                    },
                ],
            }],
            temperature=0,
            max_tokens=1000,
        )

        result = response.choices[0].message.content
        print("Chat Completions Output:\n", result)
        return result
    except Exception as e:
        print(f"Error analyzing image with OpenAI Vision API: {e}")
        return None

def parse_plaintext_to_dataframe(text):
    # Match invoice IDs like "112401" and "112926-P".
    pattern = r"^\s*(?:[-*]\s*)?(\d{6}(?:[-_/ \t]*[A-Za-z]+(?:_[A-Za-z]+)?)?):\s*\$?\s*([^\r\n]+)"
    entries = re.findall(pattern, text, flags=re.MULTILINE)
    
    # Build list of dictionaries from entries
    data = []
    for invoice, amount in entries:
        invoice = re.sub(r"\s+", "-", invoice.strip())
        data.append({"invoice #": invoice.strip(), "net invoice amount": amount.strip()})
    
    # Create DataFrame
    df = pd.DataFrame(data, columns=["invoice #", "net invoice amount"])
    return df

def sort_invoices(df):
    df = df.copy()
    df["invoice #"] = df["invoice #"].astype(str).str.strip()
    sort_parts = df["invoice #"].str.extract(r"^(\d+)(.*)$")
    df["_invoice_number"] = pd.to_numeric(sort_parts[0], errors="coerce")
    df["_invoice_suffix"] = sort_parts[1].fillna("")
    df = df.sort_values(
        by=["_invoice_number", "_invoice_suffix", "invoice #"],
        na_position="last",
    )
    return df.drop(columns=["_invoice_number", "_invoice_suffix"])



if __name__ == "__main__":
    # Run the invoice processing to create the cropped image
    cropped_image_path = process_invoice()

    # Verify that the cropped image was created
    if not os.path.exists(cropped_image_path):
        print(f"Input file does not exist: {cropped_image_path}")
    else:
        analysis_result = analyze_image_with_openai(cropped_image_path)

        if analysis_result:
            try:
                data = json.loads(analysis_result)
                df = pd.DataFrame(data, columns=["invoice #", "net invoice amount"])
            except json.JSONDecodeError:
                df = parse_plaintext_to_dataframe(analysis_result)

            if df is not None and not df.empty:
                df["invoice #"] = df["invoice #"].astype(str).str.strip()
                has_suffix = df["invoice #"].str.contains(r"\D", regex=True).any()

                if has_suffix:
                    df = sort_invoices(df)
                else:
                    df["invoice #"] = pd.to_numeric(df["invoice #"], errors='coerce')
                    df = df.sort_values(by="invoice #")
                    df.set_index("invoice #", inplace=True)
                    full_range = range(int(df.index.min()), int(df.index.max()) + 1)
                    df = df.reindex(full_range)
                    df.reset_index(inplace=True)
                    df.rename(columns={"index": "invoice #"}, inplace=True)

                df["net invoice amount"] = df["net invoice amount"].fillna("")

                print(df)
                df.to_csv("output.csv", index=False)
                print("Saved analysis to output.csv")
            else:
                print("No data was parsed into a DataFrame.")
