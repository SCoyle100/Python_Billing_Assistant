
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
            max_tokens=1000,
        )

        result = response.choices[0].message.content
        print("Chat Completions Output:\n", result)
        return result
    except Exception as e:
        print(f"Error analyzing image with OpenAI Vision API: {e}")
        return None

def parse_plaintext_to_dataframe(text):
    # Use regex to find lines with pattern: "- <invoice>: $<amount>"
    pattern = r"-\s*(\d+):\s*\$(.*)"
    entries = re.findall(pattern, text)
    
    # Build list of dictionaries from entries
    data = []
    for invoice, amount in entries:
        data.append({"invoice #": int(invoice), "net invoice amount": amount.strip()})
    
    # Create DataFrame
    df = pd.DataFrame(data, columns=["invoice #", "net invoice amount"])
    return df



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
