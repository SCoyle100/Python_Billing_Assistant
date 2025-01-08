import os
import base64
import json
import re
from PIL import Image
import pandas as pd
from openai import OpenAI

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def analyze_image_with_openai(image_path):
    client = OpenAI()
    base64_image = encode_image(image_path)

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Please provide the following columns from the table: invoice # and net invoice amount",
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}",
                                "detail": "high",
                            },
                        },
                    ],
                }
            ],
            max_tokens=1000,
        )

        result = response.choices[0].message.content
        print("Chat Completions Output:\n", result)  # Print output to the terminal
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
    input_file = r"cropped_image_final.png"

    if not os.path.exists(input_file):
        print(f"Input file does not exist: {input_file}")
    else:
        analysis_result = analyze_image_with_openai(input_file)

        if analysis_result:
            # Attempt JSON parsing first; if it fails, fall back to plaintext parsing.
            try:
                # Try parsing as JSON
                data = json.loads(analysis_result)
                df = pd.DataFrame(data, columns=["invoice #", "net invoice amount"])
            except json.JSONDecodeError:
                # If JSON parsing fails, parse plaintext manually
                df = parse_plaintext_to_dataframe(analysis_result)

            if df is not None and not df.empty:
                # Convert 'invoice #' column to numeric values for proper numeric sorting.
                df["invoice #"] = pd.to_numeric(df["invoice #"], errors='coerce')
                # Sort the DataFrame by the 'invoice #' column.
                df = df.sort_values(by="invoice #")

                # Set the invoice # column as the index for reindexing.
                df.set_index("invoice #", inplace=True)

                # Create a complete range of invoice numbers from the minimum to maximum.
                full_range = range(int(df.index.min()), int(df.index.max()) + 1)

                # Reindex the DataFrame to include all invoice numbers in the range.
                df = df.reindex(full_range)

                # Reset the index to bring 'invoice #' back as a column.
                df.reset_index(inplace=True)
                df.rename(columns={"index": "invoice #"}, inplace=True)

                # Replace NaN values in 'net invoice amount' with blank strings.
                df["net invoice amount"] = df["net invoice amount"].fillna("")

                print(df)
                df.to_csv("output.csv", index=False)
                print("Saved analysis to output.csv")
            else:
                print("No data was parsed into a DataFrame.")

