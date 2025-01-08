import os
import base64
import json
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

def create_dataframe_from_analysis(analysis_result):
    try:
        data = json.loads(analysis_result)  # Use JSON parsing for safety
        df = pd.DataFrame(data, columns=["invoice #", "net invoice amount"])
        return df
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        print(f"Raw analysis result: {analysis_result}")
    except Exception as e:
        print(f"Error creating DataFrame: {e}")
        print(f"Raw analysis result: {analysis_result}")
    return None

if __name__ == "__main__":
    input_file = r"cropped_image_final.png"

    if not os.path.exists(input_file):
        print(f"Input file does not exist: {input_file}")
    else:
        analysis_result = analyze_image_with_openai(input_file)

        if analysis_result:
            df = create_dataframe_from_analysis(analysis_result)
            if df is not None:
                print(df)
                df.to_csv("output.csv", index=False)
                print("Saved analysis to output.csv")


