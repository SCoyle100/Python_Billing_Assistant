


import cv2
import numpy as np
import os
from PIL import Image

file_path = r"C:\Users\seanc\Downloads\Remittance_Advice (2).tif"

# Check if file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File not found: {file_path}")

# Attempt to read the image
if file_path.lower().endswith(('.tif', '.tiff')):
    # Use Pillow for TIFF files
    image = Image.open(file_path)
    image = np.array(image)  # Convert to numpy array
else:
    # Use OpenCV for other formats
    image = cv2.imread(file_path, cv2.IMREAD_UNCHANGED)

# Verify the image was loaded
if image is None:
    raise ValueError("Failed to load the image. Please check the file format or path.")

# Check the data type and convert to uint8 if necessary
if image.dtype == bool:
    image = image.astype(np.uint8) * 255  # Convert boolean to 0/255 uint8

# Check the number of channels
if len(image.shape) == 2:
    # Already grayscale
    gray = image
else:
    # Convert to grayscale
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# Threshold the image to separate content from background
_, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)

# Find contours
contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

# Filter contours by size (remove very small ones)
filtered_contours = [cnt for cnt in contours if cv2.contourArea(cnt) > 1000]

# Merge all filtered contours into one bounding box
x_min, y_min, x_max, y_max = float('inf'), float('inf'), 0, 0
for cnt in filtered_contours:
    x, y, w, h = cv2.boundingRect(cnt)
    x_min = min(x_min, x)
    y_min = min(y_min, y)
    x_max = max(x_max, x + w)
    y_max = max(y_max, y + h)

# Crop the image to the bounding box
cropped_image = image[y_min:y_max, x_min:x_max]

# Save the cropped image
output_image_path = "cropped_image.png"
cv2.imwrite(output_image_path, cropped_image)

print(f"Cropped image saved at {output_image_path}")




