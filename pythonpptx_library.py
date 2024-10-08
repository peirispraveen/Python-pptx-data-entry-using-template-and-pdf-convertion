import pptx  # pip install python-pptx
from pptx.util import Inches
from pptxtopdf import convert  # pip install pptxtopdf
import requests
from io import BytesIO
import os
import json
from html2image import Html2Image


def download_image(image_url):
    """
    Downloads an image from a URL and returns it as a BytesIO object.

    Args:
        image_url: URL of the image to download.

    Returns:
        BytesIO: In-memory file-like object containing the image.
    """
    response = requests.get(image_url)
    if response.status_code == 200:
        print(f"Image downloaded from {image_url}")
        return BytesIO(response.content)
    else:
        raise Exception(f"Failed to download image from {image_url}")


def input_data_and_save_pdf(template_path, output_path, data, image_data):
    """
    Inputs data into placeholders in a PowerPoint template and saves it as a PDF.
    Also, inserts images into specified shapes with given text.

    Args:
        template_path: Path to the PowerPoint template file.
        output_path: Path to save the output PowerPoint file.
        data: Dictionary containing text data to be inserted into placeholders.
        image_data: Dictionary containing image paths or URLs where the key is the shape text and
                    the value is either an image file path or URL.
    """

    # Open the template presentation
    prs = pptx.Presentation(template_path)

    # Iterate through slides and find placeholders
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        placeholder_text = run.text
                        if placeholder_text in data:
                            run.text = data[placeholder_text]
                        # Check for the specific text "image_1" and replace it with an image
                        elif placeholder_text in image_data:
                            print(f"Found placeholder for image: {placeholder_text}")
                            image_source = image_data[placeholder_text]

                            # Determine if image is from a URL or a local path
                            if image_source.startswith("http"):
                                # Download the image from the URL
                                image_stream = download_image(image_source)
                            else:
                                # Use a local image path
                                with open(image_source, "rb") as img_file:
                                    image_stream = BytesIO(img_file.read())

                            # Clear the text from the shape
                            run.text = ''
                            # Insert the image using the placeholder shape dimensions
                            left = shape.left
                            top = shape.top
                            width = shape.width
                            height = shape.height
                            print(f"Inserting image: {image_source} at {left}, {top}")
                            # Add the image with the same dimensions as the shape
                            slide.shapes.add_picture(image_stream, left, top, width, height)


    prs.save(output_path)


template_path = "input_files/File Format for Control Union.pptx"
output_path = "output_files/CU_report_from_pptx.pptx"

# Text placeholders data
data = json.load(open("input_files/input_data.json"))

# Image placeholder data
image_data = json.load(open("input_files/input_img.json"))

# Directory to save the screenshots
output_dir = 'output_files'

# Create the output directory if it doesn't exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

html_placeholders = {}

hti = Html2Image()

# Directory to save the screenshots
output_dir = 'output_images'

# Create the output directory if it doesn't exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Set the output directory for Html2Image
hti.output_path = output_dir

# Dictionary to store HTML file placeholders
html_placeholders = {}

# Iterate through the image_data dictionary
for placeholder, image_path in image_data.items():
    # Check if the image_path ends with ".html"
    if image_path.endswith('.html'):
        # Convert HTML file to an image and save it
        output_image_filename = f'{placeholder}.png'
        hti.screenshot(html_file=image_path, save_as=output_image_filename)

        # Add the new image path (within output directory) to the html_placeholders dictionary
        html_placeholders[placeholder] = os.path.normpath(os.path.join(output_dir, output_image_filename)).replace("\\",
                                                                                                                   "/")

# Replace the HTML paths in the original image_data with the generated image paths
for placeholder, new_image_path in html_placeholders.items():
    image_data[placeholder] = new_image_path

# Output the updated image_data
print("Updated image_data:", image_data)

input_data_and_save_pdf(template_path, output_path, data, image_data)

# Convert PowerPoint to PDF
input_dir = output_path
output_dir = r"output_files"

convert(input_dir, output_dir)
