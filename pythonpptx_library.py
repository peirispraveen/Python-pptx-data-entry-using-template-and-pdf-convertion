import pptx  # pip install python-pptx
from pptx.util import Inches
from pptxtopdf import convert  # pip install pptxtopdf

def input_data_and_save_pdf(template_path, output_path, data):
    """
    Inputs data into placeholders in a PowerPoint template and saves it as a PDF.

    Args:
        template_path: Path to the PowerPoint template file.
        output_path: Path to save the output PDF file.
        data: Dictionary containing data to be inserted into placeholders.
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

    # Save the presentation as a PDF
    prs.save(output_path)


# Example usage
template_path = "input_files/File Format for Control Union.pptx"
output_path = "output_files/CU_report_from_pptx.pptx"
data = {
    "deforestation_text": "Deforestation risk is low",
    "encroachment_text1": "Encroachment risk is low ",
    "deforestation_text2": "Deforestation risk is low",
    "encroachment_text2": "Encroachment risk is low ",
    "deforestation_text3": "Deforestation risk is low",
    "encroachment_text3": "Encroachment risk is low ",
    "deforestation_text4": "Deforestation risk is low",
    "encroachment_text4": "Encroachment risk is low ",
    "total_area_val": "0.12 ha",
    "potec_val": "0.00 ha",
    "def_val": "0.00ha",
    "eligible_area_val": "0.12ha",
    "tec_val": "20M"
}

input_data_and_save_pdf(template_path, output_path, data)


input_dir = output_path
output_dir = r"output_files"

convert(input_dir, output_dir)
