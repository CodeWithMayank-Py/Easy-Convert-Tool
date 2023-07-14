from PIL import Image
from fpdf import FPDF
from pdf2image import convert_from_path
import PyPDF2

poppler_path = r"C:\\Users\\paliw\\Downloads\\poppler-23.06.0\\poppler-23.06.0\\utils"

# Button 1: PNG to JPEG

def convert_png_to_jpeg(png_path, jpeg_path):
    try:
        # Open the PNG image
        png_image = Image.open(png_path)

        # Convert the image to JPEG format
        png_image.convert("RGB").save(jpeg_path, "JPEG")

        return True

    except Exception as e:
        print("Error converting PNG to JPEG:", str(e))
        return False


# Button 2; JPEG to PNG

def convert_jpeg_to_png(jpeg_path, png_path):
    try:
        image = Image.open(jpeg_path)
        image.save(png_path, 'PNG')
        return True
    except Exception as e:
        print(f"Error converting JPEG to PNG: {str(e)}")
        return False


# Button 3 : JPEG or PNG to PDF

def convert_image_to_pdf(image_path, pdf_path):
    try:
        # Create a new PDF object
        pdf = FPDF()

        # Add a new page to the PDF
        pdf.add_page()

        # Load the image
        image = Image.open(image_path)

        # Convert the image to RGB mode (if necessary)
        if image.mode != "RGB":
            image = image.convert("RGB")

        # Get the image size
        image_width, image_height = image.size

        # Calculate the maximum width and height for the PDF page
        max_width = pdf.w - (pdf.l_margin + pdf.r_margin)
        max_height = pdf.h - (pdf.t_margin + pdf.b_margin)

        # Calculate the scaling factor for the image
        scale = min(max_width / image_width, max_height / image_height)

        # Calculate the scaled dimensions
        scaled_width = image_width * scale
        scaled_height = image_height * scale

        # Calculate the position to center the image on the page
        x = (pdf.w - scaled_width) / 2
        y = (pdf.h - scaled_height) / 2

        # Add the image to the PDF
        pdf.image(image_path, x=x, y=y, w=scaled_width, h=scaled_height)

        # Save the PDF to the specified output path
        pdf.output(pdf_path, "F")

        return True

    except Exception as e:
        print(f"Error converting image to PDF: {str(e)}")
        return False


# Button 4 : PDF to JPEG or PNG format

def convert_pdf_to_image(pdf_path, image_path):
    try:
        # Convert the PDF to a list of PIL.Image objects
        images = convert_from_path(pdf_path)

        # Save each image as a separate file
        for i, image in enumerate(images):
            image.save(f"{image_path}_{i + 1}.png", "PNG")

        return True

    except Exception as e:
        print(f"Error converting PDF to image: {str(e)}")
        return False
