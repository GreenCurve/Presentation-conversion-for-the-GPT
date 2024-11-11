from pptx import Presentation
import pdfplumber

def pptx_converter(presentation):

    slide_text = []

    for slide in presentation.slides:
        text = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                # Store or process the text
                text.append(shape.text_frame.text)
            # elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            #     # Extract image or process it
            #     pass
        slide_text.append(text)

    slides_content = []
    for slide_num, slide_content in enumerate(presentation.slides):
        # print(slide_content)
        slides_content.append({
            "slide_number": slide_num + 1,
            "text": slide_text[slide_num],
            # "images": images_list
        })

    return slides_content

def pdf_converter(pdf):

    num_slides = len(pdf.pages)

    for page in pdf.pages:
        text = page.extract_text()
        # Process images and non-text elements, as needed

        slides_content = []
    for slide_num, slide_content in enumerate(slides):
        slides_content.append({
            "slide_number": slide_num + 1,
            "text": slide_text,
            "images": images_list
        })


presentation = Presentation("N_28_Explitsitnaya_pamyat.pptx")
pdf = pdfplumber.open("Class06-GradientDescent-New.pdf")

a = pptx_converter(presentation)

for each in a:
    print(each)