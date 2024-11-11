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

    page_text = []
    for page in pdf.pages:
        text = page.extract_text()
        page_text.append(text)
        # Process images and non-text elements, as needed

    pages_content = []
    for page_num, page_content in enumerate(pdf.pages):
        pages_content.append({
            "slide_number": page_num + 1,
            "text": page_text[page_num],
            # "images": images_list
        })

    return pages_content


presentation = Presentation("N_28_Explitsitnaya_pamyat.pptx")
pdf = pdfplumber.open("Class06-GradientDescent-New.pdf")

# a = pptx_converter(presentation)
a = pdf_converter(pdf)

with open("report.txt", "w",encoding="utf-8") as my_file:
    for each in a:
        my_file.write(str(each) + '\n')