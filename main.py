from pptx import Presentation
import pymupdf

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



def pdf_converter(pdf_file):
    '''
    Splice pdf file into list with entries corrseponding to the page number, all the text , paths to all images. 
    Creates all the images in the Logs\
    '''

    pages_text = [] #will contain all the text
    pages_images = [] #will contain all the image paths

    for page_num in range(len(pdf_file)): #iterate over every page
        page = pdf_file[page_num]

        text = page.get_text("text") #get all text from page
        pages_text.append(text)#write all text


        image_list = page.get_images(full=True)  #get all images from page
        for img_index, img in enumerate(image_list):#iterate over all images on the page
            xref = img[0]
            base_image = pdf_file.extract_image(xref)
            image_bytes = base_image["image"]

            # Save the image to a file
            image_filename = f"page_{page_num + 1}_image_{img_index + 1}.png"
            with open("Logs/" + image_filename, "wb") as img_file:#create image in the Logs/
                img_file.write(image_bytes)

            pages_images.append(image_filename)#write image


    pages_content = []
    for page_num in range(len(pdf_file)):
        pages_content.append({
            "slide_number": page_num + 1,
            "text": pages_text[page_num],
             "images": pages_images[page_num]
        })

    return pages_content


presentation = Presentation("N_28_Explitsitnaya_pamyat.pptx")
pdf = pymupdf.open("Class06-GradientDescent-New.pdf")

# a = pptx_converter(presentation)
a = pdf_converter(pdf)

with open("Logs/report.txt", "w",encoding="utf-8") as my_file:
    for each in a:
        my_file.write(str(each) + '\n')