from pptx import Presentation
import pyttsx3

def extract_text_from_ppt(ppt_file):
    presentation = Presentation(ppt_file)

    extracted_text = ""

    for slide_number, slide in enumerate(presentation.slides):
        extracted_text += f"\nSlide {slide_number + 1}:\n"

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                extracted_text += f"{shape.text}\n"
    return extracted_text

ppt_file_path = 'test.pptx'
extracted_text_result = extract_text_from_ppt(ppt_file_path)
print(extracted_text_result)
bot = pyttsx3.init()
bot.say(extracted_text_result)
bot.runAndWait()