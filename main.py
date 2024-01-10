from pptx import Presentation
import pyttsx3
import threading
import time

def extract_text_from_ppt(ppt_file):
    presentation = Presentation(ppt_file)
    extracted_text = ""
    for slide_number, slide in enumerate(presentation.slides):
        extracted_text += f"\nSlide {slide_number + 1}:\n"
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                extracted_text += f"{shape.text}\n"
    return extracted_text

def speak(text_to_speak, pause_flag):
    bot = pyttsx3.init()
    for text in text_to_speak.split('\n'):
        if pause_flag.is_set():
            while pause_flag.is_set():
                time.sleep(1)
        bot.say(text)
        bot.runAndWait()

def main():
    ppt_file_path = 'test.pptx'
    extracted_text_result = extract_text_from_ppt(ppt_file_path)
    print(extracted_text_result)

    # Create a threading.Event() to signal pause/play
    pause_flag = threading.Event()

    # Create threads
    extraction_thread = threading.Thread(target=lambda: print(extracted_text_result))
    speak_thread = threading.Thread(target=lambda: speak(extracted_text_result, pause_flag))

    # Start threads
    extraction_thread.start()
    speak_thread.start()

    while True:
        user_input = input("Enter 'pause' to pause TTS, 'resume' to resume, or 'exit' to quit: ")
        if user_input.lower() == 'pause':
            pause_flag.set()
            print("TTS paused")
        elif user_input.lower() == 'resume':
            pause_flag.clear()
            print("TTS resumed")
        elif user_input.lower() == 'exit':
            break
        else:
            print("Invalid input. Please enter 'pause', 'resume', or 'exit'.")

    # Wait for threads to finish
    extraction_thread.join()
    speak_thread.join()

if __name__ == "__main__":
    main()
