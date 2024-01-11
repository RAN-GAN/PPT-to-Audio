from pptx import Presentation
import pyttsx3
import threading
import tkinter as tk
from tkinter import ttk

class TTSController:
    def __init__(self, text_to_speak):
        self.text_to_speak = text_to_speak
        self.pause_flag = threading.Event()
        self.extraction_thread = None
        self.speak_thread = None

    def start(self):
        self.extraction_thread = threading.Thread(target=lambda: print(self.text_to_speak))
        self.speak_thread = threading.Thread(target=lambda: self._speak())

        self.extraction_thread.start()
        self.speak_thread.start()

    def _speak(self):
        bot = pyttsx3.init()
        for text in self.text_to_speak.split('\n'):
            if self.pause_flag.is_set():
                while self.pause_flag.is_set():
                    time.sleep(1)
            bot.say(text)
            bot.runAndWait()

    def handle_pause(self):
        self.pause_flag.set()
        print("TTS paused")

    def handle_resume(self):
        self.pause_flag.clear()
        print("TTS resumed")

    def stop(self):
        if self.extraction_thread:
            self.extraction_thread.join()
        if self.speak_thread:
            self.speak_thread.join()

def extract_text_from_ppt(ppt_file):
    presentation = Presentation(ppt_file)
    extracted_text = ""
    for slide_number, slide in enumerate(presentation.slides):
        extracted_text += f"\nSlide {slide_number + 1}:\n"
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                extracted_text += f"{shape.text}\n"
    return extracted_text

def on_pause_button_click():
    tts_controller.handle_pause()

def on_resume_button_click():
    tts_controller.handle_resume()

def on_close():
    tts_controller.stop()
    root.destroy()

ppt_file_path = 'test.pptx'
extracted_text_result = extract_text_from_ppt(ppt_file_path)

root = tk.Tk()
root.title("Presentation Reader")

# Text area
text_area = tk.Text(root, wrap="word", width=40, height=10)
text_area.insert("1.0", extracted_text_result)
text_area.config(state="disabled")
text_area.grid(row=0, column=0, padx=10, pady=10, columnspan=2)

# Pause button
pause_button = ttk.Button(root, text="Pause", command=on_pause_button_click)
pause_button.grid(row=1, column=0, padx=5, pady=5)

# Resume button
resume_button = ttk.Button(root, text="Resume", command=on_resume_button_click)
resume_button.grid(row=1, column=1, padx=5, pady=5)

# Bind the close event to on_close function
root.protocol("WM_DELETE_WINDOW", on_close)

# Create TTS controller
tts_controller = TTSController(extracted_text_result)
tts_controller.start()

root.mainloop()
