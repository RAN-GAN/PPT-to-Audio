from pptx import Presentation
import pyttsx3
import tkinter as tk
from tkinter import ttk
import threading
import time

class TTSController:
    def __init__(self, text_to_speak):
        self.text_to_speak = text_to_speak
        self.pause_flag = False
        self.exit_flag = False
        self.speak_thread = None

    def start(self):
        self.speak_thread = threading.Thread(target=self.speak)
        self.speak_thread.start()

    def speak(self):
        bot = pyttsx3.init()
        for text in self.text_to_speak.split('\n'):
            if self.pause_flag:
                while self.pause_flag:
                    time.sleep(1)
            bot.say(text)
            bot.runAndWait()

            if self.exit_flag:
                break

    def pause(self):
        self.pause_flag = True
        print("TTS paused")

    def resume(self):
        self.pause_flag = False
        print("TTS resumed")

    def stop(self):
        self.exit_flag = True
        if self.speak_thread:
            self.speak_thread.join()

# Make tts_controller a global variable
tts_controller = None

def extract_text_from_ppt(ppt_file):
    presentation = Presentation(ppt_file)
    extracted_text = ""
    for slide_number, slide in enumerate(presentation.slides):
        extracted_text += f"\nSlide {slide_number + 1}:\n"
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                extracted_text += f"{shape.text}\n"
    print(extracted_text)
    return extracted_text

def on_pause_button_click():
    tts_controller.pause()

def on_resume_button_click():
    tts_controller.resume()

def on_stop_button_click():
    print("Stop button clicked")
    tts_controller.stop()
    control.destroy()

def on_close():
    tts_controller.stop()
    control.destroy()

def create_gui():
    global tts_controller  # Ensure that tts_controller is recognized as a global variable
    ppt_file_path = 'test.pptx'
    extracted_text_result = extract_text_from_ppt(ppt_file_path)

    global control 
    control = tk.Tk()
    control.title("Presentation Reader")

    # Text area
    text_area = tk.Text(control, wrap="word", width=50, height=20)
    text_area.insert("1.0", extracted_text_result)
    text_area.config(state="disabled")
    text_area.grid(row=0, column=0, padx=10, pady=10, columnspan=3,rowspan=15)

    # Pause button
    pause_button = ttk.Button(control, text="Pause", command=on_pause_button_click)
    pause_button.grid(row=0, column=3)

    # Resume button
    resume_button = ttk.Button(control, text="Resume", command=on_resume_button_click)
    resume_button.grid(row=1, column=3)

    # Stop button
    stop_button = ttk.Button(control, text="Exit", command=on_stop_button_click)
    stop_button.grid(row=2, column=3)

    # Create TTS controller
    tts_controller = TTSController(extracted_text_result)
    tts_controller.start()

    # Bind the close event to on_close function
    control.protocol("WM_DELETE_WINDOW", on_close)

    control.mainloop()

def main_gui():
    root = tk.Tk()
    root.title("Presentation Reader")
    
create_gui()

