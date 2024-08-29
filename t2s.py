import tkinter as tk
from tkinter import filedialog, scrolledtext
import pyttsx3
import PyPDF2
import docx
import threading

class TextToSpeechApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Local Text-to-Speech Tool")
        self.master.geometry("600x450")

        self.engine = pyttsx3.init()
        self.is_paused = False
        self.is_stopped = False

        self.create_widgets()

    def create_widgets(self):
        # Text area
        self.text_area = scrolledtext.ScrolledText(self.master, wrap=tk.WORD, width=70, height=20)
        self.text_area.pack(padx=10, pady=10)

        # Buttons
        button_frame = tk.Frame(self.master)
        button_frame.pack(pady=5)

        tk.Button(button_frame, text="Upload File", command=self.upload_file).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Speak All", command=self.speak_all).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Speak Selected", command=self.speak_selected).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Pause/Resume", command=self.toggle_pause).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Stop", command=self.stop_speech).pack(side=tk.LEFT, padx=5)

        # Voice and Speed controls
        control_frame = tk.Frame(self.master)
        control_frame.pack(pady=5)

        tk.Label(control_frame, text="Voice:").pack(side=tk.LEFT, padx=5)
        self.voice_var = tk.StringVar(value="Male")
        tk.OptionMenu(control_frame, self.voice_var, "Male", "Female", command=self.change_voice).pack(side=tk.LEFT, padx=5)

        tk.Label(control_frame, text="Speed:").pack(side=tk.LEFT, padx=5)
        self.speed_scale = tk.Scale(control_frame, from_=50, to=300, orient=tk.HORIZONTAL, command=self.change_speed)
        self.speed_scale.set(150)  # Default speed
        self.speed_scale.pack(side=tk.LEFT, padx=5)

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf"), ("Word Documents", "*.docx"), ("Text Files", "*.txt")])
        if file_path:
            if file_path.endswith('.pdf'):
                self.read_pdf(file_path)
            elif file_path.endswith('.docx'):
                self.read_docx(file_path)
            else:
                with open(file_path, 'r') as file:
                    self.text_area.delete(1.0, tk.END)
                    self.text_area.insert(tk.END, file.read())

    def read_pdf(self, file_path):
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, text)

    def read_docx(self, file_path):
        doc = docx.Document(file_path)
        text = ""
        for para in doc.paragraphs:
            text += para.text + "\n"
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, text)

    def speak_text(self, text):
        self.is_stopped = False
        self.engine.say(text)
        self.engine.runAndWait()

    def speak_all(self):
        text = self.text_area.get(1.0, tk.END)
        threading.Thread(target=self.speak_text, args=(text,)).start()

    def speak_selected(self):
        try:
            selected_text = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text:
                threading.Thread(target=self.speak_text, args=(selected_text,)).start()
            else:
                print("No text selected")
        except tk.TclError:
            print("No text selected or selection error")

    def toggle_pause(self):
        if not self.is_paused:
            self.engine.pause()
            self.is_paused = True
        else:
            self.engine.resume()
            self.is_paused = False

    def stop_speech(self):
        self.is_stopped = True
        self.engine.stop()

    def change_voice(self, choice):
        voices = self.engine.getProperty('voices')
        if choice == "Female":
            self.engine.setProperty('voice', voices[1].id)
        else:
            self.engine.setProperty('voice', voices[0].id)

    def change_speed(self, speed):
        self.engine.setProperty('rate', int(speed))

if __name__ == "__main__":
    root = tk.Tk()
    app = TextToSpeechApp(root)
    root.mainloop()
