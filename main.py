import tkinter as tk
from tkinter import ttk
import pytesseract
import mss
import time
import threading
import cv2
import numpy as np
import win32com.client

# Set Tesseract path if needed
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

class WindowsNarrator:
    def __init__(self):
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.rate = 20  # Default rate; can range approx -10 to +10
        self.voice = self.speaker.Voice  # current voice

    def set_rate(self, rate):
        self.rate = rate
        self.speaker.Rate = self.rate

    def get_rate(self):
        return self.speaker.Rate

    def get_voices(self):
        voices = []
        for i in range(self.speaker.GetVoices().Count):
            v = self.speaker.GetVoices().Item(i)
            voices.append((v.GetAttribute("Name"), v))
        return voices

    def set_voice(self, voice):
        self.voice = voice
        self.speaker.Voice = voice

    def speak(self, text):
        self.speaker.Speak(text)

class OCRReader:
    def __init__(self, region, voice):
        self.region = region
        self.last_text = ""
        self.running = True
        self.voice = voice

    def start(self):
        threading.Thread(target=self.read_loop, daemon=True).start()

    def stop(self):
        self.running = False

    def read_loop(self):
        with mss.mss() as sct:
            while self.running:
                img = np.array(sct.grab(self.region))
                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
                text = pytesseract.image_to_string(thresh, lang='eng').strip()

                if text and text != self.last_text:
                    print("Reading:", text)
                    self.voice.speak(text)
                    self.last_text = text

                time.sleep(1)

class TransparentOverlay:
    def __init__(self, region, reader, voice):
        self.region = region
        self.reader = reader
        self.voice = voice

        btn_height = 25
        win_width = region['width']
        win_height = region['height'] + btn_height  # extra space for buttons at bottom

        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.geometry(f"{win_width}x{win_height}+{region['left']}+{region['top']}")
        self.root.configure(bg='magenta')
        self.root.wm_attributes('-transparentcolor', 'magenta')

        # Canvas fills the top part
        self.canvas = tk.Canvas(self.root, bg='magenta', highlightthickness=2, highlightbackground='red')
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Draggable bottom bar frame
        self.bottom_bar = tk.Frame(self.root, bg='gray20', height=btn_height)
        self.bottom_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # Close button on right
        close_btn = tk.Button(self.bottom_bar, text="X", command=self.close, bg='red', fg='white', bd=0)
        close_btn.pack(side=tk.RIGHT, padx=2)

        # Speed down button on left
        minus_btn = tk.Button(self.bottom_bar, text="-", command=self.slow_down, bg='gray30', fg='white', bd=0)
        minus_btn.pack(side=tk.LEFT, padx=2)

        # Speed up button next to minus
        plus_btn = tk.Button(self.bottom_bar, text="+", command=self.speed_up, bg='gray30', fg='white', bd=0)
        plus_btn.pack(side=tk.LEFT, padx=2)

        # Voices dropdown fills remaining bottom bar space
        self.voice_var = tk.StringVar()
        voices = self.voice.get_voices()
        self.voices_list = voices
        voice_names = [v[0] for v in voices]

        # Try to select Microsoft Zira Desktop by default
        preferred_voice = "Microsoft Zira Desktop"
        matched_voice = next((v for v in voice_names if preferred_voice.lower() in v.lower()), voice_names[0])
        self.voice_var.set(matched_voice)

        # Immediately set the narrator to use that voice
        for name, voice_obj in voices:
            if name == matched_voice:
                self.voice.set_voice(voice_obj)
                print(f"Voice set to: {matched_voice}")
                break


        self.dropdown = ttk.Combobox(self.bottom_bar, textvariable=self.voice_var, values=voice_names, state="readonly")
        self.dropdown.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.dropdown.bind("<<ComboboxSelected>>", self.voice_changed)

        # Drag functions for bottom bar
        def start_move(event):
            self.root.x = event.x
            self.root.y = event.y

        def do_move(event):
            dx = event.x - self.root.x
            dy = event.y - self.root.y
            x = self.root.winfo_x() + dx
            y = self.root.winfo_y() + dy
            self.root.geometry(f"+{x}+{y}")

        self.bottom_bar.bind("<ButtonPress-1>", start_move)
        self.bottom_bar.bind("<B1-Motion>", do_move)

        self.root.mainloop()

    def slow_down(self):
        rate = self.voice.get_rate()
        if rate > -10:
            self.voice.set_rate(rate - 1)
            print(f"Speed down: {rate - 1}")

    def speed_up(self):
        rate = self.voice.get_rate()
        if rate < 10:
            self.voice.set_rate(rate + 1)
            print(f"Speed up: {rate + 1}")

    def voice_changed(self, event):
        selected_name = self.voice_var.get()
        for name, voice_obj in self.voices_list:
            if name == selected_name:
                self.voice.set_voice(voice_obj)
                print(f"Voice changed to: {selected_name}")
                break

    def close(self):
        self.reader.stop()
        self.root.destroy()


class RegionSelector:
    def __init__(self):
        self.region = None
        self.rect = None
        self.start_x = self.start_y = None

    def get_region(self):
        root = tk.Tk()
        root.attributes("-fullscreen", True)
        root.attributes("-alpha", 0.3)
        root.configure(bg='black')
        canvas = tk.Canvas(root, cursor="cross", bg="black")
        canvas.pack(fill=tk.BOTH, expand=True)

        def on_press(event):
            self.start_x = canvas.canvasx(event.x)
            self.start_y = canvas.canvasy(event.y)
            self.rect = canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

        def on_drag(event):
            if self.rect:
                cur_x = canvas.canvasx(event.x)
                cur_y = canvas.canvasy(event.y)
                canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

        def on_release(event):
            x1 = int(self.start_x)
            y1 = int(self.start_y)
            x2 = int(canvas.canvasx(event.x))
            y2 = int(canvas.canvasy(event.y))

            self.region = {
                "left": min(x1, x2),
                "top": min(y1, y2),
                "width": abs(x2 - x1),
                "height": abs(y2 - y1)
            }
            root.destroy()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        root.mainloop()
        return self.region

def main():
    selector = RegionSelector()
    region = selector.get_region()
    if not region:
        return

    voice = WindowsNarrator()
    voice.set_rate(0)  # default rate

    reader = OCRReader(region, voice)
    reader.start()

    TransparentOverlay(region, reader, voice)

if __name__ == "__main__":
    main()
