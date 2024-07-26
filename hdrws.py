import tkinter as tk
from tkinter import filedialog, messagebox, Scrollbar, Scale
import pyttsx3
from docx import Document
import PyPDF2

def load_file():
    global loaded_text
    loaded_text = ""
    file_path = filedialog.askopenfilename(filetypes=[("Text Documents", ".txt;.pdf;*.docx")])
    if not file_path:
        return

    try:
        if file_path.endswith('.docx'):
            doc = Document(file_path)
            loaded_text = ' '.join([para.text for para in doc.paragraphs if para.text])
        elif file_path.endswith('.pdf'):
            with open(file_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                loaded_text = ' '.join([page.extract_text() or "" for page in reader.pages])
        elif file_path.endswith('.txt'):
            with open(file_path, 'r', encoding='utf-8') as file:
                loaded_text = file.read()
        else:
            messagebox.showerror("Unsupported file", "The selected file type is not supported.")
            return
    except Exception as e:
        messagebox.showerror("Error reading file", str(e))
        return

    doc_text.delete('1.0', tk.END)
    doc_text.insert(tk.END, loaded_text)

def convert_to_speech():
    if not loaded_text:
        messagebox.showinfo("Error", "Please load a document first!")
        return

    try:
        voice_id = voice_var.get()
        engine.setProperty('voice', voice_id)
        engine.setProperty('rate', rate_scale.get())
        engine.setProperty('volume', volume_scale.get() / 100)
        engine.say(loaded_text)
        engine.runAndWait()
    except Exception as e:
        messagebox.showerror("Speech Error", str(e))

def update_voices():
    try:
        voices = engine.getProperty('voices')
        voice_menu['menu'].delete(0, 'end')
        for voice in voices:
            voice_menu['menu'].add_command(label=voice.name, command=lambda v=voice.id: voice_var.set(v))
    except Exception as e:
        messagebox.showerror("Voice Error", str(e))

root = tk.Tk()
root.title("Преобразователь документов в речь")

engine = pyttsx3.init()
loaded_text = ""

frame = tk.Frame(root)
frame.pack(pady=20)

doc_text = tk.Text(frame, height=10, width=50, wrap="none")  # wrap="none" disables word wrapping
doc_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scroll_y = Scrollbar(frame, orient="vertical", command=doc_text.yview)
scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
doc_text.config(yscrollcommand=scroll_y.set)

scroll_x = Scrollbar(frame, orient="horizontal", command=doc_text.xview)
scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
doc_text.config(xscrollcommand=scroll_x.set)

load_button = tk.Button(root, text="Загрузить документ", command=load_file)
load_button.pack(pady=10)

voice_var = tk.StringVar(root)
voice_menu = tk.OptionMenu(root, voice_var, "")
voice_menu.pack()
update_voices()

rate_scale = Scale(root, from_=100, to=300, orient=tk.HORIZONTAL, label="Скорость речи")
rate_scale.set(200)
rate_scale.pack()

volume_scale = Scale(root, from_=0, to=100, orient=tk.HORIZONTAL, label="Громкость")
volume_scale.set(100)
volume_scale.pack()

convert_button = tk.Button(root, text="Преобразовать в речь", command=convert_to_speech)
convert_button.pack(pady=10)

root.mainloop()