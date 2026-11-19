
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx2pdf import convert
def pdf_to_word():
    pdf_file = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not pdf_file:
        return
    output_file = filedialog.asksaveasfilename(
        title="Save Word File",
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not output_file:
        return
    converter = None
    try:
        status_label.config(text="Converting PDF to Word...")
        root.update_idletasks()
        converter = Converter(pdf_file)
        converter.convert(output_file)
        converter.close()
        converter = None
        status_label.config(text="Conversion completed!")
        messagebox.showinfo(
            "Success",
            "PDF converted to Word successfully!"
        )
    except Exception as e:
        status_label.config(text="Conversion failed!")
        messagebox.showerror("Error", str(e))
    finally:
        if converter is not None:
            converter.close()

def word_to_pdf():

    word_file = filedialog.askopenfilename(
        title="Select Word File",
        filetypes=[("Word Documents", "*.docx")]
    )

    if not word_file:
        return

    output_file = filedialog.asksaveasfilename(
        title="Save PDF File",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not output_file:
        return

    try:
        status_label.config(text="Converting Word to PDF...")
        root.update_idletasks()

        convert(word_file, output_file)

        status_label.config(text="Conversion completed!")
        messagebox.showinfo(
            "Success",
            "Word converted to PDF successfully!"
        )

    except Exception as e:
        status_label.config(text="Conversion failed!")
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("PDF & Word Converter")
root.geometry("500x350")
root.resizable(False, False)
root.configure(bg="#f2f5fa")

heading = tk.Label(
    root,
    text="PDF & WORD CONVERTER",
    font=("Arial", 20, "bold"),
    bg="#f2f5fa",
    fg="#243b64"
)
heading.pack(pady=30)

description = tk.Label(
    root,
    text="Choose a conversion option",
    font=("Arial", 12),
    bg="#f2f5fa",
    fg="#555555"
)
description.pack(pady=5)

# PDF to Word button
pdf_button = tk.Button(
    root,
    text="PDF to Word",
    font=("Arial", 12, "bold"),
    bg="#2864b4",
    fg="white",
    width=22,
    height=2,
    command=pdf_to_word,
    cursor="hand2"
)
pdf_button.pack(pady=12)
# Word to PDF button
word_button = tk.Button(
    root,
    text="Word to PDF",
    font=("Arial", 12, "bold"),
    bg="#e07835",
    fg="white",
    width=22,
    height=2,
    command=word_to_pdf,
    cursor="hand2"
)
word_button.pack(pady=12)
# Status label
status_label = tk.Label(
    root,
    text="Ready to convert files",
    font=("Arial", 10),
    bg="#f2f5fa",
    fg="#333333"
)
status_label.pack(pady=15)
# Run application
root.mainloop()
