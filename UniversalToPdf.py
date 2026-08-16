# MIT License - Copyright (c) 2026 Giacomo Rosatelli - vedi LICENSE
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os
from fpdf import FPDF
from PIL import Image
from pypdf import PdfWriter 

# Lista di estensioni da escludere (Software/Eseguibili/Sistema)
EXCLUDED_EXTENSIONS = [
    '.exe', '.msi', '.bat', '.cmd', '.sh', '.bin', '.dll', 
    '.sys', '.app', '.dmg', '.pkg', '.com', '.vbs', '.js'
]


def process_files_to_pdf(input_paths, output_pdf):
    """Crea un PDF (singolo o unito) dai file in input (escludendo i pdf nativi che vengono gestiti separatamente)."""
    try:
        pdf = FPDF()
        for path in input_paths:
            ext = os.path.splitext(path)[1].lower()
            if ext in EXCLUDED_EXTENSIONS: continue

            # --- IMMAGINI ---
            if ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp"]:
                try:
                    img = Image.open(path).convert("RGB")
                    pdf.add_page()
                    pdf.image(path, x=10, y=10, w=190)
                    continue 
                except: pass
            
            # --- TESTO/ALTRO ---
            if ext != ".pdf":
                pdf.add_page()
                pdf.set_font("Arial", 'B', 12)
                pdf.cell(0, 10, f"File: {os.path.basename(path)}", ln=True)
                pdf.set_font("Arial", size=10)
                pdf.ln(5)
                try:
                    with open(path, "r", encoding="utf-8", errors="replace") as f:
                        text = f.read(8000)
                    for line in text.splitlines():
                        clean_line = line.encode('latin-1', 'replace').decode('latin-1')
                        pdf.multi_cell(0, 5, clean_line)
                except:
                    pdf.cell(0, 10, "[Content not readable]", ln=True)
        
        pdf.output(output_pdf)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"PDF saving failed:\n{e}")
        return False

def convert_pdf_to_docx(pdf_path, docx_path):
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert {pdf_path}:\n{e}")

def merge_pdfs(pdf_paths, output_path):
    try:
        merger = PdfWriter()
        for pdf in pdf_paths:
            merger.append(pdf)
        merger.write(output_path)
        merger.close()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to merge PDFs:\n{e}")

def select_files():
    files = filedialog.askopenfilenames(title="Select file", filetypes=[("All files", "*.*")])
    if files:
        filtered_files = [f for f in files if os.path.splitext(f)[1].lower() not in EXCLUDED_EXTENSIONS]
        if len(filtered_files) < len(files):
            messagebox.showwarning("Filter", "Software and executables have been excluded.")
        entry_files.delete(0, tk.END)
        entry_files.insert(0, "; ".join(filtered_files))

def clear_list():
    entry_files.delete(0, tk.END)

def convert():
    input_paths = [p.strip() for p in entry_files.get().split("; ") if p.strip()]
    if not input_paths:
        messagebox.showwarning("Attention", "Select at least one file.")
        return

    should_merge_others = var_merge.get()
    pdf_action = var_pdf_action.get() # 1 = Convert DOCX, 2 = Merge PDF
    
    pdfs = [f for f in input_paths if os.path.splitext(f)[1].lower() == ".pdf"]
    others = [f for f in input_paths if os.path.splitext(f)[1].lower() != ".pdf"]

    # 1. Gestione dei file NON-PDF (Immagini, Testo, ecc.)
    if others:
        if should_merge_others:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                title="Choose the name for the merged PDF file (from not PDFs)"
            )
            if save_path:
                process_files_to_pdf(others, save_path)
        else:
            for path in others:
                save_path = filedialog.asksaveasfilename(
                    initialfile=os.path.splitext(os.path.basename(path))[0] + ".pdf",
                    defaultextension=".pdf",
                    title=f"Save PDF As (for {os.path.basename(path)})"
                )
                if save_path: process_files_to_pdf([path], save_path)

    # 2. Gestione dei file PDF Nativi
    if pdfs:
        if pdf_action == 1:
            # Converti tutti i PDF in DOCX separatamente
            for p in pdfs:
                docx_path = filedialog.asksaveasfilename(
                    initialfile=os.path.splitext(os.path.basename(p))[0] + ".docx",
                    defaultextension=".docx",
                    title=f"Save Word conversion for: {os.path.basename(p)}"
                )
                if docx_path: convert_pdf_to_docx(p, docx_path)
        elif pdf_action == 2:
            # Unisci i PDF
            if len(pdfs) > 1:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF", "*.pdf")],
                    title="Choose the name for the MERGED PDFs file"
                )
                if save_path: merge_pdfs(pdfs, save_path)
            else:
                messagebox.showinfo("Info", "Only one PDF was selected. Merging requires at least two PDF files.")
                
    messagebox.showinfo("End", "Operation completed!")

# --- GUI ---
root = tk.Tk()
root.title("Universal To Pdf")
root.geometry("700x420")
root.configure(bg="#f4f4f4")

tk.Label(root, text="Select files to convert or merge", font=("Arial", 12, "bold"), bg="#f4f4f4").pack(pady=10)

frame_input = tk.Frame(root, bg="#f4f4f4")
frame_input.pack(padx=20, fill="x")
entry_files = tk.Entry(frame_input, font=("Arial", 10))
entry_files.pack(side="left", fill="x", expand=True, padx=(0, 5))
tk.Button(frame_input, text="Browse", command=select_files).pack(side="right")

tk.Button(root, text="Empty list", command=clear_list).pack(pady=5)

# Sezione per i file non-pdf
frame_others = tk.LabelFrame(root, text="Options for not PDF files", bg="#f4f4f4")
frame_others.pack(padx=20, pady=5, fill="x")
var_merge = tk.IntVar()
tk.Checkbutton(frame_others, text="Merge files into a single PDF", variable=var_merge, bg="#f4f4f4").pack(anchor="w", padx=10, pady=5)

# Sezione per i file PDF
frame_pdf = tk.LabelFrame(root, text="Options for loaded PDF files", bg="#f4f4f4")
frame_pdf.pack(padx=20, pady=5, fill="x")
var_pdf_action = tk.IntVar(value=1)
tk.Radiobutton(frame_pdf, text="Convert to DOCX", variable=var_pdf_action, value=1, bg="#f4f4f4").pack(side="left", padx=10, pady=5)
tk.Radiobutton(frame_pdf, text="Merge into a single PDF", variable=var_pdf_action, value=2, bg="#f4f4f4").pack(side="left", padx=10, pady=5)

tk.Button(root, text="START AND CHOOSE NAMES", command=convert, 
          fg="black", font=("Arial", 11, "bold"), padx=20, pady=10).pack(pady=15)

root.mainloop()
