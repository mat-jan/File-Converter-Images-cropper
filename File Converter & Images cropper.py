import os
import time
from tkinter import Tk, Label, Button, Entry, Frame, filedialog, messagebox, StringVar, OptionMenu, BooleanVar, Checkbutton, Radiobutton
from tkinter import ttk
from PIL import Image, ImageOps
from docx2pdf import convert as convert_docx
from PyPDF2 import PdfMerger

selected_files = []

# --- KONWERTER ---
def wybierz_folder_konwert():
    folder = filedialog.askdirectory(title="Choose folder with documents")
    if folder:
        file_entry.delete(0, "end")
        file_entry.insert(0, folder)
        global selected_files
        selected_files = []

def wybierz_pliki_konwert():
    files = filedialog.askopenfilenames(
        title="Choose files",
        filetypes=[("All supported", "*.jpg *.jpeg *.png *.bmp *.tiff *.docx *.pdf")]
    )
    if files:
        file_entry.delete(0, "end")
        file_entry.insert(0, f"{len(files)} files selected")
        global selected_files
        selected_files = list(files)

def wybierz_zapis_konwert():
    path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[
            ("PDF", "*.pdf"),
            ("JPEG", "*.jpg"),
            ("PNG", "*.png"),
            ("BMP", "*.bmp"),
            ("TIFF", "*.tiff")
        ],
        title="Save as"
    )
    if path:
        pdf_entry.delete(0, "end")
        pdf_entry.insert(0, path)

def konwertuj():
    folder_or_files = file_entry.get().strip()
    output_file = pdf_entry.get().strip()
    if not folder_or_files or not output_file:
        messagebox.showerror("Error!", "Select folder/file and specify output file.")
        return

    ext_output = os.path.splitext(output_file)[1].lower()
    if ext_output not in [".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tiff"]:
        messagebox.showwarning("Warning!", "Unsupported output format, default PDF.")
        output_file += ".pdf"
        ext_output = ".pdf"

    pdf_files = []
    files_to_process = selected_files if selected_files else [
        os.path.join(folder_or_files, f) for f in os.listdir(folder_or_files)
        if f.lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".docx", ".pdf"))
    ]
    files_to_process.sort()
    if not files_to_process:
        messagebox.showwarning("No files", "No files found to process.")
        return

    progress_konvert['value'] = 0
    progress_konvert['maximum'] = len(files_to_process)
    status_konvert_var.set("Starting conversion...")
    root.update_idletasks()
    start_time = time.time()
    processed = 0

    for f in files_to_process:
        ext = os.path.splitext(f)[1].lower()
        try:
            if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff"]:
                img = Image.open(f).convert("RGB")
                if ext_output == ".pdf":
                    tmp_pdf = os.path.splitext(f)[0] + "_tmp.pdf"
                    img.save(tmp_pdf)
                    pdf_files.append(tmp_pdf)
                else:
                    img.save(output_file)
                    messagebox.showinfo("Success", f"File saved as {output_file}")
                    return
            elif ext == ".docx":
                if ext_output != ".pdf":
                    messagebox.showwarning("Warning!", "Word can only be converted to PDF.")
                    ext_output = ".pdf"
                    output_file = os.path.splitext(output_file)[0] + ".pdf"
                tmp_pdf = os.path.splitext(f)[0] + "_tmp.pdf"
                convert_docx(f, tmp_pdf)
                pdf_files.append(tmp_pdf)
            elif ext == ".pdf":
                pdf_files.append(f)
        except Exception as e:
            print(f"Failed to process {f}: {e}")

        processed += 1
        progress_konvert['value'] = processed
        elapsed = time.time() - start_time
        eta = elapsed / processed * (len(files_to_process) - processed)
        status_konvert_var.set(f"Processed {processed}/{len(files_to_process)} | ETA: {int(eta)}s")
        root.update_idletasks()

    if ext_output == ".pdf":
        try:
            merger = PdfMerger()
            for pdf in pdf_files:
                merger.append(pdf)
            merger.write(output_file)
            merger.close()
            for pdf in pdf_files:
                if pdf.endswith("_tmp.pdf") and os.path.exists(pdf):
                    os.remove(pdf)
            messagebox.showinfo("Success", f"PDF saved as:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error!", f"Could not merge PDFs:\n{e}")
    status_konvert_var.set("Conversion complete!")

# --- PRZYCINANIE ---
def wybierz_folder_przycinanie():
    folder = filedialog.askdirectory(title="Select source folder for cropping")
    if folder:
        crop_input_entry.delete(0, "end")
        crop_input_entry.insert(0, folder)

def wybierz_folder_wyjscie():
    folder = filedialog.askdirectory(title="Select the destination folder")
    if folder:
        crop_output_entry.delete(0, "end")
        crop_output_entry.insert(0, folder)

def przytnij_obrazy():
    folder_or_files = crop_input_entry.get().strip()
    output_folder = crop_output_entry.get().strip()
    if not folder_or_files or not output_folder:
        messagebox.showerror("Error!", "Select the source and destination folders for the cropped images.")
        return

    try:
        target_w = int(width_entry.get())
        target_h = int(height_entry.get())
    except ValueError:
        messagebox.showerror("Error!", "Please enter the correct numbers for width and height.")
        return

    do_przyciecia = not (target_w == 0 and target_h == 0)

    try:
        start_num = int(num_start_entry.get()) if num_start_entry.get().strip() else 1
        end_num = int(num_end_entry.get()) if num_end_entry.get().strip() else None
    except ValueError:
        messagebox.showerror("Error!", "Provide correct numbers for numbering.")
        return

    files_to_process = [
        os.path.join(folder_or_files, f) for f in os.listdir(folder_or_files)
        if f.lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".tiff"))
    ]

    reverse_order_val = reverse_var.get()
    if sort_option.get() == "Alphabetically":
        files_to_process.sort(reverse=reverse_order_val)
    else:
        files_to_process.sort(key=lambda x: os.path.getmtime(x), reverse=reverse_order_val)

    if not files_to_process:
        messagebox.showwarning("No files", "No images found to crop.")
        return

    progress_crop['value'] = 0
    progress_crop['maximum'] = len(files_to_process)
    status_crop_var.set("Starting to crop...")
    root.update_idletasks()
    start_time = time.time()
    processed = 0
    num = start_num

    for f in files_to_process:
        if end_num and num > end_num:
            break
        try:
            img = Image.open(f).convert("RGB")
            img = ImageOps.exif_transpose(img)
            if do_przyciecia:
                src_ratio = img.width / img.height
                target_ratio = target_w / target_h
                if src_ratio > target_ratio:
                    new_height = target_h
                    new_width = int(new_height * src_ratio)
                else:
                    new_width = target_w
                    new_height = int(new_width / src_ratio)
                img = img.resize((new_width, new_height), Image.LANCZOS)
                left = (new_width - target_w) // 2
                top = (new_height - target_h) // 2
                right = left + target_w
                bottom = top + target_h
                img = img.crop((left, top, right, bottom))

            if naming_mode.get() == "oryginal":
                out_name = os.path.basename(f)
            else:
                out_name = f"{num}.png"

            output_path = os.path.join(output_folder, out_name)
            img.save(output_path, "PNG", quality=95)

            # zachowanie dat modyfikacji
            stat = os.stat(f)
            os.utime(output_path, (stat.st_atime, stat.st_mtime))

        except Exception as e:
            print(f"Error przy {f}: {e}")

        processed += 1
        num += 1
        progress_crop['value'] = processed
        elapsed = time.time() - start_time
        eta = elapsed / processed * (len(files_to_process) - processed)
        status_crop_var.set(f"Processed {processed}/{len(files_to_process)} | ETA: {int(eta)}s")
        root.update_idletasks()

    status_crop_var.set("Crop finished!")
    messagebox.showinfo("Finished", f"Cropped {processed} images.\nSaved in: {output_folder}")

# --- GUI ---
root = Tk()
root.title("File converter & Image cropper")
root.geometry("720x850")
root.resizable(False, False)

# KONWERTER
Label(root, text="File converter ", font=("Arial", 12, "bold")).pack(pady=(10,0))
file_entry = Entry(root, width=65)
file_entry.pack(pady=5)
Frame_buttons = Frame(root)
Frame_buttons.pack(pady=5)
Button(Frame_buttons, text="Select folder", command=wybierz_folder_konwert).grid(row=0, column=0, padx=5)
Button(Frame_buttons, text="Select files", command=wybierz_pliki_konwert).grid(row=0, column=1, padx=5)
pdf_entry = Entry(root, width=55)
pdf_entry.pack(pady=5)
Button(root, text="Choose location to save", command=wybierz_zapis_konwert).pack(pady=5)
Button(root, text="Convert", command=konwertuj, bg="#4CAF50", fg="white", font=("Arial",12,"bold")).pack(pady=5)

progress_konvert = ttk.Progressbar(root, orient='horizontal', length=600, mode='determinate')
progress_konvert.pack(pady=5)
status_konvert_var = StringVar()
Label(root, textvariable=status_konvert_var,font=("Arial",10)).pack(pady=(0,10))

# PRZYCINANIE
Label(root, text="Image cropper", font=("Arial", 12, "bold")).pack(pady=(10,0))
Label(root, text="Source folder:").pack()
crop_input_entry = Entry(root, width=60)
crop_input_entry.pack(pady=5)
Button(root, text="Select destination folder", command=wybierz_folder_przycinanie).pack(pady=5)
Label(root, text="Select destination folder:").pack()
crop_output_entry = Entry(root, width=60)
crop_output_entry.pack(pady=5)
Button(root, text="Select destination folder:", command=wybierz_folder_wyjscie).pack(pady=5)

size_frame = Frame(root)
size_frame.pack(pady=5)
Label(size_frame, text="width").grid(row=0,column=0)
width_entry = Entry(size_frame, width=6)
width_entry.insert(0,"0")
width_entry.grid(row=0,column=1,padx=5)
Label(size_frame, text="Hight").grid(row=0,column=2)
height_entry = Entry(size_frame, width=6)
height_entry.insert(0,"0")
height_entry.grid(row=0,column=3,padx=5)

num_frame = Frame(root)
num_frame.pack(pady=5)
Label(num_frame, text="From").grid(row=0,column=0)
num_start_entry = Entry(num_frame,width=6)
num_start_entry.grid(row=0,column=1,padx=5)
Label(num_frame,text="To").grid(row=0,column=2)
num_end_entry = Entry(num_frame,width=6)
num_end_entry.grid(row=0,column=3,padx=5)

# --- Tryb nazewnictwa ---
naming_mode = StringVar(value="oryginal")
Label(root, text="Filename Mode:").pack(pady=(5,0))
Frame_naming = Frame(root)
Frame_naming.pack()
Radiobutton(Frame_naming, text="Keep original dates", variable=naming_mode, value="oryginal").grid(row=0, column=0, padx=10)
Radiobutton(Frame_naming, text="Numeruj (od X do Y)", variable=naming_mode, value="numeration").grid(row=0, column=1, padx=10)

sort_option = StringVar(root)
sort_option.set("Alphabetically")
OptionMenu(root, sort_option, "Alphabetically", "By modification date").pack(pady=5)
reverse_var = BooleanVar()
Checkbutton(root,text="Reverse order (newest to oldest)", variable=reverse_var).pack(pady=5)

progress_crop = ttk.Progressbar(root, orient='horizontal', length=600, mode='determinate')
progress_crop.pack(pady=5)
status_crop_var = StringVar()
Label(root, textvariable=status_crop_var,font=("Arial",10)).pack(pady=(0,10))
Button(root,text="Crop images", command=przytnij_obrazy, bg="#2196F3", fg="white", font=("Arial",12,"bold")).pack(pady=10)

root.mainloop()
