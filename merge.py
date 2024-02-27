import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import win32com.client
import os

input_folder_path = "C:\\Users\\Pasca\\Desktop\\Cross To Harmony\\PPP"
output_file_name = "merged.pptx"
file_paths = [os.path.join(input_folder_path, f) for f in os.listdir(input_folder_path) if os.path.isfile(os.path.join(input_folder_path, f)) and f.endswith('.pptx')]

def merge_presentations(selected_files, output_path):
    ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
    output_presentation = ppt_instance.Presentations.Add()

    for path in selected_files:
        presentation = ppt_instance.Presentations.Open(os.path.abspath(path), True, False, False)
        presentation.Slides.Range(range(1, presentation.Slides.Count + 1)).Copy()
        output_presentation.Application.Windows(1).Activate()
        output_presentation.Application.CommandBars.ExecuteMso("PasteSourceFormatting")
        presentation.Close()

    output_presentation.SaveAs(os.path.abspath(output_path))
    output_presentation.Close()
    ppt_instance.Quit()

def browse_files():
    global file_paths
    for path in file_paths:
        var = tk.BooleanVar()
        checkboxes.append(var)
        filename = os.path.basename(path)
        cb = ttk.Checkbutton(frame, text=filename, variable=var)
        cb.pack(anchor=tk.W)

def search_files():
    search_text = search_entry.get().strip()
    if search_text:
        matching_files = [f for f in file_paths if search_text.lower() in os.path.basename(f).lower()]
        for cb in frame.winfo_children():
            cb.destroy()  # Clear previous checkboxes
        for path in matching_files:
            var = tk.BooleanVar()
            checkboxes.append(var)
            filename = os.path.basename(path)
            cb = ttk.Checkbutton(frame, text=filename, variable=var)
            cb.pack(anchor=tk.W)
    else:
        messagebox.showwarning("Search", "Please enter a search term.")

def reset_files():
    search_entry.delete(0, tk.END)  # Löschen Sie den gesamten Text aus dem Suchfeld
    for cb in frame.winfo_children():
        cb.destroy()  # Clear all checkboxes
    browse_files()  # Reload all files

def merge_files():
    global file_paths
    selected_files = [file for file, var in zip(file_paths, checkboxes) if var.get()]
    if selected_files:
        output_path = os.path.join(os.getcwd(), output_file_name)
        if os.path.exists(output_path):
            confirm_overwrite = messagebox.askyesno("File Exists", "The merged.pptx file already exists. Do you want to overwrite it?")
            if not confirm_overwrite:
                return
        merge_presentations(selected_files, output_path)
        messagebox.showinfo("Merge Complete", "Presentations merged successfully!")
    else:
        messagebox.showwarning("No Files Selected", "Please select at least one file to merge.")

# Create main window
root = tk.Tk()
root.title("Merge PowerPoint Presentations")

# Frame to hold checkboxes and buttons
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Search frame
search_frame = tk.Frame(root)
search_frame.pack(padx=10, pady=(0, 10), fill=tk.X)

# Search field and button
search_label = tk.Label(search_frame, text="Search:")
search_label.pack(side=tk.LEFT)
search_entry = tk.Entry(search_frame)
search_entry.pack(side=tk.LEFT, padx=(0, 5))
search_button = tk.Button(search_frame, text="Search", command=search_files)
search_button.pack(side=tk.LEFT)
reset_button = tk.Button(search_frame, text="Reset", command=reset_files)
reset_button.pack(side=tk.LEFT)

checkboxes = []
browse_files()

# Button to merge files
merge_button = tk.Button(root, text="Merge Files", command=merge_files)
merge_button.pack(pady=10)

root.mainloop()
