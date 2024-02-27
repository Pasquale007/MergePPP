import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import win32com.client
import os

class DragDropListbox(tk.Listbox):
    def __init__(self, master, **kw):
        kw['selectmode'] = tk.SINGLE
        kw['activestyle'] = 'none'
        tk.Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.setCurrent)
        self.bind('<B1-Motion>', self.shiftSelection)
        self.curIndex = None

    def setCurrent(self, event):
        self.curIndex = self.nearest(event.y)

    def shiftSelection(self, event):
        i = self.nearest(event.y)
        if i < self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i+1, x)
            self.curIndex = i
        elif i > self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i-1, x)
            self.curIndex = i

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

def move_items(source_listbox, target_listbox):
    selected_indices = source_listbox.curselection()
    for index in selected_indices:
        target_listbox.insert(tk.END, source_listbox.get(index))
    for index in selected_indices[::-1]:
        source_listbox.delete(index)

def browse_files():
    global file_paths
    file_paths = [os.path.join(input_folder_path, f) for f in os.listdir(input_folder_path) if os.path.isfile(os.path.join(input_folder_path, f)) and f.endswith('.pptx')]
    for path in file_paths:
        filename = os.path.basename(path)
        left_listbox.insert(tk.END, filename)

def merge_files():
    selected_files = right_listbox.get(0, tk.END)
    selected_file_paths = [file_paths[right_listbox.get(0, tk.END).index(item)] for item in selected_files]
    if selected_file_paths:
        output_path = os.path.join(os.getcwd(), output_file_name)
        if os.path.exists(output_path):
            confirm_overwrite = messagebox.askyesno("File Exists", "The merged.pptx file already exists. Do you want to overwrite it?")
            if not confirm_overwrite:
                return
        merge_presentations(selected_file_paths, output_path)
        messagebox.showinfo("Merge Complete", "Presentations merged successfully!")
    else:
        messagebox.showwarning("No Files Selected", "Please select at least one file to merge.")

# Create main window
root = tk.Tk()
root.title("Merge PowerPoint Presentations")

# Frame to hold listboxes and buttons
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Left listbox
left_listbox = DragDropListbox(frame)
left_listbox.pack(side=tk.LEFT, padx=5)

# Buttons frame
button_frame = tk.Frame(frame)
button_frame.pack(side=tk.LEFT, padx=5)

# Button to move items to the right
move_right_button = tk.Button(button_frame, text=">>", command=lambda: move_items(left_listbox, right_listbox))
move_right_button.pack(side=tk.TOP, pady=5)

# Button to move items to the left
move_left_button = tk.Button(button_frame, text="<<", command=lambda: move_items(right_listbox, left_listbox))
move_left_button.pack(side=tk.TOP, pady=5)

# Right listbox
right_listbox = DragDropListbox(frame)
right_listbox.pack(side=tk.LEFT, padx=5)

# Button to merge files
merge_button = tk.Button(root, text="Merge Files", command=merge_files)
merge_button.pack(pady=10)

browse_files()
root.mainloop()
