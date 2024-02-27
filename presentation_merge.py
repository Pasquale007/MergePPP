import win32com.client
import os

def merge_presentations(file_paths, output_file_path):
    ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
    output_presentation = ppt_instance.Presentations.Add()

    for path in file_paths:
        presentation = ppt_instance.Presentations.Open(os.path.abspath(path), True, False, False)
        presentation.Slides.Range(range(1, presentation.Slides.Count + 1)).Copy()
        output_presentation.Application.Windows(1).Activate()
        output_presentation.Application.CommandBars.ExecuteMso("PasteSourceFormatting")
        presentation.Close()
    
    output_path = os.path.abspath(output_file_path)
    output_presentation.SaveAs(output_path)
    output_presentation.Close()
    ppt_instance.Quit()
