# PDF Tool Developer Documentation

## I. Overview
This tool is a Python-based desktop application that uses the `tkinter` library to create a Graphical User Interface (GUI). It provides the functions of converting multiple PDF files into Word documents and merging multiple PDF files.

## II. Environment Requirements
- **Python Version**: Python 3.x
- **Dependent Libraries**:
  - `tkinter`: A built-in Python library used for creating GUIs.
  - `pdf2docx`: Used to convert PDF files into Word documents.
  - `PyPDF2`: Used to merge multiple PDF files.
  - `ctypes`: A built-in Python library used to solve the font blurring problem.

## III. Usage Instructions
1. Run the program, and a window will pop up.
2. Click the "Select PDF" button to select one or more PDF files.
3. After selecting the files, the "Convert to Word" and "Merge PDFs" buttons will become available.
    - Click the "Convert to Word" button to convert the selected PDF files into Word documents.
    - Click the "Merge PDFs" button, select the save path, and merge the selected PDF files into one PDF file.
4. In the file list box, you can adjust the file order by dragging the files.
5. Right-click on a file in the file list box and select "Delete" to delete the file.

## IV. Code Structure
### 1. Class Definition
The `PDFTool` class is the core of the entire application, responsible for creating the GUI interface and handling various operations.

#### Initialization Method `__init__(self, root)`
- Solve the font blurring problem and enable DPI awareness.
- Set the window title, size, and position, and disable window resizing.
- Create interface elements, including labels, buttons, list boxes, etc.
- Bind mouse events and right-click events.
- Initialize variables, such as `pdf_path` used to store the paths of the selected PDF files, and `dragged_index` used to record the index of the dragged file.

#### Window Centering Method `center_window(self)`
- Force an update of the window's geometric information.
- Obtain the width and height of the screen and the window.
- Calculate the centered position of the window and set the window position.

#### Select PDF File Method `select_pdf(self)`
- Open a file selection dialog box, allowing the user to select one or more PDF files.
- If files are selected, enable the conversion and merge buttons, and update the file list display.
- Display the information of the selected files.

#### Update File List Method `update_file_listbox(self)`
- Clear the current file list.
- Iterate through the `pdf_path` list and insert the file names into the list box.

#### Convert PDF to Word Method `convert_pdf_to_word(self)`
- Check if PDF files are selected. If not, display an error message.
- Iterate through the selected PDF files and use the `pdf2docx` library to convert them into Word documents.
- Display the success or failure information of the conversion.

#### Merge PDF Files Method `merge_pdfs(self)`
- Check if PDF files are selected. If not, display an error message.
- Open a save file dialog box, allowing the user to specify the save path of the merged PDF file.
- Use the `PyPDF2` library to merge the selected PDF files and save them to the specified path.
- Display the success or failure information of the merge.

#### Start Dragging Method `on_start_drag(self, event)`
- Obtain the file index corresponding to the mouse click position.
- Record the index of the dragged file and set the file as the selected state.

#### Dragging Process Method `on_drag(self, event)`
- If a file is being dragged, obtain the file index corresponding to the current mouse position.
- If the target index is different from the dragged index, swap the positions in the file list and update the file list display.
- Update the dragged index and set the target file as the selected state.

#### Dragging End Method `on_drop(self, event)`
- Clear the record of the dragged index.

#### Display Context Menu Method `show_context_menu(self, event)`
- Obtain the file index corresponding to the mouse click position.
- Set the file as the selected state.
- Display the context menu at the mouse position.

#### Delete Selected File Method `delete_selected_file(self)`
- Obtain the index of the selected file.
- If a file is selected, delete the file from the `pdf_path` list and update the file list display.
- If the file list is empty, disable the conversion and merge buttons.

### 2. Main Program
```python
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFTool(root)
    root.mainloop()
```
Create the `tkinter` main window, instantiate the `PDFTool` class, and start the main event loop.

## V. Notes
- The conversion and merge operations may take some time, depending on the file size and system performance.
- If an error occurs during the conversion or merge process, an error message box will pop up to display the specific error content.
- Currently, some PDF files cannot be converted into Word documents. For details, refer to the documentation of the `pdf2docx` library at https://github.com/ArtifexSoftware/pdf2docx. 
