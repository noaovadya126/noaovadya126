from docx import Document
import os
import win32com.client as win32

def convert_doc_to_docx(doc_path):
    try:
        if not os.path.isfile(doc_path):
            raise FileNotFoundError(f"File not found: {doc_path}")
        output_path = doc_path + "x"
        word = win32.Dispatch("Word.Application")
        word.Visible = False  # Run in background
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(output_path, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
        doc.Close()
        word.Quit()
        return output_path
    except Exception as e:
        return f"Error converting file: {e}"

def extract_approval(file_path):
        doc = Document(file_path)
        for table in doc.tables:  # Iterate through all tables in the document
            for row in table.rows:  # Access each row in the table
                cells = row.cells  # Get all cells in the row
                for i, cell in enumerate(cells):  # Enumerate to get the index
                    if "Report approved by Neteera Head of Operations" in cell.text:
                        # Ensure there's a next cell to access
                        if i + 1 < len(cells):
                            return cells[i + 1].text.strip()  # Return the text in the adjacent cell
                        else:
                            return "Adjacent cell not found"


file_path = "C:/Users/noa.ovadya/Downloads/CVR-81.doc"  # Path to the original .doc file
if file_path.endswith(".doc"):  # Check if it's a .doc file
    file_path = convert_doc_to_docx(file_path)  # Convert to .docx

if file_path and os.path.isfile(file_path):
    approval = extract_approval(file_path)
    print(f"Report approved by Neteera Head of Operations: {approval}")
else:
    print("Failed to process the file.")
