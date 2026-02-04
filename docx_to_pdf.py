import os
import win32com.client as win32

# Function with parameter
def convert_docx_to_pdf_batch(input_dir, output_dir):
  # create Output directory
    os.makedirs(output_dir, exist_ok=True)

    word = win32.Dispatch("Word.Application")
    word.Visible = False

    # Loop through a filename with .docx in the input directory
    for filename in os.listdir(input_dir):
        if filename.lower().endswith(".docx"):
            docx_path = os.path.abspath(os.path.join(input_dir, filename))
            pdf_path = os.path.abspath(
                os.path.join(output_dir, filename.replace(".docx", ".pdf"))
            )

            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # PDF
            doc.Close()

    word.Quit()

# Call function with input and output directory
convert_docx_to_pdf_batch("input_docx", "output_pdf")
