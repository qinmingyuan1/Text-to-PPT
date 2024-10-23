import os
import comtypes.client

def ppt_to_pdf(ppt_file_path, pdf_file_path):
    # Ensure the file paths are absolute
    ppt_file_path = os.path.abspath(ppt_file_path)
    pdf_file_path = os.path.abspath(pdf_file_path)

    # Check if the PPT file exists
    if not os.path.exists(ppt_file_path):
        raise FileNotFoundError(f"The file {ppt_file_path} does not exist.")

    # Create PowerPoint application object
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    try:
        # Open the presentation
        presentation = powerpoint.Presentations.Open(ppt_file_path)

        # Save as PDF
        presentation.SaveAs(pdf_file_path, 32)  # 32 is the formatType for PDF

        # Close the presentation
        presentation.Close()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit PowerPoint
        powerpoint.Quit()

ppt_to_pdf("./temp_ppt.pptx", "./tempoutput.pdf")

# # 示例用法
# ppt_file_path = "path/to/your/F1.pptx"
# pdf_file_path = "path/to/your/F1.pdf"
# ppt_to_pdf(ppt_file_path, pdf_file_path)
# print(f"PDF saved to {pdf_file_path}")