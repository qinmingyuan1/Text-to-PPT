import streamlit as st
import base64
import pptx
from pptx.util import Inches, Pt
import os
import toml,json, requests

import pandas as pd
import docx2txt
import fitz  # PyMuPDF
from pdf2image import convert_from_bytes
from PIL import Image, ImageDraw

import tempfile
import comtypes.client

# from pptx import Presentation
# from pdf2image import convert_from_path
# from PIL import Image
# import io
# import os
# from comtypes.client import CreateObject


#openai.api_key = os.getenv('OPENAI_API_KEY')  # Replace with your actual API key
file_path = './credential.txt'
if os.path.exists(file_path):
    with open(file_path, 'r') as f:
        secrets = toml.load(f)
else:
    st.warning("Credentials file not found. Please upload the credentials file.")
OPENROUTER_API_KEY = secrets['OPENROUTER']['OPENROUTER_API_KEY']
# Define custom formatting options
TITLE_FONT_SIZE = Pt(30)
SLIDE_FONT_SIZE = Pt(16)


def generate_slide_titles(topic):
    prompt = f"Generate 5 slide titles for the topic '{topic}'."
    msg = [
        {"role": "system", "content":prompt},
        {"role":"user", "content":topic}
    ]
    response = requests.post(
        url = "https://openrouter.ai/api/v1/chat/completions",
        headers = {"Authorization": f"Bearer {OPENROUTER_API_KEY}"},
        data = json.dumps({
            "messages": msg,
            "model": "openai/gpt-4o-mini-2024-07-18"
        })
    )
    #print(response.json())
    resp = response.json()['choices'][0]['message']['content'].split("\n")
    #resp = response.json()['choices'][0]['message']
    #print("resp",resp)
    return resp

def generate_slide_content(slide_title):
    prompt = f"Generate content for the slide: '{slide_title}'."
    msg = [
        {"role": "system", "content":prompt},
        {"role":"user", "content":slide_title}
    ]
    response = requests.post(
        url = "https://openrouter.ai/api/v1/chat/completions",
        headers = {"Authorization": f"Bearer {OPENROUTER_API_KEY}"},
        data = json.dumps({
            "messages": msg,
            "model": "openai/gpt-4o-mini-2024-07-18"
        })
    )
    #print(response.json())
    resp = response.json()['choices'][0]['message']['content']
    #print("contentresp",resp)
    #resp = response.json()['choices'][0]['message']
    return resp

def create_presentation(topic, slide_titles, slide_contents):
    prs = pptx.Presentation()
    slide_layout = prs.slide_layouts[1]

    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic

    for slide_title, slide_content in zip(slide_titles, slide_contents):
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = slide_title
        slide.shapes.placeholders[1].text = slide_content
        #slide.shapes.placeholders[1].text = json.dumps(slide_content)

        # Customize font size for titles and content
        slide.shapes.title.text_frame.paragraphs[0].font.size = TITLE_FONT_SIZE
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    paragraph.font.size = SLIDE_FONT_SIZE

    prs.save(f"./{topic}_presentation.pptx")
    
def get_ppt_download_link(topic):
    ppt_filename = f"./{topic}_presentation.pptx"

    with open(ppt_filename, "rb") as file:
        ppt_contents = file.read()

    b64_ppt = base64.b64encode(ppt_contents).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64_ppt}" download="{ppt_filename}">Download the PowerPoint Presentation</a>'

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

def empdf(pdf_file_path):
    with open(pdf_file_path, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800px"></iframe>'
    return pdf_display

def embed_pdf(file):
    base64_pdf = base64.b64encode(file.read()).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800px"></iframe>'
    return pdf_display

def main():
    st.markdown(
        """
        <style>
        .main .block-container {
            width: 100% !important;
            max-width: 100% !important;
            padding-left: 2rem;
            padding-right: 2rem;
        }
        </style>
    """,
    unsafe_allow_html=True
    )
    st.title("PowerPoint Presentation Generator with GPT-3.5-turbo")
    # Create two columns
    col1, col2 = st.columns(2)
    df = None
    pf = None
    md = None
    ppt = None
    with col2:
        topic = st.text_input("Enter the topic for your presentation:")
        preferences = st.text_input("Enter your preferences for the presentation:")
        uploaded_file = st.file_uploader("Choose a file",type=['md','docx','csv','pdf','pptx'])
        chat = st.text_area("Chat with the AI")
        GenerateButton = st.button("Generate Presentation")
        ShareButton = st.button("Share Presentation")
        if uploaded_file is not None:
            st.success(uploaded_file.type)
            if uploaded_file.type == 'text/csv':
                df = pd.read_csv(uploaded_file)
                st.success("csv uploaded successfully")
            elif uploaded_file.type == 'application/octet-stream':
                md = uploaded_file.getvalue().decode("utf-8")
                st.success("md uploaded successfully")
            elif uploaded_file.type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                text = docx2txt.process(uploaded_file)
                st.success("docx uploaded successfully")
            elif uploaded_file.type == 'application/pdf':
                pf = uploaded_file
                st.success("pdf uploaded successfully")
            elif uploaded_file.type == 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
                ppt = uploaded_file
                st.success("pptx uploaded successfully")
            else:
                st.write("File not supported")
    with col1:
        # if upload file is csv 
        if df is not None:
            st.write(df)
        # if upload file is pdf        
        if pf is not None:
            # st.write(pf)
            st.markdown(embed_pdf(pf), unsafe_allow_html=True)
            # images = display_pdf_as_images(pf)
            # for image in images:
            #     st.image(image, use_column_width=True)
        # if upload file is md
        if md is not None:
            st.markdown(md)
        # if upload file is pptx
        if ppt is not None:
            if ppt is not None:
                # Save the uploaded file to a temporary location
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt:
                    tmp_ppt.write(ppt.getbuffer())
                    tmp_ppt_path = tmp_ppt.name
                    tmp_pdf_path = tmp_ppt_path.replace(".pptx", ".pdf")
                # Convert the PowerPoint file to pdf
                ppt_to_pdf(tmp_ppt_path, tmp_pdf_path)
                # Display the PDF file
                st.markdown(empdf(tmp_pdf_path), unsafe_allow_html=True)
                # Clean up the temporary PPTX file
                os.remove(tmp_ppt_path)
    if GenerateButton and topic:
        # st.info("Generating presentation... Please wait.")
        # slide_titles = generate_slide_titles(topic)
        # filtered_slide_titles= [item for item in slide_titles if item.strip() != '']
        # #print("Slide Title: ", filtered_slide_titles)

        # slide_contents = [generate_slide_content(title) for title in filtered_slide_titles]
        # #print("Slide Contents: ", slide_contents)
        # create_presentation(topic, filtered_slide_titles, slide_contents)
        # #print("Presentation generated successfully!")


        st.success("Presentation generated successfully!")
        st.markdown(get_ppt_download_link(topic), unsafe_allow_html=True)
        st.suceess("Outline Preview")
        st.success("Slides Preview")


if __name__ == "__main__":
    main()


