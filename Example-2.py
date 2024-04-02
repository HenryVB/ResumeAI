import json
import openai

from docx import Document
import docx2txt as d2t_reader
import fitz as pdf_reader  # PyMuPDF
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PIL import Image
from io import open

from dotenv import dotenv_values


def extract_resume_word(file):
    text = d2t_reader.process(file)
    return text


def extract_resume_pdf(file):
    pdf = pdf_reader.open(file)
    text = "\n".join([page.get_text() for page in pdf])
    return text


def extract_text_from_resume(file):
    if file.endswith(".docx"):
        return extract_resume_word(file)
    elif file.endswith(".pdf"):
        return extract_resume_pdf(file)
    else:
        raise Exception('Invalid resume format. Please upload a valid format')


def add_header_with_image(document, image_path, text):
    header_section = document.sections[0]
    header = header_section.header
    header_text = header.paragraphs[0]
    header_text.text = text
    header_text.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Add a blank paragraph at the beginning to insert header
    # document.add_paragraph()

    # Add image to header
    # header_paragraph = document.paragraphs[0]
    # run = header_paragraph.add_run()
    # run.add_picture(image_path, width=Inches(6))

    # Set text properties
    # header_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    # header_paragraph.paragraph_format.space_after = Inches(0.2)

    # Add text to header
    # header_text = document.add_paragraph(text)
    # header_text.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT


def add_heading(document, text):
    document.add_heading(text, level=1)


def add_paragraph(document, text):
    document.add_paragraph(text)


def add_bullet_list(document, items):
    for item in items:
        document.add_paragraph(item, style='List Bullet')


def add_experience(document, experiences):
    for exp in experiences:
        document.add_paragraph(exp['title'], style='Heading 2')
        # document.add_paragraph(f"{exp['company']}\n{exp['duration']}") puede reemplazarse por duration = datestar y dateend juntos
        document.add_paragraph(f"{exp['company']}\n{exp['date_start']} - {exp['date_end']}")
        document.add_paragraph(exp['description'])


def create_document(data):
    doc = Document()

    add_header_with_image(doc, 'Template/soaint-cv-header.png',
                          f"{data["name"]}\n{data["nationality"]}\n{data["address"]}")

    # Add content
    add_heading(doc, "RESUMEN")
    add_paragraph(doc, data["summary"])

    add_heading(doc, "FORMACIÓN")
    add_bullet_list(doc, data["education"])

    add_heading(doc, "CURSOS Y CERTIFICACIONES")
    add_bullet_list(doc, data["certification"])

    add_heading(doc, "TECNOLOGÍAS")
    add_bullet_list(doc, data["skill"])

    add_heading(doc, "EXPERIENCIA")
    add_experience(doc, data["experience"])

    doc.save('Template/resume.docx')


def convert_resume_info(data):
    config = dotenv_values(".env")
    openai.api_key = config["OPENAI_API_KEY"]

    messages = [
        {
            "role": "system",
            "content": """You will be provided with a resume information of a candidate in large text.
       Your task is to find and extract the required data from resume information and then return a JSON format.
       Here an example for desired JSON Format and explanation of each field:
        {
            "name": <Name and last name of candidate>,
            "nationality": <Nationality or Country of candidate>,
            "address": <Address of candidate>,
            "summary": <Abstract/About me/Profile paragraph of the candidate>,
            "education": <List of degree obtained in schools 
                        (Ex: Bachelor of Science in Computer Science, University of Technology, 2015-2018)>,
            "certification": <List of certificates obtained, courses of the candidate>
            "skill": <List of hard or soft skills/tools of the candidate 
                      (Ex: programming languages, methodologies, tools)>
            "experience": <List of work experience in companies of the candidate. Include actual position>
                        Consider these fields in each element of experience: 
                        (title: <Position in the Job>, company: <Name of the Company>, 
                        date_start: <start date in Month-Year format>, date end: <start date in Month-Year format>,
                        description: <Description of the job/ position, participated or in course projects and achievements>). 
        }
        In case information is not found for any field of the JSON output, set value to "N/A".     
       """,
        },
        {
            "role": "user",
            "content": data,
        }
    ]

    response = openai.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=messages
    )

    print("*********resultado extraccion**********")
    print(response)

    resume_info = json.loads(response.choices[0].message.content)
    return resume_info


def get_summary_ai(data):
    config = dotenv_values(".env")
    openai.api_key = config["OPENAI_API_KEY"]

    messages = [
        {
            "role": "system",
            "content": f"""You will be provided with a resume information of a candidate in large text or json format.
        Your task is to summarize the resume highlighting critical information about specific requirements of positions or projects.
        Your will return the summary in spanish. 
        """,
        },
        {
            "role": "user",
            "content": data,
        }
    ]

    response = openai.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=messages
    )
    print("*********resultado summary**********")
    print(response)
    print(response.choices[0].message.content)


def get_tech_experience_insight(data):
    return "Insight 1"


def get_featured_clients_projects_insight(data):
    return "Insight 2"


def get_competencies_skills_analysis_insight(data):
    return "Insight 3"


def get_sectorial_experience_insight(data):
    return "Insight 4"


def main():
    resume_file = 'CV/Harold_Portillo_945962801.pdf'
    print("***Lectura CV***")
    full_resume_text = extract_text_from_resume(resume_file)
    print(full_resume_text.strip())
    print("***Conversion a JSON***")
    json_data = convert_resume_info(full_resume_text)
    print(json_data)
    print("***Creacion de word plantilla***")
    create_document(json_data)
    print("***Resumen con OpenAI***")
    get_summary_ai(full_resume_text)


if __name__ == "__main__":
    main()
