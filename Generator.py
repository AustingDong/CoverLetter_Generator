import os
from openai import OpenAI
from openai.types.beta.threads.message_create_params import (
    Attachment,
    AttachmentToolFileSearch,
)
import json
import os
from docx import Document
import win32com.client
from dotenv import load_dotenv


def Coverletter(api_key_var_name, resume_path, job_description_path='./job_description.txt'):

    with open('./config.json', 'r') as f:
        info = json.load(f)

    api_key = os.environ.get(api_key_var_name)
    if api_key is None:
        load_dotenv('/my/envs/.env')
        api_key = os.environ.get(api_key_var_name)
        
    client = OpenAI(api_key = api_key)

    assistant = client.beta.assistants.create(
        model="gpt-4o",
        description=info["setting"],
        tools=[{"type": "file_search"}]
    )

    thread = client.beta.threads.create()
    resume_file = client.files.create(file=open(resume_path, "rb"), purpose="assistants")
    job_description_file = client.files.create(file=open(job_description_path, "rb"), purpose="assistants")

    client.beta.threads.messages.create(
        thread_id=thread.id,
        role="user",
        attachments=[
            Attachment(
                file_id=resume_file.id, tools=[AttachmentToolFileSearch(type="file_search")]
            ),

            Attachment(
                file_id=job_description_file.id, tools=[AttachmentToolFileSearch(type="file_search")]
            ),

        ],
        content=info["setting"],
    )

    run = client.beta.threads.runs.create_and_poll(
        thread_id=thread.id, assistant_id=assistant.id, timeout=1000
    )

    # Error handling
    if run.status != "completed":
        raise Exception("Run failed:", run.status)

    messages_cursor = client.beta.threads.messages.list(thread_id=thread.id)
    messages = [message for message in messages_cursor]

    # Output text
    res_txt = messages[0].content[0].text.value
    return res_txt


def save_to_word(text, filename):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(filename)


def convert_to_pdf(input_path, output_path):

    word_app = win32com.client.Dispatch("Word.Application")

    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    doc = word_app.Documents.Open(input_path)

    doc.SaveAs(output_path, FileFormat=17)

    doc.Close()
    word_app.Quit()
