import openai
import json
import os
from docx import Document
import win32com.client

def GPT(msg):
    
    with open('./config.json', 'r') as f:
        info = json.load(f)

    
    openai.api_key = info['API_key']
    response = openai.ChatCompletion.create(
        model = "gpt-4",
        messages = msg,
        temperature = 0
    )
    
    res = response["choices"][0]["message"]["content"]
    return res


def Coverletter(company, job, date):

    with open('./config.json', 'r') as f:
        info = json.load(f)

    msg = [
        {"role":"system", "content":info["setting"]},
    ]

    text = f"Date: {date}, Company name: {company}, Job: {job}"

    msg.append({"role": "user", "content": text})

    res = GPT(msg)

    return res


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