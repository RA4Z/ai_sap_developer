import google.generativeai as genai
from gemini_data import data
from language_translation import Language
import os

genai.configure(api_key=os.environ['GEMINI_API_KEY'])

generation_config = {
    "temperature": 0.9,
    "top_p": 0.95,
    "top_k": 64,
    "max_output_tokens": 8192,
    "response_mime_type": "text/plain",
}

safety_settings = [
    {
        "category": "HARM_CATEGORY_HARASSMENT",
        "threshold": "BLOCK_NONE",
    },
    {
        "category": "HARM_CATEGORY_HATE_SPEECH",
        "threshold": "BLOCK_NONE",
    },
    {
        "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
        "threshold": "BLOCK_NONE",
    },
    {
        "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
        "threshold": "BLOCK_NONE",
    },
]

model = genai.GenerativeModel(
    model_name="gemini-1.5-flash-latest",
    safety_settings=safety_settings,
    generation_config=generation_config
)


def develop_code(prompt: str):
    lang = Language()
    response = model.generate_content([
        f"""Sempre irei criar o código em Python e entregar uma documentação sobre o código em {lang.search('idioma')}. 
  ao terminar de escrever o código eu irei escrever "----fimpython----",
  Todo código que eu criar a primeira linha será: # Default model for SAP automations, developed by Robert Aron 
  Zimmermann, using Google AI Studio tuned prompt model;
  Após escrever todo o código, eu irei escrever uma documentação detalhada sobre o mesmo""",
        "\n".join(data),
        f"input: {prompt}",
        "output: "
    ])
    return response.text
