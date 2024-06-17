import shutil
from docx import Document

def retornar_codigo(texto: str):
    indice_inicio = 0
    if indice_inicio == -1:
        return None

    indice_fim = texto.find("----fimpython----", indice_inicio)
    if indice_fim == -1:
        return None

    return texto[indice_inicio:indice_fim]


def pegar_texto_externo(texto: str):
    indice_fim = texto.find("----fimpython----")
    if indice_fim == -1:
        return None

    return texto[texto.find("\n", indice_fim) + 1:]

def copiar_pasta(origin: str, destiny: str):
    try:
        shutil.copytree(origin, destiny)
        print(f"Pasta copiada com sucesso de '{origin}' para '{destiny}'")
    except FileExistsError:
        print(f"A pasta '{destiny}' já existe. A cópia foi cancelada.")
    except Exception as e:
        print(f"Erro ao copiar pasta: {e}")

def cadastrar_doc(documentation: str, path: str):
    doc = Document()
    doc.add_paragraph(documentation)
    doc.save(f'{path}\\Documentação.docx')
