import shutil


def retornar_codigo(texto: str):
    indice_inicio = texto.find("```python")
    if indice_inicio == -1:
        return None

    indice_inicio = texto.find("\n", indice_inicio) + 1
    if indice_inicio == 0:  # Caso "```python" esteja na última linha
        return None

    indice_fim = texto.find("```", indice_inicio)
    if indice_fim == -1:
        return None

    return texto[indice_inicio:indice_fim]


def pegar_texto_externo(texto: str):
    indice_inicio = texto.find("```python")
    if indice_inicio == -1:
        return None

    indice_inicio = texto.find("\n", indice_inicio) + 1
    if indice_inicio == 0:  # Caso "```python" esteja na última linha
        return None

    indice_fim = texto.find("```", indice_inicio)
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
