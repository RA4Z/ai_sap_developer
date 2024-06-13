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
