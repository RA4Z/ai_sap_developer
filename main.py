from gemini_config import develop_code
from functions import retornar_codigo

if __name__ == '__main__':
    codigo = develop_code("COHV\nvariante /PS\nexecutar\narmazenar todos os dados da coluna \"PROJN\"\n\nCN47N\ncolar todos os dados salvos no campo \"Elemento PEP\"\nescrever \"/INI_MONT\" no campo \"Layout\"\nexecutar\nescrever dados da tabela na planilha chamada \"Esquema.xlsm\"")
    with open('script/main.py', 'w', encoding="UTF-8") as f:
        f.write(f"{retornar_codigo(codigo)}")
