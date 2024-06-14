from gemini_config import develop_code
from functions import retornar_codigo, pegar_texto_externo, copiar_pasta, cadastrar_doc

if __name__ == '__main__':
    output = develop_code("COHV\nvariante /PS\nexecutar\narmazenar todos os dados da coluna \"PROJN\"\n\nCN47N\ncolar todos os dados salvos no campo \"Elemento PEP\"\nescrever \"/INI_MONT\" no campo \"Layout\"\nexecutar\nescrever dados da tabela na planilha chamada \"Esquema.xlsm\"")
    codigo = retornar_codigo(output)
    documentation = pegar_texto_externo(output)

    origem = "script"
    destino = "C:\\IAron\\script"
    copiar_pasta(origem, destino)

    open('C:\\IAron\\script\\main.py', 'w', encoding="UTF-8").write(f"{codigo}")
    cadastrar_doc(documentation, destino)

    print('Código Criado')
