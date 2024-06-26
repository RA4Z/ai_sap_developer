from gemini_config import develop_code
from functions import retornar_codigo, pegar_texto_externo, copiar_pasta, cadastrar_doc, cadastrar_worksheet
import time
import subprocess


def run_automation(user_input: str):
    output = develop_code(user_input)
    if output is not None:
        try:
            codigo = retornar_codigo(output)
            documentation = pegar_texto_externo(output)
            file = f"script {time.strftime('%d_%m_%y_%H-%M-%S', time.localtime())}"

            origem = "script"
            destino = f"C:\\IAron\\{file}"
            copiar_pasta(origem, destino)

            open(f'C:\\IAron\\{file}\\main.py', 'w', encoding="UTF-8").write(f"{codigo}")
            cadastrar_doc(documentation, destino)
            cadastrar_worksheet(output, destino)
            subprocess.run(['explorer', destino])
            return True
        except Exception as e:
            print(e)
            return False

    else:
        return False


if __name__ == '__main__':
    run_automation('Teste')
    print('Código Criado')
