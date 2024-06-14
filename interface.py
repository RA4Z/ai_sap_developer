import flet as ft
from main import run_automation


def main(page: ft.Page):
    page.title = "IAron - SAP Automation's Developer"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    page.add(
        ft.Text(
            "IAron - SAP Automation's Developer",
            size=24,
            weight=ft.FontWeight.BOLD,
            text_align=ft.TextAlign.CENTER,
        )
    )

    input_container = ft.Container(
        height=575,
        padding=ft.padding.all(15),
        content=ft.TextField(
            multiline=True,
            expand=True,
            border_radius=ft.border_radius.all(5),
            hint_text="Prompt para IAron",
            value="""Título: Python Default Script
Descrição: Default model for SAP automations, developed by Robert Aron Zimmermann, using Google AI Studio tuned prompt model
Solicitado por: Robert Aron Zimmermann
Desenvolvido por: Robert Aron Zimmermann

Observações:
Adicionar Tratativas de erro para evitar que o código trave


Procedimento:

Transação COHV
Escrever Layout "/usin_exce"
flegar "Ordens de produção"
executar
coletar dados da coluna "AUFNR"
"""
        ),
    )

    info_container = ft.Container(
        content=ft.Text(
            expand=True,
            size=20,
            text_align=ft.TextAlign.CENTER,
            value="""
Para executar esse sistema o usuário deve preencher no campo ao lado todo o procedimento que deverá ser realizado pela automatização. 
Após pressionar o botão para criar a automatização o usuário deverá acessar a pasta \"C:\\IAron\"
            """,

        ),
    )

    button = ft.ElevatedButton(
        text="Criar Automatização",
        on_click=lambda e: run_automation(input_container.content.value),
        width=200,  # Define a largura do botão
        height=50,  # Define a altura do botão
    )

    page.add(
        info_container,
        input_container,
        ft.Container(height=10),
        button,
        ft.Container(height=10),
        create_footer()
    )
    page.scroll = "always"
    page.update()


def create_footer():
    footer_panel = ft.Container(
        content=ft.Row(
            [
                ft.Image(src="./images/logo.png", width=100, height=50),  # Substitua por seu caminho da imagem
                ft.Text(
                    "© 2024 PCP WEN - Desenvolvido e Prototipado por Robert Aron Zimmermann.",
                    size=12,
                    weight=ft.FontWeight.NORMAL,
                    text_align=ft.TextAlign.CENTER,
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
        ),
        padding=ft.padding.only(top=10, bottom=10, left=20, right=20),
    )
    return footer_panel


ft.app(target=main)
