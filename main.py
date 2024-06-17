import flet as ft
from gemini import run_automation
from language_translation import Language
lang = Language()

def main(page: ft.Page):
    page.title = lang.search('title')
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    page.add(
        ft.Text(
            lang.search('title'),
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
            hint_text=lang.search('hint_text'),
            value=lang.search('prompt_model')
        ),
    )

    info_container = ft.Container(
        content=ft.Text(
            expand=True,
            size=20,
            text_align=ft.TextAlign.CENTER,
            value=lang.search('exec_title'),
        ),
    )

    button = ft.ElevatedButton(
        text=lang.search('create'),
        on_click=lambda e: executar(),
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

    def executar():
        status = run_automation(input_container.content.value)
        if not status:
            dlg = ft.AlertDialog(
                title=ft.Text(lang.search('error'), color=ft.colors.RED),
                on_dismiss=lambda e: print("Dialog dismissed!")
            )
        else:
            dlg = ft.AlertDialog(
                title=ft.Text(lang.search('success')),
                on_dismiss=lambda e: print("Dialog dismissed!")
            )

        page.dialog = dlg
        dlg.open = True
        page.update()


def create_footer():
    footer_panel = ft.Container(
        content=ft.Row(
            [
                ft.Image(src="./images/logo.png", width=100, height=50),  # Substitua por seu caminho da imagem
                ft.Text(
                    lang.search('creator'),
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


if __name__ == '__main__':
    ft.app(target=main)
