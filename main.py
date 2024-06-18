import flet as ft
from gemini import run_automation
from language_translation import Language

lang = Language()


def change_language(language):
    global lang
    with open('languages/selected.txt', 'w', encoding='utf-8') as file:
        file.write(language)
    lang = Language()


def main(page: ft.Page):
    page.title = lang.search('title')
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    main_title = ft.Text(
        lang.search('title'),
        size=24,
        weight=ft.FontWeight.BOLD,
        text_align=ft.TextAlign.CENTER,
    )
    page.add(
        main_title
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

    button = ft.ElevatedButton(
        text=lang.search('create'),
        icon=ft.icons.CODE_SHARP,
        on_click=lambda e: executar(),
        width=300,  # Define a largura do botão
        height=50,  # Define a altura do botão
    )

    def dropdown_changed(e):
        change_language(lang_selector.value)
        page.title = lang.search('title')
        main_title.value = lang.search('title')
        info_container.content.value = lang.search('exec_title')
        button.text = lang.search('create')
        input_container.content.hint_text = lang.search('hint_text')
        input_container.content.value = lang.search('prompt_model')
        footer_panel.content.controls[1].value = lang.search('creator')
        page.update()

    lang_selector = ft.Dropdown(
        on_change=dropdown_changed,
        hint_text=lang.search('select_language'),
        options=[
            ft.dropdown.Option("EN"),
            ft.dropdown.Option("PT"),
            ft.dropdown.Option("ES"),
            ft.dropdown.Option("FR"),
            ft.dropdown.Option("DE"),
        ],
        width=200,
    )

    footer_panel = ft.Container(
        content=ft.Row(
            controls=[
                ft.Image(src="./images/logo.png", width=100, height=50),
                ft.Text(
                    value=lang.search('creator'),
                    size=12,
                    weight=ft.FontWeight.NORMAL,
                    text_align=ft.TextAlign.CENTER,
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
        ),
        padding=ft.padding.only(top=10, bottom=10, left=20, right=20),
    )

    page.add(
        info_container,
        lang_selector,
        input_container,
        ft.Container(height=10),
        button,
        ft.Container(height=10),
        footer_panel
    )
    page.scroll = "always"
    page.update()


if __name__ == '__main__':
    ft.app(target=main)
