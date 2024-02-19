import excel_to_insert_statements_tool_engine as engine
import flet as ft
import clipboard as c
from typing import Dict
from flet import (
    Column,
    Row,
    ElevatedButton,
    FilePicker,
    FilePickerResultEvent,
    Page,
    ProgressRing,
    Ref,
    Text,
    icons,
)


def main(page: Page):
    page.window_left = 400
    page.window_top = 200
    page.window_width = 680
    page.window_height = 400
    page.scroll = "adaptive"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.MainAxisAlignment.CENTER
    page.window_resizable = False
    page.title = "Excel to Insert Statements Tool"
    page.update()

    prog_bars: Dict[str, ProgressRing] = {}
    files = Ref[Column]()
    button_select_excel = Ref[ElevatedButton]()
    global is_dropdown_visible
    is_dropdown_visible = False
    global path_to_file
    path_to_file = ""
    output_view = ft.Text("")
    radio_group_sheets = ft.RadioGroup(content=ft.Column([]))

    def get_radio_buttons(list_of_groups):
        group = []
        for item in list_of_groups:
            group.append(ft.Radio(value=item, label=item))
        return group

    snack_bar_text = Text("Copied SQL Insert Statements")

    page.snack_bar = ft.SnackBar(content=snack_bar_text)

    def clear_output():
        output_view.value = ""

    def load_file():
        global path_to_file
        sheets = engine.get_sheets(path_to_file)
        radio_group_sheets.content = ft.Column(get_radio_buttons(sheets))
        radio_group_sheets.value = sheets[0]

    def reload():
        load_file()
        clear_output()
        page.update()

    def file_picker_result(e: FilePickerResultEvent):
        global path_to_file
        button_select_excel.current.disabled = True if e.files is None else False
        prog_bars.clear()
        files.current.controls.clear()
        if e.files is not None:
            for f in e.files:
                prog = ProgressRing(value=0, bgcolor="#eeeeee", width=20, height=20)
                prog_bars[f.name] = prog
                files.current.controls.append(Row([prog, Text(f.name)]))
                path_to_file = f.path
                load_file()
        page.update()

    def process_excel(e):
        print(file_picker.result.files)
        if file_picker.result is not None and file_picker.result.files is not None:
            print("PATH: " + path_to_file)
            output = engine.run(path_to_file, radio_group_sheets.value)
            c.copy(output)
            output_view.value = output 
            page.snack_bar.open = True
            page.snack_bar.content = snack_bar_text
            page.update()

    file_picker = FilePicker(on_result=file_picker_result)

    def select_file():
        file_picker.pick_files(allow_multiple=False)
        clear_output()

    button_select_excel = ElevatedButton(
        "Select Input Excel file...",
        icon=icons.FOLDER_OPEN,
        on_click=lambda _: select_file(),
    )

    button_refresh = ElevatedButton(
        "Reload",
        icon=icons.REFRESH_OUTLINED,
        on_click=lambda _: reload(),
    )

    # hide dialog in a overlay
    page.overlay.append(file_picker)

    page.add(
        Row(
            controls=[
                button_select_excel,
                button_refresh,
                Column(ref=files),
            ]
        ),
        radio_group_sheets,
        ElevatedButton(
            "Get Insert Statements!",
            ref=button_select_excel,
            icon=icons.UPLOAD,
            on_click=process_excel,
            disabled=True,
        ),
        output_view,
    )


ft.app(target=main)
