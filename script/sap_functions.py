import win32com.client
from tkinter import messagebox
import re
import os
import time
import subprocess
import pyautogui
from language_dict import Language


# SAP Scripting Documentation:
# https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/a2e9357389334dc89eecc1fb13999ee3.html

# module SAP Functions, development started in 2024/03/01
class SAP:

    # Initializes the SAP object with a specified window index.
    def __init__(self, window: int, scheduled_execution, language: str):
        self.side_index = None
        self.desired_operator = None
        self.desired_text = None
        self.field_name = None
        self.target_index = None
        self.scheduled_execution = scheduled_execution
        self.language = Language(language)
        self.idiom = language
        self.connection = self.__verify_sap_open()

        if self.connection.Children(0).info.user == '':
            messagebox.showerror(title=self.language.search('sap_logon_err_title'),
                                 message=self.language.search('sap_logon_err_body'))
            exit()

        if self.connection.Children(0).info.systemName == 'EQ0':
            print(self.language.search('sap_system_err_body'))

        if self.connection.Children(0).info.language != self.idiom:
            print(self.language.search('sap_language_err_body').replace('$language', self.idiom))

        self.__count_sap_screens(window)
        self.session = self.connection.Children(window)
        self.window = self.__active_window()

    # Verify if SAP is open
    def __verify_sap_open(self):
        try:
            sapguiauto = win32com.client.GetObject('SAPGUI')
            application = sapguiauto.GetScriptingEngine
            return application.Children(0)
        except:
            if self.scheduled_execution['scheduled?']:
                return self.__open_sap()
            else:
                messagebox.showerror(title=self.language.search('sap_open_err_title'),
                                     message=self.language.search('sap_open_err_body'))
                exit()

    def __open_sap(self):
        path = "C:/Program Files (x86)/SAP/FrontEnd/SapGui/saplgpad.exe"
        subprocess.Popen(path)
        while not pyautogui.getActiveWindowTitle().startswith("SAP Logon"):
            time.sleep(1)

        sapguiauto = win32com.client.GetObject('SAPGUI')
        application = sapguiauto.GetScriptingEngine
        connection = application.OpenConnection("EP0 - ECC Produção", True)
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = self.scheduled_execution['principal']
        session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = self.scheduled_execution['username']
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = self.scheduled_execution['password']
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = self.idiom
        session.findById("wnd[0]").sendVKey(0)

        if session.activewindow.name == 'wnd[1]':
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

        return application.Children(0)

    # Count the number of open SAP screens
    def __count_sap_screens(self, window: int):
        while len(self.connection.sessions) < window + 1:
            self.connection.Children(0).createSession()
            time.sleep(3)

    # Finds the active window index.
    def __active_window(self):
        regex = re.compile('[0-9]')
        matches = regex.findall(self.session.ActiveWindow.name)
        for match in matches:
            return match

    # Scrolls through tabs within a specified area.
    def __scroll_through_tabs(self, area, extension, selected_tab):
        children = area.Children
        for child in children:
            if child.Type == "GuiTabStrip":
                extension = extension + "/tabs" + child.name
                return self.__scroll_through_tabs(self.session.findById(extension), extension, selected_tab)
            if child.Type == "GuiTab":
                extension = extension + "/tabp" + str(children[selected_tab].name)
                return self.__scroll_through_tabs(self.session.findById(extension), extension, selected_tab)
            if child.Type == "GuiSimpleContainer":
                extension = extension + "/sub" + child.name
                return self.__scroll_through_tabs(self.session.findById(extension), extension, selected_tab)
            if child.Type == "GuiScrollContainer" and 'tabp' in extension:
                extension = extension + "/ssub" + child.name
                area = self.session.findById(extension)
                return area
        return area

    # Scrolls through a grid based on its extension.
    def __scroll_through_grid(self, extension):
        if self.session.findById(extension).Type == 'GuiShell':
            try:
                var = self.session.findById(extension).RowCount
                return self.session.findById(extension)
            except:
                pass
        children = self.session.findById(extension).Children
        result = False
        for i in range(len(children)):
            if result:
                break
            if children[i].Type == 'GuiCustomControl':
                result = self.__scroll_through_grid(extension + '/cntl' + children[i].name)
            if children[i].Type == 'GuiSimpleContainer':
                result = self.__scroll_through_grid(extension + '/sub' + children[i].name)
            if children[i].Type == 'GuiScrollContainer':
                result = self.__scroll_through_grid(extension + '/ssub' + children[i].name)
            if children[i].Type == 'GuiTableControl':
                result = self.__scroll_through_grid(extension + '/tbl' + children[i].name)
            if children[i].Type == 'GuiTab':
                result = self.__scroll_through_grid(extension + '/tabp' + children[i].name)
            if children[i].Type == 'GuiTabStrip':
                result = self.__scroll_through_grid(extension + '/tabs' + children[i].name)
            if children[
                i].Type in ("GuiShell GuiSplitterShell GuiContainerShell GuiDockShell GuiMenuBar GuiToolbar "
                            "GuiUserArea GuiTitlebar"):
                result = self.__scroll_through_grid(extension + '/' + children[i].name)
        return result

    # Scrolls through a table based on its extension.
    def __scroll_through_table(self, extension):
        if 'tbl' in extension:
            try:
                return self.session.findById(extension)
            except:
                pass
        children = self.session.findById(extension).Children
        result = False
        for i in range(len(children)):
            if result:
                break
            if children[i].Type == 'GuiCustomControl':
                result = self.__scroll_through_table(extension + '/cntl' + children[i].name)
            if children[i].Type == 'GuiSimpleContainer':
                result = self.__scroll_through_table(extension + '/sub' + children[i].name)
            if children[i].Type == 'GuiScrollContainer':
                result = self.__scroll_through_table(extension + '/ssub' + children[i].name)
            if children[i].Type == 'GuiTableControl':
                result = self.__scroll_through_table(extension + '/tbl' + children[i].name)
            if children[i].Type == 'GuiTab':
                result = self.__scroll_through_table(extension + '/tabp' + children[i].name)
            if children[i].Type == 'GuiTabStrip':
                result = self.__scroll_through_table(extension + '/tabs' + children[i].name)
            if children[
                i].Type in ("GuiShell GuiSplitterShell GuiContainerShell GuiDockShell GuiMenuBar GuiToolbar "
                            "GuiUserArea GuiTitlebar"):
                result = self.__scroll_through_table(extension + '/' + children[i].name)
        return result

    # Scrolls through fields within a specified extension and objective.
    def __scroll_through_fields(self, extension, objective, selected_tab):
        children = self.session.findById(extension).Children
        result = False
        for i in range(len(children)):
            if not result:
                result = self.__generic_conditionals(i, children, objective)
            if result:
                break
            if not result and children[i].Type == "GuiTabStrip" and 'ssub' not in extension:
                result = self.__scroll_through_fields(extension + "/tabs" + children[i].name, objective, selected_tab)
            if not result and children[i].Type == "GuiTab" and 'tabp' not in extension:
                result = self.__scroll_through_fields(extension + "/tabp" + str(children[selected_tab].name), objective,
                                                      selected_tab)
            if not result and children[i].Type == "GuiSimpleContainer":
                result = self.__scroll_through_fields(extension + "/sub" + children[i].name, objective, selected_tab)
            if not result and children[i].Type == "GuiScrollContainer":
                result = self.__scroll_through_fields(extension + "/ssub" + children[i].name, objective, selected_tab)
            if not result and children[i].Type == "GuiCustomControl":
                result = self.__scroll_through_fields(extension + "/cntl" + children[i].name, objective, selected_tab)
            if not result and children[
                i].Type in ("GuiShell GuiSplitterShell GuiContainerShell GuiDockShell GuiMenuBar GuiToolbar "
                            "GuiUserArea GuiTitlebar"):
                result = self.__scroll_through_fields(extension + "/" + children[i].name, objective, selected_tab)
        return result

    # Contains generic conditional statements for different objectives.
    def __generic_conditionals(self, index, children, objective):
        if objective == 'write_text_field':
            if children(index).Text == self.field_name:
                if self.target_index == 0:
                    try:
                        children(index + 1).Text = self.desired_text
                        return True
                    except Exception as e:
                        print(self.language.search('sap_error').replace('$error', str(e)))
                    return
                else:
                    self.target_index -= 1

        if objective == 'write_text_field_until':
            if children(index).Text == self.field_name:
                if self.target_index == 0:
                    try:
                        children(index + 3).Text = self.desired_text
                        return True
                    except Exception as e:
                        print(self.language.search('sap_error').replace('$error', str(e)))
                    return
                else:
                    self.target_index -= 1

        if objective == 'find_text_field':
            if self.field_name in children(index).Text:
                try:
                    return True
                except Exception as e:
                    print(self.language.search('sap_error').replace('$error', str(e)))
                return

        if objective == 'multiple_selection_field':
            if children(index).Text == self.field_name:
                if self.target_index == 0:
                    try:
                        campo = children(index).name
                        posicaoInicial = campo.find("%") + 1
                        posicaoFinal = campo.find("-", posicaoInicial)
                        campo = campo[posicaoInicial:posicaoFinal] + "-VALU_PUSH"
                        for j in range(index, len(children)):
                            Obj = children[j]
                            if campo in Obj.name:
                                Obj.press()
                                return True
                    except Exception as e:
                        print(self.language.search('sap_error').replace('$error', str(e)))
                    return
                else:
                    self.target_index -= 1

        if objective == 'flag_field':
            if children(index).Text == self.field_name:
                if self.target_index == 0:
                    try:
                        children(index).Selected = self.desired_operator
                        return True
                    except Exception as e:
                        print(self.language.search('sap_error').replace('$error', str(e)))
                    return
                else:
                    self.target_index -= 1

        if objective == 'flag_field_at_side':
            if children(index).Text == self.field_name:
                if self.target_index == 0:
                    try:
                        children(index + self.side_index).Selected = self.desired_operator
                        return True
                    except Exception as e:
                        print(self.language.search('sap_error').replace('$error', str(e)))
                    return
                else:
                    self.target_index -= 1

        if objective == 'option_field':
            if children(index).Text == self.field_name:
                if self.target_index == 0:
                    try:
                        children(index).Select()
                        return True
                    except Exception as e:
                        print(self.language.search('sap_error').replace('$error', str(e)))
                    return
                else:
                    self.target_index -= 1

        if objective == 'press_button':
            try:
                if self.field_name in children(index).Text or self.field_name in children(index).Tooltip:
                    children(index).press()
                    return True
                if self.session.info.transaction == 'CJ20N' or self.session.info.transaction == 'MD04':
                    try:
                        for i in range(101):
                            if children(index).GetButtonTooltip(i) != '':
                                id_button = children(index).GetButtonId(i)
                                tooltip_button = children(index).GetButtonTooltip(i)
                                if self.field_name in tooltip_button:
                                    children(index).pressButton(id_button)
                    except:
                        pass
            except Exception as e:
                if str(e) != 'index out of range':
                    print(self.language.search('sap_error').replace('$error', str(e)))
            return

        if objective == 'choose_text_combo':
            if children(index).Text == self.field_name:
                if self.target_index == 0:
                    try:
                        entries = children(index + 1).Entries
                        for cont in range(len(entries)):
                            entry = entries.Item(cont)
                            if self.desired_text == str(entry.Value):
                                children(index + 1).key = entry.key
                                return True
                    except Exception as e:
                        print(self.language.search('sap_error').replace('$error', str(e)))
                    return

        if objective == 'get_text_at_side':
            if children(index).Text == self.field_name:
                if self.target_index == 0:
                    try:
                        self.found_text = children(index + self.side_index).Text
                        return True
                    except Exception as e:
                        print(self.language.search('sap_error').replace('$error', str(e)))
                    return

        return False

    # Selects a transaction within the SAP session.
    def select_transaction(self, transaction: str):
        self.session.startTransaction(transaction)
        if self.session.activeWindow.name == 'wnd[1]' and 'CN' in transaction:
            self.session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        if not self.session.info.transaction == transaction:
            messagebox.showerror(title=self.language.search('sap_transaction_err_title'),
                                 message=self.get_footer_message())
            exit()

    # Selects the main screen of the SAP session.
    def select_main_screen(self):
        if not self.session.info.transaction == "SESSION_MANAGER":
            self.session.startTransaction('SESSION_MANAGER')
            if self.session.ActiveWindow.name == "wnd[1]":
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Cleans all fields within the SAP session.
    def clean_all_fields(self, selected_tab=0):
        self.window = self.__active_window()
        area = self.__scroll_through_tabs(self.session.findById(f"wnd[{self.window}]/usr"), f"wnd[{self.window}]/usr",
                                          selected_tab)
        children = area.Children
        for child in children:
            if child.Type == "GuiCTextField":
                try:
                    child.Text = ""
                except Exception as e:
                    print(self.language.search('sap_error').replace('$error', str(e)))

    # Run the active transaction in the SAP screen
    def run_actual_transaction(self):
        self.window = self.__active_window()
        screen_title = self.session.activeWindow.text
        self.session.findById(f'wnd[{self.window}]').sendVKey(0)
        try:
            if screen_title == self.session.activeWindow.text: self.session.findById(f'wnd[{self.window}]').sendVKey(8)
        except:
            pass

    # Inserts a variant into the SAP session.
    def insert_variant(self, variant_name: str):
        try:
            self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
            if self.session.activeWindow.name == 'wnd[1]':
                self.session.findById("wnd[1]/usr/txtV-LOW").Text = variant_name
                self.session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
                self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                if self.session.activewindow.name == 'wnd[1]':
                    pass
        except Exception as e:
            print(self.language.search('sap_error').replace('$error', str(e)))

    # Changes the active tab within the SAP session.
    def change_active_tab(self, selected_tab: int):
        self.window = self.__active_window()
        area = self.__scroll_through_tabs(self.session.findById(f"wnd[{self.window}]/usr"), f"wnd[{self.window}]/usr",
                                          selected_tab)
        try:
            area.Select()
        except Exception as e:
            print(self.language.search('sap_error').replace('$error', str(e)))
        return

    # Writes text into a text field within the SAP session.
    def write_text_field(self, field_name: str, desired_text: str, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.desired_text = desired_text
        self.target_index = target_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]/usr", 'write_text_field', selected_tab)

    # Writes text into a text field until a certain index within the SAP session.
    def write_text_field_until(self, field_name: str, desired_text: str, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.desired_text = desired_text
        self.target_index = target_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]/usr", 'write_text_field_until', selected_tab)

    # Choose an option inside a SAP combo box
    def choose_text_combo(self, field_name: str, desired_text: str, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.desired_text = desired_text
        self.target_index = target_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]/usr", 'choose_text_combo', selected_tab)

    # Flags a field within the SAP session.
    def flag_field(self, field_name: str, desired_operator: bool, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.desired_operator = desired_operator
        self.target_index = target_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]/usr", 'flag_field', selected_tab)

    # Flags a field within the SAP session.
    def flag_field_at_side(self, field_name: str, desired_operator: bool, side_index=0, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.desired_operator = desired_operator
        self.target_index = target_index
        self.side_index = side_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]/usr", 'flag_field_at_side', selected_tab)

    # Selects an option within a field in the SAP session.
    def option_field(self, field_name: str, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.target_index = target_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]/usr", 'option_field', selected_tab)

    # Presses a button within the SAP session.
    def press_button(self, field_name: str, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.target_index = target_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]", 'press_button', selected_tab)

    # Finds a text field within the SAP session.
    def find_text_field(self, field_name: str, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]/usr", 'find_text_field', selected_tab)

    # Get the text that is at the side of the field_name.
    def get_text_at_side(self, field_name, side_index: int, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.target_index = target_index
        self.side_index = side_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        if self.__scroll_through_fields(f"wnd[{self.window}]", 'get_text_at_side', selected_tab):
            return self.found_text

    # Selects multiple entries within a field in the SAP session.
    def multiple_selection_field(self, field_name: str, target_index=0, selected_tab=0):
        self.window = self.__active_window()
        self.field_name = field_name
        self.target_index = target_index
        if selected_tab > 0:
            self.change_active_tab(selected_tab)
        return self.__scroll_through_fields(f"wnd[{self.window}]/usr", 'multiple_selection_field', selected_tab)

    # Pastes data into multiple selection fields in the SAP session.
    def multiple_selection_paste_data(self, data: str):
        try:
            with open('C:/Temp/temp_paste.txt', 'w') as arquivo:
                arquivo.write(data)
            self.session.findById("wnd[1]/tbar[0]/btn[23]").press()
            self.session.findById("wnd[2]/usr/ctxtDY_PATH").text = 'C:/Temp'
            self.session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "temp_paste.txt"
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
            if os.path.exists('C:/Temp/temp_paste.txt'):
                os.remove('C:/Temp/temp_paste.txt')
        except Exception as e:
            print(self.language.search('sap_error').replace('$error', str(e)))

    # Navigate around the menu in the SAP header
    def navigate_into_menu_header(self, path: str):
        try:
            id_path = 'wnd[0]/mbar'
            if ';' not in path:
                messagebox.showerror(title=self.language.search('sap_menu_err_title'),
                                     message=self.language.search('sap_menu_err_body'))
                exit()

            list_of_paths = path.split(';')
            for active_path in list_of_paths:
                children = self.session.findById(id_path).Children
                for i in range(children.Count):
                    Obj = children(i)
                    if active_path in Obj.Text:
                        menu_address = Obj.ID.split("/")[-1]
                        id_path += f'/{menu_address}'
                        break
            self.session.findById(id_path).Select()
        except Exception as e:
            print(self.language.search('sap_error').replace('$error', str(e)))

    # Saves a file in the SAP session.
    def save_file(self, file_name: str, path: str, option=0, type_of_file='txt'):
        if 'xls' in type_of_file:
            self.session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").Select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        else:
            self.session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
            self.session.findById(
                f"wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[{option},0]").Select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

        self.session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = f'{file_name}.{type_of_file}'
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Views data in list form within the SAP session.
    def view_in_list_form(self):
        myGrid = self.get_my_grid()
        myGrid.pressToolbarContextButton("&MB_VIEW")
        myGrid.SelectContextMenuItem("&PRINT_BACK_PREVIEW")

    # Retrieves the table object within the SAP session.
    def get_my_table(self):
        self.window = self.__active_window()
        return self.__scroll_through_table(f'wnd[{self.window}]/usr')

    # Retrieves a value from a cell
    def my_table_get_cell_value(self, my_table, row_index: int, column_index: int):
        try:
            return my_table.getCell(row_index, column_index).Text
        except:
            return 'Empty'

    # my_table tips:
    # VisibleRowCount => Count the number of Visible Rows in the table
    # RowCount => Count the number of Rows inside the table

    # Retrieves the grid object within the SAP session.
    def get_my_grid(self):
        self.window = self.__active_window()
        return self.__scroll_through_grid(f'wnd[{self.window}]/usr')

    # Select a Layout after accesses the table
    def my_grid_select_layout(self, layout: str):
        my_grid = self.get_my_grid()
        my_grid.selectColumn("VARIANT")
        my_grid.contextMenu()
        my_grid.selectContextMenuItem("&FILTER")
        self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = layout
        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
        self.session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        self.session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Count the total number of rows inside the Grid
    def get_my_grid_count_rows(self, my_grid):
        self.window = self.__active_window()
        rows = my_grid.RowCount
        if rows > 0:
            visiblerow = my_grid.VisibleRowCount
            visiblerow0 = my_grid.VisibleRowCount
            npagedown = rows // visiblerow0
            if npagedown > 1:
                for j in range(1, npagedown + 1):
                    try:
                        my_grid.firstVisibleRow = visiblerow - 1
                        visiblerow += visiblerow0
                    except:
                        break
            my_grid.firstVisibleRow = 0
        return rows

    # Retrieves the footer message within the SAP session.
    def get_footer_message(self):
        return self.session.findById("wnd[0]/sbar").Text
