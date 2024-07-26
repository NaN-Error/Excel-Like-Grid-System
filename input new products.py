import tkinter as tk
from tksheet import Sheet
import pyperclip

class ExcelLikeApp:
    def __init__(self, root):
        self.root = root
        self.selected_cell = None  # Initialize selected cell

        self.create_widgets()
        self.setup_bindings()

    def create_widgets(self):
        # Initialize an empty 10x5 grid with distinct cells
        data = [["" for _ in range(5)] for _ in range(10)]
        self.sheet = Sheet(self.root, width=800, height=600, data=data)
        self.sheet.grid(row=0, column=0, sticky="nswe")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.sheet.headers(["Product Image", "Product Name", "Order Date", "Fair Market Value", "Order details"])
        self.sheet.set_column_widths([100, 200, 100, 100, 100])
        self.sheet.enable_bindings((
            "single_select", 
            "column_select", 
            "row_select", 
            "toggle_select",
            "drag_select", 
            "column_drag_and_drop", 
            "row_drag_and_drop",
            "row_height_resize",
            "column_width_resize",
            "double_click_column_resize",
            "double_click_row_resize",
            "right_click_popup_menu",
            "rc_select",
            "rc_insert_column",
            "rc_delete_column",
            "rc_insert_row",
            "rc_delete_row",
            "copy",
            "cut",
            "paste",
            "delete",
            "undo",
            "redo",
            "edit_cell"))

    def setup_bindings(self):
        self.sheet.extra_bindings([("cell_select", self.handle_select)])
        self.root.bind('<Control-v>', self.handle_paste)

    def handle_select(self, event):
        selection = self.sheet.get_currently_selected()
        if selection and isinstance(selection, tuple):
            self.selected_cell = (selection[0], selection[1])
        print(f"Selected cell: {self.selected_cell}")

    def preprocess_clipboard_data(self, data):
        # Replace newline followed by "Order details" with tab and "Order details"
        data = data.replace('\nOrder details', '\tOrder details').strip()
        # Ensure there are no carriage return characters
        data = data.replace('\r', '')
        return data

    def parse_text_content(self, text):
        # Split the text into rows based on the newline character
        rows = text.split('\n')
        content = []

        for row in rows:
            if row.strip():  # Skip empty rows
                columns = row.split('\t')
                # Ensure the row has exactly 5 columns
                if len(columns) == 4:
                    columns = [''] + columns
                elif len(columns) > 5:
                    columns = columns[:4] + [' '.join(columns[4:])]
                content.append(columns)
        
        print("Parsed Content:", content)  # Debugging: print parsed content
        return content

    def display_content(self, content, start_row, start_col):
        for row_offset, row_data in enumerate(content):
            for col_offset, cell in enumerate(row_data):
                row = start_row + row_offset
                col = start_col + col_offset
                self.sheet.set_cell_data(row, col, cell)

    def handle_paste(self, event=None):
        clipboard_data = pyperclip.paste()
        print("Original Clipboard Data:", clipboard_data)  # Debugging: print original clipboard data
        processed_data = self.preprocess_clipboard_data(clipboard_data)
        content = self.parse_text_content(processed_data)
        print("Content to Display:", content)  # Debugging: print content to be displayed
        start_row, start_col = self.selected_cell if self.selected_cell else (0, 0)
        self.display_content(content, start_row, start_col)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel-like Grid System")
    app = ExcelLikeApp(root)
    root.mainloop()
