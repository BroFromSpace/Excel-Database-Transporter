import re
import tkinter as tk
from tkinter import messagebox as mb
from typing import NoReturn, List

from main import get_part_list, copy_part


class App:
    def __init__(self):
        self.part_list = get_part_list()
        self.part_names_list = [product[0].lower().strip() for product in self.part_list]

        self.root = tk.Tk()
        self.root.title("xlsx app")
        self.root.geometry("480x360")

        top_frame = tk.Frame(self.root, width=480, height=100)  # top frame for additional information
        top_frame.pack(fil=tk.Y, expand=True)

        main_frame = tk.Frame(self.root, width=480, height=160)  # main frame for input fields
        main_frame.pack(fill=tk.Y, expand=True)

        bottom_frame = tk.Frame(self.root, width=480, height=100)  # bottom frame for submit and exit buttons
        bottom_frame.pack(fill=tk.Y, expand=True)

        # Top Frame content
        app_info_label = tk.Label(
            top_frame,
            text="Enter product name and click submit button,\n or if you want to exit application press exit button",
            font=20,
            justify=tk.CENTER,
            padx=10,
            pady=10
        )
        app_info_label.grid()

        # Main Frame content
        part_name_label = tk.Label(
            main_frame,
            text="Enter product name",
            font=20,
            justify=tk.CENTER,
            padx=10,
            pady=10
        )
        part_name_label.grid()

        self.part_name_entry = tk.Entry(
            main_frame,
            font=20,
            width=35,
        )
        self.part_name_entry.grid()

        part_count_label = tk.Label(
            main_frame,
            text="Enter number of parts you want to copy",
            font=20,
            justify=tk.CENTER,
            padx=10,
            pady=10
        )
        part_count_label.grid()

        self.part_count_entry = tk.Entry(
            main_frame,
            font=20,
            width=15
        )
        self.part_count_entry.grid()

        # Bottom Frame content
        exit_btn = tk.Button(
            bottom_frame,
            text="Exit",
            font=20,
            width=10,
            height=1,
            command=lambda: self.exit()
        )
        exit_btn.grid(row=2, column=0, padx=15)

        submit_btn = tk.Button(
            bottom_frame,
            text="Submit",
            font=20,
            width=10,
            height=1,
            command=lambda: self.check_part_in_list()
        )
        submit_btn.grid(row=2, column=1, padx=15)

        self.root.mainloop()

    def exit(self):
        self.root.destroy()

    def check_part_in_list(self) -> NoReturn:
        """
        Check whether part in parts list or not
        """
        # get part name
        part_name = self.part_name_entry.get()
        part_name = part_name.lower().strip()
        if part_name in self.part_names_list:
            # get number of parts
            try:
                part_count = self.part_count_entry.get()
                part_count = (int(part_count) if int(part_count)!=0 else 1) if part_count else 1
            except ValueError:
                # show warning if number of parts is not integer
                mb.showwarning(title="Warning", message="Enter the valid number of parts you want to copy!")
            else:
                # copy row from start excel file to new
                if copy_part(self.part_list[self.part_names_list.index(part_name)], part_count):
                    # show info about successful copying
                    mb.showinfo(title="Success", message="Data was successfully copied!")
                    self.part_name_entry.delete(0, tk.END)
                    self.part_count_entry.delete(0, tk.END)
                else:
                    # show error about unsuccessful copying
                    mb.showerror(title="Error", message="Something got wrong!")
        else:
            # show warning if part was not found
            mb.showwarning(title="Warning", message=f"Item {part_name} was not found!")


if __name__ == "__main__":
    main_app = App()






