from cProfile import label
from tkinter import Label, Entry, Button, Menu, StringVar, filedialog, Tk, END
from tkinter import messagebox as mb
import webbrowser

from click import command

import build_report


class App(Tk):
    def __init__(self):
        super().__init__()

        ## App Settings
        self.resizable(False, False)
        self.iconbitmap('icon.ico')

        self.version = "0.1"
        self.author = "PRunco"

        ## Menu
        self.title("Supply Line Report Utility")
        self.menubar = Menu(self)

        self.filemenu = Menu(self.menubar, tearoff=False)
        self.filemenu.add_command(label="Exit", command=self.quit)
        self.menubar.add_cascade(label="File", menu=self.filemenu)
        
        self.helpmenu = Menu(self.menubar, tearoff=False)
        self.helpmenu.add_command(label="About", command=self.open_about)
        self.helpmenu.add_command(label="Documentation", command=self.open_docs)
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)

        self.config(menu=self.menubar)

        ## Order Input
        # Label
        self.order_status_label = Label(
            self,
            text="Order Status View",
            ).pack(padx=(5,0), anchor='w')
        # Entry
        self.order_status_entry = Entry(
            self, 
            width=80, 
            textvariable=StringVar
            )
        self.order_status_entry.pack(padx=5, pady=(0, 5))
        # Button
        self.order_status_button = Button(
            self, text="Browse", 
            command=lambda: self.browse_for("order_status")
            ).pack(padx=5, pady=(0, 5), anchor='e')

        ## Template Input
        # Label
        self.template_label = Label(
            self,
            text="Supply Line Template"
            ).pack(padx=(5,0), anchor='w')
        # Entry
        self.template_entry = Entry(
            self,
            width=80,
            textvariable=StringVar
            )
        self.template_entry.pack(padx=5, pady=(0, 5))
        # Button
        self.template_button = Button(
            self, text="Browse", 
            command=lambda: self.browse_for("template")
            ).pack(padx=5, pady=(0, 5), anchor='e')

        self.generate_button = Button(
            self, text="Update report",
            command=lambda: self.on_click_create_report()
        ).pack(fill='x', padx=5, pady=5)

    def browse_for(self, target):
        file_name = filedialog.askopenfilename(
            filetypes=(("Excel files", "*xlsx"), ("All files", "*"))
        )
        if target == "order_status":
            self.order_status_entry.config(background='white')
            self.order_status_entry.delete(0, END)
            self.order_status_entry.insert(0, file_name)
        if target == 'template':
            self.template_entry.config(background='white')
            self.template_entry.delete(0, END)
            self.template_entry.insert(0, file_name)

    def open_about(self):
        mb.showinfo(
            title="About",
            message=f"Version {self.version} | {self.author}"
            )

    def open_docs(self):
        webbrowser.open('https://www.github.com/paulrunco')
        
    def ask_save_as(self):
        save_as = filedialog.asksaveasfilename(
            title="Save as",
            filetypes=(("Excel files", "*xlsx"), ("All files", "*")),
            defaultextension="xlsx",
            initialfile="test")
        if save_as:
            print(save_as)
            return save_as

    def on_click_create_report(self):
        path_to_order_status_report = self.order_status_entry.get()
        if path_to_order_status_report == "":
            mb.showwarning(title="Warning: ID-10T", message="Please select an Order Status View")
            self.order_status_entry.config(background= 'red')
            return

        path_to_template = self.template_entry.get()
        if path_to_template == "":
            mb.showwarning(title="Warning: ID-10T", message="Please select the Supply Line Template")
            self.template_entry.config(background='red')
            return
        
        #save_as = self.ask_save_as()
        build_report.build_report(self, path_to_order_status_report, path_to_template)


if __name__ == "__main__":
    app = App()
    app.mainloop()