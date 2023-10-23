from tkinter import Tk, ttk
import webbrowser

class TemplateWindow(Tk):
    def __init__(self):
        super().__init__()
        
        self.create_frame()

        self.my_credit = ttk.Frame(self)
        self.my_credit.pack(side='bottom', pady=10)

        self.credit_label = ttk.Label(self.my_credit, font=('Segoe UI', 8), text='\nApp developed by\nProf. Nithin H M\nAssistant Professor\nDepartment of Physics\nAMC Engineering College\nBangalore - 560083', justify='center')
        self.credit_label.pack()

        self.gitlink = ttk.Label(self.my_credit, text="My GitHub", foreground="blue", cursor="hand2")
        self.gitlink.pack()
        self.gitlink.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/nithinhm"))

        self.dinlink = ttk.Label(self.my_credit, text="My LinkedIn", foreground="blue", cursor="hand2")
        self.dinlink.pack()
        self.dinlink.bind("<Button-1>", lambda e: webbrowser.open_new("https://linkedin.com/in/nithinhm13"))

    def create_frame(self):
        pass