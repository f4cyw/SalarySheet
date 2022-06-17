import tkinter as tk
from tkinter import ttk

class SalaryGui(tk.LabelFrame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.salaryMonth = tk.StringVar()
        self.salaryMonth.set("급여대장")
        self.salaryPerson = tk.StringVar()
        self.salaryPerson.set("급여 대상자")


        monthlySalary_button = ttk.Button(self, text="급여대장 선택", command = self.openSalarySheet )
        monthlySalary_label = ttk.Label(self, textvariable=self.salaryMonth)
        salaryPerson_button = ttk.Button(self, text="급여 대상자 명단  선택", command = self.openPersonalSheet )
        salaryPerson_label = ttk.Label(self, textvariable=self.salaryPerson)

        

        monthlySalary_button.grid(row=0, column=0, sticky=(tk.W+tk.E))
        monthlySalary_label.grid(row=0, column=1, sticky=(tk.W+tk.E))

        salaryPerson_button.grid(row=1, column=0, sticky=(tk.W+tk.E))
        salaryPerson_label.grid(row=1, column=1, sticky=(tk.W+tk.E))

        

    def openSalarySheet(self):
        self.salaryMonth.set("이라ㅓㄴㅇ라ㅓ니ㅏㅇ러니ㅏ얼")

    def openPersonalSheet(self):
        self.salaryPerson.set("dslkfjaslkdfjalksfj")


class MyApplication(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("급여명세서 작성기")
        self.geometry("800x600")
        SalaryGui(self).grid(sticky=(tk.E+tk.W+tk.N+tk.S))


if __name__=='__main__':
    app = MyApplication()
    app.mainloop()
    