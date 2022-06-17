import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class FileSelection(tk.LabelFrame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.salaryMonth = tk.StringVar()
        self.salaryMonth.set("급여대장")
        self.salaryPeople = tk.StringVar()
        self.salaryPeople.set("급여 대상자")
        self.salaryPersonalSheet = tk.StringVar()
        self.salaryPersonalSheet.set("급여명세서 양식")

        monthlySalary_button = ttk.Button(self, text="급여대장 선택", command = self.openSalarySheet )
        monthlySalary_label = ttk.Label(self, textvariable=self.salaryMonth)

        salaryPeople_button = ttk.Button(self, text="급여 대상자 명단  선택", command = self.openPeopleSheet)
        salaryPeople_label = ttk.Label(self, textvariable=self.salaryPeople)

        salaryPersonalSheet_button  = ttk.Button(self, text="급여 명세서 양식 선택", command = self.openPersonalSheet )
        salaryPersonalSheet_label = ttk.Label(self, textvariable=self.salaryPersonalSheet)

        monthlySalary_button.grid(padx = 10,pady=5, row=0, column=0, sticky=(tk.W+tk.E))
        monthlySalary_label.grid(padx = 10, pady=5,  row=0, column=1, sticky=(tk.W+tk.E))

        salaryPeople_button.grid(padx = 10,pady=5, row=1, column=0, sticky=(tk.W+tk.E))
        salaryPeople_label.grid(padx = 10,pady=5, row=1, column=1, sticky=(tk.W+tk.E))

        salaryPersonalSheet_button.grid(padx = 10,pady=5, row=2, column=0, sticky=(tk.W+tk.E))
        salaryPersonalSheet_label.grid(padx = 10,pady=5, row=2, column=1, sticky=(tk.W+tk.E))
        

    def openSalarySheet(self):
        self.salaryMonth.set("이라ㅓㄴㅇ라ㅓ니ㅏㅇ러니ㅏ얼")

    def openPeopleSheet(self):
        self.salaryPeople.set("dslkfjaslkdfjalksfj")

    def openPersonalSheet(self):
        self.salaryPersonalSheet.set("dslkfjaslkdfjalksfj")

    # def open():
    #     global my_image # 함수에서 이미지를 기억하도록 전역변수 선언 (안하면 사진이 안보임)
    #     root.filename = filedialog.askopenfilename(initialdir='', title='파일선택', filetypes=(
    # ('png files', '*.png'), ('jpg files', '*.jpg'), ('all files', '*.*')))
 
    #     Label(root, text=root.filename).pack() # 파일경로 view
    #     my_image = ImageTk.PhotoImage(Image.open(root.filename))
    #     Label(image=my_image).pack() #사진 view
 



class FileMaking(tk.Frame):
   def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)


class MyApplication(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("급여명세서 작성기")
        self.geometry("800x600")
        FileSelection(self).grid(sticky=(tk.E+tk.W+tk.N+tk.S))
        FileMaking(self).grid(sticky=(tk.E+tk.W+tk.N+tk.S))
        


if __name__=='__main__':
    app = MyApplication()
    app.mainloop()
    