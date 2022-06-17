from curses import BUTTON_SHIFT
import openpyxl
import tkinter as tk
# from tkinter import ttk

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        group_1 = tk.LabelFrame(self, padx=15, pady=10, text="사전작업")
        group_1.pack(padx=10, pady=5)
        buttons = ['급여대장 불러오기', '명세서 불러오기', '대상자 불러오기']
        btns = [tk.Button(group_1, text=t, command= lambda :  self.openD(t) ) for t in buttons]
        labels = [tk.Label(group_1, text = t.split()[0]) for t in buttons] 
        self.widgets = list(zip(btns, labels))

        for i, (btn, label) in enumerate(self.widgets):
            btn.grid(row=i)
            label.grid(row=i, column=1)
    



        # tk.Label(group_1, text = "이름").grid(row=0)
        # tk.Label(group_1, text = "전화번호").grid(row=1)
        # tk.Entry(group_1).grid(row=0, column=1, sticky=tk.W)
        # tk.Entry(group_1).grid(row=1, column=1, sticky=tk.W)

        # group_2 = tk.LabelFrame(self, padx=15, pady=10, text="대상자")
        # group_2.pack(padx=10, pady=5)
        # chks =[]
        # names = []
        # depts =[]
        # personalNums = []
        # mails=[]
        # buttons = ['급여대장 불러오기', '명세서 불러오기', '대상자 불러오기']
        # btns = [tk.Button(group_2, text= f) for f in buttons]
        # labels = [tk.Label(group_2, text = t.split()[0]) for t in buttons] 
        # self.widgets = list(zip(btns, labels))

        # for i, (btn, label) in enumerate(self.widgets):
        #     btn.grid(row=i)
        #     label.grid(row=i, column=1)

        # # tk.Label(group_2, text = "상세주소").grid(row=0)
        # # tk.Label(group_2, text = "도시명").grid(row=1)
        # # tk.Entry(group_2).grid(row=0, column=1, sticky=tk.W)
        # # tk.Entry(group_2).grid(row=1, column=1, sticky=tk.W)

        # self.btn_submit = tk.Button(self, text = "submit")
        # self.btn_submit.pack(padx=10, pady=10)

    def openD(self, sender):
        print("all")
        print(sender)

if __name__ =='__main__':
    app = App()
    app.mainloop()

