import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox as msgbox
import subprocess as subprocess
from pathlib import Path
from SearchExcel import SearchExcel
from SearchResult import SearchResult
from AppWarnException import AppWarnException

#from tkinter import filedialog

class Application(tk.Frame):    
    def __init__(self, master=None):
        super().__init__(master)

        # ラベル
        self.DummyLabel = tk.Label(
            self.master,
            width=2,
            text=""
        )
        self.DummyLabel.grid(row=0,column=0)
        
        self.SearchValueLabel = tk.Label(
            self.master,
            width=18,
            font = ('Meiryo UI', 12),
            text="検索ワード１",
            bg='green4',
            fg='white'
        )
        #self.SearchValueLabel.grid(row=1,column=1,rowspan=3,sticky = tk.N+tk.S)
        self.SearchValueLabel.grid(row=1,column=1)

        self.SearchValueLabe2 = tk.Label(
            self.master,
            width=18,
            font = ('Meiryo UI', 12),
            text="検索ワード２",
            bg='green4',
            fg='white'
        )
        self.SearchValueLabe2.grid(row=2,column=1)

        self.SearchValueLabel3 = tk.Label(
            self.master,
            width=18,
            font = ('Meiryo UI', 12),
            text="検索ワード３",
            bg='green4',
            fg='white'
        )
        self.SearchValueLabel3.grid(row=3,column=1)
        
        self.SearchCondLabel = tk.Label(
            self.master,
            width=18,
            font = ('Meiryo UI', 12),
            text='組合せ条件',
            bg='green4',
            fg='white'
        )
        self.SearchCondLabel.grid(row=4,column=1)

        self.PathLabel = tk.Label(
            self.master,
            width=18,
            font = ('Meiryo UI', 12),
            text='対象ディレクトリ',
            bg='green4',
            fg='white'
        )
        self.PathLabel.grid(row=5,column=1)

        # テキストボックス
        self.SearchValue1Text = tk.Entry(
            self.master,
            width=40,
            font = ('Meiryo UI', 11),
            justify=tk.LEFT,
            textvariable=""
            )
        self.SearchValue1Text.grid(row=1,column=2,ipadx=1,ipady=1,padx=1,pady=1,columnspan=2,sticky = tk.W+tk.E)
        self.SearchValue2Text = tk.Entry(
            self.master,
            width=40,
            font = ('Meiryo UI', 11),
            justify=tk.LEFT,
            textvariable="",
            )
        self.SearchValue2Text.grid(row=2,column=2,ipadx=1,ipady=1,padx=1,pady=1,columnspan=2,sticky = tk.W+tk.E)
        self.SearchValue3Text = tk.Entry(
            self.master,
            width=40,
            font = ('Meiryo UI', 11),
            justify=tk.LEFT,
            textvariable="",
            state=tk.DISABLED
            )
        self.SearchValue3Text.grid(row=3,column=2,ipadx=1,ipady=1,padx=1,pady=1,columnspan=2,sticky = tk.W+tk.E)
        self.TargetPathText = tk.Entry(
            self.master,
            width=40,
            font = ('Meiryo UI', 11),
            justify=tk.LEFT,
            textvariable="",
            )
        self.TargetPathText.grid(row=5,column=2,ipadx=1,ipady=1,padx=1,pady=1,columnspan=3,sticky = tk.W+tk.E)

        # 検索
        self.SearchButton= tk.Button(
            self.master,
            width=20,
            font = ('Meiryo UI', 12),
            text="検索",
            bg='blue4',
            fg='white',
            command=self.SearchClickHandle
            )
        self.SearchButton.grid(row=1,column=4,rowspan=4,sticky = tk.N+tk.S)
        
        # ラジオボタン
        self.AndOraRadioValue = tk.IntVar()
        self.AndRadio = tk.Radiobutton(self.master,
                                       text="AND",
                                       font = ('Meiryo UI', 10),
                                       command=self.AndOraRadioClickHandle,
                                       variable=self.AndOraRadioValue,
                                       value=0)
        self.AndRadio.grid(row=4,column=2)#,ipadx=1,ipady=1,padx=1,pady=1)
        self.OrRadio = tk.Radiobutton(self.master,
                                      text="OR",
                                      font = ('Meiryo UI', 10),
                                      command=self.AndOraRadioClickHandle,
                                      variable=self.AndOraRadioValue,
                                      value=1,
                                      state=tk.DISABLED)
        self.OrRadio.grid(row=4,column=3)#,ipadx=1,ipady=1,padx=1,pady=1)
         
        # 表
        self.ResultTable=ttk.Treeview(
            self.master, 
            columns=(1,2,3,4),
            show='headings',
            height=15)
                
        style = ttk.Style()
        style.theme_use("default")
        # D3D3D3
        style.configure("Treeview", 
                        background="#FFE4C4",
                        foreground="black",
                        rowheight=25,
                        fieldbackground="#FFE4C4",
                        bordercolor="#FFE4C4",
                        borderwidth=1)
        # 347083
        style.map("Treeview",
                background=[("selected", "#411445")],
                foreground=[("selected", "white")])
#        style.configure("Treeview.Heading", font=('Meiryo UI', 12))        
        style.configure("Treeview.Heading", font=("Meiryo UI", 12, "bold"), borderwidth=1, relief="solid")
        style.configure("Treeview", font=("Meiryo UI", 11), borderwidth=1, relief="solid")
        
        #scrollbar = ttk.Scrollbar(self.master, orient=ttk.VERTICAL, command=self.ResultTable.xview)
        #self.ResultTable.configure(yscrollcommand=scrollbar, fill=ttk.Y)
        # 列の見出し設定
        self.ResultTable.heading(1, text='#0')
        self.ResultTable.heading(2, text='No')
        self.ResultTable.heading(3, text='Path')
        self.ResultTable.heading(4, text='Sheet')

        # 列の設定
        self.ResultTable.column(1, width=0, stretch='no')
        self.ResultTable.column(2, anchor='center', width=40, stretch='no')
        self.ResultTable.column(3, anchor='w', width=620)
        self.ResultTable.column(4, anchor='w', width=80)

        # レコードの追加
#        self.ResultTable.insert(parent='', index='end', iid=0 ,values=(1, 'C:\work\main.xlsx'))
#        self.ResultTable.insert(parent='', index='end', iid=1 ,values=(2,'C:\backup\work.xlsx'))
#        self.ResultTable.insert(parent='', index='end', iid=2 ,values=(3,'C:\wok\doc\abc.xlsx'))
                
        self.ResultTable.grid(row=20,column=1,columnspan=4,sticky = tk.W+tk.E)
        self.ResultTable.bind("<Double-1>", self.ResultTableDoubleClickHandle)
                
        # 閉じるボタン
        self.CloseButton = tk.Button(
            self.master,
            height=2,
            width=20,
            font = ('Meiryo UI', 12),
            bg='blue4',
            fg='white',
            text="閉じる",
            command=self.master.destroy)
        self.CloseButton.grid(row=21,column=4)

        # 初期値をセット
        pathObj = Path()
        curDir = str(pathObj.cwd())
        # self.SearchValue1Text.insert(tk.END,"")
        self.TargetPathText.insert(tk.END,curDir)
   
    def SearchClickHandle(self):
        # コントールから値を取得
        searchWord1 = self.SearchValue1Text.get()
        searchWord2 = self.SearchValue2Text.get()
        targetPath = self.TargetPathText.get()
        radioValue = self.AndOraRadioValue.get()
        
#        # 2のみに値が入っていた場合は、2を1にコピーする。
#        if searchWord1 == "" and searchWord2 != "":
#            searchWord1 = searchWord2
#            searchWord2 = ""
#            self.SearchValue1Text.insert(tk.END, searchWord2)
#            self.SearchValue2Text.delete(0, tk.END)

        # 入力チェック
        if searchWord1 == "":
            target = self.SearchValueLabel.cget("text")
            msgbox.showwarning('Excel Grep', f'{target}を入力してください。')
            return

        if targetPath == "":
            target = self.PathLabel.cget("text")
            msgbox.showwarning('Excel Grep', f'{target}を入力してください。')
            return

        # 表の行をすべて削除する。
        for row in self.ResultTable.get_children():
            self.ResultTable.delete(row)
        
        searchExcelObj = SearchExcel()
        try:
            dataList = searchExcelObj.search(targetPath, searchWord1, searchWord2, radioValue)
        except AppWarnException as ex:
            msgbox.showwarning('Excel Grep', ex.ErrMsg)
            return        
        
        index = 0
        for data in dataList:            
            self.ResultTable.insert(parent='', index=index, iid=index ,values=(index, (index + 1), data.bookPath, data.sheetName))
            index = index + 1

    def ResultTableDoubleClickHandle(self, event):
        item = self.ResultTable.selection()[0]
        values = self.ResultTable.item(item, "values")        
#        print("Selected item:", values[1])
        print("Selected item:", values[2])
#        print("Selected item:", values[3])
        # 指定されたファイルを既定のプログラムで実行。
        subprocess.Popen(['start', values[2]], shell=True)
        
    def AndOraRadioClickHandle(self):
        value = self.AndOraRadioValue.get()
        print(f"ラジオボタンの値は {value} です")        

root = tk.Tk()
root.title('Excel Grep')
root.geometry('800x640')
app = Application(master=root)

# TODO unhadle exception
root.mainloop()
    
