import os
import sys
from pathlib import Path
import tkinter as tk

os.chdir(Path(sys.argv[0]).resolve().parents[3])

class App(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.pack()
    
    def create_widgets(self):
        label1 = tk.Label(root, text='URL')
        entry_box = tk.Entry(root, width=100, fg='black', bg='white')
        button1 = tk.Button(root, text='実行', bg='#ff7f50', command=lambda:info_slot_main(entry_box.get()))
        label2 = tk.Label(root, text='進行状況')
        label3 = tk.Label(root, text=f'出力フォルダ：{Path.cwd()}\n', height=5, justify='left', fg='black', bg='lavender', anchor=tk.NW)
        button2 = tk.Button(root, text='終了', bg='#ff7f50', command=root.destroy)

        label1.grid(row=0, column=0, columnspan=1)
        entry_box.grid(row=0, column=1, columnspan=1, sticky=tk.NE+tk.SW)
        button1.grid(row=0, column=2, columnspan=1)
        label2.grid(row=1, column=0, columnspan=1)
        label3.grid(row=1, column=1, columnspan=2, sticky=tk.NE+tk.SW)
        button2.grid(row=2, column=2, columnspan=1)

        self.rowconfigure(1, weight=1)

root = tk.Tk()
root.title('info_slot')
root.resizable(False, True)

app = App(master=root)
app.mainloop()
