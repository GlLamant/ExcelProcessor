from tkinter import *

i = 0
root = Tk()
str_var = StringVar()
str_var.set(i)

entry1 = Entry(root)
entry1.grid(row=0, column=2)


def on_click():
    global i, str_var, entry1
    # print(++i)
    # i_str = str(++i)
    # print(i_str)
    # str_var.set(i_str)
    # print(str_var.get())

    print (entry1.get())


entry = Entry(root, textvariable=str_var)
entry.grid(row=0, column=0)

Button(root, text='click me', width=20, command=on_click).grid(row=0, column=1)

root.mainloop()
