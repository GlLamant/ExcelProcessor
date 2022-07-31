import time
from tkinter import *

root = Tk()
root.geometry('200x200')
root.resizable(width=False, height=False)
root.title('Button Test')


def click_me():
    """修改btn控件的text属性以及state，3秒后，恢复原有状态"""
    btn["text"] = "不可点"
    btn['state'] = DISABLED
    time.sleep(3)
    btn["text"] = "点我测试"
    btn['state'] = NORMAL


Label(root, text="button test").pack()

# btn被点击后，调用click_me函数，将btn文案修改为：不可点，状态修改为DISABLED（置灰状态，不可点击），3秒后，恢复原有状态
btn = Button(root, text="点我测试", command=click_me)
btn.pack()
btn1 = Button(root, text="我不可被点击", state=DISABLED)
btn1.pack()

root.mainloop()