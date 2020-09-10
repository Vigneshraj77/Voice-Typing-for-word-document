from tkinter import *

def myWindow():
    myWindow=Tk()
    myWindow.title("Menu Page")
    myWindow.geometry("500x350")
    #event listener for voice typing
    def voice():
        myWindow.destroy()
        exec(open('voice.py'.read()))
        return
    #event listener for dictate option
    def dictate():
        myWindow.destroy()
        exec(open('dictate.py').read())
        return
    def excel():
        myWindow.destroy()
        exec(open('excel.py').read())
    Label(myWindow, text="").pack()
    Label(myWindow, text="").pack()
    Label(myWindow, text="").pack()
    #button for voice typing 
    vc=Button(myWindow, text="voice typing for word", width=20, height=2,command=voice).pack()
    Label(myWindow, text="").pack()
    Label(myWindow, text="").pack()
    #button for dictate
    dt=Button(myWindow, text="dictate", width=20, height=2,command=dictate).pack()
    Label(myWindow, text="").pack()
    Label(myWindow, text="").pack()
    Button(myWindow, text="Voice typing for Excel", width=20, height=2,command=excel).pack()
    myWindow.mainloop()
             
    
myWindow()
