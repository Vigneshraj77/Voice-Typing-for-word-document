from tkinter import *
from gtts import gTTS
from playsound import playsound
import os 
import docx

def dictatefile():
        try:
            global filename
            #to get file location from text box 
            loc=filename.get()
            doc = docx.Document(loc)  # Creating word reader object.
            data = ""
            fullText = []
            for para in doc.paragraphs:
                fullText.append(para.text)
                print(fullText)
                data = '\n'.join(fullText)
            print(data)
        except IOError:
                print('There was an error opening the file!')
        tts = gTTS(text=data, lang='en')
        tts.save("audio.mp3")
        tts = gTTS(text="project testing", lang='en')
        tts.save("audio.mp3")
        playsound("audio.mp3")
    

myWindow=Tk()
myWindow.title("Dictate")
myWindow.geometry("500x350")
Label(myWindow, text="").pack()
Label(myWindow, text="").pack()
#input box for entering file location
Label(myWindow, text="Enter File Location*").pack()
filename=StringVar()
name=Entry(myWindow,textvariable=filename).pack()
Label(myWindow, text="").pack()
Label(myWindow, text="").pack()
dic=Button(myWindow, text="Dictate", width=20, height=2,command=dictatefile).pack()
Label(myWindow, text="").pack()
Label(myWindow, text="Line which is in process").pack()
txt=StringVar()
Entry(myWindow,textvariable=txt).pack()
Label(myWindow, text="").pack()
myWindow.mainloop()
