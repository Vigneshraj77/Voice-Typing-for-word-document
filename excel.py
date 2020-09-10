from tkinter import *
import speech_recognition as sr 
import pyttsx3
from tkinter import filedialog
# import xlsxwriter module 
import xlsxwriter 
#
def browse_button():
    filename = filedialog.askdirectory()
    print(filename)
    global i
    i.set(filename)
    return

#saving the excel document
def save():
    global workbook
    workbook.close()
    return

#creating the interface
myWindow=Tk()
myWindow.title("Voice Typing for Excel")
myWindow.geometry("700x500")

# Start from the first cell.
# Rows and columns are zero indexed. 
row = 0
column = 0


#function for listening to voice
def listen():
    def create(txt):
        content=txt
        global row
        global worksheet
        global column
        # iterating through content list 
        for item in content :
            # write operation perform 
            worksheet.write(row, column, item)
            # incrementing the value of row by one with each iteratons. 
            row += 1
            return
    # Initialize the recognizer 
    r = sr.Recognizer()
    try:
        # Loop infinitely for user to speak 
        while(1):
            with sr.Microphone() as source :
                # wait for a second to let the recognizer 
                # adjust the energy threshold based on 
                # the surrounding noise level 
                r.adjust_for_ambient_noise(source, duration=0.2) 
                #listens for the user's input 
                audio = r.listen(source)            
                # Using google to recognize audio 
                MyText = r.recognize_google(audio) 
                MyText = MyText.lower()
                global s
                s.set(MyText)
                MyList=list(MyText.split(" "))
                create(MyList)
                return
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
Label(myWindow, text="").pack()
Label(myWindow, text="").pack()
Label(myWindow, text="Enter Filename*").pack()
Button(myWindow, text ="browse",compound = LEFT,command=browse_button).pack()
Label(myWindow, text="").pack()
# Creating a photoimage object to use image 
photo = PhotoImage(file = r"C:\Users\Hp\Documents\Python Scripts\mic.png")
# Resizing image to fit on button 
photoimage = photo.subsample(3, 3)
name=StringVar()
e1 = Entry(myWindow,textvariable=name).pack()
Label(myWindow, text="").pack()
loc=str(name.get())
loc=loc+".xlsx"
workbook = xlsxwriter.Workbook("Example.xlsx") 
worksheet = workbook.add_worksheet()
mic=Button(myWindow, text = "Click to start Voice typing", image = photoimage,compound = LEFT,command=listen).pack()
Label(myWindow, text="").pack()
s=StringVar()
Entry(myWindow,textvariable=s).pack()
Label(myWindow, text="").pack()
Button(myWindow,text="save document",command=save).pack()
myWindow.mainloop()
