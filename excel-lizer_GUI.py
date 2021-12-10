from tkinter import PhotoImage, Tk, Button, Label, Text
from tkinter.filedialog import askopenfilename
from excellizer import ExcelLizer
from PIL import Image, ImageTk

class ExcelLizerGUI:
    def __init__(self):
        self.imageFileName = ''
        self.imageChosen: PhotoImage = None

        self.window = Tk("ExcelLizer - Velasco, Franco")
        self.window.title('ExcelLizer - Velasco, Franco')
        self.window.geometry('640x480+2+15')

        self.helloLabel = Label(self.window, text = 'Welcome to ExcelLizer 1.0! Pick an image file below:')
        self.fileLocationLabel = Label(self.window, text = '')

        self.selectFileButton = Button(self.window, text = 'Select File', command=self.askFile)

        self.runExcelLizerButton = Button(self.window, text='ExcelLize!', command=self.runExcelLizer)
        self.imageLabel = Label(self.window, image=self.imageChosen)

        self.helloLabel.pack()
        self.fileLocationLabel.pack()
        self.selectFileButton.pack()
        self.imageLabel.pack()
        self.runExcelLizerButton.pack()

        self.window.mainloop()

    def askFile(self):
        self.imageFileName = askopenfilename()
        self.fileLocationLabel.config(text = self.imageFileName)

        selectedImage = Image.open(self.imageFileName)
        selectedImage.thumbnail((400, 400), Image.ANTIALIAS)

        self.imageChosen = ImageTk.PhotoImage(selectedImage.rotate(180.0))
        self.imageLabel.config(image=self.imageChosen)

    def runExcelLizer(self):
        self.window.withdraw()
        if self.imageFileName == '':
            print('Error - Make sure you chose a proper file.')
            self.window.destroy()
            return
        excellizer = ExcelLizer(self.imageFileName)
        excellizer.convertImage()
        self.window.destroy()
        

if __name__ == '__main__':
    ExcelLizerGUI()
