#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Mark
#
# Created:     29/11/2014
# Copyright:   (c) Mark 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from Tkinter import *

import sys
from BookOrigamiLib import *

def main():
    #sys.argv[0] argument is script name
    #sys.argv[1] is bitmap name.
    print 'Number of arguments:', len(sys.argv), 'arguments.'
    print 'Argument List:', str(sys.argv)
    filename = None

    if len(sys.argv) == 2: #convert bitmap from argument
        filename = sys.argv[1]

    else: #open GUI
        root = Tk()
        app = Application(parent=root)
        app.mainloop()
        root.destroy()

        #by this point, the user will have either:
        #   -selected or created a file, in which case filename is defined
        #   -has not, in which case, file name is not defined.


    if filename:
        #filename is defined, thus can continue
        print "Opening Image"
        from PIL import Image;
        myImage=Image.open(filename);
        #myImage.convert("1")
        #from tkMessageBox import showinfo
        #showinfo("Mode", myImage.mode)
        #px = myImage.load();
        #print "Image Opened"

        #what I'd like to be able to do is:
            #provide a standard bitmap
            #automatically convert to monochrome bitmap (thought bubble - choose monochrome from red, green or blue channels? provide preview?)
            #display monochrome bitmap for sanity check
            #get user input for height of book and number of sheets
            #automatically scale image appropriately for pixels per mm for height
            #automatically scale image appropriately for 1 pixel per sheet for width

        document = OpenAndInitialiseDocX(filename)
        print "Document opened"
        CalculateAndWriteDocX(document, myImage)

        document.save(filename + ".docx")
        #todo add exception handling and ask for filename.
        print "Complete"
    else:
        #filename is not defined, program should exit
        print "No Filename provided to Open"
    exit()

class Application(Frame):

    def createBMPofText(self):

        self.GenerateFromText = Toplevel(self.parent) #self = app (instance of Application). self.parent = root
        self.GenerateFromText.transient(self.parent)
        self.nextLevel=GenerateFromText(self.GenerateFromText) #nextLevel = instance of GenerateFromText


    def selectFile(self):
        from tkFileDialog import askopenfilename
        #Tk().withdraw()
        global filename
        file_opt = options = {}
        options['defaultextension'] = '.bmp'
        options['filetypes'] = [('bitmap files', '.bmp'), ('all files', '.*')]
        options['title'] = 'Select a Bitmap file to open'

        loadretry = True #retry for loop to get value filename selected
        while loadretry:
            filename = askopenfilename(**file_opt)
            if filename:
                #filename is defined, thus can continue
                print "Opening Image"
                from PIL import Image;
                myImage=Image.open(filename);

                sheets = myImage.size[0]
                height = myImage.size[1]

                from tkSimpleDialog import askinteger
                askoptions = options = {}
                options['initialvalue'] = height
                options['parent'] = self
                options['minvalue'] = 50
                options['maxvalue'] = 1000
                newheight = askinteger("Confirm Bookheight", "Bookheight\n(in mm)", **askoptions)

                options['initialvalue'] = sheets
                options['minvalue'] = 10
                options['maxvalue'] = 2000
                newsheets = askinteger("Confirm sheets", "Number of sheets", **askoptions)

                if((sheets != newsheets) and (height != newheight)):
                    myImage = myImage.resize((newsheets, newheight))

                from tkMessageBox import showinfo
                #showinfo("Mode", myImage.mode)
                #myImage.convert("L")
                #showinfo("Mode", myImage.mode)

                #px = myImage.load();
                #print "Image Opened"

                document = OpenAndInitialiseDocX(filename)
                print "Document opened"
                CalculateAndWriteDocX(document, myImage)

                saveRetry=False
                try:
                    document.save(filename + ".docx")
                    print "Complete"
                    self.status.set("Complete")
                    break
                except:
                    from tkMessageBox import askretrycancel
                    saveretry = askretrycancel("Error", "Cannot save to selected file. Try again?")

                while saveretry:
                    from tkMessageBox import asksaveasfile
                    file_opt = options = {}
                    options['defaultextension'] = '.docx'
                    options['filetypes'] = [('Word DOCX', '.docx')]
                    options['title'] = 'Save File As:'
                    filename = askopenfilename(**file_opt)
                    try:
                        document.save(filename)
                        print "Complete"
                        status.set("Complete")
                        break
                    except:
                        from tkMessageBox import askretrycancel
                        status.set("Could not save file")
                        saveretry = askretrycancel("Error", "Cannot save to selected file. Try again?")

            else:
                from tkMessageBox import askretrycancel
                #statusvar.set("Could not Open File for Import")
                loadretry = askretrycancel("Error", "File for Import was not selected. Try again?")
                if loadretry==False:
                    self.status.set("Not able to Import")

        from tkMessageBox import showinfo
        #showinfo("", self.status.get())

    def createWidgets(self):
        from PIL import Image, ImageTk
        image = Image.open("Love Book.png")
        photo = ImageTk.PhotoImage(image)
        self.IMAGE = Label(self, image = photo)
        self.IMAGE.image = photo
        self.IMAGE.pack(side=TOP, padx=10, pady=5)

        self.STATUSFIELD = Label(self, textvariable=self.status)
        self.STATUSFIELD.pack(padx=10, pady=5)

        self.CREATEBMP = Button(self, text="Create Image from Text", command=self.createBMPofText)
        self.CREATEBMP.pack(side=LEFT, padx=10, pady=5)

        self.SELECTFILE = Button(self, text="Select Image", command=self.selectFile)
        self.SELECTFILE.pack(side=LEFT, padx=10, pady=5)

        self.QUIT = Button(self, text="Quit", fg="red", command=self.quit)
        self.QUIT.pack(side=RIGHT, padx=10, pady=5)


    def centreWindow(self, width, height):
        self.parent.geometry(str(width) +"x" + str(height) + "+" + str((self.parent.winfo_screenwidth()-width)/2) + "+" + str((self.parent.winfo_screenheight()-height)/2))

    def __init__(self, parent=None):
        Frame.__init__(self, parent)

        self.status = StringVar()
        self.status.set("Welcome to the Book Origami Dimensions Generator")


        self.parent = parent
        self.parent.title("Book Origami Dimensions Generator")
        self.pack(fill=BOTH, expand=1)
        self.centreWindow(450, 450)
        self.createWidgets()



class GenerateFromText(Frame):

    def generateImage(self):
        bookheight = int(self.BOOKHEIGHT.get())
        booksheets = int(self.BOOKSHEETS.get())
        text = str(self.RENDERTEXT.get())
        font = self.SelectedFont.get()

        print text + " using font " + font
        print str(bookheight) + "mm high, " + str(booksheets) + " sheets wide"

        from PIL import Image, ImageFont, ImageDraw
        IMAGEMODE = "1" #monochrome bitmap
        #create monochrome bitmap image, 10*booksheets wide, bookheight high, 0=black, 1=white
        image = Image.new(IMAGEMODE, (booksheets*10,bookheight), 1)
        from tkMessageBox import showinfo
        #showinfo("Mode", image.mode)

        draw = ImageDraw.Draw(image)
        #fontheight = bookheight*4/5 #scale text to 80% of book height
        fontheight = bookheight

        renderfont = ImageFont.truetype(font,fontheight)
        print "Text Size " + str(draw.textsize(text, renderfont))
        print "Image Size " + str(image.size)

        #offset to attempt to centre font rather than bottom aligned (upward shift by 5%)
        #print (bookheight/-20)
        draw.text((0,(bookheight/-20)), text, fill=0, font=renderfont)

        filename = text + " using " + font + " - " + str(bookheight) + "mm, " + str(booksheets) + "p.bmp"
        image.save(filename)

        image = image.convert("L")
        image = RemoveWhiteColumns(image)
        image = image.resize((booksheets,bookheight))
        image.save(filename)


        image.show()
        from tkMessageBox import askyesno
        if(askyesno("Continue?", "Preview ok? Continue?")):
        #print px
            document = OpenAndInitialiseDocX(filename)
            print "Document opened"

            CalculateAndWriteDocX(document, image)

            document.save(filename + ".docx")
        #else:
            #filename is not defined, program should exit
            #print "No Filename provided to Open"


        #todo add exception handling and ask for filename.
        #print "Complete"
        #image.save("blahoutput.bmp")

        if not askyesno("Another?", "Create another from Text?"):
            self.quit()





    def createWidgets(self):

        self.RENDERTEXTLABEL = Label(self, text="Text: ")
        self.RENDERTEXTLABEL.pack()
        self.RENDERTEXT = Entry(self)
        self.RENDERTEXT.pack(padx=5, pady=5)

        self.SelectedFont = StringVar()
        import os
        fontlist = os.listdir("c:\\windows\\fonts\\")
        fontlist = filter(fontNameContainsTTF, fontlist) #remove non TTF
        self.FONTBOXLABEL = Label(self, text="Font: ").pack()
        self.SelectedFont.set(fontlist[147])
        self.FONTSELECT = apply(OptionMenu, (self, self.SelectedFont) + tuple(fontlist))
        self.FONTSELECT.pack()

        self.BOOKHEIGHTLABEL = Label(self, text="Height of book in mm: ").pack()
        self.BOOKHEIGHT = Entry(self)
        self.BOOKHEIGHT.pack(padx=5, pady=5)

        self.BOOKSHEETSLABEL = Label(self, text="Number of Sheets in book: ").pack()
        self.BOOKSHEETS = Entry(self)
        self.BOOKSHEETS.pack(padx=5, pady=5)

        self.CANCEL = Button(self, text="Cancel", fg="red", command=self.quit)
        self.CANCEL.pack(side=RIGHT, padx=5, pady=5)

        self.PREVIEWBUTTON = Button(self, text="Generate Image", command=self.generateImage)
        self.PREVIEWBUTTON.pack(side=RIGHT)


##        from PIL import Image, ImageTk
##        image = Image.open("Love Book.png")
##        photo = ImageTk.PhotoImage(image)
##        self.IMAGE = Label(self, image = photo)
##        self.IMAGE.image = photo
##        self.IMAGE.pack(side=TOP, padx=10, pady=5)
##
##        self.STATUSFIELD = Label(self, textvariable=self.status)
##        self.STATUSFIELD.pack(padx=10, pady=5)
##
##        self.CREATEBMP = Button(self, text="Create Image from Text", fg="blue", command=self.createBMPofText, state=DISABLED)
##        self.CREATEBMP.pack(side=LEFT, padx=10, pady=5)
##
##        self.SELECTFILE = Button(self, text="Select Image", command=self.selectFile)
##        self.SELECTFILE.pack(side=LEFT, padx=10, pady=5)
##

    def centreWindow(self, width, height):
        self.parent.geometry(str(width) +"x" + str(height) + "+" + str((self.parent.winfo_screenwidth()-width)/2) + "+" + str((self.parent.winfo_screenheight()-height)/2))



    def __init__(self, parent):
        Frame.__init__(self, parent)

        self.status = StringVar()
        self.status.set("Generating from Text")
        #from PIL import Image, ImageTk
        #self.bmp = Image.open("blank.bmp")
        #self.previewImage = ImageTk.BitmapImage(self.bmp)




        self.parent = parent
        self.parent.title("Generate from Text")
        self.pack(fill=BOTH, expand=1)
        self.centreWindow(450, 450)
        self.createWidgets()







def fontNameContainsTTF(fontname):
    return ".ttf" in fontname.lower()





if __name__ == '__main__':
    main()


