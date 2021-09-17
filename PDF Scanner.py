#import Libraries
import os
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import ImageTk,Image
import cv2 as cv
import sys
import numpy as np
import img2pdf    
from tkinter.scrolledtext import ScrolledText   #scrollable text box
import PyPDF2  #pdf merging
import win32com.client # for Word and ppd to pdf
import pytesseract    #OCR
# initiate tkinter
root=tk.Tk()
root.title("PDF Converter")
#geometry for window 
root.geometry('800x770') ##(w x h)
root.minsize(800,770)
root.maxsize(800,770)
root.configure(bg='#26A65B',bd=6)

#Image to pdf converting
def IMGTOPDF():
    global f3
    if f3:
        f3.destroy()
    f3=Frame(root,bg="green",bd=4)
    f3.grid(row=1,column=0,rowspan=10,columnspan=9)

    list1=[]# it stores link of import item
    thumb=[]# it store rezized thumbnail of pictures
    listcrop=[]# it stores the product images
    listpdf=[]#final list of image before pdf conversion

    n=0 #no of images
    i=0 #count for current image in view
    # function for thumbnail generation
    def tumbnail(imgx):
            
            imgx = imgx.resize((600,700), Image.ANTIALIAS)
            photoImg =  ImageTk.PhotoImage(imgx)
            return photoImg

    #Import function
    def insertfun():
          global f3
          nonlocal list1
          nonlocal thumb
          nonlocal n
          nonlocal i
          root.withdraw()
          files=filedialog.askopenfilenames(title="insert",filetypes=(("png files","*.png"),("allfiles","*.*")))
          
         
          filez = root.tk.splitlist(files)
          root.update()
          root.deiconify()
          
          for filex in filez:
               list1.append(filex)
               img = Image.open(filex)
               imggg=tumbnail(img)
               listcrop.append(img)
               thumb.append(imggg)
               
          i=0
          n=len(thumb)
          #redefining canvas
          canva=Canvas(f3,height=700,width=650,bg="lightsteelblue2")
          canva.grid(row=1,column=1,columnspan=6,rowspan=9)
          imago=canva.create_image(0,0,anchor=NW,image=thumb[i])

    #forward button function
    def forw():
          nonlocal list1
          nonlocal thumb
          nonlocal i
          global f3
          if i<len(thumb)-1:
              i=i+1
              canva=Canvas(f3,height=700,width=650,bg="lightsteelblue2")
              canva.grid(row=1,column=1,columnspan=6,rowspan=9)
              imago=canva.create_image(0,0,anchor=NW,image=thumb[i])
              l3 = Button(f3,text="<<",state=NORMAL,command=backw)
              l3.grid(row=5,column=0)
          else:
              l4 = Button(f3,text=">>",state=DISABLED) 
              l4.grid(row=5,column=7)

    #Backward button function
    def backw():
          nonlocal list1
          nonlocal thumb
          nonlocal i
          global f3
          if i>=0 :
              i=i-1
              canva=Canvas(f3,height=700,width=650,bg="lightsteelblue2")
              canva.grid(row=1,column=1,columnspan=6,rowspan=9)
              imago=canva.create_image(0,0,anchor=NW,image=thumb[i])
              l4 = Button(f3,text=">>",state=NORMAL,command=forw) 
              l4.grid(row=5,column=7)
          else:
              l3 = Button(f3,text="<<",state=DISABLED)
              l3.grid(row=5,column=0)



    #delete function to delete at count image
    def deletee():
             nonlocal list1
             nonlocal thumb
             nonlocal i
             nonlocal n
             global f3
             list1.pop(i)
             thumb.pop(i)
             listcrop.pop(i)
             n=len(thumb)
             if i==n-1:
                 i-=1
             if i==0:
                 if thumb.empty():

                     canva=Canvas(f3,height=700,width=650,bg="lightsteelblue2")
                     canva.grid(row=1,column=1,columnspan=6,rowspan=9)
                 else:
                   i
             canva=Canvas(f3,height=700,width=650,bg="lightsteelblue2")
             canva.grid(row=1,column=1,columnspan=6,rowspan=9)
             imago=canva.create_image(0,0,anchor=NW,image=thumb[i])

    #to crop image (main function for project)
    def cropp():
              nonlocal list1
              nonlocal i
              nonlocal thumb
              nonlocal n
              nonlocal listcrop
              nonlocal listpdf
              global f3
              listxx=[]
              img=cv.imread(cv.samples.findFile(list1[i]))

              if img is None:
                  sys.exit("Could not read the image")

              screen_res = 1280, 720
              scale_width = screen_res[0] / img.shape[1]
              scale_height = screen_res[1] / img.shape[0]
              scale = min(scale_width, scale_height)
                  #resized window width and height
              window_width = int(img.shape[1] * scale)
              window_height = int(img.shape[0] * scale)

              cv.namedWindow('Image Display',cv.WINDOW_NORMAL)
              cv.resizeWindow('Image Display',window_width,window_height)
              cv.imshow("Image Display",img)
              def click_event(event, x, y, flags, params):
                   if len(listxx)<=3:
                             # checking for left mouse clicks
                       if event == cv.EVENT_LBUTTONDOWN:
                           listxx.append([x,y])
                                 # displaying the coordinates
                                 # on the Shell
                           
                           cv.imshow('Image Display', img)
                   if len(listxx)==4:
                          
                          pts = np.array(listxx, np.int32)
                          while True:
                                pts=np.float32(pts)
                                
                                pts2 = np.float32([[0, 0], [959, 0], [0, 1279], [959, 1279]])
                    
                                # Apply Perspective Transform Algorithm
                                matrix = cv.getPerspectiveTransform(pts, pts2)
                                result = cv.warpPerspective(img, matrix, (img.shape[1], img.shape[0]))
                                
                                # Wrap the transformed image
                                cv.namedWindow('Image',cv.WINDOW_NORMAL)
                                cv.resizeWindow('Image',window_width,window_height)
                                
                                cv.imshow('Image', result) # Transformed Capture
                                result= cv.cvtColor(result, cv.COLOR_BGR2RGB)
                                im_pil = Image.fromarray(result)
                                listcrop[i]=im_pil
                                thumb[i]=tumbnail(im_pil)
                                #redefining canvas
                                canva=Canvas(f3,height=700,width=650,bg="lightsteelblue2")
                                canva.grid(row=1,column=1,columnspan=6,rowspan=9)
                                imago=canva.create_image(0,0,anchor=NW,image=thumb[i])
                                #waiting for key input to end terminal
                                zz=cv.waitKey(0)
                                if zz==ord('e'):
                                      cv.destroyAllWindows()
                                      break
                                if zz==ord('s'):
                                      listcrop[i]=result
                                      cv.destroyAllWindows()
                                      break
                          
                          
              #to set terminal in loopback
              cv.setMouseCallback("Image Display", click_event)


                    
              

    #function to rotate image              
    def rotatee():
            nonlocal i
            nonlocal listcrop
            nonlocal thumb
            nonlocal n
            global f3
            imgho=listcrop[i]
            imgho=imgho.rotate(90)
            listcrop[i]=imgho
            thumb[i]=tumbnail(imgho)
            canva=Canvas(f3,height=700,width=650,bg="lightsteelblue2")
            canva.grid(row=1,column=1,columnspan=6,rowspan=9)
            imago=canva.create_image(0,0,anchor=NW,image=thumb[i])



    #function to save image
    def makepdf():
         nonlocal listcrop
         nonlocal list1

         if len(thumb)!=0:
             lixrt=listcrop[1:]
             
             
             im1 = listcrop[0].convert('RGB')
             for xx in range(n-1):
                 lixrt[xx]=lixrt[xx].convert('RGB')
             file=filedialog.asksaveasfilename()

             im1.save(file+".pdf","PDF",save_all=True, append_images=lixrt)
             print("saved")
             # #pdf_bytes = img2pdf.convert(im1,filename)
             # im1=Image.open(lixrt[0].filename())  
             # im1 = im1.convert('RGB')
             # im1.save(r'C:/Users/a/OneDrive/Pictures/l.pdf',"PDF")
             # print("saved")
    #all Gui widgets for IMGtoPDF
    l3 = Button(f3,text="<<",command=backw)
    l3.grid(row=5,column=0,sticky=N)
    l4 = Button(f3,text=">>",command=forw)
    l4.grid(row=5,column=7)
    canva=Canvas(f3,height=700,width=584,bg="lightsteelblue2")
    canva.grid(row=1,column=1,columnspan=6,rowspan=9)
    l5 = Button(f3,text="Import",command=insertfun,width=10)
    l5.grid(row=1,column=8)
    l6 = Button(f3,text="Add Item",width=10)
    l6.grid(row=2,column=8)
    l11 = Button(f3,text="Save PDF",command=makepdf,width=10)
    l11.grid(row=10,column=8)
    l7 = Button(f3,text="crop",command=cropp,width=10)
    l7.grid(row=10,column=1)
    l8 = Button(f3,text="Rotate",command=rotatee,width=10)
    l8.grid(row=10,column=2)
    l10 = Button(f3,text="Color",width=10)
    l10.grid(row=10,column=3)
    l9 = Button(f3,text="Delete",command=deletee,width=10)
    l9.grid(row=10,column=4)         





# Tesseract OCR library
def OCR():
    global f3
    if f3:
       f3.destroy()
    f3=Frame(root,bg="green",bd=8)
    f3.grid(row=1,column=0,rowspan=10,columnspan=9)

    file=""
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'

    frame=Frame(f3,width=600,height=300)
    frame.grid(row=1,column=0,columnspan=8,rowspan=5)
    
    canvas=Canvas(frame,bg='#FFFFFF',width=584,height=350,scrollregion=(0,0,1000,1000))
    
    hbar=Scrollbar(frame,orient=HORIZONTAL)
    hbar.pack(side=BOTTOM,fill=X)
    hbar.config(command=canvas.xview)
    vbar=Scrollbar(frame,orient=VERTICAL)
    vbar.pack(side=RIGHT,fill=Y)
    vbar.config(command=canvas.yview)
    frame1=Frame(f3,width=300,height=300)
    frame1.grid(row=6,column=0,columnspan=8,rowspan=4)

    text_area=ScrolledText(frame1,height =20,width=73,font=("TkFixedFont",11))
    text_area.pack(side=LEFT,anchor=NW,expand=True,fill=BOTH)


    canvas.config(width=584,height=300) 
    canvas.config(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
    canvas.pack(side=LEFT,expand=True,fill=BOTH)
    # to add files
    def impo():
        nonlocal file
        file=""
        root.update()
        root.withdraw()
        file=filedialog.askopenfilename()
        root.update()
        root.deiconify()
        io=Image.open(file)
        root.photoImg =  ImageTk.PhotoImage(io)
        imo=canvas.create_image(0,0, anchor=NW, image=root.photoImg)

        #to convert
    def conv(): 
        nonlocal file
        if file!="":
            image = cv.imread(file)
            gray = cv.cvtColor(image, cv.COLOR_BGR2GRAY)
            # check to see if we should apply thresholding to preprocess the
            # image

            gray = cv.threshold(gray, 0, 255,
                    cv.THRESH_BINARY | cv.THRESH_OTSU)[1]
            # make a check to see if median blurring should be done to remove
            # noise

            gray1 = cv.medianBlur(gray, 3)
            # write the grayscale image to disk as a temporary file so we can
            # apply OCR to it
            filename = "{}.png".format(os.getpid())
            cv.imwrite(filename, gray)
            # load the image as a PIL/Pillow image, apply OCR, and then delete
            # the temporary file
            text = pytesseract.image_to_string(Image.open(filename))
            os.remove(filename)
            print(text)
            text_area.insert(INSERT,"maxo")
            text_area.insert(END, text)

            print( text_area.get(1.0, END) )
            # show the output images
            cv.imshow("Image", image)
            cv.imshow("Output", gray)
            cv.waitKey(0)  

    ll1=Button(f3,text="IMPORT IMG",command=impo,width=10)
    ll1.grid(row=1,column=8)
    #ll2=Button(root,text="ADD File",width=10)
    #ll1.grid(row=2,column=8)
    ll3=Button(f3,text="OCR",command=conv,width=10)
    ll3.grid(row=10,column=8)
    #ll4=Button(root,text="IMPORT IMG",command=,width=10)
    #ll4.grid(row=9,column=8)

#def CLEAR():
#converts bulk PPT to respactive pdfs

def PPTtoPDF():
      global f3
      if f3:

         f3.destroy()
      f3=Frame(root,bg="green",bd=8)
      f3.grid(row=1,column=0,rowspan=10,columnspan=9)

      wdFormatPDF = 32
      height=700
      width=600
      canvas=Canvas(f3,height=height,width=width,bg="lightsteelblue2")
      canvas.grid(row=1,column=1,columnspan=6,rowspan=9)
      listtt=[]
      listtt1=[]
      options=["None"]
      text_area=ScrolledText(f3,height =41,width=73,font=("TkFixedFont",11))
      text_area.grid(row=1,column=1,rowspan=9,columnspan =7)
      def importtt():
           listtt.clear()
           listtt1.clear()
           root.update()
           root.withdraw()
           files=filedialog.askopenfilenames(title="insert",filetypes=(("ppt files","*.pptx"),("allfiles","*.*")))
           root.update()
           root.deiconify()
           print("succes")
           
           for imp in files:
     
     
                 imp=imp.replace("/","\\")
                 name=imp.rsplit('\\')[-1].rsplit(".")[0]
                 listtt.append(imp)
                 listtt1.append(name)
           print(listtt,listtt1)
           monoo()
           #addends files in array
      def addfile():
           root.update()
           root.withdraw()
           files=filedialog.askopenfilenames(title="insert",filetypes=(("ppt files","*.pptx"),("allfiles","*.*")))
           root.update()
           root.deiconify()
           
           
           for imp in files:
     
     
                 imp=imp.replace("/","\\")
                 name=imp.rsplit('\\')[-1].rsplit(".")[0]
                 listtt.append(imp)
                 listtt1.append(name) 
           monoo()
      def monoo():
           nonlocal options
           nonlocal listtt1
           options=["None"]
           for i in range(1,len(listtt1)+1):
                options.append(i)
           dropp.set("None")
           inse.set("None")
           drop=OptionMenu(f3,dropp,*options)
           ins=OptionMenu(f3,inse,*options)
           ins.grid(row=3,column=8)
           drop.grid(row=6,column=8)
           more()
      #aranges text in text widget
      def more():
           text_area=ScrolledText(f3,height =41,width=73,font=("TkFixedFont",11))
           text_area.grid(row=1,column=1,rowspan=9,columnspan =7)
           for i in listtt1:
               text_area.insert(INSERT,i+"\n")

           

      def inser():
           if inse.get()!="None":
                  ji=int(inse.get())
                  files=filedialog.askopenfilename()
                  if files!="":
                      i=files.replace("/","\\")
                      name=i.rsplit('\\')[-1].rsplit(".")[0]
                 
                      listtt.insert(ji-1,i)
                      listtt1.insert(ji-1,name)
                      print(listtt,listtt1)
                      monoo()
                  
      def dele():
           nonlocal listtt
           nonlocal listtt1
           if dropp.get()!="None":
                joo=int(dropp.get())
             
                listtt.pop(joo-1)
                listtt1.pop(joo-1)
           monoo()
             
     

      def conv():
           root.update()
           root.withdraw()
           m=filedialog.askdirectory()
           m=m.replace("/","\\")
           root.update()
           root.deiconify()
           power = win32com.client.Dispatch('Powerpoint.Application')

           for jo in range(len(listtt)):
                 powerpoint = power.Presentations.Open(listtt[jo])
                 out_file=listtt1[jo]
                 print(out_file)
                 
                 powerpoint.SaveAs(m+out_file, FileFormat=wdFormatPDF)
                 powerpoint.Close()
      
           power.Quit()
      
      lll1=Button(f3,text="IMPORT PPT",width=10,command=importtt).grid(row=1,column=8)
      lll2=Button(f3,text="ADD FILE",width=10,command=addfile).grid(row=2,column=8)
      lll3=Button(f3,text="CONVERT",width=10,command=conv).grid(row=10,column=8)
      llll4=Button(f3,text="DEL",width=10,command=dele).grid(row=7,column=8)
      llll5=Button(f3,text="Insert",width=10,command=inser).grid(row=4,column=8)
      dropp=StringVar()
      dropp.set("None")
      inse=StringVar()
      inse.set("None")
      ins=OptionMenu(f3,inse,*options)
      ins.grid(row=3,column=8)
      drop=OptionMenu(f3,dropp,*options)
      drop.grid(row=6,column=8)
      #lll1=Button(root,text="IMPORT IMG",width=10)

# document WORD to PDF

def DOCtoPDF():
        global f3
        if f3:
           f3.destroy()
        f3=Frame(root,bg="green",bd=8)
        f3.grid(row=1,column=0,rowspan=10,columnspan=9)
        wdFormatPDF = 17
        height=700
        width=600
        canvas=Canvas(f3,height=height,width=width,bg="lightsteelblue2")
        canvas.grid(row=1,column=1,columnspan=6,rowspan=9)
        text_area=ScrolledText(f3,height =41,width=73,font=("TkFixedFont",11))
        text_area.grid(row=1,column=1,rowspan=9,columnspan =7)
        lis=[]
        li=[]
        options=["None"]

        def importtt():

             lis.clear()
             li.clear()
             root.update()
             root.withdraw()
             files=filedialog.askopenfilenames()
             
             root.update()
             root.deiconify()
             for i in files:
                  
                  
                  i=i.replace("/","\\")
                  name=i.rsplit('\\')[-1].rsplit(".")[0]
                  lis.append(i)
                  li.append(name)
             monoo()  

             print("import",lis,li)
             #maxo()
        def addfile():
             root.update()
             root.withdraw()
             files=filedialog.askopenfilenames()
             
             root.update()
             root.deiconify()
             for i in files:
                  
                  
                  i=i.replace("/","\\")
                  name=i.rsplit('\\')[-1].rsplit(".")[0]
                  lis.append(i)
                  li.append(name)
             monoo() 
             print("add file",lis,li)  
        def more():
           global f3
           text_area=ScrolledText(f3,height =41,width=73,font=("TkFixedFont",11))
           text_area.grid(row=1,column=1,rowspan=9,columnspan =7)
           for i in li:
               text_area.insert(INSERT,i+"\n")

        # def maxo():

        #      imo=filedialog.askopenfilename()
        #      immo=Image.open(imo)
        #      immmo=immo.resize((50,50), Image.ANTIALIAS)
        #      print(os.getcwd())
        #      global canva
        #      img=ImageTk.PhotoImage(immo)
        #      print(imo,immo,immmo,img)
             
             
        #    imago=canva.create_image(0,0,image=img,anchor=NW)

        def monoo():
             nonlocal li
             nonlocal options
             global f3
             options=["None"]
             for i in range(1,len(li)+1):
                  options.append(i)
             dropp.set("None")
             inse.set("None")
             drop=OptionMenu(f3,dropp,*options)
             ins=OptionMenu(f3,inse,*options)
             ins.grid(row=3,column=8)
             drop.grid(row=6,column=8)
             print("monno",lis,li)
             more()
        def inser():
             if inse.get()!="None":
                  ji=int(inse.get())
                  files=filedialog.askopenfilename()
                  if files!="":
                      i=files.replace("/","\\")
                      name=i.rsplit('\\')[-1].rsplit(".")[0]
                 
                      lis.insert(ji-1,i)
                      li.insert(ji-1,name)
                      monoo()
                      print("inser",lis,li)
        def dele():
             nonlocal lis
             nonlocal li
             if dropp.get()!="None":
                joo=int(dropp.get())
             
                lis.pop(joo-1)
                li.pop(joo-1)
             monoo()
             print("dele",lis,li)
         #convert function    
        def cont():
             root.update()
             root.withdraw()
             m=filedialog.askdirectory()
             m=m.replace("/","\\")
             root.update()
             root.deiconify()
             
             word = win32com.client.Dispatch('Word.Application')
             for i in range(len(lis)):
                  doc = word.Documents.Open(lis[i])
                  out_file=m+li[i]
                  print(out_file)
                  doc.SaveAs(out_file, FileFormat=wdFormatPDF)
                  doc.Close()
             word.Quit()
        llll1=Button(f3,text="IMPORT DOC",width=10,command=importtt).grid(row=1,column=8)
        llll2=Button(f3,text="ADD FILE",width=10,command=addfile).grid(row=2,column=8)
        llll3=Button(f3,text="CONVERT",width=10,command=cont).grid(row=10,column=8)
        llll4=Button(f3,text="DEL",width=10,command=dele).grid(row=7,column=8)
        llll5=Button(f3,text="Insert",width=10,command=inser).grid(row=4,column=8)
        dropp=StringVar()
        dropp.set("None")
        inse=StringVar()
        inse.set("None")
        ins=OptionMenu(f3,inse,*options)
        ins.grid(row=3,column=8)
        drop=OptionMenu(f3,dropp,*options)
        drop.grid(row=6,column=8)


def PDFMERGER():
        global f3
        if f3:
           f3.destroy()
        f3=Frame(root,bg="#3A6596",bd=8)
        f3.grid(row=1,column=0,rowspan=10,columnspan=9)
        pdf2merge=[]
        pdfWriter = PyPDF2.PdfFileWriter()
        height=700
        width=600
        canvas=Canvas(f3,height=height,width=width,bg="lightsteelblue2")
        canvas.grid(row=1,column=1,columnspan=6,rowspan=9)
        text_area=ScrolledText(f3,height =41,width=76,font=("TkFixedFont",11))
        text_area.grid(row=1,column=1,rowspan=9,columnspan =7)
        options=["None"]
        #imports files
        def addo():
               pdf2merge.clear()
               root.update()
               root.withdraw()                           # asks users where the PDFs are
               fileso = filedialog.askopenfilenames()
               root.update()
               root.deiconify()
               for izo in fileso:
                   
                   pdf2merge.append(izo)

               monoo()
#adds files
        def addfile():
               root.update()
               root.withdraw()                           # asks users where the PDFs are
               fileso = filedialog.askopenfilenames()
               root.update()
               for izo in fileso:
                   
                   pdf2merge.append(izo) 
               monoo() 

        def more():
           global f3
           text_area=ScrolledText(f3,height =41,width=73,font=("TkFixedFont",11))
           text_area.grid(row=1,column=1,rowspan=9,columnspan =7)
           for i in pdf2merge:
               text_area.insert(INSERT,i+"\n")

        #resets list and rearrangments of file
        def monoo():
                nonlocal pdf2merge
                nonlocal options
                global f3
                options=["None"]
                for i in range(1,len(pdf2merge)+1):
                         options.append(i)
                dropp.set("None")
                inse.set("None")
                drop=OptionMenu(f3,dropp,*options)
                ins=OptionMenu(f3,inse,*options)
                ins.grid(row=3,column=8)
                drop.grid(row=6,column=8)
                more()
                #inserts file at selected index
        def inser():
                if inse.get()!="None":
                      ji=int(inse.get())
                      files=filedialog.askopenfilename()
                      if files!="":
                          pdf2merge.insert(ji-1,files)
                          
                          monoo()
                          print("inser",lis,li)
        #deletes selected file
        def dele():
               nonlocal pdf2merge
               if dropp.get()!="None":
                     joo=int(dropp.get())
                     pdf2merge.pop(joo-1)
               monoo()
                     

        # Ask user for the name to save the file as

        def convo():
        # loop through all PDFs
                for filename in pdf2merge:
                    # rb for read binary
                    pdfFileObj = open(filename, 'rb')
                    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                    # Opening each page of the PDF
                    for pageNum in range(pdfReader.numPages):
                        pageObj = pdfReader.getPage(pageNum)
                        pdfWriter.addPage(pageObj)
                # save PDF to file, wb for write binary
                userfilename = filedialog.asksaveasfilename()
                pdfOutput = open(userfilename + '.pdf', 'wb')
                # Outputting the PDF
                pdfWriter.write(pdfOutput)
                # Closing the PDF writer
                pdfOutput.close()
        
        lllll1=Button(f3,text="IMPORT PDF",width=10,command=addo).grid(row=1,column=8)
        lllll2=Button(f3,text="ADD FILE",width=10,command=addfile).grid(row=2,column=8)
        lllll3=Button(f3,text="CONVERT",width=10,command=convo).grid(row=10,column=8)
        lllll4=Button(f3,text="DEL",width=10,command=dele).grid(row=7,column=8)
        lllll5=Button(f3,text="Insert",width=10,command=inser).grid(row=4,column=8)
        dropp=StringVar()
        dropp.set("None")
        inse=StringVar()
        inse.set("None")
        ins=OptionMenu(f3,inse,*options)
        ins.grid(row=3,column=8)
        drop=OptionMenu(f3,dropp,*options)
        drop.grid(row=6,column=8)

#all buttons and other widgets initializing
#main features Button container frame 
fr=Frame(root,bg="#0F307D",bd=2)
fr.grid(row=0,column=0,rowspan=1,columnspan=5,sticky=N)

l1 = Button(fr,text="IMG to PDF", fg="black", bg="#3A6596",height=1,width=10,bd=2,relief=GROOVE,command=IMGTOPDF)
l2 = Button(fr,text="OCR", fg="black", bg="#3A6596",width=10,command=OCR)
l14=Button(fr,text="PPT to PDF", fg="black", bg="#3A6596",width=10,command=PPTtoPDF)
l14.grid(row=0,column=3)
l1.grid(row=0,column=1)
l2.grid(row=0,column=2)

l15=Button(fr,text="Doc to PDF", fg="black", bg="#3A6596",width=10,command=DOCtoPDF)
l15.grid(row=0,column=4)
l16=Button(fr,text="PDF Merger", fg="black", bg="#3A6596",width=10,command=PDFMERGER)
l16.grid(row=0,column=5)
#l17=Button(root,text="All to PDF ", fg="black", bg="white",width=10)
#l17.grid(row=0,column=6)
lo=os.getcwd()
root.iconbitmap(lo+'\\pdf.ico')




#display frame 
f3=Frame(root,bg="green",bd=2)
f3.grid(row=1,column=0,rowspan=10,columnspan=9)


# to stuck gui in loop
root.mainloop()
