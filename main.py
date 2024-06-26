from tkinter import *
from tkinter.font import Font
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfile
from PIL import Image
import win32com.client
import os

# window
root = Tk()
root.title('File Converter')
root.configure(background="#8080ff")
root.geometry("740x400")

#berpindahan frame
def raise_frame(frame):
    frame.tkraise()

lucida = Font(
    family = "Lucida Console",
    size = 16,
    weight = "normal",
    )
    
# HOME PAGE
converterpage = Frame(root)
converterpage.place(x=0, y=0, width=740, height=400)
converterpage.configure(background="#8080ff")

# LOGO
logo2 = PhotoImage(file="logo1.png")
labellogo2 = Label(converterpage, image = logo2)
labellogo2.grid(row=0, column=2,pady=10)

# label text
txtjudul = Label (converterpage, text="FILE CONVERTER", width="15", font=lucida, bg="#8080ff", fg="white",)
txtjudul.place(x=270 , y = 200)

# BUTTON WORD TO PDF
wo2pd =Button(converterpage, text="WORD TO PDF", width="15", font=lucida, bg="#ff6e40", fg="white", command=lambda: raise_frame(wo2pdfpage))
wo2pd.grid(row=2, column=1, padx = 20 , pady = 100)

# BUTTON PPT TO PDF
pp2pd =Button(converterpage, text="PPT TO PDF", width="15", font=lucida, bg="#ff6e40", fg="white", command=lambda: raise_frame(ppt2pdfpage))
pp2pd.grid(row=2, column=2, padx = 20)

# BUTTON IMAGE TO PDF
img2pd =Button(converterpage, text="IMAGE TO PDF", width="15", font=lucida, bg="#ff6e40", fg="white", command=lambda: raise_frame(img2pdfpage))
img2pd.grid(row=2, column=3, padx = 20)


# FRAME WORD TO PDF
wo2pdfpage = Frame(root)
wo2pdfpage.place(x=0, y=0, width=740, height=400)
wo2pdfpage.configure(background="#8080ff")

#fungsi pengkonversian word to pdf
def openWord():
    fileword = askopenfile(filetypes = [("Word Files", "*.docx")])
    getfwname = fileword.name
    newwordname = getfwname.replace("/", "\\")
    out_filew = os.path.splitext(newwordname)[0]
    worddoc = win32com.client.Dispatch('Word.Application')
    doc = worddoc.Documents.Open(newwordname)
    doc.SaveAs(out_filew, FileFormat = 17)
    doc.Close()
    worddoc.Quit()
    showinfo("Complete !", "Word to PDF Succesfully Converted")

# logo
logo3 = PhotoImage(file="4.png")
labellogo3 = Label(wo2pdfpage, image = logo3)
labellogo3.grid(row=0, column=2,pady=10)

# label text
txtjudul = Label (wo2pdfpage, text="Mengkonversi File Word ke File PDF", width="34", font=lucida, bg="#8080ff", fg="white",)
txtjudul.place(x=135 , y = 180)

# LABEL CHOOSE FILE
choose_file = Label(wo2pdfpage, text="File", width="15", font=lucida, bg="#1e3d59", fg="white")
choose_file.grid(row=1, column=1, padx = (50,0) , pady = 50)

# TITIK CHOOSE FILE
titik_choose_file = Label(wo2pdfpage, text=":", width="15", font=lucida, bg="#1e3d59", fg="white")
titik_choose_file.grid(row=1, column=2)

# BUTTON BROWSE FILE
btn_browse = Button(wo2pdfpage, text="Browse File...", width="15", font=lucida, bg="#ffc13b", fg="black", relief=RAISED, command=openWord)
btn_browse.grid(row=1, column=3)

# BUTTON BACK TO FILE CONVERTER PAGE
btn_back2 = Button(wo2pdfpage, text="Back", width="5", font="Arial,(12)", command=lambda: raise_frame(converterpage), bg="#1e3d59", fg="white")
btn_back2.grid(row=3, column=2, pady=30)


#FRAME PPT TO PDF
ppt2pdfpage = Frame(root)
ppt2pdfpage.place(x=0, y=0, width=740, height=400)
ppt2pdfpage.configure(background="#8080ff")

#fungsi pengkonversian ppt to pdf
def openPpt():
    fileppt = askopenfile(filetypes=[("Presentation Files", "*.pptx")])
    getfname = fileppt.name
    newpptname = getfname.replace("/", "\\")
    out_file = os.path.splitext(newpptname)[0]
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    pdf = powerpoint.Presentations.Open(newpptname, WithWindow=False)
    pdf.SaveAs(out_file, 32)
    pdf.Close()
    powerpoint.Quit()
    showinfo("Complete !", "PPT to PDF Succesfully Converted")

# LOGO
logo4 = PhotoImage(file="3.png")
labellogo4 = Label(ppt2pdfpage, image = logo4)
labellogo4.grid(row=0, column=2, pady=10)

# label text
txtjudul3 = Label (ppt2pdfpage, text="Mengkonfersi File PowerPoint ke File PDF", width="40", font=lucida, bg="#8080ff", fg="white",)
txtjudul3.place(x=115 , y = 180)

# LABEL CHOOSE FILE
choose_file = Label(ppt2pdfpage, text="File", width="15", font=lucida, bg="#1e3d59", fg="white")
choose_file.grid(row=1, column=1, padx = (50,0) , pady = 50)

# TITIK CHOOSE FILE
titik_choose_file = Label(ppt2pdfpage, text=":", width="15", font=lucida, bg="#1e3d59", fg="white")
titik_choose_file.grid(row=1, column=2)

# BUTTON BROWSE FILE
btn_browse = Button(ppt2pdfpage, text="Browse File...", width="15", font=lucida, bg="#ffc13b", fg="black", relief=RAISED, command=openPpt)
btn_browse.grid(row=1, column=3)

# BUTTON BACK TO FILE CONVERTER PAGE
btn_back2 = Button(ppt2pdfpage, text="Back", width="5", font="Arial,(12)", command=lambda: raise_frame(converterpage), bg="#1e3d59", fg="white")
btn_back2.grid(row=3, column=2, pady=30)


#FRAME IMAGE TO PDF
img2pdfpage = Frame(root)
img2pdfpage.place(x=0, y=0, width=740, height=400)
img2pdfpage.configure(background="#8080ff")

#fungsi pengkonversian image to pdf
def openImg():
    fileimg = askopenfile(filetypes=[("Image Files", ".jpg .jpeg .png")])
    getImgname = fileimg.name
    pdf = Image.open(getImgname)
    if pdf.mode == "RGBA" :
        pdf = pdf.convert("RGB")
    if ".jpg" in getImgname:
        newImgPdfname = getImgname.replace(".jpg", ".pdf")
    if ".jpeg" in getImgname:
        newImgPdfname = getImgname.replace(".jpeg", ".pdf")
    if ".png" in getImgname:
        newImgPdfname = getImgname.replace(".png", ".pdf")
    pdf.save(newImgPdfname)
    showinfo("Complete !", "Image to PDF Succesfully Converted")

# LOGO
logo5 = PhotoImage(file="2.png")
labellogo5 = Label(img2pdfpage, image = logo5)
labellogo5.grid(row=0, column=2, pady=10)

# label text
txtjudul4 = Label (img2pdfpage, text="Mengkonfersi File Image ke File PDF", width="40", font=lucida, bg="#8080ff", fg="white",)
txtjudul4.place(x=100 , y = 180)

# LABEL CHOOSE FILE
choose_file = Label(img2pdfpage, text="File", width="15", font=lucida, bg="#1e3d59", fg="white")
choose_file.grid(row=1, column=1, padx = (50,0) , pady = 50)

# TITIK CHOOSE FILE
titik_choose_file = Label(img2pdfpage, text=":", width="15", font=lucida, bg="#1e3d59", fg="white")
titik_choose_file.grid(row=1, column=2)

# BUTTON BROWSE FILE
btn_browse = Button(img2pdfpage, text="Browse File...", width="15", font=lucida, bg="#ffc13b", fg="black", relief=RAISED, command=openImg)
btn_browse.grid(row=1, column=3)

# BUTTON BACK TO FILE CONVERTER PAGE 
btn_back3 = Button(img2pdfpage, text="Back", width="5", font="Arial,(12)", command=lambda: raise_frame(converterpage), bg="#1e3d59", fg="white")
btn_back3.grid(row=3, column=2, pady=30)

raise_frame(converterpage)

root.mainloop()
