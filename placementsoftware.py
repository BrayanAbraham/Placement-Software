from tkinter import *
from files.classes import Software

if __name__ == "__main__":
    top = Tk()
    height = top.winfo_screenheight()
    width = top.winfo_screenwidth()
    top.geometry('%sx%s' % (int(width*0.5), 650))
    top.resizable(0, 0)
    canvas = Canvas(top, width=int(width*0.5), height=650)
    canvas.pack()
    render = PhotoImage(file="bg.gif", master=top)
    canvas.create_image(350, 450, image=render)
    canvas.image = render
    root = Frame(top, height=750, width=500)
    root.place(x=100, y=25)
    Software(root)
    top.title("Placement Software")
    root.mainloop()
