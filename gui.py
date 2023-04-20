from klasa import *
from tkinter import *
root=Tk()
root.title('Ljudi')
def dodaj_osobu():
    t=Toplevel(root)
    Label(t,text='JMBG').grid(row=0,column=0)
    Label(t,text='Ime').grid(row=1,column=0)
    Label(t,text='Prezime').grid(row=2,column=0)
    Label(t,text='Godine').grid(row=3,column=0)
    Label(t,text='Pol').grid(row=4,column=0)

    e1_t=Entry(t)
    e1_t.grid(row=0,column=1,columnspan=2)
    e2_t=Entry(t)
    e2_t.grid(row=1,column=1,columnspan=2)
    e3_t=Entry(t)
    e3_t.grid(row=2,column=1,columnspan=2)
    e4_t=Entry(t)
    e4_t.grid(row=3,column=1,columnspan=2)

    var2=StringVar(t)
    var2.set('M')
    r1=Radiobutton(t,text='M',variable=var2,value='M')
    r2=Radiobutton(t,text='Z',variable=var2,value='Z')
    r1.grid(row=4,column=1)
    r2.grid(row=4,column=2)

    b_t=Button(t,text='Dodaj',command=lambda:L.dodaj_coveka(e1_t.get(),e2_t.get(),e3_t.get(),e4_t.get(),var2.get()))
    b_t.grid(row=6,column=0,columnspan=3)

def export_table(upit,naziv):
    L.kreiraj_upit(upit)
    L.get_sql()
    L.export_excel(naziv)

menubar=Menu(root)
dodaj_menu=Menu(menubar,tearoff=0)
dodaj_menu.add_command(label='Dodaj coveka',command=lambda:dodaj_osobu())
menubar.add_cascade(label='Dodaj',menu=dodaj_menu)

m="SELECT * FROM LJUDI WHERE POL='M' "
b_m=Button(root,text='Export M',height=3,command=lambda:export_table(m,'MUSKARCI'))
b_m.grid(row=0,column=0,rowspan=2)

z="SELECT * FROM LJUDI WHERE POL='Z' "
b_z=Button(root,text='Export Z',height=3,command=lambda:export_table(z,'ZENE'))
b_z.grid(row=3,column=0,rowspan=2)


var1=StringVar()
var1.set('ORDER BY JMBG ASC')
r1=Radiobutton(root,text='Ime',variable=var1,value='ORDER BY IME ASC')
r2=Radiobutton(root,text='Prezime',variable=var1,value='ORDER BY PREZIME ASC')
r3=Radiobutton(root,text='Godine',variable=var1,value='ORDER BY GODINE ASC')

r1.grid(row=0,column=1)
r2.grid(row=1,column=1)
r3.grid(row=2,column=1)

b_e=Button(root,text="Export",command=lambda:export_table("SELECT * FROM LJUDI {}".format(var1.get()),"Ljudi_Sort"))
b_e.grid(row=3,column=1)

root.config(menu=menubar)

mainloop()