from tkinter import *
import os
from main import unidor

origen = r"D:\\Documentos\\Programas\\testeos\\UPDF\origen\\"
destino = r"D:\\Documentos\\Programas\\testeos\\UPDF\\destino\\"
temporal = r"D:\\Documentos\\Programas\\testeos\\UPDF\temporal\\"


unPdf = unidor(origen,destino,temporal)

raiz = Tk()

raiz.title('Unificador de PDF')
frameUno = Frame(raiz)
frameUno.pack()
#pngs =  [archivo for archivo in os.listdir('./') if archivo.endswith('.png')]

labelCarpetaOrigen = Label(frameUno, text = 'Siniestros a unir')
labelCarpetaOrigen.pack()

listaCarpetaOrigen = Listbox(frameUno)
listaCarpetaOrigen.insert (0,*unPdf.subOrigen)
listaCarpetaOrigen.pack()


labelOrigen = Label(frameUno, text = 'Archivos a unir')
labelOrigen.pack()

listaOrigen = Listbox(frameUno)

listaOrigen.pack()

labelDestino = Label(frameUno, text = 'Archivos unidos')
labelDestino.pack()

listaDestino = Listbox(frameUno)
listaDestino.insert (0,*unPdf.listarArchivos())
listaDestino.pack()


botonIniciar = Button(frameUno,text= 'Unir', command = unPdf.unirTodos)
botonIniciar.pack()
botonMover = Button(frameUno,text= 'Mover', command = unPdf.moverCarpeta)
botonMover.pack()
def completaListbox (*x):

    listaOrigen.delete(0,END)
    try:
        orden = unPdf.subOrigen[listaCarpetaOrigen.curselection()[0]]
        archi = [archivo for archivo in unPdf.listarArchivos2(origen + orden)]
        
        listaOrigen.insert (0,*archi)
    except:
        pass

    
listaCarpetaOrigen.bind("<<ListboxSelect>>",completaListbox)


raiz.mainloop()
