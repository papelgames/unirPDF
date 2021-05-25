from PyPDF2 import PdfFileMerger
from PIL import Image
from datetime import datetime
from pikepdf import _cpphelpers 

import img2pdf
import os, shutil
import comtypes.client
import docx
from win32com import client
import random

class unidor:
    
    def __init__ (self,path_origen,path_destino, path_temporal):
        self.path_origen = path_origen
        self.path_destino = path_destino
        self.path_temporal = path_temporal
        self.subOrigen = [archivo.name for archivo in os.scandir(self.path_origen) if archivo.is_dir()]
        self.extensiones_validas = ['.doc','.docx','.png','.tif','.jpg','.pdf','.jpeg','.tiff']
        self.secuencia = random.randint(1,99)


    def controlExtensiones(self):
        print("Inicio: controlExtensiones")
        secuencia_nombre = 0
        #recorro todas las carpetas y si hay algún archivo no soportado lo guardo en destino  
        #con el nombre de la carpeta mas un numero secuencial, para que se distinga al momento de ser subido
        for carpetas in self.subOrigen:
            for archivo in os.listdir(self.path_origen + carpetas + '\\'):
                if os.path.splitext(archivo)[1] not in self.extensiones_validas:
                    shutil.move(self.path_origen + carpetas + '\\' + archivo, self.path_destino + '\\'+ carpetas + str(secuencia_nombre) + os.path.splitext(archivo)[1] )
                    secuencia_nombre += 1
                    
    def pngToJpg(self):  
        print("Inicio: pngToJpg")
        #recorro cada un de las carpetas dentro de origen 
        for carpetas in self.subOrigen:
            #creo una lista de los archvos png y tiff        
            pngs =  [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith('.png') or archivo.endswith('.tif') or  archivo.endswith('.tiff')]
            #recorro cada uno de los .png y los transformo en .jpg
            if len(pngs) > 0:
                #recorro la lista con los png y los transformo en jpg
                for png_y_tiff in pngs:
                    imagen = Image.open(self.path_origen + carpetas + '\\' + png_y_tiff)
                    rgb_im = imagen.convert('RGB')
                    rgb_im.save( self.path_origen + carpetas + '\\' + png_y_tiff + '.jpg', quality=95)
        pngs =  [archivo for archivo in os.listdir(self.path_origen + '\\') if archivo.endswith('.png') or archivo.endswith('.tif') or  archivo.endswith('.tiff')]
            #recorro cada uno de los .png y los transformo en .jpg
        if len(pngs) > 0:
            #recorro la lista con los png y los transformo en jpg
            for png_y_tiff in pngs:
                imagen = Image.open(self.path_origen  + '\\' + png_y_tiff)
                rgb_im = imagen.convert('RGB')
                rgb_im.save( self.path_origen  + '\\' + png_y_tiff + '.jpg', quality=95)
    
    
    def docxToPdf (self):
        print("Inicio: docxToPdf")
        #recorro cada un de las carpetas dentro de origen
        for carpetas in self.subOrigen:
            #creo una lista de los archvos docx        
            docxs = [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith('.docx') or archivo.endswith('.doc')]
            #controlo que la lista no esté vacía
            if len(docxs) > 0:
                #recorro cada uno de los .docx y .doc y los transformo en .pdf
                for archivo_docx in docxs:
                    out_file = os.path.abspath(self.path_origen + carpetas + '\\' + archivo_docx + '.pdf')
                    word = comtypes.client.CreateObject('Word.Application')
                    doc = word.Documents.Open(self.path_origen + carpetas + '\\' + archivo_docx)
                    doc.SaveAs(out_file, FileFormat=17)                
                    doc.Close()
        docxs = [archivo for archivo in os.listdir(self.path_origen + '\\') if archivo.endswith('.docx') or archivo.endswith('.doc')]
        #recorro la lista imagenes de la carpeta origen. 
        if len(docxs) > 0:
            for archivo_docx in docxs:
                out_file = os.path.abspath(self.path_destino+ '\\' + archivo_docx + str(self.secuencia) + '.pdf')
                word = comtypes.client.CreateObject('Word.Application')
                doc = word.Documents.Open(self.path_origen + '\\' + archivo_docx)
                doc.SaveAs(out_file, FileFormat=17)                
                doc.Close()
                self.secuencia +=1
                       

    def redimensionarJpg (self, path_jpg):
        print("Inicio: redimensionarJPG: " + path_jpg )
        #capturo el tamaño del archivo
        sizefile = os.path.getsize(path_jpg)
        #abro el archivo .jpg
        imagen = Image.open(path_jpg)
        #reduzco la imagen las veces que sea necesario hasta que sizefile se menor a 250kb
        while sizefile>250000:
            imagen = imagen.resize((int(imagen.size[0]*0.90), int(imagen.size[1]*0.90)), Image.ANTIALIAS)
            quality_val = 85
            imagen.save(path_jpg, 'jpeg', quality=quality_val)
            imagen = Image.open(path_jpg)
            sizefile = os.path.getsize(path_jpg)
        
    def jpgToPdf (self):
        print("Inicio: jpgToPdf")
        #recorro cada un de las carpetas dentro de origen
        for carpetas in self.subOrigen:
            #creo una lista con los todos los archivos jpeg y jpeg
            imagenes = sorted([self.path_origen + carpetas + '\\' + archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith(".jpg") or archivo.endswith(".jpeg")])
            #controlo que la lista no esté vacía
            if len(imagenes) > 0:
                #controlo que las archivos .jpg no sean superiores a 250kb
                for imagen_jpg in imagenes:
                    if os.path.getsize(imagen_jpg) >250000:
                        self.redimensionarJpg(imagen_jpg)
                #Uno todos los .jpg los uno en un solo .pdf
                with open(self.path_origen + carpetas + '\\' + 'ZfromJpg.pdf', "wb") as documento:
                    documento.write(img2pdf.convert(imagenes))
        #creo un lista con los archivos .jpg de la carpeta origen. 
        imagenes = sorted([self.path_origen + '\\' + archivo for archivo in os.listdir(self.path_origen + '\\') if archivo.endswith(".jpg") or archivo.endswith(".jpeg")])
        #recorro la lista imagenes de la carpeta origen. 
        if len(imagenes) > 0:
            for imagen in imagenes:
                if os.path.getsize(imagen) >250000:
                    self.redimensionarJpg(imagen)
                #Uno todos los .jpg los uno en un solo .pdf
                nombre_archivo = os.path.basename(imagen)
                with open(self.path_destino + str(self.secuencia) + '- OO' + os.path.splitext(nombre_archivo)[0] + '.pdf', "wb") as documento:
                    documento.write(img2pdf.convert(imagen))
                self.secuencia += 1 


    def unirPdf (self):
        print("Inicio: unirPdf")
        #recorro cada un de las carpetas dentro de origen
        for carpetas in self.subOrigen:
            #creo una lista por cada carpeta con solo los pdf
            pdfs = [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith(".pdf")]
            #creo un objeto para fusionar
            fusionador = PdfFileMerger()
            #ordeno la lista
            pdfs_ordenados = sorted(pdfs)
            #hago el merge con el objeto fusionador
            if len(pdfs) > 0:
                for pdf in pdfs_ordenados: 
                    fusionador.append(open(self.path_origen + carpetas + '\\' + pdf,'rb'))
                with open(self.path_destino + carpetas +'.pdf','wb') as salida: 
                    fusionador.write(salida)
        pdfs = [archivo for archivo in os.listdir(self.path_origen + '\\') if archivo.endswith(".pdf")]
        if len(pdfs) > 0:
            for pdf in pdfs:
                shutil.move(self.path_origen + pdf, self.path_destino + str(self.secuencia) + '-OO' + pdf)
                self.secuencia += 1

    def listarArchivos (self):
        archivos = [archivo for archivo in os.listdir(self.path_destino)]
        return archivos
    
    def listarArchivos2 (self, path_variable):
        archivos = [archivo for archivo in os.listdir(path_variable)]
        return archivos

    def moverCarpeta(self):
        
        for carpetas in self.subOrigen:
            shutil.move(self.path_origen + carpetas, self.path_temporal + carpetas + str(self.secuencia))
            self.secuencia += 1
        archivos = [archivo for archivo in os.listdir(self.path_origen + '\\')]    
        for archivo in archivos:
            shutil.move(self.path_origen + archivo, self.path_temporal + str(self.secuencia) + archivo)
            self.secuencia += 1

    def unirTodos(self):
        self.controlExtensiones()
        self.pngToJpg()
        self.docxToPdf()
        self.jpgToPdf()
        self.unirPdf()
        