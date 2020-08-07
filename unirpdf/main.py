from PyPDF2 import PdfFileMerger
from PIL import Image
from datetime import datetime

import img2pdf
import os, shutil


class unidor:
    
    def __init__ (self,path_origen,path_destino, path_temporal):
        self.path_origen = path_origen
        self.path_destino = path_destino
        self.path_temporal = path_temporal
        self.subOrigen = [archivo.name for archivo in os.scandir(self.path_origen) if archivo.is_dir()]
    
    def pngToJpg(self):  #listo
        ya_nombre = 0
        #recorro cada un de las carpetas de origen
        for b in self.subOrigen:
            pngs =  [archivo for archivo in os.listdir(self.path_origen + b + '\\') if archivo.endswith('.png')]
            #recorro cada uno de los .png y los transformo en .jpg
            for i in pngs:
                
                im = Image.open(self.path_origen + b + '\\' + i)
                rgb_im = im.convert('RGB')
                rgb_im.save( self.path_origen + b + '\\' + 'imgmod_' + str(ya_nombre) + '.jpg', quality=95)
                ya_nombre += 1

    def jpgToPdf (self):
        #recorro todas los .jpg 
        ya_nombre = 50
        for b in self.subOrigen:
            imagenes = [self.path_origen + b + '\\' + archivo for archivo in os.listdir(self.path_origen + b + '\\') if archivo.endswith(".jpg")]
            #Uno todos los .jpg los uno en un solo .pdf
            with open(self.path_origen + b + '\\' + "documento_" + str(ya_nombre)+".pdf", "wb") as documento:
                documento.write(img2pdf.convert(imagenes))
                

            ya_nombre += 1

    def unirPdf (self):
        for b in self.subOrigen:
            #print ('por aca: ' + b)
            pdfs = [archivo for archivo in os.listdir(self.path_origen + b + '\\') if archivo.endswith(".pdf")]

            fusionador = PdfFileMerger()

            for pdf in pdfs: 
                
                fusionador.append(open(self.path_origen + b + '\\' + pdf,'rb'))
                
            with open(self.path_destino + b +'.pdf','wb') as salida: 
                fusionador.write(salida)

    def listarArchivos (self):
        archivos = [archivo for archivo in os.listdir(self.path_destino)]
        return archivos
    
    def listarArchivos2 (self, path_variable):
        archivos = [archivo for archivo in os.listdir(path_variable)]
        return archivos

    
    def moverCarpeta(self):
        for i in self.subOrigen:
            #print("por aca: " + i)
            shutil.move(self.path_origen + i, self.path_temporal)
            

    def unirTodos(self):
        self.pngToJpg()
        self.jpgToPdf()
        self.unirPdf()
        self.moverCarpeta()