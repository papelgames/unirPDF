from PyPDF2 import PdfFileMerger
from PIL import Image
from datetime import datetime
from pikepdf import _cpphelpers 

import img2pdf
import os, shutil
import comtypes.client
import docx

class unidor:
    
    def __init__ (self,path_origen,path_destino, path_temporal):
        self.path_origen = path_origen
        self.path_destino = path_destino
        self.path_temporal = path_temporal
        self.subOrigen = [archivo.name for archivo in os.scandir(self.path_origen) if archivo.is_dir()]
    
    def pngToJpg(self):  #listo
        print("pngToJpg")
        #recorro cada un de las carpetas dentro de origen 
        for carpetas in self.subOrigen:
            #creo una lista de los archvos png y tiff        
            pngs =  [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith('.png') or archivo.endswith('.tif')]
            #recorro cada uno de los .png y los transformo en .jpg
            if len(pngs) > 0:
                #recorro la lista con los png y los transformo en jpg
                for png_y_tiff in pngs:
                    imagen = Image.open(self.path_origen + carpetas + '\\' + png_y_tiff)
                    rgb_im = imagen.convert('RGB')
                    rgb_im.save( self.path_origen + carpetas + '\\' + png_y_tiff + '.jpg', quality=95)
               
    def docxToPdf (self):
        print("docxToPdf")
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

    def jpgToPdf (self):
        print("jpgToPdf")
        #recorro cada un de las carpetas dentro de origen
        for carpetas in self.subOrigen:
            #creo una lista con los todos los archivos jpeg y jpeg
            imagenes = sorted([self.path_origen + carpetas + '\\' + archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith(".jpg") or archivo.endswith(".jpeg")])
            
            #controlo que la lista no esté vacía
            if len(imagenes) > 0:
                #Uno todos los .jpg los uno en un solo .pdf
                with open(self.path_origen + carpetas + '\\' + 'ZfromJpg.pdf', "wb") as documento:
                    documento.write(img2pdf.convert(imagenes))
                  
    def unirPdf (self):
        print("unirPdf")
        #recorro cada un de las carpetas dentro de origen
        for carpetas in self.subOrigen:
            #creo una lista por cada carpeta con solo los pdf
            pdfs = [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith(".pdf")]
            #creo un objeto para fusionar
            fusionador = PdfFileMerger()
            #ordeno la lista
            pdfs_ordenados = sorted(pdfs)
            #hago el merge con el objeto fusionador
            for pdf in pdfs_ordenados: 
                fusionador.append(open(self.path_origen + carpetas + '\\' + pdf,'rb'))
            with open(self.path_destino + carpetas +'.pdf','wb') as salida: 
                fusionador.write(salida)

    def listarArchivos (self):
        archivos = [archivo for archivo in os.listdir(self.path_destino)]
        return archivos
    
    def listarArchivos2 (self, path_variable):
        archivos = [archivo for archivo in os.listdir(path_variable)]
        return archivos

    
    def moverCarpeta(self):
        for carpetas in self.subOrigen:
            shutil.move(self.path_origen + carpetas, self.path_temporal)
            

    def unirTodos(self):
        self.pngToJpg()
        self.docxToPdf()
        self.jpgToPdf()
        self.unirPdf()
        
        #self.moverCarpeta()