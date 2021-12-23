from PyPDF2 import PdfFileMerger
from PIL import Image, ImageSequence, ImageFile
from datetime import datetime
from pikepdf import _cpphelpers
from fpdf import FPDF
 
import img2pdf
import os, shutil
import comtypes.client
import docx
from win32com import client
import random
ImageFile.LOAD_TRUNCATED_IMAGES = True

class unidor:
    def __init__ (self,path_origen,path_destino, path_temporal):
        self.path_origen = path_origen
        self.path_destino = path_destino
        self.path_temporal = path_temporal
        self.subOrigen = [archivo.name for archivo in os.scandir(self.path_origen) if archivo.is_dir()]
        self.extensiones_validas = ['.txt','.TXT','.doc','.docx','.png','.tif','.jpg','.pdf','.jpeg','.tiff','.DOC','.DOCX','.PNG','.TIF','.JPG','.PDF','.JPEG','.TIFF']
        self.secuencia = random.randint(1,99)

    def actualizoSubOrigen(self):
        self.subOrigen = [archivo.name for archivo in os.scandir(self.path_origen) if archivo.is_dir()]
    
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
            #creo una lista de los archvos png         
            pngs =  [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith('.png') or archivo.endswith('.PNG')]
            #recorro cada uno de los .png y los transformo en .jpg
            if len(pngs) > 0:
                #recorro la lista con los png y los transformo en jpg
                for png in pngs:
                    imagen = Image.open(self.path_origen + carpetas + '\\' + png)
                    rgb_im = imagen.convert('RGB')
                    rgb_im.save( self.path_origen + carpetas + '\\' + png + '.jpg', quality=95)
        pngs =  [archivo for archivo in os.listdir(self.path_origen + '\\') if archivo.endswith('.png') or archivo.endswith('.PNG')]
        #recorro cada uno de los .png y los transformo en .jpg
        if len(pngs) > 0:
            #recorro la lista con los png y los transformo en jpg
            for png in pngs:
                imagen = Image.open(self.path_origen  + '\\' + png)
                rgb_im = imagen.convert('RGB')
                rgb_im.save( self.path_origen  + '\\' + png + '.jpg', quality=95)
    
    def tiffToJpg (self):
        print("Inicio: tiffToJpg")
        #recorro cada un de las carpetas dentro de origen 
        for carpetas in self.subOrigen:
            #creo una lista de los archvos tiff        
            tiffs =  [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith('.tif') or  archivo.endswith('.tiff') or archivo.endswith('.TIF') or  archivo.endswith('.TIFF')]
            #controlo que haya tiff
            if len(tiffs) > 0:
                #recorro la lista con los .tiff tambien si tienen mas de una pagina, y los transformo en jpg
                for tiff in tiffs:
                    imagen = Image.open(self.path_origen + carpetas + '\\' + tiff)
                    imagen_multiple = ImageSequence.Iterator(imagen)
                    for imagen_unica in imagen_multiple:
                        print ('entro subcarpeta')
                        rgb_im = imagen_unica.convert('RGB')
                        rgb_im.save( self.path_origen + carpetas + '\\' + str(self.secuencia) + tiff + '.jpg', quality=95)
                        self.secuencia += 1
        tiffs =  [archivo for archivo in os.listdir(self.path_origen + '\\') if archivo.endswith('.tif') or  archivo.endswith('.tiff') or archivo.endswith('.TIF') or  archivo.endswith('.TIFF') ]
            #controlo que haya tiff
        if len(tiffs) > 0:
            #recorro la lista con los .tiff tambien si tienen mas de una pagina, y los transformo en jpg
            for tiff in tiffs:
                imagen = Image.open(self.path_origen  + '\\' + tiff)
                imagen_multiple = ImageSequence.Iterator(imagen)
                nombre_archivo = os.path.basename(tiff)
                carpeta = str(self.secuencia) + os.path.splitext(nombre_archivo)[0]
                os.mkdir(self.path_origen + carpeta)
                for imagen_unica in imagen_multiple:
                    rgb_im = imagen_unica.convert('RGB')
                    rgb_im.save( self.path_origen  + carpeta + '\\' + str(self.secuencia) + tiff + '.jpg', quality=95)
                    self.secuencia += 1
                    
    
    def txtToPdf (self):
        print("Inicio: txtToPdf")
        #recorro cada un de las carpetas dentro de origen
        for carpetas in self.subOrigen:
            #creo una lista de los archvos txt
            print (carpetas)
            #pngs =  [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith('.png') or archivo.endswith('.PNG')]
            txts=sorted([self.path_origen + carpetas + '\\'+ archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith(".txt") or archivo.endswith(".TXT")])
            #controlo que la lista no este vacia
            if len(txts)>0:
                for txt in txts:
                    #ABRO ARCHIVO ORIGEN
                    contenido = open(txt,"r")

                    #OBJETO PDF
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", size = 12)
                    pdf.set_auto_page_break("auto", margin=0)
                    line=1

                    #LEO EL TXT
                    for linea in contenido:
                        if linea[-1]==("\n"):
                            linea=linea[:-1]
                            line+=1
                        pdf.cell(200,10,txt=linea, ln=line, align="L")
                    #CREO EL PDF
                    pdf.output(self.path_origen + carpetas + '\\' + "XfromTXT" + str(self.secuencia) + ".pdf")
                    self.secuencia +=1
                    contenido.close
                    
    
    def docxToPdf (self):
        print("Inicio: docxToPdf")
        #recorro cada un de las carpetas dentro de origen
        for carpetas in self.subOrigen:
            #creo una lista de los archvos docx        
            docxs = [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith('.docx') or archivo.endswith('.doc') or archivo.endswith('.DOCX') or archivo.endswith('.DOC')]
            #controlo que la lista no esté vacía
            if len(docxs) > 0:
                #recorro cada uno de los .docx y .doc y los transformo en .pdf
                for archivo_docx in docxs:
                    out_file = os.path.abspath(self.path_origen + carpetas + '\\' + archivo_docx + '.pdf')
                    word = comtypes.client.CreateObject('Word.Application')
                    doc = word.Documents.Open(self.path_origen + carpetas + '\\' + archivo_docx)
                    doc.SaveAs(out_file, FileFormat=17)                
                    doc.Close()
        docxs = [archivo for archivo in os.listdir(self.path_origen + '\\') if archivo.endswith('.docx') or archivo.endswith('.doc') or archivo.endswith('.DOCX') or archivo.endswith('.DOC')]
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
            imagenes = sorted([self.path_origen + carpetas + '\\' + archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith(".jpg") or archivo.endswith(".jpeg") or archivo.endswith(".JPG") or archivo.endswith(".JPEG")])
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
        imagenes = sorted([self.path_origen + '\\' + archivo for archivo in os.listdir(self.path_origen + '\\') if archivo.endswith(".jpg") or archivo.endswith(".jpeg") or archivo.endswith(".JPG") or archivo.endswith(".JPEG")])
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
            pdfs = [archivo for archivo in os.listdir(self.path_origen + carpetas + '\\') if archivo.endswith(".pdf") or archivo.endswith(".PDF") ]
            print (pdfs)
            #creo un objeto para fusionar
            fusionador = PdfFileMerger(strict=False)
            #ordeno la lista
            pdfs_ordenados = sorted(pdfs)
            #hago el merge con el objeto fusionador
            con_errores = []
            if len(pdfs) > 0:
                for pdf in pdfs_ordenados: 
                    try:
                        fusionador.append(open(self.path_origen + carpetas + '\\' + pdf,'rb'))
                    except:
                        con_errores.append(pdf)
                        
                        
                with open(self.path_destino + carpetas +'.pdf','wb') as salida: 
                    fusionador.write(salida)
                for con_error in con_errores:
                    shutil.copy(self.path_origen + carpetas  + '\\' + con_error, self.path_destino + carpetas + 'CON_ERROR' + str(self.secuencia) +'.pdf' )
                    self.secuencia +=1
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
        self.txtToPdf()
        self.tiffToJpg()
        self.docxToPdf()
        self.actualizoSubOrigen()
        self.jpgToPdf()
        self.unirPdf()