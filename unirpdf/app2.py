import os
from motor import unidor
import sys

#D:\\Documentos\\Programas\\testeos\\UPDF\\
origen = sys.argv[2] + "\\origen\\"
destino = sys.argv[2] + "\\destino\\"
temporal = sys.argv[2] + "\\temporal\\"

unPdf = unidor(origen,destino,temporal)
if sys.argv[1] == "up":
    unPdf.unirTodos()
elif sys.argv[1] == "mv":
    unPdf.moverCarpeta()
