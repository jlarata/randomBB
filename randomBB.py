import shelve
from openpyxl import Workbook
import openpyxl
from random import randrange


#variables de openpyxl
path = "C:/Users/usuario/Desktop/ariel/Hackerwoman/Proyectos/randomBB/biblioteca.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
m_row = sheet_obj.max_row

#objeto y lista de objetos
class Libro:
    def __init__(self, numero, autor, titulo, genero, seccion, isLeido):
        self.numero = numero
        self.autor = autor
        self.titulo = titulo
        self.genero = genero
        self.seccion = seccion
        self.isLeido = isLeido
biblioteca = []

print("")
input("Hola, vamos a elegir un número al azar y elegir un libro de la biblioteca, presiona enter para continuar")
print("")


my_shelve = shelve.open("mydata.db")


"""método general para crear una lista, tomar todos los datos del excel y guardarlos en la shelve
### crea la lista de objetos

for i in range(870):
    biblioteca.append(Libro(i+1, "autor", "titulo", "genero", "seccion", False))

###itera la columna 2
for i in range(1, m_row + 1):
    cell_obj = sheet_obj.cell(row = i, column = 2)
    ###modifica el atributo "autor" de cada elemento de [biblioteca]
    if(cell_obj.value):
        biblioteca[i-1].autor = str(cell_obj.value)
    else:
        biblioteca[i-1].autor = "..."
    
    ###print(str(biblioteca[i-1].numero)+biblioteca[i-1].autor)

###itera la columna 3
for i in range(1, m_row + 1):
    cell_obj = sheet_obj.cell(row = i, column = 3)
    ###modifica el atributo "autor" de cada elemento de [biblioteca]
    if(cell_obj.value):
        biblioteca[i-1].titulo = str(cell_obj.value)
    else:
        biblioteca[i-1].titulo = "..."

###itera la columna 4
for i in range(1, m_row + 1):
    cell_obj = sheet_obj.cell(row = i, column = 4)
    ###modifica el atributo "autor" de cada elemento de [biblioteca]
    if(cell_obj.value):
        biblioteca[i-1].genero = str(cell_obj.value)
    else:
        biblioteca[i-1].genero = "..."

###itera la columna 5
for i in range(1, m_row + 1):
    cell_obj = sheet_obj.cell(row = i, column = 5)
    ###modifica el atributo "autor" de cada elemento de [biblioteca]
    if(cell_obj.value):
        biblioteca[i-1].seccion = str(cell_obj.value)
    else:
        biblioteca[i-1].seccion = "..."

### asigna la lista a la shelve
my_shelve["biblioteca"] = biblioteca

### iprime la lista guardada en la shelve con dos valores: número y autor
for libro in my_shelve["biblioteca"]:
    print(str(libro.numero)+", "+(libro.autor)+", "+(libro.titulo)+", "+((my_shelve["y"]).genero)+", "+((my_shelve["y"]).seccion)+", "+str(((my_shelve["y"]).isLeido)))"""

numeroAlAzar = randrange(len(my_shelve["biblioteca"])+1)

print("Bien, te salió sorteado el número "+str(numeroAlAzar)+", al que le corresponde el siguiente libro: ")
print("")
print(" "+(my_shelve["biblioteca"][numeroAlAzar].autor)+", "+(my_shelve["biblioteca"][numeroAlAzar].titulo)+", "+(my_shelve["biblioteca"][numeroAlAzar].genero)+", "+((my_shelve["biblioteca"][numeroAlAzar].seccion)))


my_shelve.close()
print("")
input("cerrando el programa, espero que te haya servido. presiona Enter para salir.")