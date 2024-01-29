import shelve
import openpyxl
from random import randrange
import os
import time

"""variables de openpyxl, no se usan salvo que cambie el excel, en ese caso revisar abajo de todo.
path = "C:/...path/bibliotecaDeEjemplo.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
m_row = sheet_obj.max_row"""

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
BBProvisoriaDeIsLeidos = []
numeroAlAzar = 0
yes_choices = ['yes', 'y', 'YES', 'Y', 'Yes', 'Si', 'S', 's', 'si']


def print_slow(str):
    for letter in str:
        print(letter, end='', flush=True)
        time.sleep(0.01)

def saludaAlUsuario():
    print("")
    print_slow("\n Hola, vamos a usar un número al azar para elegir un libro de la biblioteca, presiona enter para continuar \n")
    input("")    

def abreLaShelve():
    global my_shelve
    #el writeback=true supuestamente no es necesario y ralentiza el programa
    #sin embargo, cuando no lo puse, sencillamente no pude modificar la shelve
    my_shelve = shelve.open("mydata.db", writeback=True)

def eligeNumeroAlAzar():
    global numeroAlAzar
    global libroAlAzar

    numeroAlAzar = randrange(len(my_shelve["biblioteca"])+1)
    libroAlAzar = my_shelve["biblioteca"][numeroAlAzar]
    
def separaIsLeidos():
    enviaTodosLosIsLeidoABBProvisoriaDeIsLeidos() if libroAlAzar.isLeido else comunicaLibroAsignadoAlAzar()

def enviaTodosLosIsLeidoABBProvisoriaDeIsLeidos():
    global BBProvisoriaDeIsLeidos
    BBProvisoriaDeIsLeidos.append(libroAlAzar) 
    eligeNumeroAlAzar()
    separaIsLeidos()

def comunicaLosLibrosLeidosQueSalieron():
    if any(BBProvisoriaDeIsLeidos):
        print_slow("\n ...bueno, primero salieron estos libros, que ya leiste \n")
        for libro in BBProvisoriaDeIsLeidos:
            print("   -> "+libro.titulo)
        detieneLaFuncionSiHayErrorEnIsLeidos()

def detieneLaFuncionSiHayErrorEnIsLeidos():
    errorEnIsLeidos = input("\n Si entre estos libros hay uno que querías leer ahora, ingresa 'Y' para detener el programa: \n")
    if errorEnIsLeidos in yes_choices:
        print_slow("\n Ok, podés leer ese libro entonces, que lo disfrutes. \n"), cierraLaShelve(), despideAlUsuario(), cierraElPrograma()

def comunicaLibroAsignadoAlAzar():
    comunicaLosLibrosLeidosQueSalieron()
    print_slow("\n ...sin contar libros ya leídos, te salió sorteado el número "+str(numeroAlAzar)+", al que le corresponde el siguiente libro: \n")
    print_slow(f"""{libroAlAzar.autor}, {libroAlAzar.titulo}, {libroAlAzar.genero}, {libroAlAzar.seccion}\n""")
    
    
    indicaSiVaALeer = input("\n Según la base de datos, aún no has leído este libro. ¿Vas a leerlo ahora? Y/N: ")
    if indicaSiVaALeer in yes_choices:
        print_slow("\n ¡Excelente! Que lo disfrutes.\n")
        modificaLibroIsLeido()
    else:
        os.system('cls')
        print("\n bueno, elijamos otro...\n"), eligeNumeroAlAzar(), separaIsLeidos()

def modificaLibroIsLeido():
    libroAlAzar.isLeido = True

def cierraLaShelve():
    my_shelve.close()
    
def despideAlUsuario():
    input("\n Cerrando el programa, espero que te haya servido. presiona enter para terminar.\n")
    cierraElPrograma()

def cierraElPrograma():
    exit(0)

if __name__ == "__main__":
    saludaAlUsuario(),
    abreLaShelve(),
    eligeNumeroAlAzar(),
    separaIsLeidos(),
    cierraLaShelve(),
    despideAlUsuario(),
    cierraElPrograma(),



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


