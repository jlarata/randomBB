import shelve
import openpyxl
from random import randrange

#variables de openpyxl, no se usan salvo que cambie el excel, en ese caso revisar abajo de todo.
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
BBProvisoriaDeIsLeidos = []
numeroAlAzar = 0

yes_choices = ['yes', 'y', 'YES', 'Y', 'Yes', 'Si', 'S', 's', 'si']


#funciones
def saludaAlUsuario():
    print("")
    input("Hola, vamos a elegir un número al azar y elegir un libro de la biblioteca, presiona enter para continuar")
    print("")

def abreLaShelve():
    global my_shelve
    #el writeback=true supuestamente no es necesario y ralentiza el programa
    #sin embargo, cuando no lo puse, sencillamente no pude modificar la shelve
    my_shelve = shelve.open("mydata.db", writeback=True)

def eligeNumeroAlAzar():
    global numeroAlAzar
    global libroAlAzar
    #global numeroParaPruebas
    #numeroParaPruebas = 5

    numeroAlAzar = randrange(len(my_shelve["biblioteca"])+1)
    libroAlAzar = my_shelve["biblioteca"][numeroAlAzar]
    

def separaIsLeidos():
    enviaTodosLosIsLeidoABBProvisoriaDeIsLeidos() if libroAlAzar.isLeido else comunicaLibroAsignadoAlAzar()
    #envialosleidosaBB para pruebas
    #enviaTodosLosIsLeidoABBProvisoriaDeIsLeidos() if my_shelve["biblioteca"][numeroParaPruebas].isLeido else comunicaLibroAsignadoAlAzar()

def enviaTodosLosIsLeidoABBProvisoriaDeIsLeidos():
    global BBProvisoriaDeIsLeidos
    BBProvisoriaDeIsLeidos.append(Libro(libroAlAzar.numero, libroAlAzar.autor, libroAlAzar.titulo, libroAlAzar.genero, libroAlAzar.seccion, True)) 
    #BBProvisoria para pruebas
    #BBProvisoriaDeIsLeidos.append(Libro(my_shelve["biblioteca"][numeroParaPruebas].numero, my_shelve["biblioteca"][numeroParaPruebas].autor, my_shelve["biblioteca"][numeroParaPruebas].titulo, my_shelve["biblioteca"][numeroParaPruebas].genero, my_shelve["biblioteca"][numeroParaPruebas].seccion, True)) 
    eligeNumeroAlAzar()
    separaIsLeidos()

def comunicaLosLibrosLeidosQueSalieron():
    if any(BBProvisoriaDeIsLeidos):
        print("...bueno, primero salieron estos libros, que ya leiste")
        for libro in BBProvisoriaDeIsLeidos:
            print("   -> "+libro.titulo)
        print("")
        detieneLaFuncionSiHayErrorEnIsLeidos()

def detieneLaFuncionSiHayErrorEnIsLeidos():
    errorEnIsLeidos = input("Si entre estos libros hay uno que querías leer ahora, ingresa 'Y' para detener el programa: ")
    if errorEnIsLeidos in yes_choices:
        print(""), print("Ok, podés leer ese libro entonces, que lo disfrutes."), cierraLaShelve(), despideAlUsuario(), cierraElPrograma()

def comunicaLibroAsignadoAlAzar():
    comunicaLosLibrosLeidosQueSalieron()
    print("...bien, luego te salió sorteado el número "+str(numeroAlAzar)+", al que le corresponde el siguiente libro: ")
    print("")
    print(" "+(libroAlAzar.autor)+", "+(libroAlAzar.titulo)+", "+(libroAlAzar.genero)+", "+((libroAlAzar.seccion)))
    print("")
    #print para pruebas
    #print("...bien, te salió sorteado el número "+str(numeroParaPruebas)+", al que le corresponde el siguiente libro: ")
    #print("")
    #print(" "+(my_shelve["biblioteca"][numeroParaPruebas].autor)+", "+(my_shelve["biblioteca"][numeroParaPruebas].titulo)+", "+(my_shelve["biblioteca"][numeroParaPruebas].genero)+", "+((my_shelve["biblioteca"][numeroParaPruebas].seccion)))
    
    
    #no_choices = ['no', 'n', 'No', 'N']
    indicaSiVaALeer = input("Según la base de datos, aún no has leído este libro. ¿Vas a leerlo ahora? Y/N: ")
    if indicaSiVaALeer in yes_choices:
        print(""), print("¡Excelente! Que lo disfrutes.")
        modificaLibroIsLeido()
    else:
        print(""), print("bueno, elijamos otro..."), print(""), eligeNumeroAlAzar(), separaIsLeidos()

def modificaLibroIsLeido():
    libroAlAzar.isLeido = True

def cierraLaShelve():
    my_shelve.close()
    
def despideAlUsuario():
    print("")
    input("Cerrando el programa, espero que te haya servido. presiona enter para terminar.")
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


