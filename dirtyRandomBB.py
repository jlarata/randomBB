import shelve
from random import randrange
import os
import time

#objeto y lista de objetos
class Libro:
    def __init__(self, numero, autor, titulo, genero, seccion, isLeido):
        self.numero = numero
        self.autor = autor
        self.titulo = titulo
        self.genero = genero
        self.seccion = seccion
        self.isLeido = isLeido

#funcion para imprimir de a poco. Menos sleep =  más rápido.
def print_slow(str):
    for letter in str:
        print(letter, end='', flush=True)
        time.sleep(0.03)
        
librosLeidos = []
#esto es un poco gracioso
yes_choices = ['yes', 'y', 'YES', 'Y', 'Yes', 'Si', 'S', 's', 'si']

#Abre La Shelve
my_shelve = shelve.open("mydata.db", writeback=True)
biblioteca = my_shelve["biblioteca"]
 
#saluda Al Usuario
print_slow("\n Hola, vamos a elegir un número al azar y elegir un libro de la biblioteca, presiona enter para continuar \n")
input()

elejido = False

while not elejido:
    #Elijo libro
    numeroAlAzar = randrange(len(biblioteca)+1)
    libroAlAzar = biblioteca[numeroAlAzar]
    
    #Si está leído, guardo y repito hasta encontrar uno no leído
    while libroAlAzar.isLeido:
        librosLeidos.append(libroAlAzar) 
        numeroAlAzar = randrange(len(biblioteca)+1)
        libroAlAzar = biblioteca[numeroAlAzar]

    #Muestro los que descarté por leidos
    if librosLeidos:
        print_slow("...bueno, primero salieron estos libros, que ya leiste \n")
        for libro in librosLeidos:
            print_slow("   -> "+libro.titulo)
        
    #Muestro el libro seleccionado
    print_slow(f"""\n...bien, luego te salió sorteado el número {numeroAlAzar}, al que le corresponde el siguiente libro:\n 
        {libroAlAzar.autor}, {libroAlAzar.titulo}, {libroAlAzar.genero}, {libroAlAzar.seccion}
        \n""")

    #Confirmo
    print_slow("\n Según la base de datos, aún no has leído este libro. ¿Vas a leerlo ahora? Y/N: \n")
    indicaSiVaALeer = input()   
    
    #Si se confirma, termino el proceso, si no vuelvo a empezar.
    if indicaSiVaALeer in yes_choices:
        print_slow("\n ¡Excelente! Que lo disfrutes. \n")
        libroAlAzar.isLeido = True
        elejido = True
        print_slow("cerrando el programa, espero que te haya servido. presiona Enter para salir.")
        input()
    else:
        #limpio la pantalla
        os.system('cls')        
        print_slow("\n bueno, elijamos otro...\n") 
        
my_shelve.close()
