# proyecto final
# andrea cutzal (1990-20-14754)
# christopher taquiej (1990-20-14790)
# doris yool (1990-20-16050)


# importar la clase para leer
from lector_excel import Lector
import os


# lista notas se inicializa vacia
lista_notas_global = []


# funcion para calcular promedio
def calcular_promedio(notas, cursos):
    # calcular promedio
    promedio = notas / cursos
    # retornar el promedio
    return promedio


# funcion para imprimir estudiante
def imprimir_estudiante(estudiante):
    # imprimiendo las propiedades del estudiante
    print("Promedio: {0:.2f}".format(estudiante['promedio']))
    print(f"Semestre mas alto: {estudiante['semestre_alto']}")
    print(f"Nota mas alta: {estudiante['nota_alta']}")

    return


# funcion con parametro de entrada
def leer_archivo(nombre_archivo):
    # crear un objeto para leer a partir de la clase
    archivo_excel = Lector(nombre_archivo)
    # leer columnas
    columna_semestres = archivo_excel.leer_columna("B")
    columna_notas = archivo_excel.leer_columna("C")
    # eliminar primer elemento porque es un string
    columna_semestres.pop(0)
    columna_notas.pop(0)
    # calculos
    total_notas = sum(columna_notas)
    total_cursos = len(columna_notas)
    semestre_mas_alto = max(columna_semestres)
    nota_mas_alta = max(columna_notas)
    promedio_notas = calcular_promedio(total_notas, total_cursos)
    # agregar notas a la lista
    lista_notas_global.extend(columna_notas)
    # asignar valores
    estudiante = {
        "promedio": promedio_notas,
        "semestre_alto": semestre_mas_alto,
        "nota_alta": nota_mas_alta
    }

    return estudiante


carpeta_entrada = "notas_estudiantes"
# listar todos los archivos de una carpeta
lista_archivos = os.listdir(carpeta_entrada)


total_estudiantes = 0

# recorriendo la lista de archivos
for nombre_archivo in lista_archivos:
    total_estudiantes = total_estudiantes + 1
    # leyendo el archivo de turno
    registro_estudiante = leer_archivo(f"{carpeta_entrada}\\{nombre_archivo}")
    # imprimir resultados
    print(f"=== Estudiante {total_estudiantes} ===")
    imprimir_estudiante(registro_estudiante)

# calculos globales
total_notas_global = sum(lista_notas_global)
total_cursos_global = len(lista_notas_global)
nota_mas_alta_global = max(lista_notas_global)
nota_mas_baja_global = min(lista_notas_global)
promedio_notas_global = calcular_promedio(total_notas_global, total_cursos_global)
# imprimiendo resultados globales
print(f"====================")
print("Promedio total: {0:.2f}".format(promedio_notas_global))
print(f"Cantidad estudiantes: {total_estudiantes}")
print(f"Nota mas baja: {nota_mas_baja_global}")
print(f"Nota mas alta: {nota_mas_alta_global}")
