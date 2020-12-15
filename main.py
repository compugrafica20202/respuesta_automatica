from pymongo import MongoClient
from procesar_cotizacion import procesar_cotizacion
from time import sleep
import os
from variables import PATH_INV_EXCEL, PATH_CODO_EXCEL

if __name__ == "__main__":

    # Conexiones a la base de datos
    client = MongoClient("mongodb+srv://graphcompuser:graphcomppass@cluster0.yvnfe.mongodb.net/"
                         "graphcompprojdb?retryWrites=true&w=majority")

    db = client.graphcompprojdb

    # Tablas de la base de datos
    cotizaciones_tb = db.cotizaciones
    salida_modelacion_codornices_tb = db.salida_modelacion_codornices
    salida_modelacion_invernadero_tb = db.salida_modelacion_invernadero
    precios_tb = db.precios

    # Se eliminan las instancias (si existen)
    script_path = ".\\VBScripts\\closeMainInstances.vbs"
    error = os.system("cscript " + script_path)
    if error != 0:
        print("No existian instancias")

    # Warm up

    # Se crean las instancias globales de Excel e Inventor
    script_path = ".\\VBScripts\\initMainInstances.vbs"
    error = os.system("cscript " + script_path)
    if error != 0:
        raise Exception("No se pudo ejectutar: " + script_path)

    # Se abren los exceles
    script_path = ".\\VBScripts\\initExcels.vbs"
    error = os.system("cscript " + script_path +
                      " " + PATH_INV_EXCEL +
                      " " + PATH_CODO_EXCEL)
    if error != 0:
        raise Exception("No se pudo ejectutar: " + script_path)

    while True:
        try:
            # Busca la primera de las cotizaciones no procesadas
            cotizacion_en_cola = cotizaciones_tb.find({"ha_sido_revisado": False})[0]
            # Si encuentra una cotizacion sin procesar no arrojara la excepcion y se puede proceder
            procesar_cotizacion(cotizacion_en_cola)
        except KeyboardInterrupt:  # Cuando se desea detener el programa
            break  # Termine el loop

        except IndexError:  # Cuando no encuentra cotizacion sin procesar, arrojado en la definicion cotizacion_en_cola
            sleep(20)  # Esperar 20 segundos antes de volver a consultar para no saturar la base de datos
            pass  # Siga

        except Exception as e:  # No se ha pensado en que ninguna otra excepcion pueda ocurrir
            print(e)

    # Se eliminan las instancias
    script_path = ".\\VBScripts\\closeMainInstances.vbs"
    error = os.system("cscript " + script_path)
    if error != 0:
        raise Exception("No se pudo ejectutar: " + script_path)
