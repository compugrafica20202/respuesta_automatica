import pandas as pd
from variables import COTIZACIONES_PATH, EXCEL_COTIZACION_PATH
import os


def calcular_costos(cotizacion, precios_tb):
    # str(cotizacion["_id"]) devuelve el id, el cual es usado como nombre para la carpeta de la cotizacion
    # Por ejemplo, una cotizacion con _id: 10001
    # Se encontrara G:\\Unidades Compartidas\\Computacion_Grafica_2020-II\\Cotizaciones\\10001
    if cotizacion["es_invernadero"]:
        calcular_costos_invernadero(precios_tb, **cotizacion)
    else:
        calcular_costos_codornices(precios_tb, **cotizacion)
    print("Cotizacion generada exitosamente")


def calcular_costos_invernadero(precios_tb, _id, **kwargs):
    file = pd.read_excel(COTIZACIONES_PATH + "\\" + str(_id) + "\\Listado.xls")

    # Num de fila, Referencia, Cantidad
    datos = pd.DataFrame(file, columns=["Reference", "Lenght[m/m^2/QTY]"])
    acumulado = 0.0
    for dato in datos.iloc:
        referencia = dato["Reference"]
        cantidad = dato["Lenght[m/m^2/QTY]"]
        precio_cantidad = precios_tb.find_one({"codigo": referencia})["precio"]
        acumulado += cantidad * precio_cantidad

    script_path = ".\\VBScripts\\formato_cotizacion.vbs"
    print("Ejecutando: " + "cscript " + script_path +
          " " + EXCEL_COTIZACION_PATH +
          " " + str(kwargs["_id"]) +
          " " + str(kwargs["nombre_cliente"]) +
          " " + str(kwargs["es_empresa"]) +
          " " + str(kwargs["cc_o_nit"]) +
          " " + str(kwargs["correo"]) +
          " " + str(kwargs["municipio"]) +
          " " + str(kwargs["departamento"]) +
          " " + str(acumulado))
    # Ruta, IdCotizacion, Cliente, Empresa, NIT, Email, City, State, Total
    # error = os.system("cscript " + script_path +
    #                   " " + EXCEL_COTIZACION_PATH +
    #                   " " + str(kwargs["_id"]) +
    #                   " " + str(kwargs["nombre_cliente"]) +
    #                   " " + str(kwargs["es_empresa"]) +
    #                   " " + str(kwargs["cc_o_nit"]) +
    #                   " " + str(kwargs["correo"]) +
    #                   " " + str(kwargs["municipio"]) +
    #                   " " + str(kwargs["departamento"]) +
    #                   " " + str(acumulado))
    # if error != 0:
    #     raise Exception("En calcular_costos.py, calcular_costos_invernadero: \n"
    #                     "No se pudo ejectutar: " + script_path)


def calcular_costos_codornices(precios_tb, _id, **kwargs):
    file = pd.read_excel(COTIZACIONES_PATH + "\\" + str(_id) + "\\Listado.xlsm", engine="openpyxl")

    # Num de fila, Referencia, Cantidad
    datos = pd.DataFrame(file, columns=["Reference", "valor"])
    acumulado = 0.0
    for dato in datos.iloc:
        referencia = dato["Reference"]
        cantidad = dato["valor"]
        precio_cantidad = precios_tb.find_one({"codigo": referencia})["precio"]
        acumulado += cantidad * precio_cantidad

    script_path = ".\\VBScripts\\formato_cotizacion.vbs"
    print("Ejecutando :" + "cscript " + script_path +
          " " + EXCEL_COTIZACION_PATH +
          " " + str(kwargs["_id"]) +
          " " + str(kwargs["nombre_cliente"]) +
          " " + str(kwargs["es_empresa"]) +
          " " + str(kwargs["cc_o_nit"]) +
          " " + str(kwargs["correo"]) +
          " " + str(kwargs["municipio"]) +
          " " + str(kwargs["departamento"]) +
          " " + str(acumulado))

    # Ruta, IdCotizacion, Cliente, Empresa, NIT, Email, City, State, Total
    # error = os.system("cscript " + script_path +
    #                   " " + EXCEL_COTIZACION_PATH +
    #                   " " + str(kwargs["_id"]) +
    #                   " " + str(kwargs["nombre_cliente"]) +
    #                   " " + str(kwargs["es_empresa"]) +
    #                   " " + str(kwargs["cc_o_nit"]) +
    #                   " " + str(kwargs["correo"]) +
    #                   " " + str(kwargs["municipio"]) +
    #                   " " + str(kwargs["departamento"]) +
    #                   " " + str(acumulado))
    # if error != 0:
    #     raise Exception("En calcular_costos.py, calcular_costos_codornices: \n"
    #                     "No se pudo ejectutar: " + script_path)
