import os
from variables import NOMBRE_INV_EXCEL, NOMBRE_CODO_EXCEL, HOJA_PARAM_INV, HOJA_PARAM_CODO
from variables import NOMBRE_MACRO_INV_EXCEL


def procesar_cotizacion(cotizacion_dict):
    # Ejemplo de cotizacion_dict:
    # {'_id': ObjectId('5fd78194449baa35f4c26fac'),
    # 'tipo_invernadero': False, 'profundidad': 100, 'altura': 120, 'ancho': 170,
    # 'cantidad_lineas': 0, 'cantidad_jaulas_por_linea': 0, 'cantidad_niveles': 0, 'cantidad_aves': 0,
    # 'lineas_enfrentadas': False, 'ha_sido_revisado': False, 'es_invernadero': True,
    # 'nombre_cliente': 'Una Empresa Tal', 'cc_o_nit': 1234567890, 'es_empresa': True,
    # 'correo': 'miempresa@correo.com', 'departamento': 'Atlántico', 'municipio': 'Barranquilla'}
    es_invernadero = cotizacion_dict["es_invernadero"]
    if es_invernadero:
        _procesar_invernadero(**cotizacion_dict)
    else:
        _procesar_alimentadora(**cotizacion_dict)


def _procesar_invernadero(profundidad: int, altura: int, ancho: int, _id, **kwargs):
    """
    Esta funcion efectua los cambios en el excel de referencia y ejecuta el macro correspondiente para la generacion
    del modelado de los INVERNADEROS.
    :param int profundidad: La profundidad en centimetros
    :param int altura: La altura en centimetros
    :param int ancho: El ancho en centimetros
    :param Class ObjectId _id: El id de la cotizacion
        False -> Modelo 1
        True -> Modelo 2
    """
    script_path = ".\\VBScripts\\runExcelInv.vbs"
    error = os.system("cscript " + script_path +
                      " " + NOMBRE_INV_EXCEL +
                      " " + HOJA_PARAM_INV +
                      " " + str(_id) +
                      " " + str(profundidad * 10) +
                      " " + str(altura * 10) +
                      " " + str(ancho * 10) +
                      " " + NOMBRE_MACRO_INV_EXCEL)
    if error != 0:
        raise Exception("En procesar_cotizacion.py, _procesar_invernadero: \n"
                        "No se pudo ejectutar: " + script_path)


def _procesar_alimentadora(cantidad_lineas: int, cantidad_jaulas_por_linea: int, cantidad_niveles: int,
                           cantidad_aves: int, lineas_enfrentadas: bool, _id, **kwargs):
    """
    Esta funcion efectua los cambios en el excel de referencia y ejecuta el macro correspondiente para la generacion
    del modelado de la ALIMENTADORA DE CODORNICES.
    :param int cantidad_lineas: Lineas en el modelado
    :param int cantidad_jaulas_por_linea: Jaulas por linea
    :param int cantidad_niveles: Cantidad de niveles
    :param int cantidad_aves: Cantidad de aves
    :param bool lineas_enfrentadas: Si la lineas son enfrentadas
    """
    script_path = ".\\VBScripts\\runExcelInv.vbs"
    error = os.system("cscript " + script_path +
                      " " + NOMBRE_CODO_EXCEL +
                      " " + HOJA_PARAM_CODO +
                      " " + str(_id) +
                      " " + str(cantidad_lineas) +
                      " " + str(cantidad_jaulas_por_linea) +
                      " " + str(cantidad_niveles) +
                      " " + str(cantidad_aves) +
                      " " + str(int(lineas_enfrentadas)))
    if error != 0:
        raise Exception("En procesar_cotizacion.py, _procesar_alimentadora: \n"
                        "No se pudo ejectutar: " + script_path)
