' Funcion que modifica el excel de la alimentadora de codornices
' Parametros: NombreLibro, NombreHoja, IdCotizacion, Lineas, JaulasxLinea, Niveles, Aves, LineasEnfrentadas
' Retorna: True si existio un error y False si se ejecuta correctamente

' Protege el codigo. Permite retornar True si existio un error y False si se ejecuta correctamente
On Error Resume Next 

Dim Args
Dim NombreLibro, NombreHoja, IdCotizacion
Dim Lineas, JaulasxLinea, Niveles, Aves
Dim LineasEnfrentadas
Dim NombreMacro

' Argumentos que se pasan al programa al momento de ejecucion
Set Args = WScript.Arguments

' Pasar los parametros a variables
'NombreLibro = Args(0)
'NombreHoja = Args(1)
'IdCotizacion = Args(2)
'Lineas = CStr(Args(3))
'JaulasxLinea = CStr(Args(4))
'Niveles = CStr(Args(5))
'Aves = CStr(Args(6))
'LineasEnfrentadas = CStr(Args(7))
'NombreMacro = Args(8)

' La aplicacion de excel es abierta en initMainInstances.vbs
Set excelApp = GetObject(, "Excel.Application")
excelApp.Workbooks(Args(0)).Activate
excelApp.Workbooks(Args(0)).Worksheets(Args(1)).Range("B7").Value = "HOLA" 'Lineas
excelApp.Workbooks(Args(0)).Worksheets(Args(1)).Range("B8").Value = "HOLA" 'JaulasxLinea
excelApp.Workbooks(Args(0)).Worksheets(Args(1)).Range("B9").Value = "Hola" 'Niveles
excelApp.Workbooks(Args(0)).Worksheets(Args(1)).Range("B10").Value = "Holaa" 'LineasEnfrentadas
excelApp.Workbooks(Args(0)).Worksheets(Args(1)).Range("A11").Value = "Hola" 'IdCotizacion

' excelApp.Run NombreMacro

' Retorna True si existio un error y False si se ejecuta correctamente
If Err.Number <> 0 Then
    WScript.Echo "Error en runInventorCodo: " & Err.Description 
    WScript.Echo Err.Source
    WScript.Quit True
Else
    WScript.Echo "runInventorCodo.vbs ejecutado exitosamente"
    WScript.Quit False
End If