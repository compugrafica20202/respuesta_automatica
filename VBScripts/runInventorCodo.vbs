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
NombreLibro = Args(0)
NombreHoja = Args(1)
IdCotizacion = Args(2)
Lineas = CStr(Args(3))
JaulasxLinea = CStr(Args(4))
Niveles = CStr(Args(5))
Aves = CStr(Args(6))
LineasEnfrentadas = CStr(Args(7))
NombreMacro = Args(8)

' La aplicacion de excel es abierta en initMainInstances.vbs
Set excelApp = GetObject(, "Excel.Application")
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("B7").Value = Lineas
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("B8").Value = JaulasxLinea
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("B9").Value = Niveles
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("B10").Value = LineasEnfrentadas
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("A11").Value = IdCotizacion
excelApp.Workbooks(NombreLibro).Activate

excelApp.Run NombreMacro

' Retorna True si existio un error y False si se ejecuta correctamente
If Err.Number <> 0 Then
    WScript.Echo "Error en runInventorCodo: " & Err.Description 
    WScript.Echo Err.Source
    WScript.Quit True
Else
    WScript.Quit False
End If