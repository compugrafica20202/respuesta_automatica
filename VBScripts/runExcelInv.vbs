' Funcion que modifica el excel de invernaderos
' Parametros: NombreLibro, NombreHoja, IdCotizacion, Profundidad, Altura, Ancho
' Retorna: True si existio un error y False si se ejecuta correctamente

' Protege el codigo. Permite retornar True si existio un error y False si se ejecuta correctamente
On Error Resume Next 

Dim Args
Dim NombreLibro, NombreHoja, IdCotizacion
Dim Profundidad, Altura, Ancho
Dim NombreMacro

' Argumentos que se pasan al programa al momento de ejecucion
Set Args = WScript.Arguments

' Pasar los parametros a variables
NombreLibro = Args(0)
NombreHoja = Args(1)
IdCotizacion = Args(2)
Profundidad = CStr(Args(3))
Altura = CStr(Args(4))
Ancho = CStr(Args(5))
NombreMacro = Args(6)

' La aplicacion de excel es abierta en initExcel.vbs
Set excelApp = GetObject(, "Excel.Application")
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("B1").Value = Ancho
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("B2").Value = Altura
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("B3").Value = Profundidad
excelApp.Workbooks(NombreLibro).Worksheets(NombreHoja).Range("B5").Value = IdCotizacion
excelApp.Workbooks(NombreLibro).Activate
excelApp.Run NombreMacro

' Retorna True si existio un error y False si se ejecuta correctamente
If Err.Number <> 0 Then
    WScript.Echo "Error en modificarExcelInv: " & Err.Description 
    WScript.Echo Err.Source
    WScript.Quit True
Else
    WScript.Quit False
End If