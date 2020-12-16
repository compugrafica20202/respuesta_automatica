' Retorna: True si existio un error y False si se ejecuta correctamente

' Protege el codigo. Permite retornar True si existio un error y False si se ejecuta correctamente
On Error Resume Next 

Dim Args
Dim NombreLibro, NombreHoja, IdCotizacion
Dim Clientes, Empresa, NIT, Email
Dim City, State
Dim Total

' Argumentos que se pasan al programa al momento de ejecucion
Set Args = WScript.Arguments

' Pasar los parametros a variables
IdCotizacion = Args(1)
Clientes = Args(2)
Empresa = Args(3)
NIT = Args(4)
Email = Args(5)
City = Args(6)
State = Args(7)
Total = Args(8)

' La aplicacion de excel es abierta en initMainInstances.vbs
Set excelApp = GetObject(, "Excel.Application")

excelApp.Workbooks("Formato_Cotizaion.xlsm").Activate

excelApp.Run "Exportar_cotizacion", IdCotizacion, Clientes, Empresa, NIT, Email, City, State, Total

' Retorna True si existio un error y False si se ejecuta correctamente
If Err.Number <> 0 Then
    WScript.Echo "Error en runInventorCodo: " & Err.Description 
    WScript.Echo Err.Source
    WScript.Quit True
Else
    WScript.Echo "runInventorCodo.vbs ejecutado exitosamente"
    WScript.Quit False
End If