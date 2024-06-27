Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'If Target.CountLarge > 1 Then Exit Sub 'this prevents bugs
If Not Intersect(Target, Range("E8:M8")) Is Nothing Then

    If Target.Column = 5 Then FormTab_EmployeeInfo
    If Target.Column = 7 Then FormTab_EmployeeDUI
    If Target.Column = 8 Then FormTab_EmployeeStudies
    If Target.Column = 10 Then FormTab_EmployeeFamily
    If Target.Column = 12 Then FormTab_EmployeeBeneficiaries
    
End If


Dim rng As Range
    Dim cell As Range
    
    ' Define el rango donde deseas aplicar el cambio de color
    Set rng = Me.Range("E86:L91")
    
    ' Restablecer el color de todas las celdas en el rango al color predeterminado
    rng.Interior.ColorIndex = xlNone
    
    ' Si la celda seleccionada está dentro del rango especificado
    If Not Intersect(Target, rng) Is Nothing Then
        ' Cambiar el color de fondo de las celdas en la misma fila dentro del rango
        For Each cell In Intersect(Target.EntireRow, rng)
            cell.Interior.Color = RGB(255, 255, 0) ' Amarillo, puedes cambiarlo al color que prefieras
        Next cell
    End If
End Sub

Sub BuscarEnTabla()

    Dim tabla As Range ' Definir el rango de la tabla
    Dim valorABuscar As String ' Valor ingresado por el usuario
    Dim columnaBusqueda As Integer ' Columna seleccionada en la lista desplegable
    Dim celdaEncontrada As Range ' Celda donde se encuentra la coincidencia
    Dim valorResultado As String ' Valor del resultado que se busca

    ' Obtener el valor ingresado por el usuario
    valorABuscar = InputBox("Ingrese el valor a buscar:")

    ' Obtener la columna seleccionada en la lista desplegable
    columnaBusqueda = Range("ListaDesplegable").Value

    ' Establecer el rango de búsqueda en función de la columna seleccionada
    Select Case columnaBusqueda
        Case 1
            Set tabla = Range("ColumnaA:ColumnaA") ' Buscar en la columna A
        Case 2
            Set tabla = Range("ColumnaB:ColumnaB") ' Buscar en la columna B
        ' ... agregar casos para otras columnas
    End Select

    ' Realizar la búsqueda utilizando el método Range.Find
    Set celdaEncontrada = tabla.Find(what:=valorABuscar, LookIn:=xlValues, _
                                      LookAt:=xlWholeWord, SearchOrder:=xlByRows, _
                                      MatchCase:=xlNonSensitive)

    If Not celdaEncontrada Is Nothing Then
        valorResultado = celdaEncontrada.Value
        MsgBox "El valor encontrado para " & valorABuscar & " en la columna " & columnaBusqueda & " es: " & valorResultado
    Else
        MsgBox "El valor no se encontró en la tabla."
    End If

End Sub
