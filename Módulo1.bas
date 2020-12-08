Attribute VB_Name = "Módulo1"
Sub Generar()
'Primera Validacion Nombre de Tienda
If Cells(4, 4).Value <> "Nombre Tienda" And Cells(4, 4).Value <> "" Then
'Llamada a sub
ValidarHoja
'Variables
Dim EncontrarSeccion, EncontrarTotal As Boolean
Dim i, Inicio, fin, Margen, j As Integer
i = 1
Margen = 4
EncontrarSeccion = True
EncontrarTotal = True
'Crear Hoja
Sheets.Add.Name = "Rendicion"
Sheets("arqueo de caja").Select


Inicio = EncontrarIni()
j = Inicio
fin = EncontrarFini(j)


'Traspaso de datos
For k = 1 To 6

For j = Inicio To fin

Sheets("Rendicion").Cells(j - Inicio + Margen, k).Value = Cells(j, k + 1).Value
Sheets("Rendicion").Cells(j - Inicio + Margen, k).Interior.ColorIndex = Cells(j, k + 1).Interior.ColorIndex

Sheets("Rendicion").Cells(j - Inicio + Margen, k).Font.Name = Cells(j, k + 1).Font.Name
Sheets("Rendicion").Cells(j - Inicio + Margen, k).Font.Bold = Cells(j, k + 1).Font.Bold
Sheets("Rendicion").Cells(j - Inicio + Margen, k).Font.Size = Cells(j, k + 1).Font.Size
Sheets("Rendicion").Cells(j - Inicio + Margen, k).Font.ColorIndex = Cells(j, k + 1).Font.ColorIndex


Next j
Next k

Sheets("Rendicion").Select


Else
MsgBox "Favor ingresar nombre de Tienda", vbExclamation, "Error nombre Tienda"

End If


End Sub


Sub eliminar()
'Eliminar Rendicion
Sheets("Rendicion").Delete

End Sub
Sub encabezado(mes As Integer)
'Realizar encabezado
Dim meses
meses = Array("", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
Dim fechaActual As Date
fechaActual = Date
Sheets("Rendicion").Cells(1, 1).Value = "Tienda:"
Sheets("Rendicion").Cells(1, 1).Font.Bold = True
Sheets("Rendicion").Cells(2, 1).Value = "Fecha:"
Sheets("Rendicion").Cells(2, 1).Font.Bold = True
Sheets("Rendicion").Cells(3, 1).Value = "Periodo:"
Sheets("Rendicion").Cells(3, 1).Font.Bold = True

Sheets("Rendicion").Cells(1, 2).Value = Cells(4, 4).Value

Sheets("Rendicion").Cells(2, 2).Value = fechaActual

Sheets("Rendicion").Cells(3, 2).Value = meses(mes)



End Sub
Sub ValidarHoja()
'Validar si una rendiciion ya esta creada y eliminarla
For i = 1 To Sheets.Count
If Sheets(i).Name = "Rendicion" Then
eliminar
Exit For

End If


Next i
End Sub

Function EncontrarIni()
Dim EncontrarSeccion As Boolean
Dim i As Integer
i = 1

EncontrarSeccion = True
Do While EncontrarSeccion
If Cells(i, 2).Value = "SECCION C: Boletas o facturas pendientes de rendir mes actual" Then
EncontrarIni = i

EncontrarSeccion = False
End If
i = i + 1
Loop

End Function


Function EncontrarFini(i)
Dim EncontrarTotal As Boolean
Dim mes As Integer


EncontrarTotal = True

Do While EncontrarTotal
If Cells(i, 2).Value = "Total Gastos" Then
EncontrarFini = i
EncontrarTotal = False
End If
If Cells(i, 2).Value = "Fecha" Then
Dim k As Integer
k = i + 1

mes = Month(Cells(i + 1, 2).Value)

        Do While Cells(k, 2).Value <> "Total Gastos"
        If (Cells(k, 2).Value <> "") Then
        If (mes <> Month(Cells(k, 2).Value)) Then
        MsgBox "Los gastos deben corresponder al mismo mes"
        Cells(k, 2).Value = ""
        Exit Do
        End If
        
        End If
        
        
        k = k + 1
        Loop
        
End If



i = i + 1
Loop

Dim j As Integer

For j = 1 To Sheets.Count
If Sheets(j).Name = "Rendicion" Then
encabezado (mes)
Exit For

End If


Next j




End Function


