VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
Dim ini, fin, k As Integer
ini = EncontrarIni()
k = ini
fin = EncontrarFini(k)


If Not Application.Intersect(Target, Range("$D$" & ini & ":$D$" & fin)) Is Nothing Then
Dim texto As String

Dim celdaMonto As Integer
 
 celdaMonto = Mid(ActiveCell.Address, 4, 3)


 If Cells(celdaMonto, 7).Value <> "" Then
 MsgBox "No puede modificar un concepto, si ya se especifico un monto" & vbCrLf & "Se eliminaran los datos"
Cells(celdaMonto, 7).Value = ""
Cells(celdaMonto, 6).Value = ""
Cells(celdaMonto, 4).Value = ""

 End If
End If







'If Target.Address = "$C$66" Then eliminar
End Sub
