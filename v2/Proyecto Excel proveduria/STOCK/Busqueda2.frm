VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Busqueda2 
   Caption         =   "Busqueda"
   ClientHeight    =   6072
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8640
   OleObjectBlob   =   "Busqueda2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Busqueda2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Private rng As Variant
'Porfavor documentar el archivo
Private Sub btn_cerrar_Click()
Unload Me
End Sub

Private Sub busq_txt_Change()

Dim i As Integer
Dim bol As Boolean
bol = False
Lista.Clear
For i = LBound(rng, 1) + 1 To UBound(rng, 1) ' se salta el encabezado por eso lleva mas 1
    If rng(i, 2) Like "*" & busq_txt & "*" Then
      '  With Lista
        '.AddItem
       ' .List(.ListCount - 1, 0) = rng(i, 1)
        '.List(.ListCount - 1, 1) = rng(i, 2)
        '.List(.ListCount - 1, 2) = rng(i, 5)
        '.List(.ListCount - 1, 3) = rng(i, 6)
        '.List(.ListCount - 1, 4) = rng(i, 8)
       ' .List(.ListCount - 1, 5) = rng(i, 10)
       ' End With
       bol = True
    End If
    'https://es.stackoverflow.com/questions/329363/c%C3%B3mo-mostrar-una-matriz-en-un-formulario-de-vba
Next i
MsgBox bol
End Sub


Private Sub listado_Click()
Dim final As Integer
Dim i As Integer
final = Hoja4.Cells(Rows.count, 1).End(xlUp).Row
For i = 2 To final
    With Lista
    .AddItem
    .List(.ListCount - 1, 0) = Hoja4.Cells(i, 1) 'ID
    .List(.ListCount - 1, 1) = Hoja4.Cells(i, 2) 'DETALLE
    .List(.ListCount - 1, 2) = Hoja4.Cells(i, 5) 'ARTICULO
    .List(.ListCount - 1, 3) = Hoja4.Cells(i, 6) 'COSTO
    .List(.ListCount - 1, 4) = Hoja4.Cells(i, 8) 'EXISTENCIA
    .List(.ListCount - 1, 5) = Hoja4.Cells(i, 10) ' EFECTIVO
    End With
Next i

End Sub

Private Sub UserForm_Initialize()
rng = Hoja4.Range("A1:Z10000").Value
    With Lista
    .ColumnCount = 6
    .ColumnWidths = "30PT;130PT;60PT;50PT;25PT;50PT"
    End With
End Sub
