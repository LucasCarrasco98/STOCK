VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Busqueda2 
   Caption         =   "Busqueda"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640.001
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
Dim matriz() As Variant
Dim f As Integer, c As Integer, a As Integer, b As Long
f = Hoja4.Range("A1", Range("A1").End(xlDown)).count
c = Hoja4.Range("A1", Range("A1").End(xlToRight)).count

    ReDim matriz(f, c)

    For a = 2 To f ' bucles para recorrer filas
        For b = 1 To c 'bucle para recorrer columnas
            matriz(a, b) = Hoja4.Cells(a, b).Value 'se carga los valores en la matriz
        Next b
    Next a
'MsgBox ("Fondo de filas: " & UBound(matriz, 1)) Verifica la cantidad de filas
'MsgBox ("Fondo de column: " & UBound(matriz, 2)) Verifica la cantidad de columnas
Lista.Clear

For i = LBound(matriz, 1) To UBound(matriz, 2) ' se salta el encabezado por eso lleva mas 1
    If matriz(i, 2) Like "*" & busq_txt & "*" Then
        With Lista
        .AddItem
        .List(.ListCount - 1, 0) = matriz(i, 1)
        .List(.ListCount - 1, 1) = matriz(i, 2)
        .List(.ListCount - 1, 2) = matriz(i, 5)
        .List(.ListCount - 1, 3) = matriz(i, 6)
        .List(.ListCount - 1, 4) = matriz(i, 8)
        .List(.ListCount - 1, 5) = matriz(i, 10)
        End With
    End If
'    'https://es.stackoverflow.com/questions/329363/c%C3%B3mo-mostrar-una-matriz-en-un-formulario-de-vba
Next i
End Sub


Private Sub listado_Click()
Dim matriz() As Variant
Dim f As Integer, c As Integer, a As Integer, b As Long
f = Hoja4.Range("A1", Range("A1").End(xlDown)).count
c = Hoja4.Range("A1", Range("A1").End(xlToRight)).count

    ReDim matriz(f, c)

    For a = 2 To f ' bucles para recorrer filas
        For b = 1 To c 'bucle para recorrer columnas
            matriz(a, b) = Hoja4.Cells(a, b).Value 'se carga los valores en la matriz
        Next b
    Next a
            With Lista
            .List = matriz
            End With
End Sub

Private Sub UserForm_Initialize()
rng = Hoja4.Range("A1:Z10000").Value
    With Lista
    .ColumnCount = 6
    .ColumnWidths = "30PT;130PT;60PT;50PT;25PT;50PT"
    End With
End Sub
