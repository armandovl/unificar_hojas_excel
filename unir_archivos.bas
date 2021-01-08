Attribute VB_Name = "Módulo1"

Sub Unir_Archivos()
Dim Hoja As Object

'********Armando Valdés********
'********Facebook : Cinta Negra en Excel********

    Application.ScreenUpdating = False
       'Definir la variable como tipo Variante
       Dim X As Variant
       'Abrir cuadro de dialogo
       X = Application.GetOpenFilename _
           ("Excel Files (*.xlsx), *.xlsx", 2, "Abrir archivos", , True)
        'Validar si se seleccionaron archivos
        If IsArray(X) Then ' Si se seleccionan
          'Crea Libro nuevo
           Workbooks.Add
          'Captura nombre de archivo destino donde se grabaran los archivos seleccionados
           A = ActiveWorkbook.Name
          
        '*/********************
       For y = LBound(X) To UBound(X)
       Application.StatusBar = "Importando Archivos: " & X(y)
         Workbooks.Open X(y)
         b = ActiveWorkbook.Name
           For Each Hoja In ActiveWorkbook.Sheets
            Hoja.Copy after:=Workbooks(A).Sheets(Workbooks(A).Sheets.Count)
           Next
           Workbooks(b).Close False
       Next
       Application.StatusBar = "Listo"
       Call Unir_Hojas
    End If
    Application.ScreenUpdating = False
   End Sub

'********Armando Valdés********
'********Facebook : Cinta Negra en Excel********
Sub Unir_Hojas()
Dim Sig As Byte, Eliminar As Boolean
    For Sig = 2 To Worksheets.Count
        Worksheets(Sig).UsedRange.Copy _
        Worksheets(1).Range("a1000000").End(xlUp).Offset(1)
    Next
       Application.DisplayAlerts = False
        
    For Sig = 2 To Worksheets.Count
        Worksheets(2).Delete
    Next
Application.DisplayAlerts = True

End Sub
