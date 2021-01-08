Attribute VB_Name = "Módulo2"
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

