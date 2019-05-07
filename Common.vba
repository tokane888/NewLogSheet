Sub SelectA1AllSheet()
    Application.ScreenUpdating = False

    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    For Each sheet In Worksheets
        sheet.Activate
        Call SelectA1
    Next
    currentSheet.Activate
    
    Application.ScreenUpdating = True
End Sub

Sub SelectA1()
    Cells(1, 1).Activate
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
End Sub

' シートが存在するか返す
Function SheetExists(sheetName As String)
    SheetExists = False
    Dim sheet As Worksheet
    For Each sheet In Sheets
        If sheet.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

'最終行を返す
Function GetLastRow(Optional col = 1)
    GetLastRow = Cells(Rows.Count, col).End(xlUp).Row
End Function
