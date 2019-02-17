Sub NewLogSheet()
    Dim sheetName As String: sheetName = Format(Date, "mmdd")
    Dim addDateCount As Integer
    addDateCount = 0
    While SheetExists(sheetName)
        addDateCount = addDateCount + 1
        Date = DateAdd("d", Date, addDateCount)
        sheetName = Format(Date, "mmdd")
    Wend

    If Weekday(Date) = 1 Or Weekday(Date) = 7 Then
        ' 土日の場合
        Worksheets("土日用").Copy After:=Worksheets(Worksheets.Count)
    Else
        ' 平日の場合
        Worksheets("平日用").Copy After:=Worksheets(Worksheets.Count)
    End If
    
    ActiveSheet.Name = sheetName
    Cells(1, 1).Value = Date
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
