' 当日分の2シートをまとめて生成
Sub NewSheets()
    Call NewPlanSheet
    Call NewDetailLogSheet
End Sub

' 30分単位の予定表生成
Sub NewPlanSheet()
    Dim sheetName As String: sheetName = GetSheetName("")

    If Weekday(Date) = 7 Then
        ' 土曜の場合
        Worksheets("土").Copy After:=Worksheets(Worksheets.Count)
        ActiveSheet.Tab.Color = 15773696
    ElseIf Weekday(Date) = 1 Then
        ' 日曜の場合
        Worksheets("日").Copy After:=Worksheets(Worksheets.Count)
        ActiveSheet.Tab.Color = 49407
    Else
        ' 平日の場合
        Worksheets("平日").Copy After:=Worksheets(Worksheets.Count)
    End If
    
    ActiveSheet.Name = sheetName
    Cells(1, 1).Value = Date
End Sub

' 詳細ログ生成
Sub NewDetailLogSheet()
    Dim sheetName As String: sheetName = GetSheetName("T")

    ' 平日の場合
    Worksheets("T").Copy After:=Worksheets(Worksheets.Count)
    
    ActiveSheet.Name = sheetName
    Cells(1, 1).Value = Date
End Sub

' シート名取得
Function GetSheetName(suffix)
    Dim sheetName As String: sheetName = Format(Date, "mmdd") + suffix
    Dim addDateCount As Integer
    addDateCount = 0
    While SheetExists(sheetName)
        addDateCount = addDateCount + 1
        Date = DateAdd("d", Date, addDateCount)
        sheetName = Format(Date, "mmdd") + suffix
    Wend
    GetSheetName = sheetName
End Function
