Attribute VB_Name = "FoodTableClean"
Option Explicit

Const START_ROW As Long = 11 '食品が最初に含まれる行番号

Sub main()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Call cleanTable(ws)
    Next ws
End Sub

Sub cleanTable(ws As Worksheet)
    Application.ScreenUpdating = False
    
    Dim i As Long, j As Long
    Dim LastRow As Long, LastCol As Long
    
    ws.Copy after:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = ws.Name & "clean"
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    LastCol = Cells(START_ROW, Columns.Count).End(xlToLeft).Column
    
    Dim ary As Variant
    ary = Range(Cells(1, 1), Cells(LastRow, LastCol))
    Columns("A:C").NumberFormatLocal = "0_ "
    
    For i = START_ROW To LastRow
        For j = 1 To LastCol
            If Left(ary(i, j), 1) = "(" Then Cells(i, j).NumberFormatLocal = "(G/標準)"
            If Left(ary(i, j), 1) = "[" Then Cells(i, j).NumberFormatLocal = "(G/標準)"
            If ary(i, j) = "-" Then Cells(i, j).NumberFormatLocal = "-"
            If ary(i, j) = "(Tr)" Then Cells(i, j).NumberFormatLocal = "(""Tr"")"
            If ary(i, j) = "Tr" Then Cells(i, j).NumberFormatLocal = """Tr"""
            If ary(i, j) = "*" Then Cells(i, j).NumberFormatLocal = """*"""
            ary(i, j) = cleanString(ary(i, j))
        Next j
    Next i
    Range(Cells(1, 1), Cells(LastRow, LastCol)) = ary
    
    Application.ScreenUpdating = True
End Sub

Function cleanString(str)
    str = Replace(str, "Tr", "0")
    str = Replace(str, "-", "0")
    str = Replace(str, "(", "")
    str = Replace(str, ")", "")
    str = Replace(str, "[", "")
    str = Replace(str, "]", "")
    str = Replace(str, "†", "")
    str = Replace(str, "*", "0")
    cleanString = str
End Function
