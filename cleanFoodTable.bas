Attribute VB_Name = "cleanFoodTable"
Option Explicit

Sub cleanTable()
    On Error GoTo Err
    Const STR_NEW_WS_NAME As String = "本表 クリーンアップ"
    Dim i As Long, j As Long
    Dim LastRow As Long, LastCol As Long
    Dim ws As Worksheet
    Dim flg As Boolean
    
    'すでにシートが存在する場合
    For Each ws In Worksheets
        If ws.Name = STR_NEW_WS_NAME Then
            MsgBox "すでにクリーンアップ済みのシートが存在します"
            Exit Sub
        End If
    Next ws
    
    '本表のシートが存在しない場合
    flg = False
    For Each ws In Worksheets
        If ws.Name = "本表" Then
            flg = True
        End If
    Next ws
    If flg = False Then
        MsgBox "本表シートが存在しません。正しいブックを選択してください。"
    End If
    
    Worksheets("本表").Copy after:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = STR_NEW_WS_NAME
    
    LastRow = Range("A10000").End(xlUp).Row
    LastCol = Range("XFD9").End(xlToLeft).Column
    
    Dim ary As Variant
    ary = Range(Cells(1, 1), Cells(LastRow, LastCol))
    Columns("A:C").NumberFormatLocal = "0_ "
    
    For i = 9 To LastRow
        For j = 1 To LastCol
            ary(i, j) = Replace(ary(i, j), "Tr", "0")
            ary(i, j) = Replace(ary(i, j), "-", "0")
            ary(i, j) = Replace(ary(i, j), "(", "")
            ary(i, j) = Replace(ary(i, j), ")", "")
        Next j
    Next i
    
    Range(Cells(1, 1), Cells(LastRow, LastCol)) = ary
    MsgBox "終了しました"
    
    Exit Sub
    
Err:
    MsgBox "予期せぬエラーが発生しました。" & vbCrLf & "処理を中断します。"
End Sub
