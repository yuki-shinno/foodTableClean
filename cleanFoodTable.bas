Attribute VB_Name = "cleanFoodTable"
Option Explicit

Sub cleanTable()
    On Error GoTo Err
    Const STR_NEW_WS_NAME As String = "�{�\ �N���[���A�b�v"
    Dim i As Long, j As Long
    Dim LastRow As Long, LastCol As Long
    Dim ws As Worksheet
    Dim flg As Boolean
    
    '���łɃV�[�g�����݂���ꍇ
    For Each ws In Worksheets
        If ws.Name = STR_NEW_WS_NAME Then
            MsgBox "���łɃN���[���A�b�v�ς݂̃V�[�g�����݂��܂�"
            Exit Sub
        End If
    Next ws
    
    '�{�\�̃V�[�g�����݂��Ȃ��ꍇ
    flg = False
    For Each ws In Worksheets
        If ws.Name = "�{�\" Then
            flg = True
        End If
    Next ws
    If flg = False Then
        MsgBox "�{�\�V�[�g�����݂��܂���B�������u�b�N��I�����Ă��������B"
    End If
    
    Worksheets("�{�\").Copy after:=Worksheets(Worksheets.Count)
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
    MsgBox "�I�����܂���"
    
    Exit Sub
    
Err:
    MsgBox "�\�����ʃG���[���������܂����B" & vbCrLf & "�����𒆒f���܂��B"
End Sub
