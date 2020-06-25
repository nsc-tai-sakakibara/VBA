Attribute VB_Name = "Module1"
Option Explicit
'�e���i���Ƃ̎擾�l��
Sub count()

    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Sheet1")
    Dim rs As Worksheet
    Set rs = wb.Worksheets("�S��")
    
    '�A�z�z��
    Dim ary As Object
    Set ary = CreateObject("Scripting.Dictionary")
    Dim tmp As String
    
    Dim last As Long
    last = rs.Cells(Rows.count, 3).End(xlUp).row
    
    If last = 1 Then
        MsgBox "��Ƀf�[�^�̃C���|�[�g���s���Ă��������B", vbExclamation
    End If
    
    ws.Range("H1") = "���i��"
    ws.Range("I1") = "�擾�Ґ�"
    ws.Rows(1).HorizontalAlignment = xlCenter
    
    '�@�������ݍs�@�@�@�ǂݎ��s
    Dim wsRow As Long, rsRow As Long
    
    For rsRow = 2 To last
        tmp = rs.Cells(rsRow, 3)
        
        '������@�u��
        Call rep(tmp)
        
        '�Q�Ɨp�@��
        ws.Cells(rsRow, 10) = tmp
        
    Next
        
    '�������݊J�n�s
    wsRow = 2
    
    For rsRow = 2 To last
    
        tmp = rs.Cells(rsRow, 3)
        
        '������@�u��
        Call rep(tmp)
        
        If Not ary.Exists(tmp) Then
            
            ary.Add tmp, tmp
            
            ws.Cells(wsRow, 8) = rs.Cells(rsRow, 3)
            
            '�Q�Ɨ�J�E���g
            ws.Cells(wsRow, 9) = WorksheetFunction. _
                CountIf(ws.Range("J2:J" & last), tmp)
            
            wsRow = wsRow + 1
            
        End If
    Next
    
    ws.Range("J2:J" & last).Value = ""
    
    Call sort(last, ws)
    
    ws.Columns.AutoFit
    ws.Rows.AutoFit

End Sub

'�ϊ��p
Sub rep(quali)
    Dim char As Long
    
    quali = Replace(quali, " ", "")
    quali = Replace(quali, "�@", "")
    
    For char = 1 To Len(quali)
        
        If Mid(quali, char, 1) Like "[�`-��]" Or Mid(quali, char, 1) Like "[�O-�X]" _
            Or Mid(quali, char, 1) Like "�|" Then
            
                quali = Replace(quali, Mid(quali, char, 1), StrConv(Mid(quali, char, 1), vbNarrow))

        End If
    Next char
   
End Sub

'���בւ�
Sub sort(last, ws)
    With ws
        .Columns("H:I").Select
        With .sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Range("H1") _
                , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Selection
            .Header = xlYes
            .MatchCase = True
            .Orientation = xlTopToBottom
            .SortMethod = xlStroke
            .Apply
        End With
        
        Selection.AutoFilter
        Rows("2:2").Select
        ActiveWindow.FreezePanes = True
        .Range("D10").Select
    
    End With

End Sub
