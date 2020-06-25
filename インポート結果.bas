Attribute VB_Name = "�C���|�[�g����"
Option Explicit
'���i�擾�҃��X�g�@�C���|�[�g�@�{�^���ɑΉ�
Sub ImportData()

    Dim FolderPath As String

    FolderPath = GetFolderPath()

    If FolderPath = "" Then GoTo Skip
    
    '�p�X��\�̒ǉ�
    FolderPath = FolderPath & Application.PathSeparator
    
    Dim fso As Scripting.FileSystemObject
    Dim folderFiles As Scripting.files
    Dim folderFile As Scripting.File
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderFiles = fso.GetFolder(FolderPath).files
    
    Dim rb As Workbook, wb As Workbook
    Set wb = ThisWorkbook
    Dim ws As Worksheet
    Set ws = wb.Worksheets("�S��")
    
    Dim sheetName As Worksheet
    Dim i As Integer, check As Boolean
    Dim name As String
    
    '�V�[�g�ԍ�
    i = 1
    
    '�V�[�g�쐬
    For Each folderFile In folderFiles
        
        If i = 1 Then
            name = "�S��"
        Else
            name = Left(folderFile.name, 7)
        End If
        
        check = False
        
        For Each sheetName In Worksheets
       
            If sheetName.name = name Then
                check = True
                Exit For
            End If
            
        Next sheetName
        
        If Not check Then
            On Error Resume Next
            wb.Worksheets.Add(After:=Worksheets(i)).name = name
            On Error GoTo 0
        
        End If
        
        i = i + 1
    Next folderFile
    
    Dim rbSheetNumber As Integer, wbSheetNumber As Integer, _
    rbRow As Integer, wbRow As Integer, wbTotalRow As Integer
    
    '�Q�Ɛ�̃V�[�g�ԍ�
    rbSheetNumber = 1
    '�Q�ƌ��̃V�[�g�ԍ�
    wbSheetNumber = 3
    '�Q�ƌ��@�e���̍s��
    wbRow = 2
    '�Q�ƌ��@�S���̍s��
    wbTotalRow = 2
    
        ws.Range("A1").Value = "�Ј��ԍ�"
        ws.Range("A1").HorizontalAlignment = xlCenter
        ws.Range("B1").Value = "���O"
        ws.Range("B1").HorizontalAlignment = xlCenter
        ws.Range("C1").Value = "���i��"
        ws.Range("C1").HorizontalAlignment = xlCenter
        ws.Range("D1").Value = "�擾��"
        ws.Range("D1").HorizontalAlignment = xlCenter
    
    '�t�H���_��
    For Each folderFile In folderFiles
        
        Set rb = Workbooks.Open(folderFile)
            
        wb.Sheets(wbSheetNumber).Range("A1").Value = "�Ј��ԍ�"
        wb.Sheets(wbSheetNumber).Range("A1").HorizontalAlignment = xlCenter
        wb.Sheets(wbSheetNumber).Range("B1").Value = "���O"
        wb.Sheets(wbSheetNumber).Range("B1").HorizontalAlignment = xlCenter
        wb.Sheets(wbSheetNumber).Range("C1").Value = "���i��"
        wb.Sheets(wbSheetNumber).Range("C1").HorizontalAlignment = xlCenter
        wb.Sheets(wbSheetNumber).Range("D1").Value = "�擾��"
        wb.Sheets(wbSheetNumber).Range("D1").HorizontalAlignment = xlCenter
            
        '�t�@�C���I��
        If InStr(fso.GetBaseName(folderFile), "041") = 0 Then
            
            '�ǂݍ��݃V�[�g�̊J�n�s��
            rbRow = 11
        
            '�V�[�g�I��
            Do While InStr(rb.Sheets(rbSheetNumber).name, "��") > 0
                
                '�Ј��ԍ�
                Do While rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value <> ""
                    
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 1).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value
                    ws.Cells(wbTotalRow, 1).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value
                    
                    '���O
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 2).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 3).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 2).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 3).MergeArea.Item(1).Value
                    
                    '���i
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 3).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 5).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 3).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 5).MergeArea.Item(1).Value

                    '�擾��
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 4).Value = rb.Sheets(rbSheetNumber).Cells(7, 6).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 4).Value = rb.Sheets(rbSheetNumber).Cells(7, 6).MergeArea.Item(1).Value

                    '�s�ԍ��X�V
                    rbRow = rbRow + 1
                    wbRow = wbRow + 1
                    wbTotalRow = wbTotalRow + 1
                    
                Loop
                
                '�����l�߂�
                rbRow = 11
                '�Q�Ɛ�y�[�W�X�V
                rbSheetNumber = rbSheetNumber + 1
            Loop
            
            '�����l�߂�
            wbRow = 2
            rbSheetNumber = 1
        
        Else
            
            '�Q�Ɛ�V�[�g�̊J�n�s��
            rbRow = 12

            Do While Len(rb.Sheets(rbSheetNumber).name) = 6 And IsNumeric(rb.Sheets(rbSheetNumber).name)
                
                '�Ј��ԍ�
                Do While rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value <> ""
                    
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 1).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value
                    ws.Cells(wbTotalRow, 1).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value
                    
                    '���O
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 2).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 3).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 2).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 3).MergeArea.Item(1).Value
                    
                    '���i
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 3).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 5).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 3).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 5).MergeArea.Item(1).Value

                    '�擾��
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 4).Value = rb.Sheets(rbSheetNumber).Cells(8, 6).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 4).Value = rb.Sheets(rbSheetNumber).Cells(8, 6).MergeArea.Item(1).Value

                    '�s�ԍ��X�V
                    rbRow = rbRow + 1
                    wbRow = wbRow + 1
                    wbTotalRow = wbTotalRow + 1
                    
                Loop
                
                '�����l�߂�
                rbRow = 12
                '�Q�Ɛ�y�[�W�X�V
                rbSheetNumber = rbSheetNumber + 1
            Loop
            
            '�����l�߂�
            wbRow = 2
            rbSheetNumber = 1
            
        End If
        
        wb.Sheets(wbSheetNumber).Columns("A:D").AutoFit
        wb.Sheets(wbSheetNumber).Columns("A:D").EntireRow.AutoFit
        
        '�Q�ƌ��y�[�W�X�V
        wbSheetNumber = wbSheetNumber + 1
        
        Call rb.Close(SaveChanges:=False)
          
    Next folderFile
    
    wb.Sheets(2).Columns("A:D").AutoFit
    wb.Sheets(2).Columns("A:D").EntireRow.AutoFit

Skip:
Exit Sub

End Sub
'�t�H���_�擾
Private Function GetFolderPath() As String

    Dim FolderPath As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        '.InitialFileName = "*���i�擾�������x�����񍐏�.xlsx"
        If .Show = -1 Then
            FolderPath = .SelectedItems(1)
        End If
    End With
    
    GetFolderPath = FolderPath
    
End Function
