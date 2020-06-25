Attribute VB_Name = "インポート結果"
Option Explicit
'資格取得者リスト　インポート　ボタンに対応
Sub ImportData()

    Dim FolderPath As String

    FolderPath = GetFolderPath()

    If FolderPath = "" Then GoTo Skip
    
    'パスに\の追加
    FolderPath = FolderPath & Application.PathSeparator
    
    Dim fso As Scripting.FileSystemObject
    Dim folderFiles As Scripting.files
    Dim folderFile As Scripting.File
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderFiles = fso.GetFolder(FolderPath).files
    
    Dim rb As Workbook, wb As Workbook
    Set wb = ThisWorkbook
    Dim ws As Worksheet
    Set ws = wb.Worksheets("全期")
    
    Dim sheetName As Worksheet
    Dim i As Integer, check As Boolean
    Dim name As String
    
    'シート番号
    i = 1
    
    'シート作成
    For Each folderFile In folderFiles
        
        If i = 1 Then
            name = "全期"
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
    
    '参照先のシート番号
    rbSheetNumber = 1
    '参照元のシート番号
    wbSheetNumber = 3
    '参照元　各期の行数
    wbRow = 2
    '参照元　全期の行数
    wbTotalRow = 2
    
        ws.Range("A1").Value = "社員番号"
        ws.Range("A1").HorizontalAlignment = xlCenter
        ws.Range("B1").Value = "名前"
        ws.Range("B1").HorizontalAlignment = xlCenter
        ws.Range("C1").Value = "資格名"
        ws.Range("C1").HorizontalAlignment = xlCenter
        ws.Range("D1").Value = "取得月"
        ws.Range("D1").HorizontalAlignment = xlCenter
    
    'フォルダ数
    For Each folderFile In folderFiles
        
        Set rb = Workbooks.Open(folderFile)
            
        wb.Sheets(wbSheetNumber).Range("A1").Value = "社員番号"
        wb.Sheets(wbSheetNumber).Range("A1").HorizontalAlignment = xlCenter
        wb.Sheets(wbSheetNumber).Range("B1").Value = "名前"
        wb.Sheets(wbSheetNumber).Range("B1").HorizontalAlignment = xlCenter
        wb.Sheets(wbSheetNumber).Range("C1").Value = "資格名"
        wb.Sheets(wbSheetNumber).Range("C1").HorizontalAlignment = xlCenter
        wb.Sheets(wbSheetNumber).Range("D1").Value = "取得月"
        wb.Sheets(wbSheetNumber).Range("D1").HorizontalAlignment = xlCenter
            
        'ファイル選別
        If InStr(fso.GetBaseName(folderFile), "041") = 0 Then
            
            '読み込みシートの開始行数
            rbRow = 11
        
            'シート選別
            Do While InStr(rb.Sheets(rbSheetNumber).name, "月") > 0
                
                '社員番号
                Do While rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value <> ""
                    
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 1).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value
                    ws.Cells(wbTotalRow, 1).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value
                    
                    '名前
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 2).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 3).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 2).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 3).MergeArea.Item(1).Value
                    
                    '資格
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 3).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 5).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 3).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 5).MergeArea.Item(1).Value

                    '取得月
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 4).Value = rb.Sheets(rbSheetNumber).Cells(7, 6).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 4).Value = rb.Sheets(rbSheetNumber).Cells(7, 6).MergeArea.Item(1).Value

                    '行番号更新
                    rbRow = rbRow + 1
                    wbRow = wbRow + 1
                    wbTotalRow = wbTotalRow + 1
                    
                Loop
                
                '初期値戻し
                rbRow = 11
                '参照先ページ更新
                rbSheetNumber = rbSheetNumber + 1
            Loop
            
            '初期値戻し
            wbRow = 2
            rbSheetNumber = 1
        
        Else
            
            '参照先シートの開始行数
            rbRow = 12

            Do While Len(rb.Sheets(rbSheetNumber).name) = 6 And IsNumeric(rb.Sheets(rbSheetNumber).name)
                
                '社員番号
                Do While rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value <> ""
                    
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 1).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value
                    ws.Cells(wbTotalRow, 1).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 2).Value
                    
                    '名前
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 2).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 3).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 2).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 3).MergeArea.Item(1).Value
                    
                    '資格
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 3).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 5).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 3).Value = rb.Sheets(rbSheetNumber).Cells(rbRow, 5).MergeArea.Item(1).Value

                    '取得月
                    wb.Sheets(wbSheetNumber).Cells(wbRow, 4).Value = rb.Sheets(rbSheetNumber).Cells(8, 6).MergeArea.Item(1).Value
                    ws.Cells(wbTotalRow, 4).Value = rb.Sheets(rbSheetNumber).Cells(8, 6).MergeArea.Item(1).Value

                    '行番号更新
                    rbRow = rbRow + 1
                    wbRow = wbRow + 1
                    wbTotalRow = wbTotalRow + 1
                    
                Loop
                
                '初期値戻し
                rbRow = 12
                '参照先ページ更新
                rbSheetNumber = rbSheetNumber + 1
            Loop
            
            '初期値戻し
            wbRow = 2
            rbSheetNumber = 1
            
        End If
        
        wb.Sheets(wbSheetNumber).Columns("A:D").AutoFit
        wb.Sheets(wbSheetNumber).Columns("A:D").EntireRow.AutoFit
        
        '参照元ページ更新
        wbSheetNumber = wbSheetNumber + 1
        
        Call rb.Close(SaveChanges:=False)
          
    Next folderFile
    
    wb.Sheets(2).Columns("A:D").AutoFit
    wb.Sheets(2).Columns("A:D").EntireRow.AutoFit

Skip:
Exit Sub

End Sub
'フォルダ取得
Private Function GetFolderPath() As String

    Dim FolderPath As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        '.InitialFileName = "*資格取得推奨制度月次報告書.xlsx"
        If .Show = -1 Then
            FolderPath = .SelectedItems(1)
        End If
    End With
    
    GetFolderPath = FolderPath
    
End Function
