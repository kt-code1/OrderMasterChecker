Attribute VB_Name = "Module1"
Option Explicit

Sub ImportCsv()

    Dim filePath As Variant
    Dim csvBook As Workbook
    Dim csvSheet As Worksheet
    Dim dataSheet As Worksheet
    
    filePath = Application.GetOpenFilename( _
        FileFilter:="CSVファイル (*.csv),*.csv", _
        Title:="CSVファイルを選択してください")
        
    If filePath = False Then
    
        Exit Sub
    
    End If
    
    Set csvBook = Workbooks.Open(filePath)
    Set csvSheet = csvBook.Worksheets(1)
    Set dataSheet = ThisWorkbook.Worksheets("データ")
    
    '「操作」シート 前回データ消去
    ThisWorkbook.Worksheets("操作").Range("B5").ClearContents
    ThisWorkbook.Worksheets("操作").Range("B10:B13").ClearContents
    
    '「データ」シート 前回データ消去
    dataSheet.Cells.Clear
    
    'csvヘッダーチェック
    If CheckCsvHeader(csvSheet, csvBook) = False Then
    
        Exit Sub
    
    End If
    
    'パス表示
    ThisWorkbook.Worksheets("操作").Range("B5").Value = filePath
    
    '「データ」シートへcsvの内容転記
    dataSheet.Range("A1").Resize( _
        csvSheet.UsedRange.Rows.Count, _
        csvSheet.UsedRange.Columns.Count _
    ).Value = csvSheet.UsedRange.Value
    
    'チェック結果用の見出しを追加
    dataSheet.Range("F1").Value = "Result"
    dataSheet.Range("G1").Value = "ErrorMessage"
    
    '各項目チェック・結果表示
    CheckData dataSheet
    
    'csv閉じる
    csvBook.Close SaveChanges:=False
    
    
End Sub

Private Function CheckCsvHeader( _
    ByVal csvSheet As Worksheet, _
    ByVal csvBook As Workbook _
) As Boolean

    If csvSheet.UsedRange.Columns.Count <> 5 Or _
        csvSheet.Cells(1, 1) <> "注文番号" Or _
        csvSheet.Cells(1, 2) <> "受注日" Or _
        csvSheet.Cells(1, 3) <> "商品コード" Or _
        csvSheet.Cells(1, 4) <> "数量" Or _
        csvSheet.Cells(1, 5) <> "単価" Then
    
        MsgBox "CSVの見出しが、このツールで想定している形式と異なります。"
        csvBook.Close SaveChanges:=False
        
        CheckCsvHeader = False
        
    Else
        
        CheckCsvHeader = True
        
    End If
    
End Function

Private Sub CheckData(ByVal dataSheet As Worksheet)

    Dim errorMsg As String
    Dim rowNo As Long
    Dim dataCount As Long
    Dim okCount As Long
    Dim ngCount As Long
    Dim totalQuantity As Long
    Dim matchResult As Variant
    Dim masterLastRow As Long
    Dim masterSheet As Worksheet
    Dim masterCodeRow As Long
    
    'データ数(UsedRangeで行数取得。ヘッダ部分は除外するため-1)
    dataCount = dataSheet.UsedRange.Rows.Count - 1
    
    '「商品マスタ」シート保持
    Set masterSheet = ThisWorkbook.Worksheets("商品マスタ")
    
    '「商品マスタ」シートの最終行
    masterLastRow = masterSheet.UsedRange.Rows.Count
    
    okCount = 0
    ngCount = 0
    totalQuantity = 0

    '各行チェック開始
    '(ループで必要なのは行番号。dataCountは「件数」なので、+1で実際の「最終行番号」に戻す)
    For rowNo = 2 To dataCount + 1
    
        'エラーメッセージ初期化
        errorMsg = ""
        
        '商品コード 「商品マスタ」上での該当する行番号初期化
        masterCodeRow = 0
            
        '注文番号 空欄チェック
        If dataSheet.Cells(rowNo, 1).Value = "" Then
        
            errorMsg = AddErrorMessage(errorMsg, "注文番号が空欄です")
        
        '注文番号 重複チェック
        ElseIf Application.WorksheetFunction.CountIf( _
            dataSheet.Range("A2:A" & (dataCount + 1)), _
            dataSheet.Cells(rowNo, 1).Value) > 1 Then

            errorMsg = AddErrorMessage(errorMsg, "注文番号が重複しています")
        
        End If
        
        '受注日 空欄チェック
        If dataSheet.Cells(rowNo, 2).Value = "" Then
        
            errorMsg = AddErrorMessage(errorMsg, "受注日が空欄です")
        
        '受注日 日付として扱えないか
        ElseIf Not IsDate(dataSheet.Cells(rowNo, 2).Value) Then
        
            errorMsg = AddErrorMessage(errorMsg, "受注日が日付ではありません")
        
        End If
        
        
        '商品コード 空欄チェック
        If dataSheet.Cells(rowNo, 3).Value = "" Then
        
            errorMsg = AddErrorMessage(errorMsg, "商品コードが空欄です")
        
        Else
        
            '商品コード「商品マスタ」シートと照合
            matchResult = Application.Match( _
                dataSheet.Cells(rowNo, 3).Value, _
                masterSheet.Range("A2:A" & masterLastRow), _
                0)
        
            '「商品マスタ」に存在しない
            If IsError(matchResult) Then
        
                errorMsg = AddErrorMessage(errorMsg, "商品コードが「商品マスタ」シートに存在しません")
                
            Else
            
                '「商品マスタ」シート上での行番号を控えておく
                '(matchResultは、.Matchで指定した範囲の中で何番目かを返している。+1することで実際のエクセルの行番号になる。)
                masterCodeRow = matchResult + 1
        
            End If
        
        End If
        
        '数量 空欄チェック
        If dataSheet.Cells(rowNo, 4).Value = "" Then
        
            errorMsg = AddErrorMessage(errorMsg, "数量が空欄です")
        
        '数量 数値として扱えないか
        ElseIf Not IsNumeric(dataSheet.Cells(rowNo, 4).Value) Then
        
            errorMsg = AddErrorMessage(errorMsg, "数量が不正な値です")
        
        '数量 整数チェック
        ElseIf CDbl(dataSheet.Cells(rowNo, 4).Value) <> _
            Fix(CDbl(dataSheet.Cells(rowNo, 4).Value)) Then
            
            errorMsg = AddErrorMessage(errorMsg, "数量が整数ではありません")
            
        '数量 負数チェック
        ElseIf CDbl(dataSheet.Cells(rowNo, 4).Value) < 0 Then
        
            errorMsg = AddErrorMessage(errorMsg, "数量が負数です")
            
        End If
        
        '単価 空欄チェック
        If dataSheet.Cells(rowNo, 5).Value = "" Then
        
            errorMsg = AddErrorMessage(errorMsg, "単価が空欄です")
        
        '単価 数値として扱えないか
        ElseIf Not IsNumeric(dataSheet.Cells(rowNo, 5).Value) Then
        
            errorMsg = AddErrorMessage(errorMsg, "単価が不正な値です")
            
        '単価 整数チェック
        ElseIf CDbl(dataSheet.Cells(rowNo, 5).Value) <> _
            Fix(CDbl(dataSheet.Cells(rowNo, 5).Value)) Then
            
            errorMsg = AddErrorMessage(errorMsg, "単価が整数ではありません")
            
        '単価 負数チェック
        ElseIf CDbl(dataSheet.Cells(rowNo, 5).Value) < 0 Then
        
            errorMsg = AddErrorMessage(errorMsg, "単価が負数です")
            
        '単価が正常
        Else
        
            '商品コード「商品マスタ」シート上で照合確認済み
            If masterCodeRow > 0 Then
            
                '読み込んだcsvの単価と「商品マスタ」の単価が一致しない
                If CDbl(dataSheet.Cells(rowNo, 5).Value) <> _
                    CDbl(masterSheet.Cells(masterCodeRow, 3).Value) Then
                    
                    errorMsg = AddErrorMessage(errorMsg, "商品マスタとの単価不一致")
                    
                End If
            
            End If
        
        End If
            
        
        '結果
        If errorMsg = "" Then
        
            dataSheet.Cells(rowNo, 6).Value = "OK"
            okCount = okCount + 1
            
            '合計数量
            totalQuantity = totalQuantity + CLng(dataSheet.Cells(rowNo, 4).Value)
        
        Else
        
            dataSheet.Cells(rowNo, 6).Value = "NG"
            dataSheet.Cells(rowNo, 7).Value = errorMsg
            ngCount = ngCount + 1
            
        End If
    
    Next rowNo
        
    ThisWorkbook.Worksheets("操作").Range("B10").Value = dataCount
    ThisWorkbook.Worksheets("操作").Range("B11").Value = okCount
    ThisWorkbook.Worksheets("操作").Range("B12").Value = ngCount
        
    If ngCount <> 0 Then
        
        ThisWorkbook.Worksheets("操作").Range("B13").Value = "集計不可"
            
    Else
        
        ThisWorkbook.Worksheets("操作").Range("B13").Value = totalQuantity
        
    End If

End Sub

Private Function AddErrorMessage( _
    ByVal currentMessage As String, _
    ByVal addMessage As String _
) As String

    If currentMessage <> "" Then
    
        currentMessage = currentMessage & " / "
        
    End If
        
    AddErrorMessage = currentMessage & addMessage

End Function

Sub ExportNgCsv()

    Dim dataSheet As Worksheet
    Dim savePath As Variant
    Dim exportBook As Workbook
    Dim exportSheet As Worksheet
    Dim dataCount As Long
    Dim exportRow As Long
    Dim rowNo As Long
    
    Set dataSheet = ThisWorkbook.Worksheets("データ")
    
    'NGデータ出力前「データ」シートチェック
    If CanExportNgData(dataSheet) = False Then
    
        Exit Sub
    
    End If
    
    '出力するcsvの保存先
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="error_data.csv", _
        FileFilter:="CSVファイル (*.csv),*.csv", _
        Title:="NGデータの保存先を指定してください")

    'キャンセルした場合は終了
    If savePath = False Then
        Exit Sub
    End If
    
    '出力用の新規ブックを開く
    Set exportBook = Workbooks.Add
    
    'シートも取得
    Set exportSheet = exportBook.Worksheets(1)
    
    '「データ」シートのヘッダ部分を転記
    exportSheet.Range("A1:G1").Value = dataSheet.Range("A1:G1").Value
    
    '「データ」シート上で件数取得
    dataCount = dataSheet.UsedRange.Rows.Count - 1
    
    '新規シート上での出力開始行
    exportRow = 2
    
    'NG行だけ転記していく
    For rowNo = 2 To dataCount + 1
    
        If dataSheet.Cells(rowNo, 6).Value = "NG" Then
        
            exportSheet.Range("A" & exportRow & ":G" & exportRow).Value = _
                dataSheet.Range("A" & rowNo & ":G" & rowNo).Value
                
            exportRow = exportRow + 1

        End If
        
    Next rowNo
    
    On Error GoTo SaveError
    
    '出力用ブックをCSV形式で保存
    exportBook.SaveAs _
        Filename:=savePath, _
        FileFormat:=xlCSV, _
        Local:=True
        
    On Error GoTo 0

    '出力用ブックを閉じる
    exportBook.Close SaveChanges:=False
    
    '出力完了を通知
    MsgBox "NGデータのCSV出力が完了しました。"
    
    Exit Sub
    
SaveError:
    
    '出力用ブックを閉じる
    exportBook.Close SaveChanges:=False

    MsgBox "CSVの保存を中止したか、保存に失敗しました。"
    
End Sub

Private Function CanExportNgData( _
    ByVal dataSheet As Worksheet) As Boolean
    
    'Result列が無い
    If dataSheet.Range("F1").Value <> "Result" Then
    
        MsgBox "csvを先に読み込ませてください"
        CanExportNgData = False
    
    'NGデータが無い
    ElseIf Application.WorksheetFunction.CountIf( _
        dataSheet.Range("F:F"), _
        "NG") = 0 Then
        
        MsgBox "NGデータがありません"
        CanExportNgData = False
        
    Else
    
        CanExportNgData = True
        
    End If

End Function

