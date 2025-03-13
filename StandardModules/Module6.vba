'//////////////////////////////////////////////////////////////////////////
'// Module6: InsertRows, InsertColumns
'//////////////////////////////////////////////////////////////////////////




Sub InsertRowsAndBorders()
    Dim rng As Range
    Dim SelectedRange As Range
    Dim i As Long
    Dim rowCount As Long
    Dim lastColumn As Long
    
    
    ' DI列の列番号（ExcelではA=1, B=2, ... なのでDI列は130番）
    lastColumn = 130
    
    ' 現在選択されている範囲を取得
    Set SelectedRange = Selection
    
    ' 選択された範囲の行数を取得
    rowCount = SelectedRange.Rows.Count
    
    ' 選択された範囲の最終行の次の行から開始
    For i = rowCount To 1 Step -1
        ' 10行空行を挿入
        SelectedRange.Rows(i).Offset(1, 0).Resize(10, SelectedRange.Columns.Count).Insert Shift = xlDown
        
        '挿入した行と次の行の間にDI列まで罫線を引く
        SelectedRange.Rows(i).Offset(10, 0).Range (Cells(1, 1)), Cells(1, lastColumn).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    Next i
End Sub



' InsertRowsAboveWithValueInColumnA Macro
' A列を範囲選択中に、A列に値のある行の上にそれぞれ30行追加する
' Keyboard Shortcut: -
Sub InsertRowsAboveWithValueInColumnA()

End Sub

