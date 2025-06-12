Attribute VB_Name = "Module7"
'//////////////////////////////////////////////////////////////////////////
'// Module7: MakeSheets
'//////////////////////////////////////////////////////////////////////////




' CreateSheetsFromCellValue Macro
' 選択中範囲の値と同名のシートを作成する
' Keyboard Shortcut: -
Sub CreateSheetsFromCellValue()

End Sub



' シート一覧を取得
' WorkbookSheetList Macro
' Keyboard Shortcut: Ctrl+Shift+W
Sub WorkbookSheetList()
Attribute WorkbookSheetList.VB_ProcData.VB_Invoke_Func = "W\n14"
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim i As Integer
    
    ' アクティブシートを取得
    Dim activeIndex As Integer
    activeIndex = ActiveSheet.Index
    
    ' 新しいシートをアクティブシートの右に挿入
    Set newSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(activeIndex))
    newSheet.Name = "シート一覧"
    
    ' シート一覧を書き出し
    With newSheet
        ' 初期化
        .Cells.Clear
        ' ヘッダーを書く
        .Cells(1, 1).Value = ActiveWorkbook.Name
        .Cells(2, 1).Value = "シート名"
        .Cells(2, 2).Value = "リンク"
        i = 3 ' データの開始行
        ' 各シート名を取得
        For Each ws In ActiveWorkbook.Sheets
            .Cells(i, 1).Value = ws.Name
            i = i + 1
        Next ws
        
        ' ハイパーリンク
        Range("B3:B" & (i - 1)).Select
        Selection.FormulaR1C1 = "=HYPERLINK(""[""&R1C1&""]""&RC[-1]&""!A1"",RC[-1])"
    
        ' スタイル
        Range("A2:B2").Interior.Color = RGB(226, 239, 218)
        ' Range("B2").Interior.Color = RGB(226, 239, 218)
        
        Range("A2:B" & (i - 1)).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
        
    Range("A1").Select
    
End Sub


