'//////////////////////////////////////////////////////////////////////////
'// Module5: Search
'//////////////////////////////////////////////////////////////////////////



' HighlightTextInCell Macro
' Keyboard Shortcut: -
' 処理概要: セル内の単語を着色
Sub HighlightTextInCell()
    
    Dim cell As Range
    Dim startPos As Long
    Dim foundPos As Long
    
    Dim searchText As String
    Dim caseStrict As Boolean
    Dim highlightColor As Long
    Dim boldFlg As Boolean
    Dim supreme As Long
    Dim intCnt As Long
    
    '★検索設定★
    searchText = "AAA"
    caseStrict = True      ' 大文字小文字を区別する場合はTrue
    
    '★結果のセル書式設定★
    highlightColor = RGB(255, 0, 0)  ' 赤色
'    highlightColor = RGB(0, 0, 255)  ' 青色
''    highlightColor = RGB(255, 249, 79)  ' 黄色
    boldFlg = False            ' 太字にする場合はTrue
    
    
    '★走査セル上限数設定（無限ループ対策）★
    supreme = 1000           ' 変更注意
    
    ' 無限ループ対策
    If Selection.Cells.Count > supreme Then
        MsgBox ("上限オーバーです。選択範囲を見直してください.")
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' 選択範囲内の各セルを処理
    intCnt = 1
    For Each cell In Selection
        ' セルが空でない場合のみ処理
        If Not IsEmpty(cell.Value) Then
            ' セル内でsearchTextが見つかる位置を検索
            startPos = 1
            Do
                ' searchText はセル内に存在するか？
                If caseStrict = True Then
                    foundPos = InStr(startPos, cell.Value, searchText)
                Else
                    foundPos = InStr(startPos, LCase(cell.Value), LCase(searchText))
                End If
                
                ' 見つかった場合
                If foundPos > 0 Then
                    ' セル内の文字列の該当部分をハイライト
                    cell.Characters(foundPos, Len(searchText)).Font.Color = highlightColor
                    ' 太字処理
                    If boldFlg = True Then
                        cell.Characters(foundPos, Len(searchText)).Font.Bold = True
                    End If
                    '// 次の検索位置をリセット
                    startPos = foundPos + Len(searchText)
                End If
            Loop While foundPos > 0
        End If
    Next cell
    
    MsgBox ("HighlightTextInCell完了")
    Exit Sub
    
ErrorHandler:
    MsgBox ("エラー終了")
    
End Sub



