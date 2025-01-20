'//////////////////////////////////////////////////////////////////////////
'// Module5: Others
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
    boldFlg = False            ' 太字にする場合はTrue
    
    
    '★走査セル上限数設定（無限ループ対策）★
    supreme = 500           ' 変更注意
    
    ' 無限ループ対策
    intCnt = intCnt + 1
    If intCnt > supreme Then
        MsgBox ("上限オーバーです。選択範囲を見直してください")
        Exit Sub
    End If
    
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
                    startPos = foundPos + Len(searchText) '// 考えて納得する
                End If
            Loop While foundPos > 0
        End If
    Next cell
    
    MsgBox ("HighlightTextInCell完了")
    
End Sub



' シート一覧を取得
' GetWorkbookAllSheets Macro
' Keyboard Shortcut: Ctrl+Shift+W
Sub GetWorkbookAllSheets()
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
        .Cells(1, 1).Value = "シート名"
        i = 2 ' データの開始行
        ' 各シート名を取得
        For Each ws In ActiveWorkbook.Sheets
            .Cells(i, 1).Value = ws.Name
            i = i + 1
        Next ws
    End With
End Sub



' 起動中のVSCodeでファイルを開く
' OpenInVSCode Macro
' Keyboard Shortcut: Ctrl+Shift+V
Sub OpenInVSCode()
    Dim filePath As String
    Dim vscodePath As String

    ' 選択中セルのファイルパスを取得
    filePath = ActiveCell.Value

    ' VSCodeのインストールパス（パスを適宜変更してください）
    vscodePath = "C:\Program Files\Microsoft VS Code\Code.exe"
'    vscodePath = "C:\Users\ユーザー名\AppData\Local\Programs\Microsoft VS Code\Code.exe"

    ' VSCodeを起動してファイルを開く
    Shell """" & vscodePath & """ """ & filePath & """", vbNormalFocus
End Sub


