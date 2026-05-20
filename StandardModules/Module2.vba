Attribute VB_Name = "Module2"
'//////////////////////////////////////////////////////////////////////////
'// Module2: AddShape
'//////////////////////////////////////////////////////////////////////////

Option Explicit
'ポップアップの名前
Private Const TITLE_SEARCH_SHAPE_TEXT As String = "オートシェイプ検索"

'@brief  : 文字検索関数
'@return : なし
Public Sub searchShapeText()

    Dim sheet As Worksheet          'ワークシート
    Dim searchWord As String        '検索ワード

    '検索ワード入力ポップアップを表示する
    searchWord = InputBox("検索したいワードを入力して下さい", TITLE_SEARCH_SHAPE_TEXT)
    
    If searchWord = "" Then
        GoTo ExitSub
    End If

     '対象のワークシートを現在開いているシートとする
    Set sheet = ActiveSheet

    '検索ワードが見つからない場合に出力
    If Not (searchReplaceShapeText(sheet.Shapes, searchWord)) Then
        MsgBox "「" & searchWord & "」が見つかりません", vbExclamation, TITLE_SEARCH_SHAPE_TEXT
    End If

ExitSub:
End Sub


'@brief : 図形内検索置換関数
'@param : worksheetObject Worksheetオブジェクト
'@param : searchWord      検索文字
'@return: searchReplaceShapeText 処理継続判定
Private Function searchReplaceShapeText(ByVal worksheetObject As Object, ByVal searchWord As String) As Boolean

    Dim targetShape  As Shape       'ワークシート内の図形
    Dim shapeText   As String       '図形内の文字
    Dim discoveryWord As Long       '検索ワード発見位置
    Dim replaceWord As String       '置換後の文字
    Dim replacePopupMsg As String   '置換ポップアップメッセージ
    Dim ret As Boolean              '処理継続判定
    Dim searchWordCnt As Long: searchWordCnt = 1 '図形内検索ワード数
    
    ret = False
    
    'ワークシートに図形が存在する間ループ
    For Each targetShape In worksheetObject
    
        'クループ化された図形の時
        If (targetShape.Type = msoGroup) Then
            
            If (searchReplaceShapeText(targetShape.GroupItems, searchWord)) Then
                ret = True
                GoTo ExitFunction
            End If
            
        'コメントの時
        ElseIf (targetShape.Type = msoComment) Then
            GoTo CONTINUE

        Else
            '指定したテキストフレームにテキストがあるかどうかを返す
            If (targetShape.TextFrame2.HasText = msoTrue) Then
                
                '図形内のテキストを取得
                shapeText = targetShape.TextFrame2.TextRange.Text
                
                '図形内の文字列から検索
                discoveryWord = InStr(shapeText, searchWord)
                
                '検索ワードが見つかったとき、置換の処理を行う
                If (discoveryWord > 0&) Then
                
                    'ウィンドウを図形の位置にスクロール
                    ActiveWindow.ScrollRow = targetShape.TopLeftCell.Row
                    ActiveWindow.ScrollColumn = targetShape.TopLeftCell.Column
                    
                    Do While (discoveryWord > 0&)
                    
                         'テキスト範囲選択を解除するため、カレントセルを選択する
                        targetShape.TopLeftCell.Select

                        targetShape.TextFrame2.TextRange.Characters(discoveryWord, Len(searchWord)).Select
                        
                        replacePopupMsg = "置換する場合、入力してください。" & vbCr & vbCr & "置換前 : " & searchWord & vbCr & "置換後"
                        
                        ' 置換入力メッセージを出力する
                        replaceWord = InputBox(replacePopupMsg, "置換")
                        
                        If replaceWord = "" Then
                          ret = True
                          GoTo CONTINUE
                        End If
                        
                        '図形内の文字列を置換する
                        targetShape.TextFrame2.TextRange.Text = Replace(shapeText, searchWord, replaceWord, 1, searchWordCnt)
                        targetShape.TopLeftCell.Select
                        

                        'もう一度検索・置換するのか
                        If (MsgBox("continue?", vbQuestion Or vbOKCancel, TITLE_SEARCH_SHAPE_TEXT) <> vbOK) Then
                            ret = True
                            
                            GoTo CONTINUE

                        '同じ図形内で文字検索
                        Else
                            discoveryWord = InStr(discoveryWord + 1&, shapeText, searchWord)
                        End If
                        
                        searchWordCnt = searchWordCnt + 1

                    Loop

                    GoTo CONTINUE
                End If
            End If
        End If
CONTINUE:
    Next

ExitFunction:
    searchReplaceShapeText = ret
ExitSub:
End Function


'
' FloatingComment Macro
' Keyboard Shortcut: Ctrl+Shift+F
Sub FloatingComment()
Attribute FloatingComment.VB_ProcData.VB_Invoke_Func = "F\n14"
    
    ActiveSheet.Shapes.AddShape(msoShapeRectangularCallout, Selection.Cells(1, 1).Left, Selection.Cells(1, 1).Top, 210, 100) _
        .Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 153, 255)
        .Transparency = 0.25
        .Solid
    End With
    With Selection.ShapeRange.TextFrame2.TextRange.Font
        .BaselineOffset = 0
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
    End With
End Sub


'
' InsertArrow Macro
' Keyboard Shortcut: Ctrl+Shift+A
Sub InsertArrow()
Attribute InsertArrow.VB_ProcData.VB_Invoke_Func = "A\n14"

    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, Selection.Cells(1, 1).Left, _
        Selection.Cells(1, 1).Top, Selection.Cells(1, 1).Left + 100, Selection.Cells(1, 1).Top + 0).Select
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 1.5
    End With
    ActiveWindow.Zoom = 85
End Sub


'
' InsertRedRectangle Macro
' Keyboard Shortcut: Ctrl+Shift+I
Sub InsertRedRectangle()
Attribute InsertRedRectangle.VB_ProcData.VB_Invoke_Func = "I\n14"
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.Cells(1, 1).Left, Selection.Cells(1, 1).Top, 144, 69.75) _
        .Select
    Selection.ShapeRange.Fill.Visible = msoFalse
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 1.5
    End With
End Sub



