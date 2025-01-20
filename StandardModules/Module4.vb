'//////////////////////////////////////////////////////////////////////////
'// ■Index■
'// Module4: TextBox
'//////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////
'// ■Index■
'   GetTextBoxText
'   ReplaceTextboxText
'   CreateTextboxFromCellValue
'   SearchTextInTextBoxes
'//////////////////////////////////////////////////////////////////////////

' テキストボックスの値をコピー（未完成）
' GetTextBoxText Macro
' Keyboard Shortcut:-
Sub GetTextBoxText()
'    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, Selection.Cells(1, 1).Left, Selection.Cells(1, 1).Top, 72, _
'        72).Select
'    Range("I13").Select
'    ActiveCell.FormulaR1C1 = "aaa"
'    ActiveSheet.Shapes.Range(Array("TextBox 2")).Select
'    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = ""
'    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 3).ParagraphFormat. _
'        FirstLineIndent = 0
'    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 3).Font
'        .NameComplexScript = "+mn-cs"
'        .NameFarEast = "+mn-ea"
'        .Fill.Visible = msoTrue
'        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
'        .Fill.ForeColor.TintAndShade = 0
'        .Fill.ForeColor.Brightness = 0
'        .Fill.Transparency = 0
'        .Fill.Solid
'        .Size = 11
'        .Name = "+mn-lt"
'    End With
'
    Dim intShape As Integer
    Dim s As String
    
    For intShape = 1 To ActiveSheet.Shapes.Count
        If ActiveSheet.Shapes(intShape).Type = msoTextBox Then
            s = s & ActiveSheet.Shapes(intShape).TextFrame.Characters.Text
        End If
    Next
    
    
'    .ActiveCell.Value = s
End Sub

'
' CreateTextboxFromCellValue Macro
' Keyboard Shortcut:-
Sub CreateTextboxFromCellValue()
    Dim intRow As Integer
    Dim s As String
    
    For intRow = 1 To Selection.Rows.Count
        s = Selection.Cells(intRow, 1).Value
        ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            Selection.Cells(intRow, 1).Left + Selection.Cells(intRow, 1).Width, Selection.Cells(intRow, 1).Top, 200, 72) _
            .TextFrame.Characters.Text = s
'            Selection.Cells(1, intRow).Left + Selection.Cells(1, intRow).Width, Selection.Cells(0, intRow).Top, 200, 72) _
'            .TextFrame.Characters.Text = s
    Next
End Sub



'
' SearchTextInTextBoxes Macro
' Keyboard Shortcut: Ctrl+Shift+T
Sub SearchTextInTextBoxes()
    frmSearchText.Show
End Sub

