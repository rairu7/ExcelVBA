'//////////////////////////////////////////////////////////////////////////
'// Module2: AddShape
'//////////////////////////////////////////////////////////////////////////

'
' FloatingComment Macro
' Keyboard Shortcut: Ctrl+Shift+F
Sub FloatingComment()
    
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



