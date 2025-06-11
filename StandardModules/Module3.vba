Attribute VB_Name = "Module3"
'//////////////////////////////////////////////////////////////////////////
'// Module3: Image
'//////////////////////////////////////////////////////////////////////////

'
' DrawImageBorder Macro
' Keyboard Shortcut: Ctrl+Shift+E
Sub DrawImageBorder()
Attribute DrawImageBorder.VB_ProcData.VB_Invoke_Func = "E\n14"

    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
End Sub

'
' ImageSixtyPercent Macro
' Keyboard Shortcut: Ctrl+Shift+Z
Sub ImageSixtyPercent()
Attribute ImageSixtyPercent.VB_ProcData.VB_Invoke_Func = "Z\n14"

    Select Case TypeName(Selection)
    '画像が選択済み
    Case "Picture"
        Selection.ShapeRange.ScaleHeight 0.6, msoFalse, msoScaleFromTopLeft
        Application.CommandBars("Format Object").Visible = False
    Case Else
    
    End Select

End Sub

