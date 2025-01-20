
'//////////////////////////////////////////////////////////////////////////
'// Module1：FontColor, BackColor, SheetColor
'//////////////////////////////////////////////////////////////////////////

'
' BackColorYellow Macro
' Keyboard Shortcut: Ctrl+Shift+Q
Sub BackColorYellow()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With
End Sub

'
' BackColorBeige Macro
' Keyboard Shortcut: Ctrl+Shift+S
Sub BackColorBeige()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

'
' BackColorTransparent Macro
' Keyboard Shortcut: Ctrl+Shift+C
Sub BackColorTransparent()
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

'
' BackColorOrange Macro
' Keyboard Shortcut: Ctrl+Shift+O
Sub BackColorOrange()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

'
' BackColorLightBlue Macro
' Keyboard Shortcut: Ctrl+Shift+B
Sub BackColorLightBlue()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16777062
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

'
' BackColorGray Macro
' Keyboard Shortcut: Ctrl+Shift+G
Sub BackColorGray()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
End Sub

'
' FontColorRed Macro
' Keyboard Shortcut: Ctrl+Shift+R
Sub FontColorRed()
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub
'
' FontColorDefault Macro
' Keyboard Shortcut: Ctrl+Shift+D
Sub FontColorDefault()
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub

'
' FontColorBlue Macro
' Keyboard Shortcut: Ctrl+Shift+X
Sub FontColorBlue()
    With Selection.Font
        .Color = -65536
        .TintAndShade = 0
    End With
End Sub

'
'SheetColorKiiro Macro
' Keyboard Shortcut: Ctrl+Shift+K
Sub SheetColorKiiro()
    With ActiveSheet.Tab
        .Color = 65535
        .TintAndShade = 0
    End With
End Sub

'
'SheetColorMushoku Macro
' Keyboard Shortcut: Ctrl+Shift+M
Sub SheetColorMushoku()
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
    End With
End Sub
