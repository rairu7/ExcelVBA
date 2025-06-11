Attribute VB_Name = "Module1"

'//////////////////////////////////////////////////////////////////////////
'// Module1：FontColor, BackColor, SheetColor
'//////////////////////////////////////////////////////////////////////////

'
' BackColorYellow Macro
' Keyboard Shortcut: Ctrl+Shift+Q
Sub BackColorYellow()
Attribute BackColorYellow.VB_ProcData.VB_Invoke_Func = "Q\n14"
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
Attribute BackColorBeige.VB_ProcData.VB_Invoke_Func = "S\n14"
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
Attribute BackColorTransparent.VB_ProcData.VB_Invoke_Func = "C\n14"
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
Attribute BackColorOrange.VB_ProcData.VB_Invoke_Func = "O\n14"
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
Attribute BackColorLightBlue.VB_ProcData.VB_Invoke_Func = "B\n14"
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
Attribute BackColorGray.VB_ProcData.VB_Invoke_Func = "G\n14"
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
Attribute FontColorRed.VB_ProcData.VB_Invoke_Func = "R\n14"
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub
'
' FontColorDefault Macro
' Keyboard Shortcut: Ctrl+Shift+D
Sub FontColorDefault()
Attribute FontColorDefault.VB_ProcData.VB_Invoke_Func = "D\n14"
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub

'
' FontColorBlue Macro
' Keyboard Shortcut: Ctrl+Shift+X
Sub FontColorBlue()
Attribute FontColorBlue.VB_ProcData.VB_Invoke_Func = "X\n14"
    With Selection.Font
        .Color = -65536
        .TintAndShade = 0
    End With
End Sub

'
'SheetColorKiiro Macro
' Keyboard Shortcut: Ctrl+Shift+K
Sub SheetColorKiiro()
Attribute SheetColorKiiro.VB_ProcData.VB_Invoke_Func = "K\n14"
    With ActiveSheet.Tab
        .Color = 65535
        .TintAndShade = 0
    End With
End Sub

'
'SheetColorMushoku Macro
' Keyboard Shortcut: Ctrl+Shift+M
Sub SheetColorMushoku()
Attribute SheetColorMushoku.VB_ProcData.VB_Invoke_Func = "M\n14"
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
    End With
End Sub
