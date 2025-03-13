'//////////////////////////////////////////////////////////////////////////
'// ModuleF: Draft
'//////////////////////////////////////////////////////////////////////////



Sub Macro2()
'
' Macro2 Macro
'

'
End Sub
Sub Macro3()
'
' Macro3 Macro
'

'
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
End Sub






'■■■
'.AAE ファイルが存在するようなファイル名の画像をtrushTempフォルダへ移動させるマクロ
'入力はC3セル
Sub TrushAAE()
    
    
    
    
    
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "G:\Data\Apple_iPadmini4\202310__\新しいフォルダー「"
    Range("C4").Select
    
    Range("C3").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
        "G:\Data\Apple_iPadmini4\202310__\新しいフォルダー", TextToDisplay:= _
        "G:\Data\Apple_iPadmini4\202310__\新しいフォルダー"
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
End Sub






