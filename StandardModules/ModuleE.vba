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






'������
'.AAE �t�@�C�������݂���悤�ȃt�@�C�����̉摜��trushTemp�t�H���_�ֈړ�������}�N��
'���͂�C3�Z��
Sub TrushAAE()
    
    
    
    
    
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "G:\Data\Apple_iPadmini4\202310__\�V�����t�H���_�[�u"
    Range("C4").Select
    
    Range("C3").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
        "G:\Data\Apple_iPadmini4\202310__\�V�����t�H���_�[", TextToDisplay:= _
        "G:\Data\Apple_iPadmini4\202310__\�V�����t�H���_�["
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
End Sub






