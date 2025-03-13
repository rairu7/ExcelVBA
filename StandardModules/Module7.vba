'//////////////////////////////////////////////////////////////////////////
'// Module7: MakeSheets
'//////////////////////////////////////////////////////////////////////////




' CreateSheetsFromCellValue Macro
' �I�𒆔͈͂̒l�Ɠ����̃V�[�g���쐬����
' Keyboard Shortcut: -
Sub CreateSheetsFromCellValue()

End Sub



' �V�[�g�ꗗ���擾
' WorkbookSheetList Macro
' Keyboard Shortcut: Ctrl+Shift+W
Sub WorkbookSheetList()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim i As Integer
    
    ' �A�N�e�B�u�V�[�g���擾
    Dim activeIndex As Integer
    activeIndex = ActiveSheet.Index
    
    ' �V�����V�[�g���A�N�e�B�u�V�[�g�̉E�ɑ}��
    Set newSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(activeIndex))
    newSheet.Name = "�V�[�g�ꗗ"
    
    ' �V�[�g�ꗗ�������o��
    With newSheet
        ' ������
        .Cells.Clear
        ' �w�b�_�[������
        .Cells(1, 1).Value = "�V�[�g��"
        i = 2 ' �f�[�^�̊J�n�s
        ' �e�V�[�g�����擾
        For Each ws In ActiveWorkbook.Sheets
            .Cells(i, 1).Value = ws.Name
            i = i + 1
        Next ws
    
        Range("A1").Interior.Color = RGB(226, 239, 218)
        
        Range("A1:A" & (i - 1)).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
        
End Sub


