'//////////////////////////////////////////////////////////////////////////
'// Module8: Open In Other App
'//////////////////////////////////////////////////////////////////////////





' �N������VSCode�Ńt�@�C�����J��
' OpenInVSCode Macro
' Keyboard Shortcut: Ctrl+Shift+V
Sub OpenInVSCode()
    Dim filePath As String
    Dim vscodePath As String

    ' �I�𒆃Z���̃t�@�C���p�X���擾
    filePath = ActiveCell.Value

    ' VSCode�̃C���X�g�[���p�X�i�p�X��K�X�ύX���Ă��������j
    vscodePath = "C:\Program Files\Microsoft VS Code\Code.exe"
'    vscodePath = "C:\Users\���[�U�[��\AppData\Local\Programs\Microsoft VS Code\Code.exe"

    ' VSCode���N�����ăt�@�C�����J��
    Shell """" & vscodePath & """ """ & filePath & """", vbNormalFocus
End Sub


