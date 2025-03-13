'//////////////////////////////////////////////////////////////////////////
'// Module8: Open In Other App
'//////////////////////////////////////////////////////////////////////////





' 起動中のVSCodeでファイルを開く
' OpenInVSCode Macro
' Keyboard Shortcut: Ctrl+Shift+V
Sub OpenInVSCode()
    Dim filePath As String
    Dim vscodePath As String

    ' 選択中セルのファイルパスを取得
    filePath = ActiveCell.Value

    ' VSCodeのインストールパス（パスを適宜変更してください）
    vscodePath = "C:\Program Files\Microsoft VS Code\Code.exe"
'    vscodePath = "C:\Users\ユーザー名\AppData\Local\Programs\Microsoft VS Code\Code.exe"

    ' VSCodeを起動してファイルを開く
    Shell """" & vscodePath & """ """ & filePath & """", vbNormalFocus
End Sub


