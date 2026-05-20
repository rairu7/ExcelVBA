Option Explicit

Sub 配下ファイル一覧を作成()

    Dim ws As Worksheet
    Dim rootPath As String
    Dim rowNum As Long
    Dim result As VbMsgBoxResult
    
    rootPath = ThisWorkbook.Path
    
    If rootPath = "" Then
        MsgBox "先にこのマクロブックを任意のフォルダに保存してください。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 既存の「ファイル一覧」シート削除
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("ファイル一覧").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' 新規シート作成
    Set ws = Worksheets.Add
    
    ws.Name = "ファイル一覧"
    
    ws.Range("A1").Value = "ファイル名"
    ws.Range("B1").Value = "フルパス"
    ws.Range("C1").Value = "フォルダ"
    
    rowNum = 2
    
    ' 再帰処理開始
    ListFilesRecursive rootPath, ws, rowNum
    
    ' 列幅自動調整
    ws.Columns("A:C").AutoFit
    
    ' Sheet1削除確認
    On Error Resume Next
    
    If Not Worksheets("Sheet1") Is Nothing Then
    
        result = MsgBox( _
            "Sheet1 を削除しますか？", _
            vbYesNo + vbQuestion, _
            "確認" _
        )
        
        If result = vbYes Then
        
            Application.DisplayAlerts = False
            Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            
        End If
        
    End If
    
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    MsgBox "ファイル一覧を作成しました。", vbInformation

End Sub

Private Sub ListFilesRecursive(ByVal folderPath As String, _
                               ByVal ws As Worksheet, _
                               ByRef rowNum As Long)

    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' ファイル一覧出力
    For Each file In folder.Files
    
        ' マクロブック自身は除外
        If file.Path <> ThisWorkbook.FullName Then
        
            ws.Hyperlinks.Add _
                Anchor:=ws.Cells(rowNum, 1), _
                Address:=file.Path, _
                TextToDisplay:=file.Name
            
            ws.Cells(rowNum, 2).Value = file.Path
            ws.Cells(rowNum, 3).Value = folderPath
            
            rowNum = rowNum + 1
            
        End If
        
    Next file
    
    ' サブフォルダを再帰処理
    For Each subFolder In folder.SubFolders
        ListFilesRecursive subFolder.Path, ws, rowNum
    Next subFolder

End Sub

