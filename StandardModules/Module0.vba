'//////////////////////////////////////////////////////////////////////////
'// ■Index■
'// Module0: Index, ExportModules
'// Module1: FontColor, BackColor, SheetColor
'// Module2: AddShape
'// Module3: Image
'// Module4: TextBox
'// Module5: Search
'// Module6: InsertRows, InsertColumns
'// Module7: MakeSheets
'// Module8: OpenInOtherApp
'// Module9: not used
'// ModuleA: not used
'// ModuleB: not used
'// ModuleC: not used
'// ModuleD: Others
'// ModuleE: Draft
'// ModuleF: Unpublished Side
'//////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////
'// ■Keyboard Shortcut Management■
' Ctrl＋Shift＋
'// A: InsertArrow
'// B: BackColorLightBlue
'// C: BackColorTransparent
'// D: FontColorDefault
'// E:
'// F: FloatingComment/セルの書式設定
'// G: BackColorGray
'// H: ReplaceTextBoxText（未完成）
'// I: InsertRedRectangle
'// J: （プロジェクト特化マクロ）
'// K: SheetColorKiiro
'// L:
'// M: SheetColorMushoku
'// N:
'// O: BackColorOrange
'// P: フォント設定
'// Q: BackColorYellow
'// R: FontColorRed
'// S: BackColorBeige
'// T: SearchTextInTextBoxes（一部機能未完成）
'// U: 数式バー開閉
'// V: OpenInVSCode/値のみ貼り付け
'// W: GetWorkbookAllSheets
'// X: FontColorBlue
'// Y:
'// Z: ImageSixtyPercent
'// 1: 書式_通貨 1,234
'// 2: 書式_上のセル値をコピー
'// 3: 書式_日付
'// 4: 書式_通貨 \1,234
'// 5: 書式_パーセンテージ
'// 6: 罫線_外枠
'// 7:
'// 8:
'// 9: セルの高さ？不明
'// 0:
'// ;: 挿入
'// :: 表示範囲を全選択
'// /:
'// \: 罫線_線なし
'//////////////////////////////////////////////////////////////////////////

Sub ExportModules()
    Dim moduleCount As Integer
    Dim moduleName As String
    Dim exportPath As String
    
    ' ログイン中Winユーザーのドキュメントフォルダのパス
    Dim docPath As String
    docPath = Environ("USERPROFILE") & "\Documents\"
'    MsgBox docPath


    ' エクスポート先
    Dim exportDir As String
    exportDir = docPath & "develop\excel_vba\sources_git\ショートカット一覧\StandardModules\"
    
    ' エクスポートするファイル形式
    Dim extension As String
'    extension = ".bas"
    extension = ".vba"
    
    ' ★要確認★
    moduleCount = 9

    For iNumber = 0 To moduleCount
        ' エクスポートするモジュールの名前と保存先
        moduleName = "Module" & iNumber
        exportPath = exportDir & moduleName & extension

        ' モジュールをエクスポート
        ThisWorkbook.VBProject.VBComponents(moduleName).Export exportPath
    Next
    
    Dim Alphas(6) As String
    Alphas(0) = "A"
    Alphas(1) = "B"
    Alphas(2) = "C"
    Alphas(3) = "D"
    Alphas(4) = "E"
    Alphas(5) = "F"
    
    For iNumber = 0 To 5
        ' エクスポートするモジュールの名前と保存先
        moduleName = "Module" & Alphas(iNumber)
        exportPath = exportDir & moduleName & extension

        ' モジュールをエクスポート
        ThisWorkbook.VBProject.VBComponents(moduleName).Export exportPath
    Next
    
    Set objShell = Nothing
    
End Sub

Sub ExportForm()
    Dim formName As String
    Dim exportPath As String
    
    ' ログイン中Winユーザーのドキュメントフォルダのパスを取得
    Dim docPath As String
    docPath = Environ("USERPROFILE") & "\Documents\"
'    MsgBox docPath
    
    ' エクスポートするフォームの名前と保存先
    formName = "frmSearchText"
    exportPath = "C:\path\to\save\frmSearchText.frm"
    
    ' フォームをエクスポート
    ThisWorkbook.VBProject.VBComponents(formName).Export exportPath
    
'    ' フォームのリソースもエクスポート　←画像やアイコンを使っていなければ、インポートの際は､.frxファイルは不要.
'    exportPath = "C:\path\to\save\Form1.frx"
'    ' リソースのエクスポートは手動で行う
End Sub
