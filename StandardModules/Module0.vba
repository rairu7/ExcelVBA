'//////////////////////////////////////////////////////////////////////////
'// ■Index■
'// Module0: Index, ExportModules
'// Module1: FontColor, BackColor, SheetColor
'// Module2: AddShape
'// Module3: Image
'// Module4: TextBox
'// Module5: InsertRowsOthers
'// Module6: Unpublished Side
'// Module7: Draft
'// Module8: not used
'// Module9: not used
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
    
    ' ログイン中Winユーザーのドキュメントフォルダのパスを取得
    Dim docPath As String
    docPath = Environ("USERPROFILE") & "\Documents\"
'    MsgBox docPath

    
    ' ★要確認★
    moduleCount = 9

    For iNumber = 0 To moduleCount
        ' エクスポートするモジュールの名前と保存先
        moduleName = "Module" & iNumber
        exportPath = docPath & "develop\excel_vba\sources_git\ショートカット一覧\StandardModules\Module" & iNumber & ".vba"

        ' モジュールをエクスポート
        ThisWorkbook.VBProject.VBComponents(moduleName).Export exportPath
    Next
    
    Set objShell = Nothing
    
End Sub

Sub ExportForm()
    Dim formName As String
    Dim exportPath As String
    
    ' エクスポートするフォームの名前と保存先
    formName = "frmSearchText"
    exportPath = "C:\path\to\save\frmSearchText.frm"
    
    ' フォームをエクスポート
    ThisWorkbook.VBProject.VBComponents(formName).Export exportPath
    
'    ' フォームのリソースもエクスポート　←画像やアイコンを使っていなければ、インポートの際は､.frxファイルは不要.
'    exportPath = "C:\path\to\save\Form1.frx"
'    ' リソースのエクスポートは手動で行う
End Sub
