' ******************************
' プログラム名: ExcelDropHandler.vbs
' バージョン: 1.0
' 作成者: あなたの名前
' 作成日: 2024年11月23日
' 最終更新日: 2024年11月23日
' 概要:
'   このスクリプトはドラッグ＆ドロップされたExcelファイル（.xlsx形式）を処理します。
'   指定されたシート「Sheet2」を対象に以下の操作を実行します:
'     - セルC5の値を取得し、セルD6に設定。
'     - 2行目から5行目を削除。
'     - C列とE列を削除。
'   処理後のファイルは、元のファイル名にタイムスタンプ（_yyyymmdd-hhmmss形式）を
'   付加して保存します。
' 注意:
'   このスクリプトはShift-JIS形式で保存してください。
'   他の形式（UTF-8など）で保存すると文字化けが発生し、正しく動作しません。
' ******************************
' バージョン: 1.1
' 概要:
'   ドラッグ＆ドロップされたExcelファイル（.xlsx形式）を処理します。
'   シート「Sheet2」を対象に指定の操作を実行します。
' ******************************
' バージョン: 1.2
' 概要:
'   ドラッグ＆ドロップされたExcelファイル（.xlsx形式）を処理します。
'   シート「Sheet2」を対象に指定の操作を実行します。
' ******************************

' ドラッグ＆ドロップされたファイルを取得
Set args = WScript.Arguments
If args.Count = 0 Then
    MsgBox "Excelファイルをスクリプトにドラッグ＆ドロップしてください。", vbExclamation, "エラー"
    WScript.Quit
End If

filePath = args(0)

' 拡張子チェック
If LCase(Right(filePath, 5)) <> ".xlsx" Then
    MsgBox "Excelファイル（.xlsx形式）のみ処理できます。", vbExclamation, "エラー"
    WScript.Quit
End If

' Excelアプリケーションを起動
On Error Resume Next
Set excelApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Excelアプリケーションを起動できませんでした。", vbCritical, "エラー"
    WScript.Quit
End If
On Error GoTo 0

excelApp.Visible = False

' ファイルを開く
On Error Resume Next
Set workbook = excelApp.Workbooks.Open(filePath)
If Err.Number <> 0 Then
    MsgBox "Excelファイルを開けませんでした。", vbCritical, "エラー"
    excelApp.Quit
    WScript.Quit
End If
On Error GoTo 0

' 指定されたシートを取得
Set sheet = Nothing
For Each ws In workbook.Sheets
    If ws.Name = "Sheet2" Then
        Set sheet = ws
        Exit For
    End If
Next

If sheet Is Nothing Then
    MsgBox "指定されたシート 'Sheet2' が見つかりません。", vbExclamation, "エラー"
    workbook.Close False
    excelApp.Quit
    WScript.Quit
End If

' 指定された行と列を削除
On Error Resume Next
sheet.Rows("1:15").Delete
sheet.Columns("A").Delete
sheet.Columns("C").Delete
If Err.Number <> 0 Then
    MsgBox "行または列の削除中にエラーが発生しました。", vbCritical, "エラー"
    workbook.Close False
    excelApp.Quit
    WScript.Quit
End If
On Error GoTo 0

' 条件処理
rowCount = sheet.UsedRange.Rows.Count
For i = 1 To rowCount
    JValue = sheet.Cells(i, "J").Value
    AAValue = sheet.Cells(i, "AA").Value
    If JValue = "田中たかし" And InStr(AAValue, "課長案件") > 0 Then
        prefix = Split(AAValue, "課長")(0)
        sheet.Cells(i, "J").Value = prefix
    End If
Next

' 保存先ファイル名を生成
Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = fso.GetParentFolderName(filePath)
fileName = fso.GetBaseName(filePath)
currentDateTime = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "-" & _
                  Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
newFileName = folderPath & "\" & fileName & "_" & currentDateTime & ".xlsx"

' ファイルを保存
On Error Resume Next
workbook.SaveAs newFileName
If Err.Number <> 0 Then
    MsgBox "ファイルの保存中にエラーが発生しました。", vbCritical, "エラー"
    workbook.Close False
    excelApp.Quit
    WScript.Quit
End If
On Error GoTo 0

workbook.Close False
excelApp.Quit

' オブジェクトを解放
Set sheet = Nothing
Set workbook = Nothing
Set excelApp = Nothing
Set fso = Nothing

' 完了通知
MsgBox "処理が完了しました。" & vbCrLf & "保存先: " & newFileName, vbInformation, "完了"
WScript.Quit
