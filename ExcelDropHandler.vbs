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

' セルC5の値を取得してセルD6に設定
On Error Resume Next
valueC5 = sheet.Range("C5").Value
sheet.Range("D6").Value = valueC5
If Err.Number <> 0 Then
    MsgBox "セルの操作中にエラーが発生しました。", vbCritical, "エラー"
    workbook.Close False
    excelApp.Quit
    WScript.Quit
End If
On Error GoTo 0

' 指定された行と列を削除
On Error Resume Next
sheet.Rows("2:5").Delete
sheet.Columns("C").Delete
sheet.Columns("E").Delete
If Err.Number <> 0 Then
    MsgBox "行または列の削除中にエラーが発生しました。", vbCritical, "エラー"
    workbook.Close False
    excelApp.Quit
    WScript.Quit
End If
On Error GoTo 0

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
