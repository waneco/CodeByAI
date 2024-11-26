' ******************************
' プログラム名: ExcelDropHandler.vbs
' バージョン: 1.6
' 作成日: 2024年11月23日
' 最終更新日: 2024年11月26日
' 概要:
'   このスクリプトはドラッグ＆ドロップされたExcelファイル（.xlsx形式）を処理します。
'   シート名が正規表現 "*eet2" に部分一致するシートを対象に指定の操作を実行します:
'     - 2行目から15行目を削除。
'     - A列とC列を削除。
'     - 条件に基づき、AA列の値からJ列を更新。
'   処理後のファイルは、元のファイル名にタイムスタンプ（_yyyymmdd-hhmmss形式）を
'   付加して保存します。
' 注意:
'   このスクリプトはShift-JIS形式で保存してください。
'   他の形式（UTF-8など）で保存すると文字化けが発生し、正しく動作しません。
' ******************************

' エラー表示関数
Sub HandleError(message)
    MsgBox message & vbCrLf & "エラーコード: " & Err.Number, vbCritical, "エラー"
    WScript.Quit
End Sub

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
    HandleError "Excelアプリケーションを起動できませんでした。"
End If
On Error GoTo 0

excelApp.Visible = False

' ファイルを開く
On Error Resume Next
Set workbook = excelApp.Workbooks.Open(filePath)
If Err.Number <> 0 Then
    HandleError "Excelファイルを開けませんでした。"
End If
On Error GoTo 0

' 正規表現を使用するための準備
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Pattern = ".*eet2$"

' 指定されたシートを取得
Set sheet = Nothing
For Each ws In workbook.Sheets
    If regEx.Test(ws.Name) Then
        Set sheet = ws
        Exit For
    End If
Next

' シートが見つからなかった場合
If sheet Is Nothing Then
    HandleError "指定されたパターンに一致するシートが見つかりません。"
End If

' 指定された行と列を削除
On Error Resume Next
sheet.Rows("2:15").Delete
If Err.Number <> 0 Then
    HandleError "行の削除中にエラーが発生しました。"
End If
sheet.Columns("A").Delete
If Err.Number <> 0 Then
    HandleError "A列の削除中にエラーが発生しました。"
End If
sheet.Columns("C").Delete
If Err.Number <> 0 Then
    HandleError "C列の削除中にエラーが発生しました。"
End If
On Error GoTo 0

' 条件処理
rowCount = sheet.UsedRange.Rows.Count
On Error Resume Next
For i = 1 To rowCount
    JValue = sheet.Cells(i, "J").Value
    AAValue = sheet.Cells(i, "AA").Value
    If JValue = "田中たかし" And InStr(AAValue, "課長案件") > 0 Then
        prefix = Split(AAValue, "課長")(0)
        sheet.Cells(i, "J").Value = prefix
    End If
    If Err.Number <> 0 Then
        HandleError "条件処理中にエラーが発生しました。"
    End If
Next
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
    HandleError "ファイルの保存中にエラーが発生しました。"
End If
On Error GoTo 0

' リソース解放処理
On Error Resume Next
If Not workbook Is Nothing Then workbook.Close False
If Not excelApp Is Nothing Then excelApp.Quit
Set sheet = Nothing
Set workbook = Nothing
Set excelApp = Nothing
Set fso = Nothing
Set regEx = Nothing
On Error GoTo 0
        
' 完了通知
MsgBox "処理が完了しました。" & vbCrLf & "保存先: " & newFileName, vbInformation, "完了"

WScript.Quit
