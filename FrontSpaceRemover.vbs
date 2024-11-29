' 指定された拡張子を持つテキストファイルを処理し、各行の先頭にあるすべての半角スペースを削除するスクリプト
' 処理後のファイルは元ファイルの末尾にタイムスタンプを付けて保存します

Option Explicit

Dim objArgs, objFSO, objFile, inputFilePath, outputFilePath
Dim validExtensions, fileExtension, timestamp

' 処理対象の有効な拡張子を定義
validExtensions = Array(".py", ".bat", ".log", ".txt", ".ps1", ".vbs")

' FileSystemObject を作成
Set objFSO = CreateObject("Scripting.FileSystemObject")

' コマンドライン引数（ドラッグアンドドロップされたファイル）を取得
Set objArgs = WScript.Arguments

' ファイルがドラッグアンドドロップされていない場合のエラーメッセージ
If objArgs.Count = 0 Then
    MsgBox "ファイルがドラッグアンドドロップされていません。スクリプトを終了します。", vbExclamation, "エラー"
    WScript.Quit
End If

' 各ファイルを処理
For Each inputFilePath In objArgs
    If objFSO.FileExists(inputFilePath) Then
        ' ファイルの拡張子を取得して小文字に変換
        fileExtension = LCase(objFSO.GetExtensionName(inputFilePath))
        fileExtension = "." & fileExtension
        
        ' 拡張子が有効かどうかを確認
        If IsInArray(fileExtension, validExtensions) Then
            On Error Resume Next
            ' ファイルを読み込む
            Set objFile = objFSO.OpenTextFile(inputFilePath, 1)
            If Err.Number <> 0 Then
                MsgBox "ファイルの読み取り中にエラーが発生しました: " & inputFilePath, vbCritical, "エラー"
                WScript.Quit
            End If
            On Error GoTo 0

            Dim fileContent, line, outputLines
            fileContent = objFile.ReadAll ' ファイル全体の内容を読み取る
            objFile.Close
            
            ' 各行を処理して先頭の半角スペースを削除
            outputLines = ""
            For Each line In Split(fileContent, vbCrLf)
                outputLines = outputLines & TrimStart(line) & vbCrLf
            Next

            ' タイムスタンプ付きの出力ファイルパスを生成
            timestamp = GetTimestamp()
            outputFilePath = objFSO.GetParentFolderName(inputFilePath) & "\" & _
                             objFSO.GetBaseName(inputFilePath) & "_" & timestamp & fileExtension

            ' 処理結果を新しいファイルに書き込む
            Set objFile = objFSO.CreateTextFile(outputFilePath, True)
            objFile.Write outputLines
            objFile.Close

            MsgBox "ファイルの処理が完了しました: " & vbCrLf & inputFilePath & vbCrLf & _
                   "出力ファイル: " & vbCrLf & outputFilePath, vbInformation, "完了"
        Else
            ' 無効な拡張子の場合のエラーメッセージ
            MsgBox "無効なファイル拡張子です: " & fileExtension & vbCrLf & _
                   "対応拡張子: " & Join(validExtensions, ", "), vbExclamation, "エラー"
        End If
    Else
        ' 指定されたファイルが見つからない場合のエラーメッセージ
        MsgBox "指定されたファイルが見つかりません: " & inputFilePath, vbCritical, "エラー"
    End If
Next

' 指定された値が配列に存在するかを確認する関数
Function IsInArray(value, arr)
    Dim i
    For i = LBound(arr) To UBound(arr)
        If value = arr(i) Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

' 文字列の先頭にあるすべての半角スペースを削除する関数
Function TrimStart(str)
    Dim i
    i = 1
    Do While i <= Len(str) And Mid(str, i, 1) = " "
        i = i + 1
    Loop
    TrimStart = Mid(str, i)
End Function

' 現在の日時をタイムスタンプ形式 (yyyymmdd-hhmmss) に変換する関数
Function GetTimestamp()
    Dim now, year, month, day, hour, minute, second
    now = Now
    year = Year(now)
    month = Right("0" & Month(now), 2)
    day = Right("0" & Day(now), 2)
    hour = Right("0" & Hour(now), 2)
    minute = Right("0" & Minute(now), 2)
    second = Right("0" & Second(now), 2)
    GetTimestamp = year & month & day & "-" & hour & minute & second
End Function
