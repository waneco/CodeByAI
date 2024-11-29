Option Explicit

Dim objArgs, objFSO, inputFilePath, outputFilePath
Dim validExtensions, fileExtension, timestamp

validExtensions = Array(".py", ".bat", ".log", ".txt", ".ps1", ".vbs")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
    MsgBox "ファイルがドラッグアンドドロップされていません。スクリプトを終了します。", vbExclamation, "エラー"
    WScript.Quit
End If

For Each inputFilePath In objArgs
    If objFSO.FileExists(inputFilePath) Then
        fileExtension = LCase("." & objFSO.GetExtensionName(inputFilePath))
        
        If IsInArray(fileExtension, validExtensions) Then
            On Error Resume Next
            Dim objFile, fileContent, line, outputLines
            Set objFile = objFSO.OpenTextFile(inputFilePath, 1)
            If Err.Number <> 0 Then
                MsgBox "ファイルの読み取り中にエラーが発生しました: " & inputFilePath, vbCritical, "エラー"
                WScript.Quit
            End If
            On Error GoTo 0

            fileContent = objFile.ReadAll
            objFile.Close

            If Len(fileContent) = 0 Then
                MsgBox "ファイルが空です: " & inputFilePath, vbExclamation, "注意"
                Continue For
            End If

            outputLines = ""
            For Each line In Split(fileContent, vbCrLf)
                outputLines = outputLines & TrimStart(line) & vbCrLf
            Next

            timestamp = GetTimestamp()
            outputFilePath = objFSO.GetParentFolderName(inputFilePath) & "\" & _
                             objFSO.GetBaseName(inputFilePath) & "_" & timestamp & fileExtension

            Set objFile = objFSO.CreateTextFile(outputFilePath, True)
            objFile.Write outputLines
            objFile.Close

            MsgBox "ファイルの処理が完了しました: " & vbCrLf & inputFilePath & vbCrLf & _
                   "出力ファイル: " & vbCrLf & outputFilePath, vbInformation, "完了"
        Else
            MsgBox "無効なファイル拡張子です: " & fileExtension, vbExclamation, "エラー"
        End If
    Else
        MsgBox "指定されたファイルが見つかりません: " & inputFilePath, vbCritical, "エラー"
    End If
Next

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

Function TrimStart(str)
    TrimStart = Replace(LTrim(str), vbTab, "")
End Function

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
