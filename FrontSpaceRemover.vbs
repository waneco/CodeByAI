' ***********************************************************************
' スクリプト名: FrontSpaceRemover.vbs
' 説明: 
'     ドラッグ＆ドロップで指定されたテキストファイルの各行の先頭にある
'     空白やタブを削除し、タイムスタンプ付きの新しいファイルとして保存します。
'
' 主な機能:
'     1. ドラッグ＆ドロップで渡されたファイルを処理します。
'     2. サポートする拡張子（.py, .bat, .log, .txt, .ps1, .vbs）のファイルに対応。
'     3. 各行の先頭にある余分な空白やタブを削除。
'     4. タイムスタンプ付きのファイル名で保存。
'
' 使用方法:
'     1. このスクリプトをダブルクリックして起動するか、ファイルにドラッグ＆ドロップします。
'     2. 処理結果が保存され、新しいファイルのパスがメッセージで通知されます。
'
' 注意事項:
'     - UTF-8形式のテキストファイルを想定しています。
'     - 元のファイルは変更されません。新しいファイルが同じフォルダに保存されます。
'     - 空のファイルやサポートされていない拡張子のファイルはスキップされます。
'
' バージョン: 1.0
' 作成者: （あなたの名前やニックネーム）
' 作成日: （作成日を記載）
' ***********************************************************************
Option Explicit

' スクリプトの引数やファイル操作関連のオブジェクトを定義
Dim objArgs, objFSO, inputFilePath, outputFilePath
Dim validExtensions, fileExtension, timestamp

' サポートするファイル拡張子を定義
validExtensions = Array(".py", ".bat", ".log", ".txt", ".ps1", ".vbs")

' FileSystemObjectの作成
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments

' 引数が渡されていない場合、エラーを表示して終了
If objArgs.Count = 0 Then
    MsgBox "ファイルがドラッグアンドドロップされていません。スクリプトを終了します。", vbExclamation, "エラー"
    WScript.Quit
End If

' ドラッグ&ドロップされたファイルごとに処理を実行
For Each inputFilePath In objArgs
    If objFSO.FileExists(inputFilePath) Then
        ' ファイルの拡張子を小文字で取得
        fileExtension = LCase("." & objFSO.GetExtensionName(inputFilePath))
        
        ' サポートされている拡張子か確認
        If IsInArray(fileExtension, validExtensions) Then
            On Error Resume Next
            Dim objFile, fileContent, line, outputLines
            
            ' ファイルをUTF-8として読み込む
            Set objFile = objFSO.OpenTextFile(inputFilePath, 1, False, -1)
            If Err.Number <> 0 Then
                MsgBox "ファイルの読み取り中にエラーが発生しました: " & inputFilePath & vbCrLf & "エラー番号: " & Err.Number, vbCritical, "エラー"
                Err.Clear
                Continue For
            End If
            On Error GoTo 0

            ' ファイル内容を読み取り
            fileContent = objFile.ReadAll
            objFile.Close

            ' ファイルが空でないか確認
            If Len(fileContent) = 0 Then
                MsgBox "ファイルが空です: " & inputFilePath, vbExclamation, "注意"
                Continue For
            End If

            ' 行ごとに先頭の空白を削除し、新しい内容を生成
            outputLines = ""
            For Each line In Split(fileContent, vbCrLf)
                outputLines = outputLines & TrimStart(line) & vbCrLf
            Next

            ' タイムスタンプ付きの出力ファイルパスを生成
            timestamp = GetTimestamp()
            outputFilePath = objFSO.GetParentFolderName(inputFilePath) & "\" & _
                             objFSO.GetBaseName(inputFilePath) & "_" & timestamp & fileExtension

            ' 出力ファイルが既に存在していればスキップ
            If objFSO.FileExists(outputFilePath) Then
                MsgBox "出力ファイルが既に存在します: " & outputFilePath, vbExclamation, "注意"
                Continue For
            End If

            ' 新しいファイルに書き込み
            Set objFile = objFSO.CreateTextFile(outputFilePath, True, True) ' UTF-8で書き込む
            objFile.Write outputLines
            objFile.Close

            ' 完了メッセージを表示
            MsgBox "ファイルの処理が完了しました: " & vbCrLf & inputFilePath & vbCrLf & _
                   "出力ファイル: " & vbCrLf & outputFilePath, vbInformation, "完了"
        Else
            ' サポートされていない拡張子のエラーを表示
            MsgBox "無効なファイル拡張子です: " & fileExtension, vbExclamation, "エラー"
        End If
    Else
        ' ファイルが存在しない場合のエラーを表示
        MsgBox "指定されたファイルが見つかりません: " & inputFilePath, vbCritical, "エラー"
    End If
Next

' 配列内に値が存在するかを確認する関数
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

' 文字列の先頭の空白やタブを削除する関数
Function TrimStart(str)
    ' タブや空白を削除
    TrimStart = Replace(LTrim(str), vbTab, "")
End Function

' 現在の日時をタイムスタンプ形式（YYYYMMDD-HHMMSS）で取得する関数
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
