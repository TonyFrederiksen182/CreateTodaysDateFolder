Option Explicit

Dim fileSystemObject ' FileSystemObject
Dim folderName       ' フォルダ名
Dim message          ' 表示用メッセージ
Dim inputString      ' キー入力用変数
Dim title            ' タイトル

Set fileSystemObject = WScript.CreateObject("Scripting.FileSystemObject")

folderName = Replace(Date(), "/", "")    '当日の日付を取得し、"/"を取り除く
title = "CreateTodaysDateFolder"
message  = "作成するフォルダ名を入力して下さい。"

inputString = InputBox(message, title, folderName)

If Len(inputString) = 0 Then ' キャンセルが押された場合
	' Do Nothing
Else
    If Err.Number = 0 Then
        If fileSystemObject.FolderExists(inputString) = True Then
            message = "フォルダ " & folderName & " は既に存在しています。"
        Else
            fileSystemObject.CreateFolder(inputString)
            If Err.Number = 0 Then
                message = "フォルダ " & inputString & " を作成しました。"
            Else
                message = "エラー: " & Err.Description
            End If
        End If
        WScript.Echo message
    Else
        WScript.Echo "エラー: " & Err.Description
    End If
End If

Set fileSystemObject = Nothing
