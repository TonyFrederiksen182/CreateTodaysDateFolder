Option Explicit

Dim fileSystemObject ' FileSystemObject
Dim folderName       ' �t�H���_��
Dim message          ' �\���p���b�Z�[�W
Dim inputString      ' �L�[���͗p�ϐ�
Dim title            ' �^�C�g��

Set fileSystemObject = WScript.CreateObject("Scripting.FileSystemObject")

folderName = Replace(Date(), "/", "")    '�����̓��t���擾���A"/"����菜��
title = "CreateTodaysDateFolder"
message  = "�쐬����t�H���_������͂��ĉ������B"

inputString = InputBox(message, title, folderName)

If Len(inputString) = 0 Then ' �L�����Z���������ꂽ�ꍇ
	' Do Nothing
Else
    If Err.Number = 0 Then
        If fileSystemObject.FolderExists(inputString) = True Then
            message = "�t�H���_ " & folderName & " �͊��ɑ��݂��Ă��܂��B"
        Else
            fileSystemObject.CreateFolder(inputString)
            If Err.Number = 0 Then
                message = "�t�H���_ " & inputString & " ���쐬���܂����B"
            Else
                message = "�G���[: " & Err.Description
            End If
        End If
        WScript.Echo message
    Else
        WScript.Echo "�G���[: " & Err.Description
    End If
End If

Set fileSystemObject = Nothing
