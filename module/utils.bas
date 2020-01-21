Attribute VB_Name = "utilsModule"
Public Function makeDirectory(ByVal currentDir As String, Optional parentDir As String)
    
    Dim fileDirTmp As String

    If Right(currentDir, 1) <> "\" Then currentDir = currentDir & "\"
    
    fileDirTmp = Mid(currentDir, InStr(currentDir, "\") + 1)
    parentDir = parentDir & Mid(currentDir, 1, InStr(currentDir, "\"))
    currentDir = fileDirTmp
    
    If Dir(parentDir, vbDirectory) = "" Then MkDir parentDir
    If currentDir <> "" Then makeDirectory currentDir, parentDir
    
    makeDirectory = currentDir
End Function


'*********************************************************
'* ログファイルを作成して、書き出す
'*    - 引数　：msg(テキスト), name(サブプロシージャ名)
'*    - 返り値：Nothing
'*********************************************************
Public Function writeLog(msg As String, name As String)
    Dim FSO As Object
    Dim LOG As Object
    Dim now As String
    
    now = Format(Now, "yyyymmddhhmmss")
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(glb_ERROR_LOG_FILE_PATH) Then
        If FSO.FileExists(glb_ERROR_LOG_FILE_PATH & "\" & today & "_log.txt") = False Then
            FSO.CreateTextFile glb_ERROR_LOG_FILE_PATH & "\" & today & "_log.txt"
        End If
    
        Set LOG = FSO.OpenTextFile(glb_ERROR_LOG_FILE_PATH & "\" & today & "_log.txt", 8)
        LOG.WriteLine Now & vbTab & name & vbTab & msg
        Set LOG = Nothing
    End If
    Set FSO = Nothing
End Function

' JIS漢字第一・第二水準判定関数
Function isSJIS(ByVal argStr As String) As Boolean
    Dim sQuestion As String
    Dim i As Long
    Dim char As String
    Dim char_code As Long
    Dim JisLevel2 As Long
    JisLevel2 = Asc("熙") ' 第二水準最後
    
    For i = 1 To Len(argStr)
        char = Mid(argStr, i, 1)
        char_code = Asc(char)
        If char_code > JisLevel2 And char <> " " And char <> "　" And _
            Not char Like "[a-zA-Z]" And Not char Like "[\!-~]" Then
            isSJIS = False
            Exit Function
        End If
    Next
    isSJIS = True
End Function
