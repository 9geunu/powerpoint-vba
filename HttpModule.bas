Attribute VB_Name = "HttpModule"
Option Explicit

Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr

'Reference : https://stackoverflow.com/questions/15981960/how-do-i-issue-an-http-get-from-excel-vba-for-mac
Function execShell(command As String, Optional ByRef exitCode As Long) As String
    Dim file As LongPtr
    file = popen(command, "r")

    If file = 0 Then
        Exit Function
    End If

    While feof(file) = 0
        Dim chunk As String
        Dim read As Long
        chunk = Space(50)
        read = fread(chunk, 1, Len(chunk) - 1, file)
        If read > 0 Then
            chunk = Left$(chunk, read)
            execShell = execShell & chunk
        End If
    Wend

    exitCode = pclose(file)
End Function

Function HTTPGet(sUrl As String) As String

    Dim sCmd As String
    Dim sResult As String
    Dim lExitCode As Long

    sCmd = "curl --get """ & sUrl & """"
    sResult = execShell(sCmd, lExitCode)

    HTTPGet = sResult

End Function

Function HTTPPost(sUrl As String, sData As String) As String

    Dim sCmd As String
    Dim sResult As String
    Dim lExitCode As Long

    sCmd = "curl -d '" & sData & "' -H ""Accept: application/json"" -H ""Content-Type: application/json"" """ & sUrl & """"
    
    sResult = execShell(sCmd, lExitCode)
    
    Debug.Print lExitCode

    HTTPPost = sResult

End Function