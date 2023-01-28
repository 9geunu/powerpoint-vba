# powerpoint-vba

## PowerPoint VBA code blocks

VBA code blocks for personal PowerPoint game development projects

### HttpModule
HttpModule contains http connection functions for mac OS. <br>
HttpModule uses `curl` for http connection, not `WinHttpRequest`.

```vb
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
```

### HttpModule Usage
Example below is uploading json data to firestore with `HTTPPost` function.
```vb
Dim URL As String, JSONString As String, Response As String
Dim UserName As String, slide As Long, choice As Long

UserName = "9geunu"
slide = 1
choice = 2

URL = "https://firestore.googleapis.com/v1beta1/projects/{project_id}/databases/(default)/documents/{collection_name}"
JSONString = "{""fields"": {""userName"": {""stringValue"": """ & UserName & """},""slide"": {""integerValue"": " & slide & "}, ""choice"": {""integerValue"": " & choice & "}}}"

Response = HTTPPost(URL, JSONString)

Debug.Print Response
```
---
## Replace string in current slide
If you need to replace string during active presentation, use below code blocks.
```vb
'Reference : https://www.access-programmers.co.uk/forums/threads/using-access-vba-to-change-a-powerpoint-text-box-value.315303/
Public Sub ReplaceStringInCurrentSlide(fWord As String, sRepWord As String)
    
    Dim shp As shape
    Dim FindWord As Variant
    Dim ReplaceWord As Variant
    Dim iWords As Integer
    Dim textLoc As PowerPoint.TextRange

    FindWord = fWord
    ReplaceWord = sRepWord
    
    For Each shp In ActivePresentation.SlideShowWindow.View.slide.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                Set textLoc = shp.TextFrame.TextRange.Find(FindWord)
                If Not (textLoc Is Nothing) Then
                    textLoc.Text = ReplaceWord
                    iWords = iWords + 1
                End If
            End If
        End If
    Next shp
      
    Debug.Print "" & fWord & " updated to " & sRepWord & " " & iWords & " times"

End Sub

Public Sub Refresh()
    ActivePresentation.SlideShowWindow.View.slide.Shapes.AddTextbox msoTextOrientationHorizontal, 1, 1, 1, 1
End Sub
```

### Usage
If you want to get user name from input box and replace placeholder with user name, you can use `ReplaceStringInCurrentSlide("placeholder", UserName)`. <br>
You should call `Refresh` after calling `ReplaceStringInCurrentSlide` to reflect changes.

```vb
Public Sub SetUserName()
    Dim UserName As String
    UserName = InputBox("Enter user name.", "OK")
    Debug.Print "You are " & UserName & "."
    Call ReplaceStringInCurrentSlide("placeholder", UserName)
    Call Refresh
    
End Sub
```
