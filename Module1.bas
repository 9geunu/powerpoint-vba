Attribute VB_Name = "Module1"
Option Explicit

Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr

Public UserName As String
Public RelationshipPoint As Long
Public YongJaeScore As Long
Public GeunwooScore As Long

Public Sub InitGlobalVariables()
    RelationshipPoint = 0
    YongJaeScore = 0
    GeunwooScore = 0
    
    Debug.Print "RelationshipPoint : " & RelationshipPoint & ", YongJaeScore : " & YongJaeScore & ", GeunwooScore : " & GeunwooScore
End Sub


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
Public Sub SetUserName()
    Dim val As String
    val = InputBox("Enter player name.", "OK")
    UserName = val
    Debug.Print "You are " & UserName & "."
    Call ReplaceStringInCurrentSlide("xx", UserName)
    Call Refresh
    
End Sub

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

Public Sub ReplaceStringInAllSlides(fWord As String, sRepWord As String)
    
    Dim shp As shape
    Dim FindWord As Variant
    Dim ReplaceWord As Variant
    Dim iWords As Integer
    Dim textLoc As PowerPoint.TextRange

    FindWord = fWord
    ReplaceWord = sRepWord

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
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
      
    Next sld
      
    Debug.Print "" & fWord & " updated to " & sRepWord & " " & iWords & " times"

End Sub

Public Sub Refresh()
    ActivePresentation.SlideShowWindow.View.slide.Shapes.AddTextbox msoTextOrientationHorizontal, 1, 1, 1, 1
End Sub

Sub OnSlideShowPageChange(ByVal SSW As SlideShowWindow)
    
    If SSW.View.CurrentShowPosition = 110 Then
        Call ReplaceStringInCurrentSlide("RP", "" & RelationshipPoint)
        Call Refresh
    End If
    
End Sub

Public Sub HttpRequest(slide As Long, choice As Long)

    Dim URL As String, JSONString As String, Response As String

    URL = "https://localhost:8080" 'replace with your server
    JSONString = "{""fields"": {""userName"": {""stringValue"": """ & UserName & """},""slide"": {""integerValue"": " & slide & "}, ""choice"": {""integerValue"": " & choice & "}}}"
    
    Response = HTTPPost(URL, JSONString)
    
    Debug.Print Response

End Sub

Public Sub HandleShapeClick(shape As shape)

    Dim sld As Long
    sld = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    
    Select Case shape.Name
        Case Is = "Rectangle 1"
            Call HttpRequest(sld, 1)
        Case Is = "Rectangle 2"
            Call HttpRequest(sld, 2)
    End Select
    
End Sub

Public Sub GoToSlide(shape As shape)

    Dim destination As Long
    destination = CLng(shape.Name)
    
    With ActivePresentation.SlideShowWindow.View
        .GoToSlide destination
    End With
    
End Sub

Public Sub GoToSlideWithLong(destination As Long)

    With ActivePresentation.SlideShowWindow.View
        .GoToSlide destination
    End With
    
End Sub

Public Sub GoToNextSlide()
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Public Sub UpdateRelationshipPoint(shp As shape)
    
    RelationshipPoint = RelationshipPoint + CLng(shp.Name)
    Debug.Print "RelationshipPoint : " & RelationshipPoint
    Call GoToNextSlide
    
End Sub

Public Sub IncrementYongJaeScore()

    YongJaeScore = YongJaeScore + 1
    Debug.Print "YongJaeScore : " & YongJaeScore
    Call GoToNextSlide
    
End Sub

Public Sub UpdateScoreAndGoTo(shp As shape)
    'G 1 43
    'Y 1 45
    'R 1 46
    'pointName score destination
    
    Dim arrSplitStrings() As String
    Dim pointName As String
    Dim score As Long
    Dim destination As Long
    
    Debug.Print shp.Name
    arrSplitStrings = Split(shp.Name, ";")
    
    pointName = arrSplitStrings(0)
    score = CLng(arrSplitStrings(1))
    destination = CLng(arrSplitStrings(2))
    
    Select Case pointName
        Case Is = "Y"
            YongJaeScore = YongJaeScore + 1
            Debug.Print "YongJaeScore : " & YongJaeScore
        Case Is = "G"
            GeunwooScore = GeunwooScore + 1
            Debug.Print "GeunwooScore : " & GeunwooScore
        Case Is = "R"
            RelationshipPoint = RelationshipPoint + score
            Debug.Print "RelationshipPoint : " & RelationshipPoint
    End Select
    
    Call GoToSlideWithLong(destination)
    
End Sub

Public Sub IncrementGeunwooScore()

    GeunwooScore = GeunwooScore + 1
    Debug.Print "GeunwooScore : " & GeunwooScore
    Call GoToNextSlide
    
End Sub

Public Sub FinalChoice()
    
    Debug.Print "YongJaeScore : " & YongJaeScore
    Debug.Print "GeunwooScore : " & GeunwooScore
    
    If YongJaeScore > GeunwooScore Then
        Call GoToSlideWithLong(102)
    ElseIf YongJaeScore < GeunwooScore Then
        Call GoToSlideWithLong(104)
    Else
        Dim x As Long
        x = Int(Rnd * (104 - 102)) + 102
        x = x + (x Mod 2)
        Debug.Print "x : " & x
        Call GoToSlideWithLong(x)
    End If
End Sub

