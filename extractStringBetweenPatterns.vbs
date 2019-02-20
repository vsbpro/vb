' Author   : Vishal B
' Date     : 20/02/2019
' Change   : Methods extractStringBetweenPatterns and testMethodExtractStringBetweenPatterns created.
Function extractStringBetweenPatterns(ByVal inputString As String, startPattern As String, endPattern As String) As String
    Dim startIndex, endStartIndex As Long
    
    startIndex = InStr(1, inputString, startPattern, vbTextCompare)
    If startIndex > 0 Then
        endStartIndex = InStr(1, inputString, endPattern, vbTextCompare)
        If endStartIndex > 0 Then
            startIndex = startIndex + Len(startPattern)
            extractStringBetweenPatterns = Mid(inputString, startIndex, (endStartIndex - startIndex))
        Else
            extractStringBetweenPatterns = ""
        End If
    Else
        extractStringBetweenPatterns = ""
    End If
    
End Function

Sub testMethodExtractStringBetweenPatterns()
    Dim inputString, output As String
    inputString = "<tag class=""code"" href=""abcdedf.html#100234"" id=""something"">"
    If inputString Like "*.html[#]######*" Then
        output = extractStringBetweenPatterns(inputString, "href=""", """ id=")
        Debug.Print "inputString   ["; inputString; "]"
        Debug.Print "output        ["; output; "]"
    End If
End Sub
