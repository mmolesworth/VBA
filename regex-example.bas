Function RegexExtract(ByVal text As String, _
                      ByVal extract_what As String, _
                      Optional seperator As String = "") As String

Dim i As Long, j As Long
Dim result As String
Dim allMatches As Object
Dim RE As Object
Set RE = CreateObject("vbscript.regexp")

RE.pattern = extract_what
RE.Global = True
Set allMatches = RE.Execute(text)

For i = 0 To allMatches.count - 1
    For j = 0 To allMatches.Item(i).submatches.count - 1
        result = result & seperator & allMatches.Item(i).submatches.Item(j)
    Next
Next

If Len(result) <> 0 Then
    result = Right(result, Len(result) - Len(seperator))
End If

RegexExtract = result

End Function
