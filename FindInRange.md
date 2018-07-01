<pre>
Function FindInRange(searchValue, searchRange, Optional useLookAt) As range
Dim startRange As range, resultRange As range
Dim findRange As range
Dim SHSearch As String, WBSearch As String

On Error GoTo gotError

10  If (IsMissing(useLookAt)) Then useLookAt = xlPart

20  SHSearch = searchRange.Parent.Name
30  WBSearch = searchRange.Parent.Parent.Name
    
    ' following line incase searchValue is in the first cell of the searchRange
40  Set startRange = searchRange.Cells(searchRange.Rows.count, searchRange.Columns.count)
    '
    '  LookAt:= xlWhole, xlPart
    '  SearchOrder:= xlByRows, xlByColumns
    '  SearchDirection:= xlNext, xlPrevious
    '  searchformat:=  True, False
    '
50  Set resultRange = searchRange.Find(What:=searchValue, _
                    After:=startRange, _
                    LookAt:=useLookAt, _
                    LookIn:=xlValues, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    searchformat:=False, _
                    MatchCase:=False)
60  If Not (resultRange Is Nothing) Then
70      startAddress = resultRange.Address(True, True, xlA1)
80      Set findRange = resultRange
90      Do
            'debug_print findRange.Address
         'Set findRange = Workbooks(WBSearch).Worksheets(SHSearch).FindNext(after:=findRange)
100         Set findRange = searchRange.FindNext(After:=findRange)

            t = findRange.Address(True, True, xlA1)
110            If (findRange.Address(True, True, xlA1) = startAddress) Then Exit Do
            'findRange.Copy
120          Set resultRange = MyUnion(resultRange, findRange)
130      Loop
140  End If
150 Set FindInRange = resultRange
160 Set resultRange = Nothing
170 Set startRange = Nothing
    Exit Function
gotError:
     MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="FindInRange"
    Stop
    Resume Next
End Function
</pre>
