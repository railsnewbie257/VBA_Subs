<b>Dependencies</b>
- [FastloadWrite](https://github.com/ppihoge/VBA_Subs/blob/master/FastloadWrite.md)
- [TDTableExists](https://github.com/ppihoge/VBA_Subs/blob/master/TDTableExists.md)
- [ColumnLastRow](https://github.com/ppihoge/VBA_Subs/blob/master/ColumnLastRow.md)
- CheckReservedWords


<pre>
Sub FastLoad(Optional tableName)
Dim wsh As Object
Dim userTableName As String     ' from user
Dim newTableName As String      ' table name after TrimReplace, user may have used invalid characters
Dim fullTableName As String     ' database table name

Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1 ' or whatever suits you best
Dim emptyColumnCount As Integer: emptyColumnCount = 1
Dim errorCode As Integer
Dim mergeTable As Boolean       ' whether a merge is necessary because the table already exists


On Error GoTo gotError

    Set wsh = CreateObject("WScript.Shell")

    WBOrig = ActiveWorkbook.Name
    SHOrig = ActiveSheet.Name
    Dim c As String
    
    If IsMissing(tableName) Then
        On Error Resume Next
        userTableName = InputBox("Table Name?", Title:="FastLoad", Default:=GLBTableName)
        If IsEmpty(userTableName) Or userTableName = "" Then Exit Sub
    Else
        userTableName = tableName
    End If
    GLBTableName = userTableName
    newTableName = TrimReplace(userTableName)
    If newTableName <> userTableName Then
        retCode = MsgBox("Modifying Table Name:" & vbNewLine & vbNewLine & userTableName & "  ->  " & newTableName, vbOKCancel, Title:="FastLoad")
        If retCode = vbCancel Then Exit Sub
    End If
    
    Call UsageTracker("FastLoad", "Start: " & newTableName)
    '
    '
    DatabaseName = "dl_oge_analytics"
    fullTableName = DatabaseName & "." & newTableName
    Call UsageTracker("FastLoad", fullTableName)
    
    If TDTableExists(fullTableName) Then
        retCode = MsgBox("Table: " & newTableName & " already EXISTS, will APPEND this table.", vbOKCancel)
        If retCode = vbCancel Then Exit Sub
        mergeTable = True
        newTableName = newTableName & "_up"
        fullTableName = DatabaseName & "." & newTableName

    Else
        retCode = MsgBox("Creating Table: " & newTableName, vbOKCancel, Title:="Fastload")
        If retCode = vbCancel Then Exit Sub
        mergeTable = False
    End If
    '
    filePath = "C:\oge\fastload\" & newTableName & ".fl"
    On Error Resume Next
    Kill filePath
    
    On Error GoTo gotError
    Call StatusbarDisplay("Fastload: Setup")
    
    Call FastLoadWrite(filePath, "LOGMECH LDAP;")
    userName = LCase(Environ$("Username"))
    Call FastLoadWrite(filePath, "LOGON TD1/" & userName & "," & Password & ";")
    Call FastLoadWrite(filePath, "DATABASE dl_oge_analytics;")
    '
    ' DROP TABLES ------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & ";")
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & "_ET;")
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & "_UV;")
    '
    ' CREATE TABLE -----------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Create Table")
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "CREATE MULTISET TABLE " & fullTableName & ",")
    Call FastLoadWrite(filePath, "NO FALLBACK,")
    Call FastLoadWrite(filePath, "NO BEFORE JOURNAL,")
    Call FastLoadWrite(filePath, "NO AFTER JOURNAL,")
    Call FastLoadWrite(filePath, "CHECKSUM = DEFAULT,")
    Call FastLoadWrite(filePath, "DEFAULT MERGEBLOCKRATIO")
    Call FastLoadWrite(filePath, "(")
    rightCol = RowLastColumn(1, SHOrig, WBOrig)
    '
    ' COLUMN NAMES -----------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "LoadDate DATE,") ' load date column
    For i = 1 To rightCol
        t = Worksheets(SHOrig).Cells(1, i)
        If t = "" Then
            t = "EmptyColumn" & emptyColumnCount
            Worksheets(SHOrig).Cells(1, i) = t
            emptyColumnCount = emptyColumnCount + 1
        End If
        t = CheckReservedWord(t)
        c = ","
        If i = rightCol Then c = ")"
        Call FastLoadWrite(filePath, t & " varchar(300)" & c)
    Next i

    t = CheckReservedWord(Worksheets(SHOrig).Cells(1, 1))
    Call FastLoadWrite(filePath, "PRIMARY INDEX(" & t & ");") 'set first column as primary index to spread processing
    '
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "BEGIN LOADING " & fullTableName)
    Call FastLoadWrite(filePath, "ERRORFILES " & newTableName & "_ET, " & newTableName & "_UV;")
    Call FastLoadWrite(filePath, "SET RECORD VARTEXT delimiter " & "'|' QUOTE YES " & "'" & """" & "'" & ";")
    '
    ' DEFINE -------------------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Define")
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "DEFINE")
    Call FastLoadWrite(filePath, "in_LoadDate (varchar(20)),")
    For i = 1 To rightCol
        t = Worksheets(SHOrig).Cells(1, i)
        t = "in_" & TrimReplace(t)
        c = ","
        If i = rightCol Then c = ""
        Call FastLoadWrite(filePath, t & " (varchar(300))" & c)
    Next i
    Call FastLoadWrite(filePath, "FILE= " & newTableName & ".txt;")
    '
    ' INSERT --------------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "INSERT INTO " & fullTableName & " (")
    Call FastLoadWrite(filePath, "LoadDate,")
    For i = 1 To rightCol
        t = Worksheets(SHOrig).Cells(1, i)
        t = CheckReservedWord(t)
        c = ","
        If i = rightCol Then c = ")"
        Call FastLoadWrite(filePath, t & c)
    Next i
    '
    ' VALUES --------------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "VALUES (")
    Call FastLoadWrite(filePath, ": in_LoadDate,")
    For i = 1 To rightCol
        t = Worksheets(SHOrig).Cells(1, i)
        t = ": in_" & TrimReplace(t)
        c = ","
        If i = rightCol Then c = ");"
        Call FastLoadWrite(filePath, t & c)
    Next i
    '
    ' END LOADING ---------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "END LOADING;")
    Call FastLoadWrite(filePath, "LOGOFF;")
    '
    ' DATA FILE -----------------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Create Data File")
    filePath = "C:\OGE\fastload\" & newTableName & ".txt"
    On Error Resume Next
    Kill filePath
    
    
    Set foundRange = FindRangeErrors(Cells)
    foundRange.Value = ""  ' clear all #N/As
    
    On Error GoTo gotError
    botRow = ColumnLastRow(1, SHOrig)
    For j = 2 To botRow
        aline = """" & format(Now(), "yyyy-mm-dd") & """" & "|"
        'aline = ""
        For i = 1 To rightCol
            If j = botRow + 10 Then
                ' get the format frm the first line of data, header may only be "General"
                aline = aline & """" & Worksheets(SHOrig).Cells(j + 1, i).NumberFormat & """"
            Else
                t = Worksheets(SHOrig).Cells(j, i).Value
                t = Replace(t, """", "")
                aline = aline & """" & t & """"
            End If
            c = "|"
            If i = rightCol Then c = ""
            aline = aline & c
        Next i
        Call FastLoadWrite(filePath, aline)
    Next j
    Call StatusbarDisplay("Fastload: Shell Run")
    '
    ' Shell DOS command ---------------------------------------------------------------------------
    '
    t = "cmd.exe /c cd /d C:\oge\fastload && fastload < " & newTableName & ".fl"
    output = ShellRun("cmd.exe /c cd /d C:\oge\fastload && fastload < " & newTableName & ".fl")
    
    filePath = "C:\OGE\fastload\" & newTableName & ".log"
    On Error Resume Next
    Kill filePath
    
    On Error GoTo gotError
    Call FastLoadWrite(filePath, output)
    '
    ' Need to MERGE?
    If mergeTable Then Call FastloadMerge(fullTableName)
    '
    ' Extract return code
    '
    LOAD TextForm
    TextForm.txtBody = output
    If InStr(output, "Highest return code encountered = '0'") > 0 Then
        TextForm.txtHeader = "Success"
    Else
        TextForm.txtHeader = "Failed"
    End If
    TextForm.Show
    Unload TextForm

    Call UsageTracker("FastLoad", "Finished")
    Exit Sub
    
gotError:
    t = Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl
    MsgBox t, Title:="Fastload"
    Call UsageTracker("FastLoad", t)
    Stop
    Resume Next
    
End Sub
</pre>
