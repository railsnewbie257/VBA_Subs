'
' This routine is a workhorse
' It checks to see if the provided object is connected
' if not it checks if the global object is connected
' if so, it uses the global connection
' otherwise, it opens a new connection and saves to the global connection
'
Function TDCheckConnection(Optional TDConn)
Dim haderror As Boolean

    Call StatusbarDisplay("TDCheckConnection: Check is Nothing.")
    haderror = False
    If Not TDConn Is Nothing Then
        Set TDCheckConnection = TDConn
        Exit Function
    End If
    If TDGlbConnection Is Nothing Then
        Call StatusbarDisplay("TDCheckConnection: Allocate New.")
        Set TDConn = New ADODB.Connection
    Else
        Set TDConn = TDGlbConnection
    End If
    
    Call StatusbarDisplay("TDCheckConnection: Check Open or Closed")
    If TDConn.State = adStateClosed Then
        userName = LCase(Environ$("Username"))
        Password = thisworkbook.Sheets("Top").Cells(1, 1)
        If (Len(userName) = 0 Or Len(Password) = 0) Or Password = "" Then
            LoginForm.Show
            If formCancel Then
                Set TDCheckConnection = Nothing
                Exit Function
            End If
        End If
        
        Call StatusbarDisplay("TDCheckConnection: Opening...")
        
        loginString = "DSN=OGE;Databasename=dbc;Uid=" & userName & ";PWD=" & Password & ";Authentication Mechanism=LDAP;"
        'loginString = "DSN=OGE2;"
        
        On Error GoTo LoginError
        TDConn.ConnectionTimeout = 0 'To wait till the query finishes without generating error
        
        TDConn.Open loginString
        Call StatusbarDisplay("TDCheckConnection: Config")
        Application.ODBCTimeout = 900
        TDConn.CommandTimeout = 1200
        '
        ' Save Password
        '
        If Not haderror Then
            With thisworkbook.Sheets("Top")
                .Cells(1, 1) = Password
                .Cells(1, 1).Font.ThemeColor = xlThemeColorDark1
                .Cells(1, 1).Font.TintAndShade = 0
            End With
        Else
            TDCheckConnection (TDConn)
            Set TDCheckConnection = Nothing
            TDConn.Close
            Set TDConn = Nothing
            Exit Function
        End If
    End If
    
    Call StatusbarDisplay("TDCheckConnection: Opened")
    Set TDGlbConnection = TDConn
    Set TDCheckConnection = TDConn
    Exit Function
    
LoginError:
    MsgBox "TDCheckConnection: " & vbNewLine & Err.Description & vbNewLine & vbNewLine & loginString, Title:="Login Error"
    thisWorkbook.Sheets("Pallette").Cells(1, 1) = "" ' only way to correct an incorrect Password
    haderror = True
    On Error GoTo 0
    Resume Next
    
End Function
</pre>
