Module Log
    Dim TabStr As String = ""
    Dim SubStr1 As String = ""
    Dim SubStr2 As String = ""
    Dim Bullit As String = ""

    Sub xtrace_init()
        My.Computer.FileSystem.WriteAllText(LogFile, "" & vbNewLine, False)
        WriteInfo(AppName & " V" & AppVer)
        xtrace("AppRoot = " & AppRoot)
        WriteInfo("LogFile = " & LogFile)

        xtrace_init_substr()
    End Sub

    Sub xtrace_init_substr()
        xtrace("TabMode = " & TabMode.ToString)
        If TabMode = 0 Then ' No tabbing
            TabStr = ""
            SubStr1 = "## "
            Bullit = " * "

        ElseIf TabMode = 1 Then ' Minimal tabbing
            TabStr = " "
            Dim SubStr1 = ">"
            Bullit = ""

        ElseIf TabMode = 2 Then ' Clear Tabbing
            TabStr = "|  "
            SubStr1 = "|->"
            Bullit = ""

        ElseIf TabMode = 3 Then ' Clear Tabbing
            TabStr = "|  "
            SubStr1 = "|->"
            SubStr2 = "|<-"
            Bullit = ""

        End If
    End Sub

    '---- xtrace
    Sub xtrace(Msg As String)
        xtrace(Msg, 2)
    End Sub

    Sub xtrace(Msg As String, TV As Integer)
        If TV <= LTrace Then
            Dim Tab As String = ""
            Dim Nr As Integer
            For Nr = 1 To SubLevel
                Tab = Tab + TabStr
            Next

            My.Computer.FileSystem.WriteAllText(LogFile, Tab + Msg & vbNewLine, True)
        End If
    End Sub

    Sub xtrace_i(Msg As String)
        xtrace(Bullit & Msg)
    End Sub

    Sub xtrace_subs(Msg As String)
        xtrace(SubStr1 & Msg)
        SubLevel = SubLevel + 1
    End Sub

    Sub xtrace_sube(Msg As String)
        SubLevel = SubLevel - 1
        If SubStr2 <> "" Then xtrace(SubStr2 & Msg)
    End Sub

    Sub xtrace_sube()
        xtrace_sube("")
    End Sub

    Sub xtrace_err(Msg As String)
        xtrace("ERROR: " & Msg)
    End Sub

    Public Sub WriteInfo(Msg)
        Form1.TextBoxInfo.AppendText(Msg & vbNewLine)
        xtrace(Msg)
    End Sub
End Module
