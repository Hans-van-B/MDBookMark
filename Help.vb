Module Help
    Dim HelpPage As String = Temp & "\help.html"
    Dim HelpPageHtm As String = AppRoot & "\" & AppName & ".html"
    Dim HelpPagePdf As String = AppRoot & "\" & AppName & ".pdf"

    Sub ShowHelp()
        xtrace("Show Help")

        If (System.IO.File.Exists(HelpPageHtm)) Then
            HelpPage = HelpPageHtm
        ElseIf (System.IO.File.Exists(HelpPagePdf)) Then
            HelpPage = HelpPagePdf
        Else
            xtrace("Did not find " & HelpPageHtm)
            xtrace("Did not find " & HelpPagePdf)

            xtrace("Create page")
            My.Computer.FileSystem.WriteAllText(HelpPage, "<html>" & vbNewLine, False)
            WriteHelp("<head>")
            WriteHelp("")
            WriteHelp("</head>")
            WriteHelp("<H2>" & AppName & " V" & AppVer & " Help Page</H2>")
            WriteHelp("MD Bookmark intends to let you store bookmarks separately for each Virtual Desktop or &quot;Task View&quot;<br>")
            WriteHelp("The bookmark view does not allow you to browse, 
so if you want to store shortcuts to folders, you need to add &quot;explorer&quot; before the path.<br>
<br>
If the window is outside the display area then you can repeatedly click the icon in the task bar.
This will reset the window location.<br>
<i>Please note that when you do this, and another app that has the same function is active,
then that app is also repeatedly activated and you will reset the window location of both apps.</i><br>

<H3>Command line arguments</H3>
--help or -h or /? : Show this help page <br>
--tab=2 : Set the active tab <br>
-v : Verbose increases the log file<br>
")
            WriteHelp("<br>")
            WriteHelp("<br>")
            WriteHelp("The " & AppName & " log can be found here: " & LogFile & "<br>")
            WriteHelp("The ini file can be found here: " & IniFile & "<br>")
            WriteHelp("</body>")
            WriteHelp("</html")
        End If

        Process.Start(HelpPage, "")
    End Sub

    Sub ShowHelpAbout()
        xtrace("Show Help, about")

        xtrace("Create page")
        My.Computer.FileSystem.WriteAllText(HelpPage, "<html>" & vbNewLine, False)
        WriteHelp("<head>")
        WriteHelp("")
        WriteHelp("</head>")
        WriteHelp("<H2>" & AppName & " V" & AppVer & " Help About</H2>")
        WriteHelp("<br>")
        WriteHelp("<br>")
        WriteHelp("<br>")
        WriteHelp("<br>")
        WriteHelp("<br>")
        WriteHelp("<font size=""-1"">")
        WriteHelp(" Dev: %OneDriveCommercial%\VS_Projects_2019\" & AppName & "<br>")
        WriteHelp(" Maint: P:\Dev\Windows\Win10\MDBookmark<br>")
        WriteHelp(" https://github.com/Hans-van-B?tab=repositories <br>")
        WriteHelp(" Inst.: " & AppRoot & "<br>")
        WriteHelp(" The " & AppName & " log can be found here: " & LogFile & "<br>")
        WriteHelp("</font>")
        WriteHelp("</body>")
        WriteHelp("</html")

        Process.Start(HelpPage, "")
    End Sub

    Sub WriteHelp(Line As String)
        My.Computer.FileSystem.WriteAllText(HelpPage, Line & vbNewLine, True)
    End Sub

End Module
