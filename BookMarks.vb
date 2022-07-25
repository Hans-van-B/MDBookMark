Imports System
Imports System.IO
Module BookMarks
    Function GetActiveBookmarksDir() As String

        xtrace_subs("GetActiveBookmarksDir")
        ' Make a reference to a directory.
        Dim di1 As New DirectoryInfo(BookMarkDir)
        ' Get a reference to each file in that directory.
        Dim DirList As DirectoryInfo() = di1.GetDirectories()
        ' List the names of the Directories.
        Dim di2 As DirectoryInfo
        For Each di2 In DirList
            xtrace_i(di2.Name)
            ActiveBookMarkDir = BookMarkDir & "\" & di2.Name
        Next di2
        xtrace_i("ActiveBookMarkDir = " & ActiveBookMarkDir)

        xtrace_sube("GetActiveBookmarksDir")
        Return ActiveBookMarkDir    ' Global
    End Function
    Sub InitBookMarks(Init As Boolean)
        xtrace_subs("InitBookMarks")
        If Init Then GetActiveBookmarksDir()
        Form1.ToolStripStatusLabel1.Text = ActiveBookMarkDir

        BookmarkDirDT1 = ActiveBookMarkDir & "\DT1"
        BookmarkDirDT2 = ActiveBookMarkDir & "\DT2"
        BookmarkDirDT3 = ActiveBookMarkDir & "\DT3"
        BookmarkDirDT4 = ActiveBookMarkDir & "\DT4"
        BookmarkDirDT5 = ActiveBookMarkDir & "\DT5"
        BookmarkDirDT6 = ActiveBookMarkDir & "\DT6"
        BookmarkDirDT7 = ActiveBookMarkDir & "\DT7"
        BookmarkDirDT8 = ActiveBookMarkDir & "\DT8"
        BookmarkDirDT9 = ActiveBookMarkDir & "\DT9"
        BookmarkDirDT10 = ActiveBookMarkDir & "\DT10"
        BookmarkDirDT11 = ActiveBookMarkDir & "\DT11"
        BookmarkDirDT12 = ActiveBookMarkDir & "\DT12"
        TabInfoTextFile = ActiveBookMarkDir & "\TabInfo.txt"

        EnableNavigate()
        InitBrowserAll()
        wait(2)
        DisableNavigate()

        xtrace_i("Load " & TabInfoTextFile)
        Try
            If My.Computer.FileSystem.FileExists(TabInfoTextFile) Then
                Form1.RichTextBox1.LoadFile(TabInfoTextFile)
            Else
                Dim Nr1, Nr2 As Integer
                For Nr1 = 0 To 9
                    If Len(DefltDesktopTitle(Nr1)) > 1 Then
                        Nr2 = Nr1 + 1
                        Form1.RichTextBox1.AppendText("DT" & Nr2.ToString & " = " & DefltDesktopTitle(Nr1) & vbNewLine)
                    End If
                Next
            End If
        Catch ex As Exception
            xtrace(ex.Message)
        End Try

        xtrace_i("Allow Navigate = " & Form1.WebBrowser1.AllowNavigation.ToString)
        xtrace_sube("InitBookMarks")
    End Sub

    '==== Init Browser ========================================================
    'Form1.WebBrowser1.Url = New Uri(BookmarkDirDT1)
    'Form1.WebBrowser1.Navigate(New Uri(BookmarkDirDT1))

    Sub InitBrowserAll()
        xtrace_subs("InitBrowserAll")
        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT1) Then
            InitBrowser1()
        Else
            ResetBrowser1()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT2) Then
            InitBrowser2()
        Else
            ResetBrowser2()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT3) Then
            InitBrowser3()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT4) Then
            InitBrowser4()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT5) Then
            InitBrowser5()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT6) Then
            InitBrowser6()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT7) Then
            InitBrowser7()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT8) Then
            InitBrowser8()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT9) Then
            InitBrowser9()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT10) Then
            InitBrowser10()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT11) Then
            InitBrowser11()
        End If

        If My.Computer.FileSystem.DirectoryExists(BookmarkDirDT12) Then
            InitBrowser12()
        End If
        xtrace_sube("InitBrowserAll")
    End Sub

    Sub InitBrowser1()
        xtrace_subs("InitBrowser1")
        Form1.WebBrowser1.Visible = True
        Form1.WebBrowser1.Dock = DockStyle.Fill
        Form1.WebBrowser1.Navigate(BookmarkDirDT1)
        xtrace_i("URL = " & BookmarkDirDT1)
        xtrace_sube("InitBrowser1")
    End Sub

    Sub InitBrowser2()
        xtrace_subs("InitBrowser2")
        Form1.WebBrowser2.Visible = True
        Form1.WebBrowser2.Dock = DockStyle.Fill
        Form1.WebBrowser2.Navigate(BookmarkDirDT2)
        xtrace_i("URL = " & BookmarkDirDT2)
        xtrace_sube("InitBrowser2")
    End Sub

    Sub InitBrowser3()
        xtrace_subs("InitBrowser3")
        Form1.WebBrowser3.Visible = True
        Form1.WebBrowser3.Dock = DockStyle.Fill
        Form1.WebBrowser3.Navigate(BookmarkDirDT3)
        xtrace_sube("InitBrowser3")
    End Sub

    Sub InitBrowser4()
        xtrace_subs("InitBrowser4")
        Form1.WebBrowser4.Visible = True
        Form1.WebBrowser4.Dock = DockStyle.Fill
        Form1.WebBrowser4.Navigate(BookmarkDirDT4)
        xtrace_sube("InitBrowser4")
    End Sub

    Sub InitBrowser5()
        xtrace_subs("InitBrowser5")
        Form1.WebBrowser5.Visible = True
        Form1.WebBrowser5.Dock = DockStyle.Fill
        Form1.WebBrowser5.Navigate(BookmarkDirDT5)
        xtrace_sube("InitBrowser5")
    End Sub

    Sub InitBrowser6()
        xtrace_subs("InitBrowser6")
        Form1.WebBrowser6.Visible = True
        Form1.WebBrowser6.Dock = DockStyle.Fill
        Form1.WebBrowser6.Navigate(BookmarkDirDT6)
        xtrace_sube("InitBrowser6")
    End Sub

    Sub InitBrowser7()
        xtrace_subs("InitBrowser7")
        Form1.WebBrowser7.Visible = True
        Form1.WebBrowser7.Dock = DockStyle.Fill
        Form1.WebBrowser7.Navigate(BookmarkDirDT7)
        xtrace_sube("InitBrowser7")
    End Sub

    Sub InitBrowser8()
        xtrace_subs("InitBrowser8")
        Form1.WebBrowser8.Visible = True
        Form1.WebBrowser8.Dock = DockStyle.Fill
        Form1.WebBrowser8.Navigate(BookmarkDirDT8)
        xtrace_sube("InitBrowser8")
    End Sub

    Sub InitBrowser9()
        xtrace_subs("InitBrowser9")
        Form1.WebBrowser9.Visible = True
        Form1.WebBrowser9.Dock = DockStyle.Fill
        Form1.WebBrowser9.Navigate(BookmarkDirDT9)
        xtrace_sube("InitBrowser9")
    End Sub

    Sub InitBrowser10()
        xtrace_subs("InitBrowser10")
        Form1.WebBrowser10.Visible = True
        Form1.WebBrowser10.Dock = DockStyle.Fill
        Form1.WebBrowser10.Navigate(BookmarkDirDT10)
        xtrace_sube("InitBrowser10")
    End Sub

    Sub InitBrowser11()
        xtrace_subs("InitBrowser11")
        Form1.WebBrowser11.Visible = True
        Form1.WebBrowser11.Dock = DockStyle.Fill
        Form1.WebBrowser11.Navigate(BookmarkDirDT11)
        xtrace_sube("InitBrowser11")
    End Sub

    Sub InitBrowser12()
        xtrace_subs("InitBrowser12")
        Form1.WebBrowser12.Visible = True
        Form1.WebBrowser12.Dock = DockStyle.Fill
        Form1.WebBrowser12.Navigate(BookmarkDirDT12)
        xtrace_sube("InitBrowser12")
    End Sub

    '==== Reset Browser =======================================================
    Sub ResetBrowser1()
        xtrace_subs("ResetBrowser1")
        Form1.WebBrowser1.Visible = False
        Form1.WebBrowser1.Dock = DockStyle.None
        Form1.WebBrowser1.Navigate("")
        xtrace_sube("ResetBrowser1")
    End Sub
    Sub ResetBrowser2()
        xtrace_subs("ResetBrowser2")
        Form1.WebBrowser2.Visible = False
        Form1.WebBrowser2.Dock = DockStyle.None
        Form1.WebBrowser2.Navigate("")
        xtrace_sube("ResetBrowser2")
    End Sub

    Sub ResetBrowser3()
        xtrace_subs("ResetBrowser3")
        Form1.WebBrowser3.Visible = False
        Form1.WebBrowser3.Dock = DockStyle.None
        Form1.WebBrowser3.Navigate("")
        xtrace_sube("ResetBrowser3")
    End Sub
    Sub ResetBrowser4()
        xtrace_subs("ResetBrowser4")
        Form1.WebBrowser4.Visible = False
        Form1.WebBrowser4.Dock = DockStyle.None
        Form1.WebBrowser4.Navigate("")
        xtrace_sube("ResetBrowser4")
    End Sub
    Sub ResetBrowser5()
        xtrace_subs("ResetBrowser5")
        Form1.WebBrowser5.Visible = False
        Form1.WebBrowser5.Dock = DockStyle.None
        Form1.WebBrowser5.Navigate("")
        xtrace_sube("ResetBrowser5")
    End Sub
    Sub ResetBrowser6()
        xtrace_subs("ResetBrowser6")
        Form1.WebBrowser6.Visible = False
        Form1.WebBrowser6.Dock = DockStyle.None
        Form1.WebBrowser6.Navigate("")
        xtrace_sube("ResetBrowser6")
    End Sub
    Sub ResetBrowser7()
        xtrace_subs("ResetBrowser7")
        Form1.WebBrowser7.Visible = False
        Form1.WebBrowser7.Dock = DockStyle.None
        Form1.WebBrowser7.Navigate("")
        xtrace_sube("ResetBrowser7")
    End Sub
    Sub ResetBrowser8()
        xtrace_subs("ResetBrowser8")
        Form1.WebBrowser8.Visible = False
        Form1.WebBrowser8.Dock = DockStyle.None
        Form1.WebBrowser8.Navigate("")
        xtrace_sube("ResetBrowser8")
    End Sub
    Sub ResetBrowser9()
        xtrace_subs("ResetBrowser9")
        Form1.WebBrowser9.Visible = False
        Form1.WebBrowser9.Dock = DockStyle.None
        Form1.WebBrowser9.Navigate("")
        xtrace_sube("ResetBrowser9")
    End Sub
    Sub ResetBrowser10()
        xtrace_subs("ResetBrowser10")
        Form1.WebBrowser10.Visible = False
        Form1.WebBrowser10.Dock = DockStyle.None
        Form1.WebBrowser10.Navigate("")
        xtrace_sube("ResetBrowser10")
    End Sub
    Sub ResetBrowser11()
        xtrace_subs("ResetBrowser11")
        Form1.WebBrowser11.Visible = False
        Form1.WebBrowser11.Dock = DockStyle.None
        Form1.WebBrowser11.Navigate("")
        xtrace_sube("ResetBrowser11")
    End Sub
    Sub ResetBrowser12()
        xtrace_subs("ResetBrowser12")
        Form1.WebBrowser12.Visible = False
        Form1.WebBrowser12.Dock = DockStyle.None
        Form1.WebBrowser12.Navigate("")
        xtrace_sube("ResetBrowser12")
    End Sub

    '==== Enable Navigate =====================================================
    Sub EnableNavigate()
        xtrace_subs("EnableNavigate")
        Form1.WebBrowser1.AllowNavigation = True
        Form1.WebBrowser2.AllowNavigation = True
        Form1.WebBrowser3.AllowNavigation = True
        Form1.WebBrowser4.AllowNavigation = True
        Form1.WebBrowser5.AllowNavigation = True
        Form1.WebBrowser6.AllowNavigation = True
        Form1.WebBrowser7.AllowNavigation = True
        Form1.WebBrowser8.AllowNavigation = True
        Form1.WebBrowser9.AllowNavigation = True
        Form1.WebBrowser10.AllowNavigation = True
        Form1.WebBrowser11.AllowNavigation = True
        Form1.WebBrowser12.AllowNavigation = True
        xtrace_sube("EnableNavigate")
    End Sub

    '==== Disable Navigate
    Sub DisableNavigate()
        xtrace_subs("DisableNavigate")
        Form1.WebBrowser1.AllowNavigation = False
        Form1.WebBrowser2.AllowNavigation = False
        Form1.WebBrowser3.AllowNavigation = False
        Form1.WebBrowser4.AllowNavigation = False
        Form1.WebBrowser5.AllowNavigation = False
        Form1.WebBrowser6.AllowNavigation = False
        Form1.WebBrowser7.AllowNavigation = False
        Form1.WebBrowser8.AllowNavigation = False
        Form1.WebBrowser9.AllowNavigation = False
        Form1.WebBrowser10.AllowNavigation = False
        Form1.WebBrowser11.AllowNavigation = False
        Form1.WebBrowser12.AllowNavigation = False

        Form1.WebBrowser1.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser2.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser3.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser4.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser5.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser6.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser7.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser8.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser9.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser10.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser11.WebBrowserShortcutsEnabled = False
        Form1.WebBrowser12.WebBrowserShortcutsEnabled = False
        xtrace_sube("DisableNavigate")
    End Sub

End Module
