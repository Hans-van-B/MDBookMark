﻿Imports Microsoft.VisualBasic.FileIO
Public Class Form1
    '---- Initialize ----------------------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        xtrace_init()
        xtrace_subs("Form1_Load")
        ReadDefaults()
        Read_Command_Line_Arg()
        LoadMyPosition()            ' Mod SaveToReg
        LoadDefaults()
        FormResize()
        InitBookMarks(True)
        ShowTab9()
        ShowTab11()
        ShowLogTab()

        If FileSystem.FileExists(XYplorer) Then
            xtrace_i("XYplorer exists")
        Else
            xtrace_i("XYplorer does not exists")
            ToolStripMenuItemOpenBookmarksFolderinXYPlorer.Enabled = False
        End If

        Me.Text = AppName & " - " & AppVer
        xtrace_sube("Form1_Load")
    End Sub

    Public Shared Sub WriteInfo(Msg)
        Form1.TextBoxInfo.AppendText(Msg & vbNewLine)
        xtrace(Msg)
    End Sub

    '---- Resize Form ---------------------------------------------------------
    Private Sub FormResize()
        xtrace_subs("FormResize")
        Dim CBD_Top = ((TabPage1.Height - ButtonCreateDT1BookmarkDir.Height) / 2) - 0
        Dim CBD_Left = ((TabPage1.Width - ButtonCreateDT1BookmarkDir.Width) / 2) - 2

        ButtonCreateDT1BookmarkDir.Top = CBD_Top
        ButtonCreateDT1BookmarkDir.Left = CBD_Left

        ButtonCreateDT2BookmarkDir.Top = CBD_Top
        ButtonCreateDT2BookmarkDir.Left = CBD_Left

        ButtonCreateDT3BookmarkDir.Top = CBD_Top
        ButtonCreateDT3BookmarkDir.Left = CBD_Left

        ButtonCreateDT4BookmarkDir.Top = CBD_Top
        ButtonCreateDT4BookmarkDir.Left = CBD_Left

        ButtonCreateDT5BookmarkDir.Top = CBD_Top
        ButtonCreateDT5BookmarkDir.Left = CBD_Left

        ButtonCreateDT6BookmarkDir.Top = CBD_Top
        ButtonCreateDT6BookmarkDir.Left = CBD_Left

        ButtonCreateDT7BookmarkDir.Top = CBD_Top
        ButtonCreateDT7BookmarkDir.Left = CBD_Left

        ButtonCreateDT8BookmarkDir.Top = CBD_Top
        ButtonCreateDT8BookmarkDir.Left = CBD_Left

        ButtonCreateDT9BookmarkDir.Top = CBD_Top
        ButtonCreateDT9BookmarkDir.Left = CBD_Left

        ButtonCreateDT10BookmarkDir.Top = CBD_Top
        ButtonCreateDT10BookmarkDir.Left = CBD_Left

        SaveMyPosition()

        If EnableTimer Then
            xtrace_i("Start Timer")
            Timer1.Start()
        Else
            xtrace_i("Stop Timer")
            Timer1.Stop()
            Timer1.Dispose()
        End If
        xtrace_sube("FormResize")
    End Sub

    Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles MyBase.ResizeEnd
        xtrace_subs("Form1_ResizeEnd")
        FormResize()
        xtrace_sube("Form1_ResizeEnd")
    End Sub

    '==== Main Menu ===========================================================
    '---- New Date
    Private Sub NewDateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewDateToolStripMenuItem.Click
        xtrace_subs("NewDateToolStripMenuItem_Click")

        CreateNewDateDir(BookMarkDir)

        xtrace_sube("NewDateToolStripMenuItem_Click")
    End Sub

    Public Shared Sub CreateNewDateDir(BMDir)
        xtrace_subs("CreateNewDateDir")

        Dim CurDate As String = Date.Now.Date
        xtrace_i(CurDate)
        ActiveBookMarkDir = BMDir & "\" & CurDate
        My.Computer.FileSystem.CreateDirectory(ActiveBookMarkDir)
        xtrace_i("ActiveBookMarkDir = " & ActiveBookMarkDir)
        InitBookMarks(True)

        xtrace_sube("CreateNewDateDir")
    End Sub

    '---- Copy to new Date
    Private Sub CopyToNewDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToNewDataToolStripMenuItem.Click

    End Sub

    '---- Open Bookmarks Folder
    Private Sub ToolStripMenuItemOpenBookmarksFolder_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemOpenBookmarksFolder.Click
        xtrace_subs("OpenBookmarksFolder")
        Process.Start(ActiveBookMarkDir)
        xtrace_sube("OpenBookmarksFolder")
    End Sub

    Private Sub ToolStripMenuItemOpenBookmarksFolderinXYPlorer_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemOpenBookmarksFolderinXYPlorer.Click
        xtrace_subs("OpenBookmarksFolderinXYPlorer")
        Dim WX As Integer = Me.Left + 100
        Dim WY As Integer = Me.Top + 110
        Dim WW As Integer = Me.Width
        Dim Param As String = "/new /readonly /win=normal," & WX.ToString & "," &
            WY.ToString & "," &
            WW.ToString & "," &
            "500" &
            " /path=""" & ActiveBookMarkDir & """ /exp"
        xtrace_i("Param = " & Param)
        Process.Start(XYplorer, Param)
        xtrace_sube("OpenBookmarksFolderinXYPlorer")
    End Sub

    Private Sub ShowLogTab()
        xtrace_subs("ShowLogTab")
        ShowLogTabMethod2()
        xtrace_sube("ShowLogTab")
    End Sub

    Private Sub ShowLogTabMethod1()
        xtrace_subs("ShowLogTabMethod1")
        If ShowLogToolStripMenuItem.Checked Then
            xtrace_i("On")
            TabPageInfo.Show()
            TabPageInfo.Text = "Log"
            TextBoxInfo.Visible = True
            TextBoxInfo.Enabled = True
            TextBoxInfo.Show()
        Else
            xtrace_i("Off")
            TabPageInfo.Hide()
            TabPageInfo.Text = ""
            TextBoxInfo.Visible = False
            TextBoxInfo.Enabled = False
            TextBoxInfo.Hide()
        End If
        xtrace_sube("ShowLogTabMethod1")
    End Sub

    Private Sub ShowLogTabMethod2()
        xtrace_subs("ShowLogTabMethod2")
        If ShowLogToolStripMenuItem.Checked Then
            xtrace_i("On")
            Dim Nr As Integer
            Nr = 9
            If ToolStripMenuItemTab9.Checked Then Nr = 11
            TabControlI1.TabPages.Insert(Nr, TabPageInfo)
            TabPageInfo.Text = "Log"
        Else
            xtrace_i("Off")
            TabControlI1.TabPages.Remove(TabPageInfo)
        End If
        xtrace_sube("ShowLogTabMethod2")
    End Sub

    Private Sub ShowLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowLogToolStripMenuItem.Click
        xtrace_subs("ShowLogToolStripMenuItem_Click")
        ShowLogTab()
        xtrace_sube("ShowLogToolStripMenuItem_Click")
    End Sub
    '---- Settings menu -------------------------------------------------------
    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        xtrace_subs("ResetToolStripMenuItem_Click")
        InitBookMarks(False)
        xtrace_subs("ResetToolStripMenuItem_Click")
    End Sub

    Private Sub OpenLogFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenLogFileToolStripMenuItem.Click
        xtrace_subs("OpenLogFile")
        Process.Start(LogFile)
        xtrace_sube("OpenLogFile")
    End Sub

    Private Sub ShowTab9()
        xtrace_subs("ShowTab9")
        Dim MyPos As Integer = 8

        If ToolStripMenuItemTab9.Checked Then
            TabControlI1.TabPages.Insert(MyPos, TabPage9)
            TabControlI1.TabPages.Insert(MyPos + 1, TabPage10)
        Else
            TabControlI1.TabPages.Remove(TabPage9)
            TabControlI1.TabPages.Remove(TabPage10)
        End If
        xtrace_sube("ShowTab9")
    End Sub

    Private Sub ShowTab11()
        xtrace_subs("ShowTab11")
        Dim MyPos As Integer = 10
        If Not ToolStripMenuItemTab9.Checked Then MyPos = MyPos - 2

        If ToolStripMenuItemTab11.Checked Then
            TabControlI1.TabPages.Insert(MyPos, TabPage11)
            TabControlI1.TabPages.Insert(MyPos + 1, TabPage12)
        Else
            TabControlI1.TabPages.Remove(TabPage11)
            TabControlI1.TabPages.Remove(TabPage12)
        End If
        xtrace_sube("ShowTab11")
    End Sub

    Private Sub ToolStripMenuItemTab9_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemTab9.Click
        ShowTab9()
    End Sub

    Private Sub ToolStripMenuItemTab11_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemTab11.Click
        ShowTab11()
    End Sub

    '==== Exit ================================================================
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        xtrace_subs("ExitToolStripMenuItem_Click")
        exit_program()
        xtrace_sube("ExitToolStripMenuItem_Click")
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        xtrace_subs("Form1_FormClosing (X)")
        exit_program()
        xtrace_sube("Form1_FormClosing (X)")
    End Sub

    '==== Buttons =============================================================
    '---- Create Bookmark directories
    Private Sub ButtonCreateDT1BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT1BookmarkDir.Click
        xtrace_subs("ButtonCreateDT1BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT1)
        InitBrowser1()
        xtrace_sube("ButtonCreateDT1BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT2BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT2BookmarkDir.Click
        xtrace_subs("ButtonCreateDT2BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT2)
        InitBrowser2()
        xtrace_sube("ButtonCreateDT2BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT3BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT3BookmarkDir.Click
        xtrace_subs("ButtonCreateDT3BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT3)
        InitBrowser3()
        xtrace_sube("ButtonCreateDT3BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT4BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT4BookmarkDir.Click
        xtrace_subs("ButtonCreateDT4BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT4)
        InitBrowser4()
        xtrace_sube("ButtonCreateDT4BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT5BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT5BookmarkDir.Click
        xtrace_subs("ButtonCreateDT5BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT5)
        InitBrowser5()
        xtrace_sube("ButtonCreateDT5BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT6BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT6BookmarkDir.Click
        xtrace_subs("ButtonCreateDT6BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT6)
        InitBrowser6()
        xtrace_sube("ButtonCreateDT6BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT7BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT7BookmarkDir.Click
        xtrace_subs("ButtonCreateDT7BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT7)
        InitBrowser7()
        xtrace_sube("ButtonCreateDT7BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT8BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT8BookmarkDir.Click
        xtrace_subs("ButtonCreateDT8BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT8)
        InitBrowser8()
        xtrace_sube("ButtonCreateDT8BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT9BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT9BookmarkDir.Click
        xtrace_subs("ButtonCreateDT9BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT9)
        InitBrowser9()
        xtrace_sube("ButtonCreateDT9BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT10BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT10BookmarkDir.Click
        xtrace_subs("ButtonCreateDT10BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT10)
        InitBrowser10()
        xtrace_sube("ButtonCreateDT10BookmarkDir_Click")
    End Sub


    Private Sub ButtonCreateDT11BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT11BookmarkDir.Click
        xtrace_subs("ButtonCreateDT11BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT11)
        InitBrowser11()
        xtrace_sube("ButtonCreateDT11BookmarkDir_Click")
    End Sub

    Private Sub ButtonCreateDT12BookmarkDir_Click(sender As Object, e As EventArgs) Handles ButtonCreateDT12BookmarkDir.Click
        xtrace_subs("ButtonCreateDT12BookmarkDir_Click")
        My.Computer.FileSystem.CreateDirectory(BookmarkDirDT11)
        InitBrowser12()
        xtrace_sube("ButtonCreateDT12BookmarkDir_Click")
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        xtrace_i("Save " & TabInfoTextFile)
        Try
            RichTextBox1.SaveFile(TabInfoTextFile)
        Catch ex As Exception
            xtrace(ex.Message)
        End Try

    End Sub

    ' Werkt niet
    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        xtrace_subs("WebBrowser1_DoubleClick")
        xtrace_sube("WebBrowser1_DoubleClick")
    End Sub

    Private Sub Form1_Move(sender As Object, e As EventArgs) Handles MyBase.Move
        ' SaveMyPosition()  See FormResizeEnd
    End Sub
    '---- Allow Browsing
    Private Sub AllowBrowsingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AllowBrowsingToolStripMenuItem.Click
        ' Get the active Tab
        Dim ActiveTab As TabPage
        ActiveTab = TabControlI1.SelectedTab

        ' Get the WebBrowser inside it
        Dim MyControl As Control
        Dim ActiveBrowser As WebBrowser

        For Each MyControl In ActiveTab.Controls
            If TypeOf MyControl Is WebBrowser Then
                ActiveBrowser = MyControl

                ' Apply the setting
                If ActiveBrowser.AllowNavigation Then
                    ActiveBrowser.AllowNavigation = False
                    ToolStripStatusLabel2.Text = "AllowNavigation = False"
                Else
                    ActiveBrowser.AllowNavigation = True
                    ToolStripStatusLabel2.Text = "AllowNavigation = True"
                End If
            End If
        Next

    End Sub

    Private Sub TabControlI1_TabIndexChanged(sender As Object, e As EventArgs) Handles TabControlI1.TabIndexChanged
    End Sub

    Private Sub TabControlI1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControlI1.SelectedIndexChanged
        ToolStripStatusLabel2.Text = " "

    End Sub

    '---- Follow other instances
    Private Sub FollowOtherInstancesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FollowOtherInstancesToolStripMenuItem.Click
        SaveDefaults()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        xtrace_subs("Timer1_Tick", TimerTraceLevel)
        xtrace_i("TimerTraceLevel = " & TimerTraceLevel.ToString, TimerTraceLevel)

        If FollowOtherInstancesToolStripMenuItem.Checked Then SaveToReg.CheckMyPosition()
        SaveToReg.LoadDefaults()   ' Always follow defaults if they are changed

        xtrace_sube("Timer1_Tick", TimerTraceLevel)
    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        xtrace_subs("Form1_Activated")
        ResetFormPosition.FormActivate()
        xtrace_sube("Form1_Activated")
    End Sub

    Private Sub HelpToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuHelp.Click
        ShowHelp()
    End Sub

    Private Sub HelpAboutToolStripHelpAbout_Click(sender As Object, e As EventArgs) Handles HelpAboutToolStripHelpAbout.Click
        ShowHelpAbout()
    End Sub
    '---- Go to date ----
    Private Sub ToolStripGoToDate_Click(sender As Object, e As EventArgs) Handles ToolStripGoToDate.Click
        My.Computer.FileSystem.WriteAllText(LogFile, "=========================================================================" & vbNewLine, True)
        xtrace_subs("ToolStripGoToDate_Click")

        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.FolderBrowserDialog1.Description = "Select the date folder"
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        Dim SDir As Environment.SpecialFolder
        SDir = Environment.SpecialFolder.UserProfile
        Me.FolderBrowserDialog1.RootFolder = SDir
        Me.FolderBrowserDialog1.SelectedPath = BookMarkDir
        Me.FolderBrowserDialog1.ShowDialog()
        ActiveBookMarkDir = FolderBrowserDialog1.SelectedPath
        xtrace_i("Selected folder = " & ActiveBookMarkDir)
        InitBookMarks(False)

        xtrace_sube("ToolStripGoToDate_Click")
    End Sub
End Class
