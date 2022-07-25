Module Glob
    Public AppName As String = "MDBookmark"
    Public AppVer As String = "0.12"
    Public AppRoot As String = Application.StartupPath
    Public CD As String = My.Computer.FileSystem.CurrentDirectory

    ' Log File
    Public Temp As String = Environment.GetEnvironmentVariable("TEMP")
    Public LTrace As Integer = 2
    Public ErrorCount As Integer = 0
    Public WarningCount As Integer = 0
    Public SubLevel As Integer = 0
    Public TabMode As Integer = 2

    ' Defaults
    Public IniFile1 As String = AppRoot & "\" & AppName & ".ini"
    Public IniFile2 As String = AppRoot & "\Data\" & AppName & ".ini"
    Public BookMarkDir As String = "C:\Temp\Bookmarks"
    Public D_FormW As Integer = 300
    Public TimerTraceLevel As Integer = 3
    Public EnableTimer As Boolean = True

    Public DefltDesktopTitle() As String = {"Default Support", "MDP Dev", "MDP Test", "TA Inst Dev", "TA Inst Test & Release", "VB Dev", "", "", "", ""}

    ' BookMarks
    Public ActiveBookMarkDir As String
    Public BookmarkDirDT1 As String
    Public BookmarkDirDT2 As String
    Public BookmarkDirDT3 As String
    Public BookmarkDirDT4 As String
    Public BookmarkDirDT5 As String
    Public BookmarkDirDT6 As String
    Public BookmarkDirDT7 As String
    Public BookmarkDirDT8 As String
    Public BookmarkDirDT9 As String
    Public BookmarkDirDT10 As String
    Public BookmarkDirDT11 As String
    Public BookmarkDirDT12 As String
    Public TabInfoTextFile As String

    Public EndDelay As Integer = 2
    Public EndPause As Boolean = False
    Public ShowGui As Boolean = False
    Public ExitProgram As Boolean = False

    Public XYplorer As String = "C:\Program Files (x86)\XYplorer\XYplorer.exe"
End Module
