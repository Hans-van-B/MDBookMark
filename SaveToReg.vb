Module SaveToReg
    Dim MyRootKey As String = "Software\TA\" & AppName
    Dim MyRootKey2 As String = "HKEY_CURRENT_USER\Software\TA\" & AppName

    ' Add to FormResizeEnd
    Sub SaveMyPosition()
        xtrace_subs("SaveMyPosition")
        Dim FormX As Integer = Form1.Left
        Dim FormY As Integer = Form1.Top
        Dim FormW As Integer = Form1.Width
        xtrace_i(FormX.ToString & "," & FormY.ToString)
        xtrace_i("To " & MyRootKey2)

        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "Form X", FormX)
        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "Form Y", FormY)
        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "Form W", FormW)

        MyCurrentXPos = FormX
        MyCurrentYPos = FormY
        MyCurrentWidth = FormW
        xtrace_sube("SaveMyPosition")
    End Sub

    Sub ResetMyPosition()
        xtrace_subs("ResetMyPosition")
        Dim FormX As Integer = 100
        Dim FormY As Integer = 100
        Dim FormW As Integer = D_FormW
        xtrace_i(FormX.ToString & "," & FormY.ToString)
        xtrace_i("To " & MyRootKey2)

        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "Form X", FormX)
        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "Form Y", FormY)
        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "Form W", FormW)

        MyCurrentXPos = FormX
        MyCurrentYPos = FormY
        MyCurrentWidth = FormW
        xtrace_sube("ResetMyPosition")
    End Sub

    ' Add to FormLoad
    Sub LoadMyPosition()
        xtrace_subs("LoadMyPosition")
        Dim FormXR As Integer
        Dim FormYR As Integer
        Dim FormWR As Integer
        FormXR = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "Form X", 100)
        FormYR = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "Form Y", 100)
        FormWR = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "Form W", D_FormW)

        Form1.Left = FormXR
        Form1.Top = FormYR
        Form1.Width = FormWR
        xtrace_sube("LoadMyPosition")
    End Sub

    ' This routine needs to be called by a time interrupt,
    ' it follows other instances if the location is changed

    ' These variables need to be updated if the form is moved, see SaveMyPosition
    Public MyCurrentXPos As Integer = 10
    Public MyCurrentYPos As Integer = 10
    Public MyCurrentWidth As Integer = D_FormW

    Sub CheckMyPosition()
        xtrace_subs("CheckMyPosition", TimerTraceLevel)
        Dim FormXR As Integer
        Dim FormYR As Integer
        Dim FormWR As Integer
        FormXR = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "Form X", 100)
        FormYR = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "Form Y", 100)
        FormWR = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "Form W", D_FormW)

        If MyCurrentXPos <> FormXR Then
            Form1.Left = FormXR
            MyCurrentXPos = FormXR
        End If

        If MyCurrentYPos <> FormYR Then
            Form1.Top = FormYR
            MyCurrentYPos = FormYR
        End If

        If MyCurrentWidth <> FormWR Then
            Form1.Width = FormWR
            MyCurrentWidth = FormWR
        End If

        xtrace_sube("CheckMyPosition", TimerTraceLevel)
    End Sub

    Sub SaveDefaults()
        xtrace_subs("SaveDefaults")
        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "FollowOtherInstances", Form1.FollowOtherInstancesToolStripMenuItem.Checked)
        xtrace_sube("SaveDefaults")
    End Sub

    Sub LoadDefaults()
        xtrace_subs("LoadDefaults", TimerTraceLevel)
        Dim Result As Boolean
        Result = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "FollowOtherInstances", True)
        Form1.FollowOtherInstancesToolStripMenuItem.Checked = Result
        xtrace_i("FollowOtherInstances = " & Result.ToString, TimerTraceLevel)
        xtrace_sube("LoadDefaults", TimerTraceLevel)
    End Sub
End Module
