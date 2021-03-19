Module SaveToReg
    Dim MyRootKey As String = "Software\TA\" & AppName
    Dim MyRootKey2 As String = "HKEY_CURRENT_USER\Software\TA\" & AppName

    ' Add to FormResizeEnd
    Sub SaveMyPosition()
        xtrace_subs("SaveMyPosition")
        Dim FormX As Integer = Form1.Left
        Dim FormY As Integer = Form1.Top
        xtrace_i("To " & MyRootKey2)

        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "Form X", FormX)
        My.Computer.Registry.SetValue(MyRootKey2 & "\Preferences", "Form Y", FormY)
        xtrace_sube("SaveMyPosition")
    End Sub

    ' Add to FormLoad
    Sub LoadMyPosition()
        xtrace_subs("LoadMyPosition")
        Dim FormX As Integer
        Dim FormY As Integer
        FormX = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "Form X", 100)
        FormY = My.Computer.Registry.GetValue(MyRootKey2 & "\Preferences", "Form Y", 100)

        Form1.Left = FormX
        Form1.Top = FormY
        xtrace_sube("LoadMyPosition")
    End Sub

End Module
