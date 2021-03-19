Module Defaults
    Sub ReadDefaults()
        xtrace_subs("ReadDefaults")
        Dim ReadFile
        Dim Template As String = ""

        If My.Computer.FileSystem.FileExists(IniFile1) Then
            WriteInfo("Read Defaults " & IniFile1)
            ReadFile = My.Computer.FileSystem.OpenTextFileReader(IniFile1)

        ElseIf My.Computer.FileSystem.FileExists(IniFile2) Then
            WriteInfo("Read Defaults " & IniFile2)
            ReadFile = My.Computer.FileSystem.OpenTextFileReader(IniFile2)

        Else
            WriteInfo("Failed to read:")
            WriteInfo("   " & IniFile1)
            WriteInfo("Or " & IniFile2)
            CreateIniFile()
            Exit Sub
        End If

        Dim Line As String
        Dim Group As String = ""
        Dim P1, P2 As Integer
        Dim DName, DVal As String

        While Not ReadFile.EndOfStream
            Try
                Line = ReadFile.ReadLine()

                '---- Remark lines ----
                If Left(Line, 1) = "#" Then
                    Continue While
                End If

                '---- Read Group ----------
                P1 = InStr(Line, "[")
                P2 = InStr(Line, "]")

                If P1 = 1 And P2 > 2 Then
                    Group = Mid(Line, 2, P2 - 2)
                    xtrace("Group = " & Group)
                    Continue While
                End If

                '---- Pick Lists ----
                ' Reset group (Optional)
                If Len(Line) < 2 Then
                    Group = ""
                    Continue While
                End If

                If Group = "COMBOBOX01" Then
                    'Form1.COMBOBOX01.Items.Add(Line)
                End If

                '---- Read Defaults
                P1 = InStr(Line, "=")
                DName = Left(Line, P1 - 1)
                DVal = Mid(Line, P1 + 1)
                xtrace("Default " & DName & "=" & DVal)

                If Group = "INIT" Then

                    If DName = "WindowState" Then
                        If UCase(DVal) = "MIN" Then
                            xtrace("Minimize form")
                            Form1.WindowState = FormWindowState.Minimized
                        End If
                    End If

                    If DName = "BookMarkDir" Then
                        BookMarkDir = DVal
                        xtrace("Set BookMarkDir = " & BookMarkDir)
                    End If

                    If DName = "TabMode" Then
                        TabMode = Val(DVal)
                        xtrace("Set TabMode = " & DVal)

                        xtrace_init_substr()
                    End If

                    If DName = "DT1TXT" Then
                        xtrace_i("Set DT1TXT = " & DVal)
                        DefltDesktopTitle(0) = DVal
                    End If
                    If DName = "DT2TXT" Then
                        xtrace_i("Set DT2TXT = " & DVal)
                        DefltDesktopTitle(1) = DVal
                    End If
                    If DName = "DT3TXT" Then DefltDesktopTitle(2) = DVal
                    If DName = "DT4TXT" Then DefltDesktopTitle(3) = DVal
                    If DName = "DT5TXT" Then DefltDesktopTitle(4) = DVal
                    If DName = "DT6TXT" Then DefltDesktopTitle(5) = DVal
                    If DName = "DT7TXT" Then DefltDesktopTitle(6) = DVal
                    If DName = "DT8TXT" Then DefltDesktopTitle(7) = DVal
                    If DName = "DT9TXT" Then DefltDesktopTitle(8) = DVal

                    If DName = "Template" Then
                            Template = DVal
                            xtrace("Set Template = " & Template)
                        End If


                    End If

            Catch ex As Exception

            End Try
        End While
        ReadFile.Dispose()
        xtrace_sube("ReadDefaults")
    End Sub

    Function StringToBoolean(Val As String) As Boolean
        Val = UCase(Val).Trim

        If (Val = "TRUE") _
        Or (Val = "1") _
        Or (Val = "Y") Then
            Return True
        Else
            Return False
        End If
    End Function

    '---- Create ini file -----------------------------------------------------
    Sub CreateIniFile()
        xtrace_subs("CreateIniFile")
        WriteInfo("Create Defaults " & IniFile1)
        My.Computer.FileSystem.WriteAllText(IniFile1, "[INIT]" & vbNewLine, False)
        My.Computer.FileSystem.WriteAllText(IniFile1, "BookMarkDir=" & vbNewLine, True)

        xtrace_sube("CreateIniFile")
    End Sub
End Module
