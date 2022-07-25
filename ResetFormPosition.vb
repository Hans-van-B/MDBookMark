Module ResetFormPosition
    Dim LastActivated As Integer = 0
    Dim ActivatedCount As Integer = 0
    Sub FormActivate()
        xtrace_subs("FormActivate")
        Dim Now As Integer = GetSeconds()
        If Now > LastActivated + 2 Then
            xtrace_i("Reset")
            LastActivated = Now
            ActivatedCount = 1
        Else
            ActivatedCount += 1
            xtrace_i("Count: " & ActivatedCount.ToString)
        End If

        If ActivatedCount > 3 Then
            xtrace_i("Reset Form Position!")
            Form1.Timer1.Stop()

            Form1.BringToFront()
            Form1.Select()

            Form1.Top = 100
            Form1.Left = 100
            Form1.Width = D_FormW
            ActivatedCount = 1

            If Form1.Top <> 100 Then
                xtrace("Error, retry")
                Form1.Top = 100
                Form1.Left = 100
                Form1.Width = D_FormW
            End If

            ResetMyPosition()
        End If
        xtrace_sube("FormActivate")
    End Sub
End Module
