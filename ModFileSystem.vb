﻿' Validate by setting the log file readonly
Module ModFileSystem
    Function WriteTxtToFile(FileName As String,
                       Msg As String,
                       Append As Boolean,
                       ErrNr As Integer,
                       ErrDetails As String,
                       Hint As String,
                       ShowDialog As Boolean,
                       Fatal As Boolean) As Boolean
        Dim Result As Boolean = True
        Try
            My.Computer.FileSystem.WriteAllText(FileName, Msg, Append)
        Catch ex As Exception
            Result = False
            CreateWarning(ErrNr, "Failed to write to: " & FileName, ErrDetails, Hint, ShowDialog, ex.Message, Fatal)
        End Try
        Return Result
    End Function

    Function DeleteFile(FileName As String,
                   ErrNr As Integer,
                   ErrDetails As String,
                   Hint As String,
                   ShowDialog As Boolean,
                   Fatal As Boolean) As Boolean
        Dim Result As Boolean = True
        xtrace_i("DeleteFile")
        Try
            My.Computer.FileSystem.DeleteFile(FileName)
        Catch ex As Exception
            Result = False
            CreateWarning(ErrNr, "Failed to delete: " & FileName, ErrDetails, Hint, ShowDialog, ex.Message, Fatal)
        End Try
        Return Result
    End Function

    Function CreateDirectory(DirPath As String,
                        ErrNr As Integer,
                        ErrDetails As String,
                        Hint As String,
                        ShowDialog As Boolean,
                        Fatal As Boolean) As Boolean
        xtrace_i("CreateDirectory")
        Dim Result As Boolean = True
        Try
            System.IO.Directory.CreateDirectory(DirPath)
        Catch ex As Exception
            Result = False
            CreateWarning(ErrNr, "Failed to create directory: " & DirPath, ErrDetails, Hint, ShowDialog, ex.Message, Fatal)
        End Try
        Return Result
    End Function

    Function CopyFile(FileName1 As String,
                      FileName2 As String, ErrNr As Integer,
                      ErrDetails As String,
                      Hint As String,
                      ShowDialog As Boolean,
                      Fatal As Boolean) As Boolean
        xtrace_i("CopyFile")
        Dim Result As Boolean = True
        Try
            FileSystem.FileCopy(FileName1, FileName2)
        Catch ex As Exception
            Result = False
            CreateWarning(ErrNr, "Failed to copy " & FileName1 & " to " & FileName2, ErrDetails, Hint, ShowDialog, ex.Message, Fatal)
        End Try
        Return Result
    End Function

    '==== Create Warning message ==================================================================

    Dim FileSystemMsgRecursion As Integer = 0
    Sub CreateWarning(ErrNr As Integer,
                      ErrMsg As String,
                      ErrDetails As String,
                      Hint As String,
                      ShowDialog As Boolean,
                      OSErr As String,
                      Fatal As Boolean)
        xtrace_i("CreateWarning, Fatal=" & Fatal.ToString)
        If Fatal Then
            CreateErrMsg(ErrNr, ErrMsg, ErrDetails, Hint, ShowDialog, OSErr)
        Else
            CreateWarning(ErrNr, ErrMsg, ErrDetails, Hint, ShowDialog, OSErr)
        End If
    End Sub
    Sub CreateWarning(ErrNr As Integer,
                      ErrMsg As String,
                      ErrDetails As String,
                      Hint As String,
                      ShowDialog As Boolean,
                      OSErr As String)
        xtrace_i("CreateWarning")
        FileSystemMsgRecursion += 1

        ' xtrace_warn may trigger another message. Don't process recursive errors
        If FileSystemMsgRecursion = 1 Then
            xtrace_warn({ErrNr.ToString,
                     ErrMsg,
                     "Details : " & ErrDetails,
                     "Hint    : " & Hint,
                     "OS Error: " & OSErr
                    })

            If ShowDialog Then

                MsgBox( ' "Warning : " & ErrNr.ToString & vbNewLine &
                       ErrMsg & vbNewLine &
                       "Details : " & ErrDetails & vbNewLine &
                       "Hint : " & Hint,
                       MsgBoxStyle.Critical,
                       "Warning")

                ' MessageBox.Show is part of Form1 and can only be used in GUI applications
            End If

            FileSystemMsgRecursion = 0
        End If
    End Sub

    Sub CreateErrMsg(ErrNr As Integer,
                      ErrMsg As String,
                      ErrDetails As String,
                      Hint As String,
                      ShowDialog As Boolean,
                      OSErr As String)
        xtrace_i("CreateErrMsg")
        FileSystemMsgRecursion += 1

        ' xtrace_warn may trigger another message. Don't process recursive errors
        If FileSystemMsgRecursion = 1 Then
            xtrace_err({ErrNr.ToString,
                     ErrMsg,
                     "Details : " & ErrDetails,
                     "Hint    : " & Hint,
                     "OS Error: " & OSErr
                    })

            If ShowDialog Then

                MsgBox( ' "Warning : " & ErrNr.ToString & vbNewLine &
                       ErrMsg & vbNewLine &
                       "Details : " & ErrDetails & vbNewLine &
                       "Hint : " & Hint,
                       MsgBoxStyle.Critical,
                       "Error")

                ' MessageBox.Show is part of Form1 and can only be used in GUI applications
            End If

            FileSystemMsgRecursion = 0
        End If

        exit_program()
    End Sub
End Module
