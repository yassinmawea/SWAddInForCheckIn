﻿Public Class Form1
    Dim MaxValue As Integer = 10
    Dim percentage As Integer
    Dim currentValue As Integer = 0


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        TopMost = True

        BackgroundWorker1.WorkerSupportsCancellation = True
        BackgroundWorker1.WorkerReportsProgress = True

        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Public Sub Open()
        Debug.Print("In Open ")
        Try


            'Show()
            Refresh()

            Do
                System.Threading.Thread.Sleep(1000)
                'currentValue = ProgressBar1.Value
                currentValue = currentValue + 2

                ProgressBar1.Value = currentValue


                Debug.Print("currentValue -> " & currentValue)


                If currentValue = 99 Then
                    currentValue = 1
                End If

            Loop Until currentValue = 111

        Catch ex As Exception
            Close()

        End Try

        'If Not BackgroundWorker1.IsBusy = True Then
        'BackgroundWorker1.RunWorkerAsync()
        'End If
    End Sub

    Private Sub Form1_Close(sender As Object, e As EventArgs) Handles MyBase.FormClosed
        Debug.Print("In Close")
        BackgroundWorker1.CancelAsync()
        'BackgroundWorker1.ReportProgress(100)
    End Sub

    Public Sub DrawingCount(num As Integer)
        MaxValue = num
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Debug.Print("In Open ")
        Try

            Show()
            Refresh()

            Do
                System.Threading.Thread.Sleep(1000)
                'currentValue = ProgressBar1.Value
                currentValue = currentValue + 2

                BackgroundWorker1.ReportProgress(currentValue)

                Refresh()

                Debug.Print("currentValue -> " & currentValue)

                If currentValue = 96 Then
                    currentValue = 0
                End If

                Debug.Print("Visible -> " & Visible)

            Loop Until Not Visible



        Catch ex As Exception
            Debug.Print("In Exception")
            Close()
            BackgroundWorker1.CancelAsync()
            e.Cancel = True
            BackgroundWorker1.ReportProgress(100)
            Debug.Print("Exception end")
        End Try


    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub
End Class