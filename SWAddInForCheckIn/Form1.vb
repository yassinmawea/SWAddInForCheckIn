Public Class Form1
    Dim MaxValue As Integer = 10
    Dim percentage As Integer
    Dim currentValue As Integer = 0


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        TopMost = True
        Debug.Print("In Open ")
        Try


            Show()
            Refresh()

            Do
                System.Threading.Thread.Sleep(1000)
                'currentValue = ProgressBar1.Value
                currentValue = currentValue + 2

                ProgressBar1.Value = currentValue
                Refresh()

                Debug.Print("currentValue -> " & currentValue)


                If currentValue = 99 Then
                    currentValue = 1
                End If

            Loop Until currentValue = 111

        Catch ex As Exception
            Close()

        End Try
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

    Public Sub CloseForm()
        Close()
    End Sub

    Public Sub DrawingCount(num As Integer)
        MaxValue = num
    End Sub

    Public Sub IncreaseValue()
        'currentValue = currentValue + 1
        'ProgressBar1.Increment(1)
    End Sub



End Class