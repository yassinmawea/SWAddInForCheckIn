Public Class Form1
    Dim MaxValue As Integer = 10
    Dim percentage As Integer
    Dim currentValue As Integer = 0


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Public Sub Open()
        Show()
        Refresh()

        Do
            System.Threading.Thread.Sleep(100)
            currentValue = ProgressBar1.Value
            currentValue = currentValue + 1

            ProgressBar1.Value = currentValue

            If currentValue = 100 Then
                currentValue = 1
            End If

        Loop Until currentValue = 111



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