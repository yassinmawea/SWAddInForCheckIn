Public Class Form1
    Dim MaxValue As Integer
    Dim percentage As Integer
    Dim currentValue As Integer = 0


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Public Sub Open()
        Show()
        Refresh()

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

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        System.Threading.Thread.Sleep(1000)
        'currentValue = currentValue + 1
        'percentage = (currentValue / MaxValue)
        'BackgroundWorker1.ReportProgress(percentage)
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        'ProgressBar1.Value = e.ProgressPercentage
    End Sub

End Class