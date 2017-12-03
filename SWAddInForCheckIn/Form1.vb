Public Class Form1
    Dim MaxValue As Integer
    Dim percentage As Integer
    Dim currentValue As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Public Sub DrawingCount(num As Integer)
        ProgressBar1.Maximum = num
    End Sub

    Public Sub IncreaseValue()
        ProgressBar1.Increment(1)
    End Sub

End Class