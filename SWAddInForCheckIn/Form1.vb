Public Class Form1
    Dim MaxValue As Integer
    Dim percentage As Integer
    Dim currentValue As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Public Sub DrawingCount(num As Integer)
        ProgressBar1.Maximum = num
    End Sub

    Public Sub increaseValue()
        ProgressBar1.Value = ProgressBar1.Value + 1
    End Sub

End Class