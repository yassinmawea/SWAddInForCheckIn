Public Class Form1
    Dim MaxValue As Integer
    Dim percentage As Integer
    Dim currentValue As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Public Sub setMaximum(ByVal max)
        ProgressBar1.Maximum = max
    End Sub

    Public Sub increaseProgress()
        currentValue = (currentValue + 1)
        Debug.Print(currentValue)
        ProgressBar1.Value = currentValue
    End Sub

End Class