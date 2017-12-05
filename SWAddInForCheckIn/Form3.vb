Public Class Form3
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
    End Sub

    Public Sub Open()
        Debug.Print("In Open Method")
        Try
            Debug.Print("This 1")
            Me.Show()
            Debug.Print("This 2")
            Me.Refresh()
            Debug.Print("This 3")

        Catch ex As Exception
            Debug.Print("Ex catched -> " + ex.ToString)
        End Try
    End Sub

End Class