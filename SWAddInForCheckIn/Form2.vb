Imports System.Windows.Forms

Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        Application.Exit()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()

    End Sub

    Private Sub Form2_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        TopMost = True
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Public Sub Open()
        Debug.Print("In Open Method")
        Try
            Debug.Print("This 1")
            Show()
            Debug.Print("This 2")
            Refresh()
            Debug.Print("This 3")

            'Do
            'If Application.OpenForms.OfType(Of Form2).Any Then
            'Debug.Print("Form2 is Opened")



            'End If

            'Loop Until Not Application.OpenForms.OfType(Of Form2).Any

        Catch ex As Exception
            Debug.Print("Ex catched -> " + ex.ToString)
        End Try
    End Sub

End Class