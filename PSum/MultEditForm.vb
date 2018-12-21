Public Class MultEditForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub MultEditForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Width = 100
        Height = 25
        Me.Location = Form1.CommonStuff.MultFormLocation
    End Sub
End Class