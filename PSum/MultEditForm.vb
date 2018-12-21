Public Class MultEditForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        SaveMult()

        Close()
    End Sub

    Private Sub MultEditForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Width = 100
        Height = 25
        Me.Location = Form1.CommonStuff.MultFormLocation

        Select Case Form1.CommonStuff.ActiveMultBtn
            Case 0
                TextBox1.Text = My.Settings.Mult0.ToString("N2")
            Case 1
                TextBox1.Text = My.Settings.Mult1.ToString("N2")
            Case 2
                TextBox1.Text = My.Settings.Mult2.ToString("N2")
            Case 3
                TextBox1.Text = My.Settings.Mult3.ToString("N2")
            Case 4
                TextBox1.Text = My.Settings.Mult4.ToString("N2")
            Case 5
                TextBox1.Text = My.Settings.Mult5.ToString("N2")
        End Select

        Me.Show()
        TextBox1.Select()
        TextBox1.SelectAll()


    End Sub

    Sub SaveMult()
        Select Case Form1.CommonStuff.ActiveMultBtn
            Case 0
                My.Settings.Mult0 = CDbl(TextBox1.Text)
            Case 1
                My.Settings.Mult1 = CDbl(TextBox1.Text)
            Case 2
                My.Settings.Mult2 = CDbl(TextBox1.Text)
            Case 3
                My.Settings.Mult3 = CDbl(TextBox1.Text)
            Case 4
                My.Settings.Mult4 = CDbl(TextBox1.Text)
            Case 5
                My.Settings.Mult5 = CDbl(TextBox1.Text)
        End Select
    End Sub

    Private Sub MultEditForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            SaveMult()
            Close()
        End If

        If e.KeyCode = Keys.Escape Then
            Close()
        End If
    End Sub
End Class