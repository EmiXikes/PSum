Imports System.Globalization

Public Class MultEditForm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        SaveMult()
        Close()

    End Sub

    Private Sub MultEditForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Width = 120
        Height = 25
        Me.Location = Form1.CommonStuff.MultFormLocation

        Select Case Form1.CommonStuff.ActiveMultBtn
            Case 0
                TextBox1.Text = My.Settings.Mult0.ToString("N2")
                Button2.BackColor = Form1.ColorFromRgbString(My.Settings.Col0)
            Case 1
                TextBox1.Text = My.Settings.Mult1.ToString("N2")
                Button2.BackColor = Form1.ColorFromRgbString(My.Settings.Col1)
            Case 2
                TextBox1.Text = My.Settings.Mult2.ToString("N2")
                Button2.BackColor = Form1.ColorFromRgbString(My.Settings.Col2)
            Case 3
                TextBox1.Text = My.Settings.Mult3.ToString("N2")
                Button2.BackColor = Form1.ColorFromRgbString(My.Settings.Col3)
            Case 4
                TextBox1.Text = My.Settings.Mult4.ToString("N2")
                Button2.BackColor = Form1.ColorFromRgbString(My.Settings.Col4)
            Case 5
                TextBox1.Text = My.Settings.Mult5.ToString("N2")
                Button2.BackColor = Form1.ColorFromRgbString(My.Settings.Col5)
        End Select

        Me.Show()
        TextBox1.Select()
        TextBox1.SelectAll()


    End Sub

    Sub SaveMult()

        Dim userMultiplier As String = TextBox1.Text
        Dim dMult As Double

        userMultiplier = Strings.Replace(userMultiplier, ".", ",")

        Try

            dMult = CDbl(userMultiplier)

            Select Case Form1.CommonStuff.ActiveMultBtn
                Case 0
                    My.Settings.Mult0 = dMult
                Case 1
                    My.Settings.Mult1 = CDbl(dMult)
                Case 2
                    My.Settings.Mult2 = CDbl(dMult)
                Case 3
                    My.Settings.Mult3 = CDbl(dMult)
                Case 4
                    My.Settings.Mult4 = CDbl(dMult)
                Case 5
                    My.Settings.Mult5 = CDbl(dMult)
            End Select
        Catch ex As Exception
            MsgBox("Nepareizs formāts.")
            Exit Sub
        End Try

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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim cdial As New ColorDialog

        Select Case Form1.CommonStuff.ActiveMultBtn
            Case 0
                cdial.Color = Form1.ColorFromRgbString(My.Settings.Col0)
                sender.BackColor = Form1.ColorFromRgbString(My.Settings.Col0)
            Case 1
                cdial.Color = Form1.ColorFromRgbString(My.Settings.Col1)
                sender.BackColor = Form1.ColorFromRgbString(My.Settings.Col1)
            Case 2
                cdial.Color = Form1.ColorFromRgbString(My.Settings.Col2)
                sender.BackColor = Form1.ColorFromRgbString(My.Settings.Col2)
            Case 3
                cdial.Color = Form1.ColorFromRgbString(My.Settings.Col3)
                sender.BackColor = Form1.ColorFromRgbString(My.Settings.Col3)
            Case 4
                cdial.Color = Form1.ColorFromRgbString(My.Settings.Col4)
                sender.BackColor = Form1.ColorFromRgbString(My.Settings.Col4)
            Case 5
                cdial.Color = Form1.ColorFromRgbString(My.Settings.Col5)
                sender.BackColor = Form1.ColorFromRgbString(My.Settings.Col5)
        End Select

        cdial.ShowDialog()

        Select Case Form1.CommonStuff.ActiveMultBtn
            Case 0
                My.Settings.Col0 = cdial.Color.R.ToString & "," & cdial.Color.G.ToString & "," & cdial.Color.B.ToString
            Case 1
                My.Settings.Col1 = cdial.Color.R.ToString & "," & cdial.Color.G.ToString & "," & cdial.Color.B.ToString
            Case 2
                My.Settings.Col2 = cdial.Color.R.ToString & "," & cdial.Color.G.ToString & "," & cdial.Color.B.ToString
            Case 3
                My.Settings.Col3 = cdial.Color.R.ToString & "," & cdial.Color.G.ToString & "," & cdial.Color.B.ToString
            Case 4
                My.Settings.Col4 = cdial.Color.R.ToString & "," & cdial.Color.G.ToString & "," & cdial.Color.B.ToString
            Case 5
                My.Settings.Col5 = cdial.Color.R.ToString & "," & cdial.Color.G.ToString & "," & cdial.Color.B.ToString
        End Select


        '' MsgBox(cdial.Color.ToString)

        SaveMult()
        Close()
    End Sub
End Class