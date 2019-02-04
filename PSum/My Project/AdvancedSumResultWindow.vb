Imports NCalc

Public Class AdvancedSumResultWindow
    Private Sub AdvancedSumResultWindow_MouseLeave(sender As Object, e As EventArgs) Handles MyBase.MouseLeave
        '' Me.Close()
    End Sub

    Private Sub AdvancedSumResultWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 25
    End Sub

    Private Sub TextBox1_MouseLeave(sender As Object, e As EventArgs) Handles TextBox1.MouseLeave
        '' Me.Close()
    End Sub

    Private Sub AdvancedSumResultWindow_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        ProcessData()
        Me.Close()
    End Sub

    Dim AdvSum As String
    Dim OrigStr As String
    Public Prefix As String

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Escape Then
            ProcessData()
            Me.Close()
        End If

        If e.KeyCode = Keys.Enter Then
            ProcessData()
        End If


    End Sub

    Private Sub ProcessData()
        Dim tempRes As Double
        OrigStr = TextBox1.Text
        AdvSum = Strings.UCase(TextBox1.Text)
        AdvSum = AdvSum.Replace(",", ".")
        AdvSum = AdvSum.Replace("X", "*")
        AdvSum = AdvSum.Replace("[", "(")
        AdvSum = AdvSum.Replace("]", ")")

        Dim exp As Expression
        Dim CorrectPart As Integer
        Dim AdvSumString As String

        For i = 0 To AdvSum.Split("=").Length
            Try
                exp = New Expression(AdvSum.Split("=")(i))
                tempRes = exp.Evaluate()
                CorrectPart = i
                Exit For
            Catch ex As Exception
                Continue For
            End Try
        Next

        AdvSumString = Prefix & OrigStr.Split("=")(CorrectPart) & "=" & tempRes & "kW"

        If Prefix = "Pn=" Then
            Form1.OverridePnResult(tempRes.ToString("N2"))
        Else
            Form1.OverridePaResult(tempRes.ToString("N2"), AdvSumString)
        End If

        TextBox1.Text = AdvSumString
    End Sub
End Class