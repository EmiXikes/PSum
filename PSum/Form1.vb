Imports NCalc
Imports DEL_acadltlib_EM.CAD_attextract
Imports DEL_acadltlib_EM.CAD_commands
Imports DEL_acadltlib_EM.DDE

Public Class Form1

    Public Class CommonStuff
        Public Shared MultFormLocation As Point

    End Class



    Function SelectedPower() As Double
        Dim er As String = Chr(10)
        Dim esc As String = Chr(27)
        Dim com As String
        Dim PowerList As List(Of DXFExtItemRef)
        Dim PowerValueS As String
        Dim PowerValueD As Double
        Dim exp As Expression

        initDDE()

        com = "m" & er & "0" & er & "0" & er
        SendCommand(com)

        PowerList = AttributeExtractAll(New List(Of String) From {"JAUDA"}, True)

        com = er & er & esc & esc & esc
        SendCommand(com)

        For Each PowerItem In PowerList

            PowerValueS = PowerItem.GetAttValue("JAUDA")
            PowerValueS = Strings.UCase(PowerValueS)
            PowerValueS = Strings.Replace(PowerValueS, ".", ".")
            PowerValueS = Strings.Replace(PowerValueS, ",", ".")
            PowerValueS = Strings.Replace(PowerValueS, " ", "")
            PowerValueS = Strings.Replace(PowerValueS, "KW", "")
            PowerValueS = Strings.Replace(PowerValueS, "X", "*")



            If PowerValueS <> "" Then

                exp = New Expression(PowerValueS)
                PowerValueS = exp.Evaluate()

                PowerValueS = Strings.Replace(PowerValueS, ".", Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                PowerValueS = Strings.Replace(PowerValueS, ",", Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)
                PowerValueD = PowerValueD + CDbl(PowerValueS)

            End If



        Next

        Return PowerValueD

    End Function


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox(SelectedPower())
    End Sub

    Private Sub Button2_MouseDown(sender As Object, e As MouseEventArgs) Handles Button2.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim mEdit As New MultEditForm
            CommonStuff.MultFormLocation = New Point
            CommonStuff.MultFormLocation.X = Location.X + 10
            CommonStuff.MultFormLocation.Y = Location.Y + 90
            mEdit.ShowDialog()
        End If
    End Sub

    Private Sub Button1_MouseDown(sender As Object, e As MouseEventArgs) Handles Button1.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim mEdit As New MultEditForm
            CommonStuff.MultFormLocation = New Point
            CommonStuff.MultFormLocation.X = Location.X + 10
            CommonStuff.MultFormLocation.Y = Location.Y + 116
            mEdit.ShowDialog()
        End If
    End Sub

    Private Sub Button3_MouseDown(sender As Object, e As MouseEventArgs) Handles Button3.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim mEdit As New MultEditForm
            CommonStuff.MultFormLocation = New Point
            CommonStuff.MultFormLocation.X = Location.X + 10
            CommonStuff.MultFormLocation.Y = Location.Y + 142
            mEdit.ShowDialog()
        End If
    End Sub
End Class
