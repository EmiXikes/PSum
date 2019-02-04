Imports NCalc
Imports DEL_acadltlib_EM.CAD_attextract
Imports DEL_acadltlib_EM.CAD_commands
Imports DEL_acadltlib_EM.DDE
Imports DEL_acadltlib_EM.DXF
Imports DEL_acadltlib_EM.AutoCADLTinfo

Public Class Form1
    Public Class CommonStuff
        Public Shared MultFormLocation As Point
        Public Shared ActiveMultBtn As Integer
    End Class

    Dim Pn As Double
    Dim Pa As Double

    Public Sub OverridePnResult(ByVal NewRes)
        Lbl_Pn.Text = "Pn=" & NewRes & "kW"
    End Sub

    Public Sub OverridePaResult(ByVal NewRes)
        Lbl_Pn.Text = "Pa=" & NewRes & "kW"
    End Sub

    Dim PnTxt As String
    Dim PaTxt As String

    Function SetColorForExtract()
        Dim er As String = Chr(10)
        Dim esc As String = Chr(27)
        Dim com As String
        Dim CadColor As String

        Select Case CommonStuff.ActiveMultBtn
            Case 0
                CadColor = My.Settings.Col0
            Case 1
                CadColor = My.Settings.Col1
            Case 2
                CadColor = My.Settings.Col2
            Case 3
                CadColor = My.Settings.Col3
            Case 4
                CadColor = My.Settings.Col4
            Case 5
                CadColor = My.Settings.Col5
        End Select

        initDDE()

        com = "change" & er & "p" & er & er & "p" & er & "c" & er & "t" & er & CadColor & er & er
        SendCommand(com)

        CloseDDE()

    End Function


    Function SelectedPower() As Double

        Dim er As String = Chr(10)
        Dim esc As String = Chr(27)
        Dim com As String
        Dim PowerList As List(Of DXFExtItemRef)
        Dim PowerValueS As String
        Dim PowerValueD As Double
        Dim exp As Expression

        initDDE()

        Try
            PowerList = ReadBlockAttributesFromDXFExtractFile(Export2DXF_Selected())
        Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
        End Try

        Dim CadColor As String

        Select Case CommonStuff.ActiveMultBtn
            Case 0
                CadColor = My.Settings.Col0
            Case 1
                CadColor = My.Settings.Col1
            Case 2
                CadColor = My.Settings.Col2
            Case 3
                CadColor = My.Settings.Col3
            Case 4
                CadColor = My.Settings.Col4
            Case 5
                CadColor = My.Settings.Col5
        End Select

        com = "change" & er & "p" & er & er & "p" & er & "c" & er & "t" & er & CadColor & er & er
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

        CloseDDE()

        Return PowerValueD

    End Function
    Dim ButtonH As Integer = 27

    Function ColorFromRgbString(stringColorFormat) As Color
        With stringColorFormat
            Return Color.FromArgb(CInt(.Split(",")(0)), CInt(.Split(",")(1)), CInt(.Split(",")(2)))
        End With
    End Function

    Private Sub Button2_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn_Mult0.MouseDown
        If e.Button = MouseButtons.Right Then
            CommonStuff.ActiveMultBtn = 0
            ShowMultEditor()
        End If
    End Sub

    Private Sub Button1_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn_Mult1.MouseDown
        If e.Button = MouseButtons.Right Then
            CommonStuff.ActiveMultBtn = 1
            ShowMultEditor()
        End If
    End Sub

    Private Sub Button3_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn_Mult2.MouseDown
        If e.Button = MouseButtons.Right Then
            CommonStuff.ActiveMultBtn = 2
            ShowMultEditor()
        End If
    End Sub

    Private Sub Button4_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn_Mult3.MouseDown
        If e.Button = MouseButtons.Right Then
            CommonStuff.ActiveMultBtn = 3
            ShowMultEditor()
        End If
    End Sub

    Private Sub Button5_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn_Mult4.MouseDown
        If e.Button = MouseButtons.Right Then
            CommonStuff.ActiveMultBtn = 4
            ShowMultEditor()
        End If
    End Sub

    Private Sub Button6_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn_Mult5.MouseDown
        If e.Button = MouseButtons.Right Then
            CommonStuff.ActiveMultBtn = 5
            ShowMultEditor()
        End If
    End Sub

    Sub ShowMultEditor()
        Dim mEdit As New MultEditForm
        CommonStuff.MultFormLocation = New Point
        CommonStuff.MultFormLocation.X = Location.X + 10
        CommonStuff.MultFormLocation.Y = Location.Y + 102 + (Me.Height - Me.ClientSize.Height) + (ButtonH * CommonStuff.ActiveMultBtn)
        mEdit.ShowDialog()
        ReloadUI()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Pn = 0
        Pa = 0

        ReloadUI()
        ReloadResult()
        initDDE()
        EnableLog()
        CloseDDE()

    End Sub

    Sub ReloadResult()

        ''TODO Set Autocad as foremost again

        Lbl_Pn.Text = "Pn = " & Pn.ToString("N2")
        Lbl_Pa.Text = "Pa = " & Pa.ToString("N2")

        TxtBxPn.Text = "Pn=" & PnTxt & "=" & Pn.ToString("N2") & "kW"
        TxtBxPa.Text = "Pa=" & PaTxt & "=" & Pa.ToString("N2") & "kW"

    End Sub

    Sub ReloadUI()
        Btn_Mult0.Text = Chr(215) & My.Settings.Mult0.ToString("N2")
        Btn_Mult1.Text = Chr(215) & My.Settings.Mult1.ToString("N2")
        Btn_Mult2.Text = Chr(215) & My.Settings.Mult2.ToString("N2")
        Btn_Mult3.Text = Chr(215) & My.Settings.Mult3.ToString("N2")
        Btn_Mult4.Text = Chr(215) & My.Settings.Mult4.ToString("N2")
        Btn_Mult5.Text = Chr(215) & My.Settings.Mult5.ToString("N2")

        Btn_Mult0.ForeColor = ColorFromRgbString(My.Settings.Col0)
        Btn_Mult1.ForeColor = ColorFromRgbString(My.Settings.Col1)
        Btn_Mult2.ForeColor = ColorFromRgbString(My.Settings.Col2)
        Btn_Mult3.ForeColor = ColorFromRgbString(My.Settings.Col3)
        Btn_Mult4.ForeColor = ColorFromRgbString(My.Settings.Col4)
        Btn_Mult5.ForeColor = ColorFromRgbString(My.Settings.Col5)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Pn = 0
        Pa = 0
        ReloadResult()
    End Sub

    Dim pls As String = "+"


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Btn_Mult0.Click
        CommonStuff.ActiveMultBtn = 0
        Dim Pwr As Double = SelectedPower()
        If Pwr = 0 Then Exit Sub
        Dim Mult As Double = My.Settings.Mult0
        Pn = Pn + Pwr
        Pa = Pa + Pwr * Mult

        If PnTxt = "" Then pls = "" Else pls = "+"
        PnTxt = PnTxt & pls & Pwr
        PaTxt = PaTxt & pls & "[" & Pwr & "]x" & Mult

        ReloadResult()

    End Sub
    Private Sub Btn_Mult1_Click(sender As Object, e As EventArgs) Handles Btn_Mult1.Click
        CommonStuff.ActiveMultBtn = 1
        Dim Pwr As Double = SelectedPower()
        If Pwr = 0 Then Exit Sub
        Dim Mult As Double = My.Settings.Mult1
        Pn = Pn + Pwr
        Pa = Pa + Pwr * My.Settings.Mult1

        If PnTxt = "" Then pls = "" Else pls = "+"
        PnTxt = PnTxt & pls & Pwr
        PaTxt = PaTxt & pls & "[" & Pwr & "]x" & Mult

        ReloadResult()
    End Sub

    Private Sub Btn_Mult2_Click(sender As Object, e As EventArgs) Handles Btn_Mult2.Click
        CommonStuff.ActiveMultBtn = 2
        Dim Pwr As Double = SelectedPower()
        If Pwr = 0 Then Exit Sub
        Dim Mult As Double = My.Settings.Mult2
        Pn = Pn + Pwr
        Pa = Pa + Pwr * My.Settings.Mult2

        If PnTxt = "" Then pls = "" Else pls = "+"
        PnTxt = PnTxt & pls & Pwr
        PaTxt = PaTxt & pls & "[" & Pwr & "]x" & Mult

        ReloadResult()
    End Sub

    Private Sub Btn_Mult3_Click(sender As Object, e As EventArgs) Handles Btn_Mult3.Click
        CommonStuff.ActiveMultBtn = 3
        Dim Pwr As Double = SelectedPower()
        If Pwr = 0 Then Exit Sub
        Dim Mult As Double = My.Settings.Mult3
        Pn = Pn + Pwr
        Pa = Pa + Pwr * My.Settings.Mult3

        If PnTxt = "" Then pls = "" Else pls = "+"
        PnTxt = PnTxt & pls & Pwr
        PaTxt = PaTxt & pls & "[" & Pwr & "]x" & Mult

        ReloadResult()
    End Sub

    Private Sub Btn_Mult4_Click(sender As Object, e As EventArgs) Handles Btn_Mult4.Click
        CommonStuff.ActiveMultBtn = 4
        Dim Pwr As Double = SelectedPower()
        If Pwr = 0 Then Exit Sub
        Dim Mult As Double = My.Settings.Mult4
        Pn = Pn + Pwr
        Pa = Pa + Pwr * My.Settings.Mult4

        If PnTxt = "" Then pls = "" Else pls = "+"
        PnTxt = PnTxt & pls & Pwr
        PaTxt = PaTxt & pls & "[" & Pwr & "]x" & Mult

        ReloadResult()
    End Sub

    Private Sub Btn_Mult5_Click(sender As Object, e As EventArgs) Handles Btn_Mult5.Click
        CommonStuff.ActiveMultBtn = 5
        Dim Pwr As Double = SelectedPower()
        If Pwr = 0 Then Exit Sub
        Dim Mult As Double = My.Settings.Mult5
        Pn = Pn + Pwr
        Pa = Pa + Pwr * My.Settings.Mult5

        If PnTxt = "" Then pls = "" Else pls = "+"
        PnTxt = PnTxt & pls & Pwr
        PaTxt = PaTxt & pls & "[" & Pwr & "]x" & Mult

        ReloadResult()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        ColorDialog1.ShowDialog()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        CloseDDE()
    End Sub

    Private Sub TxtBxPn_MouseDown(sender As Object, e As MouseEventArgs) Handles TxtBxPn.MouseDown

        'Cop1.Visible = True
        'TxtBxPn.SelectAll()
        'TxtBxPn.Copy()
        '' TxtBxPn.DeselectAll()

    End Sub

    Private Sub TxtBxPa_MouseDown(sender As Object, e As MouseEventArgs) Handles TxtBxPa.MouseDown

    End Sub

    Private Sub TxtBxPn_TextChanged(sender As Object, e As EventArgs) Handles TxtBxPn.TextChanged



    End Sub

    Private Sub TxtBxPn_MouseLeave(sender As Object, e As EventArgs) Handles TxtBxPn.MouseLeave
        ''Cop1.Visible = False
    End Sub

    Private Sub TxtBxPa_MouseLeave(sender As Object, e As EventArgs) Handles TxtBxPa.MouseLeave
        ''Cop2.Visible = False
    End Sub

    Private Sub TxtBxPn_MouseClick(sender As Object, e As MouseEventArgs) Handles TxtBxPn.MouseClick
        'Dim advSum As New AdvancedSumResultWindow
        'advSum.Left = Me.Left - 200
        'advSum.Top = Me.Top + 70
        'advSum.TextBox1.Text = TxtBxPa.Text
        'advSum.Prefix = "Pn="
        'advSum.Show()

        'Cop1.Visible = True
        'TxtBxPn.SelectAll()
        'TxtBxPn.Copy()
        'TxtBxPn.DeselectAll()
    End Sub

    Private Sub TxtBxPa_MouseClick(sender As Object, e As MouseEventArgs) Handles TxtBxPa.MouseClick

        Dim advSum As New AdvancedSumResultWindow
        advSum.Left = Me.Left - 200
        advSum.Top = Me.Top + 90
        advSum.TextBox1.Text = TxtBxPa.Text
        advSum.Prefix = "Pa="
        advSum.Show()

        'Cop2.Visible = True
        'TxtBxPa.SelectAll()
        'TxtBxPa.Copy()
        'TxtBxPa.DeselectAll()
    End Sub

    Private Sub Lbl_Pa_Click(sender As Object, e As EventArgs) Handles Lbl_Pa.Click

    End Sub
End Class
