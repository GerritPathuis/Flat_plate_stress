﻿Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged
        'http://www.roymech.co.uk/Useful_Tables/Mechanics/Plates.html
        'Rectangle simply supported
        Dim a, b, t As Double
        Dim Elas, p, σm, yt As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        TextBox4.Text = p.ToString
        a = NumericUpDown2.Value / 1000 '[m]
        b = NumericUpDown3.Value / 1000 '[m]
        t = NumericUpDown4.Value / 1000 '[m]
        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]

        If a >= b Then
            Label40.Visible = False
            σm = 0.75 * p * b ^ 2
            σm /= t ^ 2 * (1.61 * (b / a) ^ 3 + 1)
            σm /= 10 ^ 6        '[N/mm2]

            yt = 0.142 * p * b ^ 4
            yt /= Elas * t ^ 3 * (2.21 * (b / a) ^ 3 + 1)
            yt *= 1000          '[mm]
        Else
            σm = 0
            yt = 0
            Label40.Visible = True
        End If


        TextBox2.Text = σm.ToString("0")
        TextBox3.Text = yt.ToString("0.0")

        '===== checks ================
        TextBox2.BackColor = IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage3.Enter, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged
        'Rectangle clamped edges
        Dim a, b, t As Double
        Dim Elas, p, σm, yt As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        TextBox4.Text = p.ToString
        a = NumericUpDown8.Value / 1000 '[m]
        b = NumericUpDown7.Value / 1000 '[m]
        t = NumericUpDown6.Value / 1000 '[m]
        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]

        If a >= b Then
            Label39.Visible = False
            σm = p * b ^ 2
            σm /= t ^ 2 * (2.63 * (b / a) ^ 6 + 1)
            σm /= 10 ^ 6        '[N/mm2]

            yt = 0.0284 * p * b ^ 4
            yt /= Elas * t ^ 3 * (1.056 * (b / a) ^ 5 + 1)
            yt *= 1000          '[mm]
        Else
            σm = 0
            yt = 0
            Label39.Visible = True
        End If


        TextBox6.Text = σm.ToString("0")
        TextBox5.Text = yt.ToString("0.0")
        '===== check ================
        TextBox6.BackColor = IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, TabPage4.Enter, NumericUpDown9.ValueChanged, NumericUpDown11.ValueChanged
        'Round plate simply supported
        Dim dia, r, t As Double
        Dim Elas, p, σm, yt, v As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        TextBox4.Text = p.ToString
        dia = NumericUpDown11.Value / 1000 '[m]
        r = dia / 2 '[m]
        t = NumericUpDown9.Value / 1000 '[m]
        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]

        v = 0.3 'For steel
        σm = 1.238 * p * r ^ 2 / t ^ 2
        σm /= 10 ^ 6        '[N/mm2]

        yt = 0.696 * p * r ^ 4
        yt /= Elas * t ^ 3
        yt *= 1000          '[mm]

        TextBox8.Text = σm.ToString("0")
        TextBox7.Text = yt.ToString("0.0")
        '===== check ================
        TextBox8.BackColor = IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen)
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, TabPage6.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged
        'Round with hole
        Dim a, b, t As Double
        Dim Elas, p, σm, ym As Double
        Dim x, k1, k2 As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        a = NumericUpDown17.Value / 1000 '[m]
        b = NumericUpDown16.Value / 1000 '[m]
        t = NumericUpDown15.Value / 1000 '[m]
        TextBox20.Text = a * 2000.ToString("0")
        TextBox21.Text = b * 2000.ToString("0")

        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]

        '============= determine k1, k2 =================
        x = a / b
        TextBox22.Text = x.ToString("0.0")

        k1 = 0.0067 * x ^ 4 - 0.0584 * x ^ 3 + 0.0519 * x ^ 2 + 0.6132 * x - 0.4358
        k2 = 0.0127 * x ^ 4 - 0.131 * x ^ 3 + 0.3117 * x ^ 2 + 0.6069 * x - 0.219

        If x > 5 Then k1 = 0.815
        If x > 5 Then k2 = 2.2

        TextBox17.Text = k1.ToString("0.000")
        TextBox16.Text = k2.ToString("0.000")

        'MessageBox.Show("a= " & a.ToString & " b= " & b.ToString & " t=" & t.ToString & " e=" & ee.ToString)
        σm = k2 * p * a ^ 2 / t ^ 2
        σm /= 10 ^ 6                '[N/mm2]
        TextBox12.Text = σm.ToString("0")

        ym = k1 * p * a ^ 4
        ym /= Elas * t ^ 3
        ym *= 1000          '[mm]
        TextBox11.Text = ym.ToString("0.0")
        '===== check ================
        TextBox12.BackColor = IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen)
    End Sub

End Class
