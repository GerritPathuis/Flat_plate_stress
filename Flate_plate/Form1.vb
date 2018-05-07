Public Class Form1
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


        σm = 0.75 * p * b ^ 2
        σm /= t ^ 2 * (1.61 * (b / a) ^ 3 + 1)
        σm /= 10 ^ 6        '[N/mm2]

        yt = 0.142 * p * b ^ 4
        yt /= Elas * t ^ 3 * (2.21 * (b / a) ^ 3 + 1)
        yt *= 1000          '[mm]

        TextBox2.Text = σm.ToString("0")
        TextBox3.Text = yt.ToString("0.0")
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


        σm = p * b ^ 2
        σm /= t ^ 2 * (2.63 * (b / a) ^ 6 + 1)
        σm /= 10 ^ 6        '[N/mm2]

        yt = 0.0284 * p * b ^ 4
        yt /= Elas * t ^ 3 * (1.056 * (b / a) ^ 5 + 1)
        yt *= 1000          '[mm]

        TextBox6.Text = σm.ToString("0")
        TextBox5.Text = yt.ToString("0.0")
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

        v = 0.3
        TextBox13.Text = v.ToString("0.0")

        σm = 1.238 * p * r ^ 2 / t ^ 2
        σm /= 10 ^ 6        '[N/mm2]

        yt = 0.696 * p * r ^ 4
        yt /= Elas * t ^ 3
        yt *= 1000          '[mm]

        TextBox8.Text = σm.ToString("0")
        TextBox7.Text = yt.ToString("0.0")

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage5.Enter, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown10.ValueChanged
        'Rectangle simply supported with hole
        Dim a, b, t, e1, ee As Double
        Dim Elas, p, σm, ym As Double
        Dim x, k1, k2 As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        TextBox4.Text = p.ToString
        a = NumericUpDown14.Value / 1000 '[m]
        b = NumericUpDown13.Value / 1000 '[m]
        t = NumericUpDown12.Value / 1000 '[m]
        e1 = NumericUpDown10.Value / 2000 '[m]
        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]


        If e1 < 0.5 * t Then
            ee = Math.Sqrt(1.6 * e1 ^ 2 + t ^ 2)
            ee -= 0.675 * t
        Else
            ee = e1
        End If

        TextBox18.Text = ee.ToString("0.000")

        '============= determine k1, k2 =================
        x = a / b
        TextBox19.Text = x.ToString("0.00")
        k1 = -0.00003 * x ^ 6 + 0.001 * x ^ 5 - 0.0122 * x ^ 4 + 0.0756 * x ^ 3 - 0.2549 * x ^ 2 + 0.4418 * x - 0.1238
        k2 = 0.0023 * x ^ 5 - 0.0511 * x ^ 4 + 0.4403 * x ^ 3 - 1.8174 * x ^ 2 + 3.5846 * x - 1.7042

        'If x > 4 Then k1 = 0.185
        'If x > 3 Then k2 = 1.0

        TextBox14.Text = k1.ToString("0.000")
        TextBox15.Text = k2.ToString("0.000")

        'MessageBox.Show("a= " & a.ToString & " b= " & b.ToString & " t=" & t.ToString & " e=" & ee.ToString)
        σm = 1.5 * p / (Math.PI * t ^ 2)
        σm *= (1.3 * Math.Log(2 * b / (Math.PI * ee)) + k2)

        σm /= 10 ^ 6        '[N/mm2]
        TextBox10.Text = σm.ToString("0")

        ym = k1 * p * b ^ 2
        ym /= Elas * t ^ 3
        ym *= 1000          '[mm]
        TextBox9.Text = ym.ToString("0.0")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, TabPage6.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged
        'Round with hole
        Dim a, b, t, e1, ee As Double
        Dim Elas, p, σm, ym As Double
        Dim x, k1, k2 As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        a = NumericUpDown17.Value / 1000 '[m]
        b = NumericUpDown16.Value / 1000 '[m]
        t = NumericUpDown15.Value / 1000 '[m]

        e1 = NumericUpDown10.Value / 2000 '[m]
        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]


        If e1 < 0.5 * t Then
            ee = Math.Sqrt(1.6 * e1 ^ 2 + t ^ 2)
            ee -= 0.675 * t
        Else
            ee = e1
        End If

        TextBox18.Text = ee.ToString("0.000")

        '============= determine k1, k2 =================
        x = a / b
        TextBox19.Text = x.ToString("0.00")
        k1 = 0.0067 * x ^ 4 - 0.0584 * x ^ 3 + 0.0519 * x ^ 2 + 0.6132 * x - 0.4358
        k2 = 0.0127 * x ^ 4 - 0.131 * x ^ 3 + 0.3117 * x ^ 2 + 0.6069 * x - 0.219

        'If x > 4 Then k1 = 0.185
        'If x > 3 Then k2 = 1.0

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
    End Sub
End Class
