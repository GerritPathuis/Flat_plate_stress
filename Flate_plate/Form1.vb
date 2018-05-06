Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged
        'http://www.roymech.co.uk/Useful_Tables/Mechanics/Plates.html
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
        'Rectangle clamped
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Rectangle simply supported with hole
        Dim a, b, t, e1, ee As Double
        Dim Elas, p, σm, yt As Double
        Dim k1, k2 As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        TextBox4.Text = p.ToString
        a = NumericUpDown14.Value / 1000 '[m]
        b = NumericUpDown13.Value / 1000 '[m]
        t = NumericUpDown12.Value / 1000 '[m]
        e1 = NumericUpDown10.Value / 2000 '[m]
        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]

        ee = Math.Sqrt(1.6 * e1 ^ 2 + t ^ 2)
        ee -= 0.675 * t

        '============= determine k1, k2 =================
        '====================NIET AF !!!!!!!!!!!!!!!!!!!!!!!!!!!
        k1 = 1.0
        k2 = 0.999999999

        σm = 1.5 * p / (Math.PI * t ^ 2)
        σm *= (1 + 0.3) * (Math.Log(2 * b / Math.PI * ee) + k2)
        σm /= 10 ^ 6        '[N/mm2]

        yt = 0.0284 * p * b ^ 4
        yt /= Elas * t ^ 3 * (1.056 * (b / a) ^ 5 + 1)
        yt *= 1000          '[mm]

        TextBox10.Text = σm.ToString("0")
        TextBox9.Text = yt.ToString("0.0")
    End Sub
End Class
