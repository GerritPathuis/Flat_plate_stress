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
        Elas = NumericUpDown5.Value


        σm = 0.75 * p * b ^ 2
        σm /= t ^ 2 * (1.61 * (b / a) ^ 3 + 1)


        yt = 0.142 * p * b ^ 4
        yt /= Elas * t ^ 3 * (2.21 * (b / a) ^ 3 + 1)

        TextBox2.Text = σm.ToString("0")
        TextBox3.Text = yt.ToString("0.0")
    End Sub
End Class
