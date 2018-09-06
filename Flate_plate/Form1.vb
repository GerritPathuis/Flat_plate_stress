Imports System.Math
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Threading
Imports System.Windows.Forms

Public Class Form1
    '"Name; I (strong axis)[cm4]; Profile height[mm]"
    Public Shared UNP() As String = {
     "UNP 40; 14.1; 40",
     "UNP 50; 26.4; 50",
     "UNP 65; 57.5; 65",
     "UNP 80; 106; 80",
     "UNP 100; 206; 100",
     "UNP 120; 364; 120",
     "UNP 140; 605; 140",
     "UNP 160; 925; 160",
     "UNP 180; 1350; 180",
     "UNP 200; 1910; 200",
     "UNP 220; 2690; 220",
     "UNP 240; 3600; 240",
     "UNP 260; 4820; 260",
     "UNP 280; 6280; 280",
     "UNP 300; 8030; 300",
     "UNP 320; 10870; 320",
     "UNP 350; 12840; 350",
     "UNP 380; 15760; 380",
     "UNP 400; 20350; 400",
     "HEB 100; 449; 100",
     "HEB 120; 864; 120",
     "HEB 140; 1509; 140",
     "HEB 160; 2492; 160",
     "HEB 180; 3831; 180",
     "HEB 200; 5696; 200",
     "HEB 220; 8091; 220",
     "HEB 240; 11260; 240",
     "HEB 260; 14920; 260",
     "HEB 280; 19270; 280",
     "HEB 300; 25170; 300"
    }
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        For hh = 0 To UNP.Length - 1             'Fill combobox1
            words = UNP(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 7
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged
        Calc()
    End Sub
    Private Sub Calc()
        'http://www.roymech.co.uk/Useful_Tables/Mechanics/Plates.html
        'Rectangle simply supported
        Dim a, b, t, flex As Double
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
            flex = a * 10 ^ 3 / yt
        Else
            σm = 0
            yt = 0
            Label40.Visible = True
            flex = 0
        End If


        TextBox2.Text = σm.ToString("0")
        TextBox3.Text = yt.ToString("0.0")
        TextBox25.Text = flex.ToString("0")

        '===== checks ================
        TextBox2.BackColor = IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage3.Enter, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged
        'Rectangle clamped edges
        Dim a, b, t As Double
        Dim Elas, p, σm, yt, flex As Double

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
            flex = a * 10 ^ 3 / yt
        Else
            σm = 0
            yt = 0
            Label39.Visible = True
            flex = 0
        End If



        TextBox6.Text = σm.ToString("0")
        TextBox5.Text = yt.ToString("0.0")
        TextBox24.Text = flex.ToString("0")
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage5.Enter, NumericUpDown18.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown19.ValueChanged
        'http://beamguru.com/online/beam-calculator/
        'https://www.amesweb.info/StructuralBeamDeflection/SimplySupportedBeamStressDeflectionAnalysis.aspx
        Dim l, Iy, w, Elas As Double
        Dim flex, mom_max, stress_b As Double
        Dim p, load_width, area As Double
        Dim beam_height As Double

        l = NumericUpDown18.Value * 10 ^ 3      '[m->mm]
        Iy = NumericUpDown14.Value * 10 ^ 4     '[mm4]
        Elas = NumericUpDown5.Value * 10 ^ 3    '[N/mm2]

        'MessageBox.Show(Elas.ToString)
        '===== Distributed load ============
        p = NumericUpDown1.Value * 10 ^ 2 * 10 ^ 6      '[mbar->N/mm2]
        load_width = NumericUpDown19.Value * 10 ^ 3     '[mm]
        area = load_width                               '[mm2]
        w = p * area                                    '[N/mm]

        '===== Flex, moment and stress ============
        mom_max = (w * l ^ 2) / 8
        flex = 5 * w * l ^ 4 / (384 * Elas * Iy)
        Double.TryParse(TextBox18.Text, beam_height)
        stress_b = mom_max * beam_height * 0.5 / Iy

        TextBox15.Text = (w / 10 ^ 12).ToString("0.00")     '[kN/m]
        TextBox10.Text = (flex / 10 ^ 12).ToString("0.0")   '[mm] flex
        TextBox13.Text = (mom_max / 10 ^ 18).ToString("0.0") '[Nm]
        TextBox14.Text = (stress_b / 10 ^ 12).ToString("0") '[N/mm2]
        TextBox23.Text = NumericUpDown1.Value.ToString      '[mbar]
        '===== check ================
        TextBox14.BackColor = IIf(stress_b / 10 ^ 12 > NumericUpDown10.Value, Color.Red, Color.LightGreen)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim words() As String = UNP(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            NumericUpDown14.Value = CDec(words(1))  'Inertia Iy
            TextBox18.Text = CDec(words(2))        'Beam Height
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        Calc()
    End Sub
End Class
