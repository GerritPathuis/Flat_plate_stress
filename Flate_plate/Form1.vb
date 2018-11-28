Imports System.Math
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Threading
Imports System.Windows.Forms

Public Class Form1
    '"Name; I (strong axis)[cm4]; Profile height[mm]; kg; Ey[mm]"
    Public Shared UNP() As String = {
     "Strip 60x6; 10.8; 60;     2.8;    30",
     "Strip 60x8; 14.4; 60;     3.7;    30",
     "Strip 60x10; 18.0; 60;    4.7;    30",
     "Strip 80x6; 25.6; 80;     3.7;    40",
     "Strip 80x8; 34.1; 80;     5;      40",
     "Strip 80x10; 42.7; 80;    6.2;    40",
     "Strip 100x6; 50; 100;     4.7;    50",
     "Strip 100x8; 66.7; 100;  6.2;     50",
     "Strip 100x10; 83.3; 100; 7.8;     50",
     "Strip 120x6; 86.4; 120;  5.6;     60",
     "Strip 120x8; 115.2; 120; 7.5;     60",
     "Strip 120x10; 144.0; 120;9.4;     60",
     "Strip 120x12; 172.8; 120;11.2;    60",
     "Strip 140x6; 137.2; 140;  6.6;    70",
     "Strip 140x8; 182.9; 140;  8.7;    70",
     "Strip 140x10; 228.7; 140; 10.9;   70",
     "Strip 140x12; 274.4; 140; 13.1;   70",
     "Strip 160x6; 204.8; 160;  7.5;    80",
     "Strip 160x8; 273.1; 160;  10;     80",
     "Strip 160x10; 341.3; 160;  12.5;  80",
     "Strip 160x12; 409.6; 160; 15;     80",
     "Strip 180x6; 291.6; 180;  8.4;    90",
     "Strip 180x8; 388.8; 180;  11.2;   90",
     "Strip 180x10; 486.0; 180; 14;     90",
     "Strip 180x12; 583.2; 180; 16.8;   90",
     "Strip 200x8; 533.3; 200;  12.5;   100",
     "Strip 200x10; 666.7; 200; 15.6;   100",
     "Strip 200x12; 800.0; 200; 18.7;   100",
     "Strip 200x15; 1000; 200;  23.4;   100",
     "Hoek 60x60x6; 22.8; 60;   5.4;    16.9",
     "Hoek 70x70x8; 47.5; 70;   8.4;    20.1",
     "Hoek 80x80x8; 72.2; 80;   9.6;    22.6",
     "Hoek 80x80x10; 87.5; 80;  11.9;   23.4",
     "Hoek 80x80x12; 102; 80;   14.0;   24.1",
     "Hoek 100x100x8; 145; 100; 12.2;   27.4",
     "Hoek 100x100x10; 177; 100; 15.0;  28.2",
     "Hoek 100x100x12; 207; 100; 17.8;  29",
     "Hoek 100x100x15; 249; 100; 21.9;  30.2",
     "Hoek 120x120x10; 313; 120; 18.2;  33.1",
     "Hoek 120x120x12; 368; 120; 21.6;  34",
     "Hoek 120x120x15; 445; 120; 26.6;  35.1",
     "Hoek 150x150x12; 737; 150; 27.3;  42.1",
     "Hoek 150x150x15; 898; 150; 33.8;  42.5",
     "Hoek 150x150x18; 1050; 150; 40.1; 43.7",
     "Hoek 180x180x15; 1590; 180; 40.9; 49.8",
     "Hoek 180x180x18; 1870; 180; 48.6; 51",
     "Hoek 180x180x20; 2040; 180; 53.7; 51.8",
     "Hoek 200x200x16; 2340; 200; 48.5; 55.2",
     "Hoek 200x200x20; 2850; 200; 59.9; 56.8",
     "Hoek 200x200x24; 3330; 200; 71.1; 58.4",
     "UNP 40; 14.1; 40;   4.9;     20",
     "UNP 50; 26.4; 50;   5.6;     25",
     "UNP 65; 57.5; 65;   7.2;     32.5",
     "UNP 80; 106; 80;    8.6;     40",
     "UNP 100; 206; 100;  10.6;     50",
     "UNP 120; 364; 120;  13.4;     60",
     "UNP 140; 605; 140;  16.0;     70",
     "UNP 160; 925; 160;  18.8;     80",
     "UNP 180; 1350; 180; 22.0;     90",
     "UNP 200; 1910; 200; 25.3;     100",
     "UNP 220; 2690; 220; 29.4;     110",
     "UNP 240; 3600; 240; 33.2;     120",
     "UNP 260; 4820; 260; 37.9;     130",
     "UNP 280; 6280; 280; 41.8;     140",
     "UNP 300; 8030; 300; 46.2;     150",
     "UNP 320; 10870; 320; 59.5;    160",
     "UNP 350; 12840; 350; 60.6;    175",
     "UNP 380; 15760; 380; 62.6;    190",
     "UNP 400; 20350; 400; 71.8;    200",
     "HEB 100; 449; 100;   20.4;    50",
     "HEB 120; 864; 120;   26.7;    60",
     "HEB 140; 1509; 140;  33.7;    70",
     "HEB 160; 2492; 160;  42.6;    80",
     "HEB 180; 3831; 180;  51.2;    90",
     "HEB 200; 5696; 200;  61.3;    100",
     "HEB 220; 8091; 220;  71.5;    110",
     "HEB 240; 11260; 240;  83.2;   120",
     "HEB 260; 14920; 260;  93;     130",
     "HEB 280; 19270; 280;  103;    140",
     "HEB 300; 25170; 300;  177;    150"
    }

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        For hh = 0 To UNP.Length - 1             'Fill combobox1
            words = UNP(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
            ComboBox2.Items.Add(words(0))
            ComboBox3.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 56    'UNP 140
        ComboBox2.SelectedIndex = 55
        ComboBox3.SelectedIndex = 57
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged
        Calc()
    End Sub
    Private Sub Calc()
        'http://www.roymech.co.uk/Useful_Tables/Mechanics/Plates.html
        'Rectangle simply supported
        Dim a, b, t, flex As Double
        Dim Elas, p, σm, yt As Double

        p = NumericUpDown1.Value * 100          '[mbar]->[N/m2]]
        TextBox4.Text = p.ToString              '[N/m2]
        TextBox43.Text = (p / 10 ^ 6).ToString  '[N/mm2]
        a = NumericUpDown2.Value / 1000 '[m]    'Length
        b = NumericUpDown3.Value / 1000 '[m]    'Width
        t = NumericUpDown4.Value / 1000 '[m]    'Thickness
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
        TextBox2.BackColor = CType(IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen), Color)
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
        TextBox6.BackColor = CType(IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen), Color)
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
        TextBox8.BackColor = CType(IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen), Color)
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
        TextBox20.Text = (a * 2000).ToString("0")
        TextBox21.Text = (b * 2000).ToString("0")

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
        TextBox12.BackColor = CType(IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage5.Enter, NumericUpDown18.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown24.ValueChanged
        'http://beamguru.com/online/beam-calculator/
        'https://www.amesweb.info/StructuralBeamDeflection/SimplySupportedBeamStressDeflectionAnalysis.aspx
        Dim l, Iy, w, Elas As Double
        Dim flex, mom_max, σ_bend As Double
        Dim p, load_width, area As Double
        Dim beam_height, beam_section_area As Double
        Dim σ_tension, tension As Double
        Dim σ_combi As Double

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
        σ_bend = mom_max * beam_height * 0.5 / Iy
        σ_bend /= 10 ^ 12                             '[N/mm2]
        ' MessageBox.Show(σ_bend.ToString)

        '===== Tension σ ============================
        Double.TryParse(TextBox48.Text, beam_section_area)
        tension = NumericUpDown24.Value * 10 ^ 3     '[N]
        σ_tension = tension / beam_section_area      '[N/mm2]

        '========= Combines stress =============
        σ_combi = σ_bend + σ_tension

        '===== Present =
        TextBox15.Text = (w / 10 ^ 12).ToString("0.00")     '[kN/m]
        TextBox10.Text = (flex / 10 ^ 12).ToString("0.0")   '[mm] flex
        TextBox13.Text = (mom_max / 10 ^ 18).ToString("0.0") '[Nm]
        TextBox14.Text = σ_bend.ToString("0") '[N/mm2]
        TextBox23.Text = NumericUpDown1.Value.ToString      '[mbar]
        TextBox44.Text = σ_bend.ToString("0.0")           '[N/mm2]
        TextBox45.Text = σ_tension.ToString("0.0")          '[N/mm2]
        TextBox46.Text = σ_combi.ToString("0.0")            '[N/mm2]

        '===== check ================
        TextBox14.BackColor = CType(IIf(σ_bend > NumericUpDown10.Value, Color.Red, Color.LightGreen), Color)
        TextBox46.BackColor = CType(IIf(σ_combi > NumericUpDown10.Value, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim area As Double
        Try
            Dim words() As String = UNP(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            NumericUpDown14.Value = CDec(words(1))  'Inertia Iy
            TextBox18.Text = words(2)        'Beam Height
            TextBox47.Text = words(3)        'Beam weight
            area = Math.Round(CDbl(words(3)) / 0.00785, 0)
            TextBox48.Text = area.ToString    'Beam area
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        Calc()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown20.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, RadioButton1.CheckedChanged, NumericUpDown21.ValueChanged
        Dim moi As Double       'Moment of Inertia
        Dim wi, le, th As Double
        Dim max_d As Double     'Max deflection
        Dim ω As Double         'Uniformly distrib load
        Dim Elas As Double
        Dim p_load As Double     'Point load [N]

        Elas = NumericUpDown5.Value * 10 ^ 3    '[N/mm2]

        le = NumericUpDown12.Value      '[mm]
        wi = NumericUpDown13.Value      '[mm]
        th = NumericUpDown20.Value      '[mm]
        p_load = NumericUpDown21.Value  '[N]
        moi = wi * th ^ 3 / 12          'Moment of Inertia

        '---------- Select load type -------------
        If RadioButton1.Checked Then    'Point 
            PictureBox8.Visible = False 'Distr. load
            PictureBox9.Visible = True  'Point load
            GroupBox14.Visible = False
            GroupBox17.Visible = True
            max_d = p_load * le ^ 3 / (3 * Elas * moi)

        Else                            'Distributed load
            PictureBox8.Visible = True  'Distr. load
            PictureBox9.Visible = False 'Point load
            GroupBox14.Visible = True
            GroupBox17.Visible = False

            ω = wi * th / 10 ^ 9 * 7800 * 10 '[N/mm]
            max_d = ω * le ^ 4 / (8 * Elas * moi)
        End If

        TextBox26.Text = moi.ToString("0")      'Moment of Inertia
        TextBox27.Text = max_d.ToString("0.00") 'Max deflection
        TextBox28.Text = ω.ToString("0.00")     'Distrib. Load
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown22.ValueChanged, ComboBox3.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, TabPage8.Enter
        Calc_grill()
    End Sub
    'Design of Ship Hull Structures ISBN: 978-3-642-10009-3
    'Chapter Grillage Structure Page 254 
    '
    Private Sub Calc_grill()
        Dim p As Double         'Uniform load
        Dim a, b As Double      'Longer and shorter edge
        Dim Ix, Iy As Double    'Second Moment of inertia
        Dim m, n As Double      'No Girders
        Dim ey As Double        'Distance COG to plate face
        Dim llx, lly As Double  'Girder space
        Dim σy As Double        'Midpoint Stiffener Stress
        Dim δ As Double         'Midpoint (Max) deflection
        Dim Elas As Double      'Young modulus
        Dim weight As Double    'Girder weight
        Dim gir_wht_x As Double
        Dim gir_wht_y As Double
        Dim words() As String

        '---------- get data -------------
        'p = NumericUpDown25.Value / 10 ^ 4  '[mbar]->[N/mm2]
        Double.TryParse(TextBox43.Text, p)  '[N/mm2]
        TextBox33.Text = p.ToString  '[N/mm2]

        m = NumericUpDown29.Value
        n = NumericUpDown28.Value
        a = NumericUpDown22.Value
        b = NumericUpDown23.Value

        Elas = NumericUpDown5.Value * 10 ^ 3    '[N/mm2]

        If ComboBox2.SelectedIndex > 0 And ComboBox3.SelectedIndex > 0 Then
            '--- girders in X direction
            words = UNP(ComboBox2.SelectedIndex).Split(CType(";", Char()))
            Ix = CDbl(words(1)) * (10 ^ 4)      'Inertia Iy [cm^4->mm^4]
            gir_wht_x = CDbl(words(3))          '[kg]
            TextBox38.Text = gir_wht_x.ToString
            TextBox40.Text = (Ix / 10 ^ 4).ToString     '[cm4]
            TextBox32.Text = (Ix / 10 ^ 4).ToString     '[cm4]

            '--- girders in Y direction
            words = UNP(ComboBox3.SelectedIndex).Split(CType(";", Char()))
            Iy = CDbl(words(1)) * (10 ^ 4)      'Inertia Iy [cm^4->mm^4]
            gir_wht_y = CDbl(words(3))          '[kg]
            TextBox39.Text = gir_wht_y.ToString
            TextBox41.Text = (Iy / 10 ^ 4).ToString     '[cm4]
            TextBox34.Text = (Iy / 10 ^ 4).ToString     '[cm4]

            ey = CDbl(words(2)) - CDbl(words(4))
            TextBox42.Text = ey.ToString("0.0") 'distance to plate face [mm]
        End If


        '--------- calc girder spacing -------------
        llx = b / (m + 1)   'Girder space
        lly = a / (n + 1)   'Girder space

        '------ deflection and stress @ midpoint--------
        δ = a * b * p
        δ /= PI ^ 6 * Elas / 16
        δ /= ((Ix * (m + 1) / a ^ 3) + (Iy * (n + 1) / b ^ 3))
        σy = PI ^ 2 * δ * Elas * ey / b ^ 2

        '------ Girder weight ---------------------
        weight = n * b / 1000 * gir_wht_y    'Y (vertical girder) [kg]
        weight += m * a / 1000 * gir_wht_x   'X (horizontal girder) [kg]

        '------ present ----------
        TextBox29.Text = σy.ToString("0")       '[N/mm2]
        TextBox31.Text = δ.ToString("0.0")      '[mm]
        ' TextBox33.Text = p.ToString("0.0000")   '[N/mm2]

        TextBox36.Text = llx.ToString("0")      '[mm]
        TextBox35.Text = lly.ToString("0")      '[mm]

        TextBox35.Text = lly.ToString("0")      '[mm]
        TextBox37.Text = weight.ToString("0.0") '[kg]

        '------ check -shorter/longer edge---------
        NumericUpDown22.BackColor = CType(IIf(a > b, Color.Yellow, Color.Red), Color)
        TextBox29.BackColor = CType(IIf(σy < NumericUpDown10.Value, Color.LightGreen, Color.Red), Color)
    End Sub


End Class
