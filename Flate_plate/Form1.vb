Imports System.Math
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Threading
Imports System.Windows.Forms

Public Class Form1
    'https://en.wikipedia.org/wiki/List_of_second_moments_of_area
    'https://calcresource.com/cross-section-angle.html
    'https://www.eurocodeapplied.com/design/en1993/ipe-hea-heb-hem-design-properties
    '"Name; I (strong axis)[cm4]; Profile height[mm]; [kg/m]; Ey[mm]; Zx [mm3]"
    Public Shared UNP() As String = {
     "Angle 60x60x6;    22.8;   60;   5.4;  16.9; 0.0",
     "Angle 70x70x8;    47.5;   70;   8.4;  20.1; 0.0",
     "Angle 80x80x8;    72.2;   80;   9.6;  22.6; 0.0",
     "Angle 80x80x10;   87.5;   80;  11.9;  23.4; 0.0",
     "Angle 80x80x12;   102;    80;  14.0;  24.1; 0.0",
     "Angle 100x100x8;  145;    100; 12.2;  27.4; 0.0",
     "Angle 100x100x10; 177;    100; 15.0;  28.2; 0.0",
     "Angle 100x100x12; 207;    100; 17.8;  29; 0.0",
     "Angle 100x100x15; 249;    100; 21.9;  30.2; 0.0",
     "Angle 120x120x10; 313;    120; 18.2;  33.1; 0.0",
     "Angle 120x120x12; 368;    120; 21.6;  34; 0.0",
     "Angle 120x120x15; 445;    120; 26.6;  35.1; 0.0",
     "Angle 150x150x12; 737;    150; 27.3;  42.1; 0.0",
     "Angle 150x150x15; 898;    150; 33.8;  42.5; 0.0",
     "Angle 150x150x18; 1050;   150; 40.1; 43.7; 0.0",
     "Angle 180x180x15; 1590;   180; 40.9; 49.8; 0.0",
     "Angle 180x180x18; 1870;   180; 48.6; 51; 0.0",
     "Angle 180x180x20; 2040;   180; 53.7; 51.8; 0.0",
     "Angle 200x200x16; 2340;   200; 48.5; 55.2; 0.0",
     "Angle 200x200x20; 2850;   200; 59.9; 56.8; 0.0",
     "Angle 200x200x24; 3330;   200; 71.1; 58.4; 0.0",
     "Bulb flat 160x10; 481; 160;   15.3;   92.6; 0.0", 'Holland profile
     "Bulb flat 180x10; 712; 180;   17.6;   106; 0.0",
     "Bulb flat 200x12; 1157; 200;  23.3;   117; 0.0",
     "Bulb flat 220x12; 1586; 220;  26.2;   130; 0.0",
     "Bulb flat 240x12; 2117; 240;  29.3;   144; 0.0",
     "Bulb flat 260x12; 2762; 260;  32.4;   158; 0.0",
     "Bulb flat 280x12; 3525; 280;  35.7;   172; 0.0",
     "Bulb flat 320x12; 5506; 320;  42.6;   201; 0.0",
     "Bulb flat 340x14; 7504; 340;  51.5;   211; 0.0",
     "Bulb flat 400x14; 12873; 400; 63.9;   255; 0.0",
     "HEB 100;      449; 100;   20.4;    50; 0.0",
     "HEB 120;      864; 120;   26.7;    60; 0.0",
     "HEB 140;      1509; 140;  33.7;    70; 0.0",
     "HEB 160;      2492; 160;  42.6;    80; 0.0",
     "HEB 180;      3831; 180;  51.2;    90; 0.0",
     "HEB 200;      5696; 200;  61.3;    100; 0.0",
     "HEB 220;      8091; 220;  71.5;    110; 0.0",
     "HEB 240;      11260; 240;  83.2;   120; 0.0",
     "HEB 260;      14920; 260;  93;     130; 0.0",
     "HEB 280;      19270; 280;  105;    140; 0.0",
     "HEB 300;      25170; 300;  119;    150; 0.0",
     "HEB 320;      30820; 320;  129;    160; 0.0",
     "HEB 340;      36600; 340;  137;    170; 0.0",
     "HEB 360;      43190; 360;  145;    180; 0.0",
     "HEB 400;      57680; 400;  158;    200; 0.0",
     "HEB 450;      79890; 450;  174;    225; 0.0",
     "HEB 500;      107200; 500;  190;   250; 0.0",
     "HEB 550;      136700; 550;  203;   275; 0.0",
     "HEB 600;      171000; 600;  216;   300; 0.0",
     "HEB 650;      210600; 650;  229;   325; 0.0",
     "HEB 700;      256900; 700;  245;   350; 0.0",
     "HEB 800;      359100; 800;  267;   400; 0.0",
     "Strip 20x3;   0.2;    20;  0.47;      10; 300",
     "Strip 40x4;   2.13;   40;   1.2;      20; 1600",
     "Strip 60x6;   10.8;   60;   2.8;      30; 5400",
     "Strip 60x8;   14.4;   60;   3.7;      30; 7200",
     "Strip 60x10;  18.0;   60;   4.7;      30; 9000",
     "Strip 80x6;   25.6;   80;   3.7;      40; 9600",
     "Strip 80x8;   34.1;   80;     5;      40; 12800",
     "Strip 80x10;  42.7;   80;     6.2;    40; 16000",
     "Strip 100x6;    50;  100;     4.7;    50; 15000",
     "Strip 100x8;   66.7; 100;     6.2;    50; 20000",
     "Strip 100x10;  83.3; 100;     7.8;    50; 25000",
     "Strip 120x6;   86.4; 120;     5.6;    60; 21600",
     "Strip 120x8;  115.2; 120;     7.5;    60; 28800",
     "Strip 120x10; 144.0; 120;     9.4;    60; 36000",
     "Strip 120x12; 172.8; 120;     11.2;   60; 43200",
     "Strip 140x6;  137.2; 140;     6.6;    70; 29400",
     "Strip 140x8;  182.9; 140;      8.7;   70; 39200",
     "Strip 140x10; 228.7; 140;     10.9;   70; 49000",
     "Strip 140x12; 274.4; 140;     13.1;   70; 58800",
     "Strip 160x6;  204.8; 160;      7.5;   80; 38400",
     "Strip 160x8;  273.1; 160;      10;    80; 51200",
     "Strip 160x10; 341.3; 160;     12.5;   80; 64000",
     "Strip 160x12; 409.6; 160;     15;     80; 76800",
     "Strip 180x6;  291.6; 180;      8.4;   90; 48600",
     "Strip 180x8;  388.8; 180;      11.2;  90; 64800",
     "Strip 180x10; 486.0; 180;     14;     90; 81000",
     "Strip 180x12; 583.2; 180;     16.8;   90; 97200",
     "Strip 200x8;  533.3; 200;     12.5;   100; 80000",
     "Strip 200x10; 666.7; 200;     15.6;   100; 100000",
     "Strip 200x12; 800.0; 200;     18.7;   100; 120000",
     "Strip 200x15;  1000; 200;     23.4;   100; 150000",
     "Strip 250x15;  1953; 250;     29.2;   125; 234375",
     "Strip 250x20;  2604; 250;     39.0;   125; 312500",
     "Strip 300x20;  4500; 300;     46.8;   150; 450000",
     "UNP 40; 14.1; 40;   4.9;     20; 0.0",
     "UNP 50; 26.4; 50;   5.6;     25; 0.0",
     "UNP 65; 57.5; 65;   7.2;     32.5; 0.0",
     "UNP 80; 106; 80;    8.6;     40; 0.0",
     "UNP 100; 206; 100;  10.6;     50; 0.0",
     "UNP 120; 364; 120;  13.4;     60; 0.0",
     "UNP 140; 605; 140;  16.0;     70; 0.0",
     "UNP 160; 925; 160;  18.8;     80; 0.0",
     "UNP 180; 1350; 180; 22.0;     90; 0.0",
     "UNP 200; 1910; 200; 25.3;     100; 0.0",
     "UNP 220; 2690; 220; 29.4;     110; 0.0",
     "UNP 240; 3600; 240; 33.2;     120; 0.0",
     "UNP 260; 4820; 260; 37.9;     130; 0.0",
     "UNP 280; 6280; 280; 41.8;     140; 0.0",
     "UNP 300; 8030; 300; 46.2;     150; 0.0",
     "UNP 320; 10870; 320; 59.5;    160; 0.0",
     "UNP 350; 12840; 350; 60.6;    175; 0.0",
     "UNP 380; 15760; 380; 62.6;    190; 0.0",
     "UNP 400; 20350; 400; 71.8;    200; 0.0"
     }

    Public Shared _ρ_steel As Double = 7850

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        For hh = 0 To UNP.Length - 1             'Fill combobox1
            words = UNP(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
            ComboBox2.Items.Add(words(0))
            ComboBox3.Items.Add(words(0))
            ComboBox4.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 46    'HEB 450
        ComboBox2.SelectedIndex = 94    'UNP 160
        ComboBox3.SelectedIndex = 89    'UNP 65
        ComboBox4.SelectedIndex = 63    'Strip 100x10
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged
        Calc()
    End Sub
    Private Sub Calc()
        'http://www.roymech.co.uk/Useful_Tables/Mechanics/Plates.html
        'Rectangle simply supported
        Dim a, b, t, flex As Double
        Dim Elas, p, σm, yt As Double
        Dim weight As Double

        p = NumericUpDown1.Value * 100                  '[mbar]->[N/m2]]
        TextBox4.Text = p.ToString                      '[N/m2]
        TextBox43.Text = (p / 10 ^ 6).ToString          '[N/mm2]
        TextBox49.Text = NumericUpDown1.Value.ToString  '[mbar]
        TextBox50.Text = NumericUpDown1.Value.ToString  '[mbar]

        a = NumericUpDown2.Value / 1000 '[m]    'Length
        b = NumericUpDown3.Value / 1000 '[m]    'Width
        t = NumericUpDown4.Value / 1000 '[m]    'Thickness
        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]
        weight = a * b * t * _ρ_steel           '[kg]

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
        TextBox51.Text = weight.ToString("0")

        '===== checks ================
        TextBox2.BackColor = CType(IIf(σm > NumericUpDown10.Value, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage3.Enter, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged
        'Rectangle clamped edges
        Dim a, b, t As Double
        Dim Elas, p, σm, yt, flex As Double
        Dim weight As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        TextBox4.Text = p.ToString
        a = NumericUpDown8.Value / 1000 '[m]
        b = NumericUpDown7.Value / 1000 '[m]
        t = NumericUpDown6.Value / 1000 '[m]
        Elas = NumericUpDown5.Value * 10 ^ 9    '[GPa]
        weight = a * b * t * _ρ_steel           '[kg]

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
        TextBox52.Text = weight.ToString("0")
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage5.Enter, NumericUpDown18.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown24.ValueChanged
        Calc_beam()

    End Sub
    Private Sub Calc_beam()
        'http://beamguru.com/online/beam-calculator/
        'https://www.amesweb.info/StructuralBeamDeflection/SimplySupportedBeamStressDeflectionAnalysis.aspx
        Dim l, Iy, w, Elas As Double
        Dim flex, mom_max, σ_bend As Double
        Dim p, load_width, area As Double
        Dim beam_height, beam_section_area As Double
        Dim σ_tension, tension As Double
        Dim σ_combi As Double

        l = NumericUpDown18.Value * 10 ^ 3      '[m->mm]
        Double.TryParse(TextBox41.Text, Iy)
        Iy *= 10 ^ 4     '[mm4]
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
            TextBox41.Text = words(1)               'Inertia Iy [cm4]
            TextBox18.Text = words(2)               'Beam Height [mm]
            TextBox47.Text = words(3)               'Beam weight [mm]
            area = Math.Round(CDbl(words(3)) * 10 ^ 6 / _ρ_steel, 0)
            TextBox48.Text = area.ToString          'Beam area
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
        Calc_beam()
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
        Dim press As Double             'Uniform load
        Dim a_hor As Double             'Longer horizontal edge
        Dim b_vert As Double            'Shorter vertical edge
        Dim I_hor_beam As Double        'Second Moment of inertia
        Dim I_vert_girder As Double     'Second Moment of inertia
        Dim no_vert_girders As Double   'No Girders vertical
        Dim no_hor_beams As Double      'No Beams horintal
        Dim ey_girder As Double         'Distance COG to plate face
        Dim space_beams As Double       'Beams space 
        Dim space_girders As Double     'Girder space 
        Dim σy As Double                'Midpoint Stiffener Stress
        Dim δ As Double                 'Midpoint (Max) deflection
        Dim Elas As Double              'Young modulus
        Dim weight As Double            'Total weight
        Dim gir_vert_wht As Double      'Girder_vert 
        Dim beam_hor_wht As Double      'Beam vertical
        Dim words() As String
        Dim l_opti As Double
        Dim δ1, δ2 As Double            'Temp value for readability

        '---------- get data -------------
        Double.TryParse(TextBox43.Text, press)          '[N/mm2]
        TextBox33.Text = press.ToString                 '[N/mm2]
        TextBox40.Text = NumericUpDown1.Value.ToString  '[mbar]

        '--- NOTE GIRDER is vertical , BEAM is horizontal------
        no_vert_girders = NumericUpDown29.Value    'no Girder
        no_hor_beams = NumericUpDown28.Value       'no Beams
        a_hor = NumericUpDown22.Value              'Longer edge
        b_vert = NumericUpDown23.Value             'Shortes edge

        Elas = NumericUpDown5.Value * 10 ^ 3        '[N/mm]

        If ComboBox2.SelectedIndex > -1 And ComboBox3.SelectedIndex > -1 Then
            '--- Beams Horizontal
            words = UNP(ComboBox3.SelectedIndex).Split(CType(";", Char()))
            I_hor_beam = CDbl(words(1)) * (10 ^ 4)  'Inertia Iy [cm^4->no_gird^4]
            beam_hor_wht = CDbl(words(3))           '[kg]
            TextBox39.Text = beam_hor_wht.ToString
            TextBox34.Text = (I_hor_beam / 10 ^ 4).ToString     '[cm4]

            '--- Girders VERTICAL
            words = UNP(ComboBox2.SelectedIndex).Split(CType(";", Char()))
            I_vert_girder = CDbl(words(1)) * (10 ^ 4)   'Inertia Iy [cm^4->mm^4]
            gir_vert_wht = CDbl(words(3))               '[kg]
            TextBox38.Text = gir_vert_wht.ToString
            TextBox32.Text = (I_vert_girder / 10 ^ 4).ToString     '[cm4]

            ey_girder = CDbl(words(2)) - CDbl(words(4))
            TextBox42.Text = ey_girder.ToString("0.0")      'Distance to plate face [mm]
        End If

        '--------- calc girder spacing -------------
        space_beams = a_hor / (no_hor_beams + 1)            'Beam space
        space_girders = b_vert / (no_vert_girders + 1)      'Girder space

        '------ deflection and stress @ midpoint--------
        '------ formula 6.1.1---------------------------

        δ1 = I_hor_beam * (no_hor_beams + 1) / a_hor ^ 3
        δ2 = I_vert_girder * (no_vert_girders + 1) / b_vert ^ 3

        δ = a_hor * b_vert * press
        δ /= PI ^ 6 * Elas / 16
        δ /= (δ1 + δ2)
        σy = PI ^ 2 * δ * Elas * ey_girder / b_vert ^ 2

        '------ Girder weight ---------------------
        weight = no_vert_girders * b_vert / 1000 * gir_vert_wht   'Girder vertical [kg]
        weight += no_hor_beams * a_hor / 1000 * beam_hor_wht    'Beam horizontal [kg]

        '------ Optimum girder distance (formula 6.2.6)-----------
        '------ Both ends of the girder is supported--------------
        l_opti = 0.72 * space_beams ^ (1 / 3) * b_vert ^ (2 / 3)

        '------ present ----------
        TextBox29.Text = σy.ToString("0")       '[N/mm2]
        TextBox31.Text = δ.ToString("0.0")      '[mm]

        TextBox35.Text = space_girders.ToString("0")        'Girder distance
        TextBox36.Text = space_beams.ToString("0")          'Beam distance
        TextBox37.Text = weight.ToString("0.0")             '[kg]

        Label148.Text = "If girders and beam identical Optimum Girder space= " & l_opti.ToString("0") & " [mm]"

        '------ check -shorter/longer edge---------
        NumericUpDown22.BackColor = CType(IIf(a_hor > b_vert, Color.Yellow, Color.Red), Color)
        TextBox29.BackColor = CType(IIf(σy < NumericUpDown10.Value, Color.LightGreen, Color.Red), Color)
    End Sub
    'Stress and strain table 8.13 page 260
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click, TabPage10.Enter, NumericUpDown25.ValueChanged, NumericUpDown14.ValueChanged, ComboBox4.SelectedIndexChanged
        Dim words() As String
        Dim Mp As Double            'Plastic moment
        Dim σ As Double             'max allowed stress [n/mm2]
        Dim l_beam As Double        'Beam length [mm]
        Dim p_width As Double       'Pressure width [mm]
        Dim press As Double         '[N/mm2]
        Dim area_beam, Z_plastic As Double 'Plasic modul [mm3]
        Dim Wa As Double            'Actual Uniform load [N/mm]
        Dim Wac As Double           'Collapse Uniform load [N/mm]

        '------------ get data -----------
        σ = NumericUpDown10.Value           'max allowed stress [n/mm2]
        l_beam = NumericUpDown25.Value      'beam length [mm]
        p_width = NumericUpDown14.Value     'pressure width [mm]

        If ComboBox4.SelectedIndex > -1 Then
            '--- Beams Horizontal
            words = UNP(ComboBox4.SelectedIndex).Split(CType(";", Char()))
            TextBox55.Text = words(1)  'Inertia Iy [cm^4]
            area_beam = Math.Round(CDbl(words(3)) * 10 ^ 6 / _ρ_steel, 0)
            TextBox56.Text = area_beam.ToString     'Area [cm^2]
            TextBox57.Text = words(3)               '[kg]
            TextBox58.Text = words(2)               '[mm]
            TextBox61.Text = words(5)               '[mm3]
            Z_plastic = CDbl(words(5))              '[mm3]
            press = NumericUpDown1.Value / 10 ^ 4   '[mbar]-->[N/mm2]
        End If

        '-------------- calc -------------
        Wa = press * p_width                    '[N/mm]
        Mp = σ * Z_plastic                      '[Nm]
        Wac = 16 * Mp * l_beam ^ 2 / l_beam ^ 4 '[Nm]

        '-------------- Present ----------
        TextBox53.Text = press.ToString("0.000")        '[N/mm2]
        TextBox54.Text = (Mp / 1000).ToString("0.0")    '[kNmm]
        TextBox62.Text = Wa.ToString("0.0")             '[N/mm] 
        TextBox63.Text = Wac.ToString("0.0")            '[N/mm]
        '-------------- Checks --------
        TextBox62.BackColor = CType(IIf(Wa < Wac, Color.LightGreen, Color.Red), Color)
    End Sub
End Class
