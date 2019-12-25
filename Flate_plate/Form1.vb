Imports System.Math
Imports System.Globalization
Imports System.Threading

Public Structure Grillstruct
    'Implements IComparer(Of Grillstruct)
    Public girder As Integer      'Girder no
    Public beam As Integer        'Beam no
    Public weight As Double       'Total weight

    'Public Function Compare(x As Grillstruct, y As Grillstruct) As Integer Implements IComparer(Of Grillstruct).Compare
    '    Throw New NotImplementedException()
    '    If x.weight < y.weight Then
    '        Return (-1)
    '    Else
    '        Return (+1)
    '    End If
    'End Function
End Structure

Public Class Form1
    Public grillresults(13400) As Grillstruct    'Auto grillage


    'https://en.wikipedia.org/wiki/Listofsecondmomentsofarea
    'https://calcresource.com/cross-section-angle.html
    'https://www.eurocodeapplied.com/design/en1993/ipe-hea-heb-hem-design-properties
    'https://www.eurocodeapplied.com/design/en1993/ipe-hea-heb-hem-design-properties
    '"Name; I (strong axis)[cm4]; Profile height[mm]; [kg/m]; Ey[mm]; Zx [mm3]"
    ReadOnly UNP() As String = {
     "Not required;       0;    0;     0;     0;     0",
     "Angle 60x60x6;    22.8;   60;   5.4;  16.9;  5452",
     "Angle 70x70x8;    47.5;   70;   8.4;  20.1; 9767",
     "Angle 80x80x8;    72.2;   80;   9.6;  22.6; 12923",
     "Angle 80x80x10;   87.5;   80;  11.9;  23.4; 15796",
     "Angle 80x80x12;   102;    80;  14.0;  24.1; 18548",
     "Angle 100x100x8;  145;    100; 12.2;  27.4; 20567", 'OK
     "Angle 100x100x10; 177;    100; 15.0;  28.2; 25240",
     "Angle 100x100x12; 207;    100; 17.8;  29;   29748",
     "Angle 100x100x15; 249;    100; 21.9;  30.2; 36227",
     "Angle 120x120x10; 313;    120; 18.2;  33.1; 36907",
     "Angle 120x120x12; 368;    120; 21.6;  34;   43615",
     "Angle 120x120x15; 445;    120; 26.6;  35.1; 53311",
     "Angle 150x150x12; 737;    150; 27.3;  42.1; 69415",
     "Angle 150x150x15; 898;    150; 33.8;  42.5; 85186",
     "Angle 150x150x18; 1050;   150; 40.1; 43.7;  100402",
     "Angle 180x180x15; 1590;   180; 40.9; 49.8;  124562",
     "Angle 180x180x18; 1870;   180; 48.6; 51;    147202",
     "Angle 180x180x20; 2040;   180; 53.7; 51.8;  161923",
     "Angle 200x200x16; 2340;   200; 48.5; 55.2;  164541",
     "Angle 200x200x20; 2850;   200; 59.9; 56.8;  201924",
     "Angle 200x200x24; 3330;   200; 71.1; 58.4;  237989",
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
     "HEB 100;      449; 100;   20.4;    50;  104000", 'Strong axis OK
     "HEB 120;      864; 120;   26.7;    60;  165000",
     "HEB 140;      1509; 140;  33.7;    70;  245000",
     "HEB 160;      2492; 160;  42.6;    80;  354000",
     "HEB 180;      3831; 180;  51.2;    90;  481000",
     "HEB 200;      5696; 200;  61.3;    100; 642000",
     "HEB 220;      8091; 220;  71.5;    110; 827000",
     "HEB 240;      11260; 240;  83.2;   120; 1053000",
     "HEB 260;      14920; 260;  93;     130; 1283000",
     "HEB 280;      19270; 280;  105;    140; 1534000",
     "HEB 300;      25170; 300;  119;    150; 1869000",
     "HEB 320;      30820; 320;  129;    160; 2149000",
     "HEB 340;      36600; 340;  137;    170; 2408000",
     "HEB 360;      43190; 360;  145;    180; 2683000",
     "HEB 400;      57680; 400;  158;    200; 3232000",
     "HEB 450;      79890; 450;  174;    225; 3982000",
     "HEB 500;      107200; 500;  190;   250; 4815000",
     "HEB 550;      136700; 550;  203;   275; 5591000",
     "HEB 600;      171000; 600;  216;   300; 6425000",
     "HEB 650;      210600; 650;  229;   325; 7320000",
     "HEB 700;      256900; 700;  245;   350; 8327000",
     "HEB 800;      359100; 800;  267;   400; 10259000",
     "Strip 20x3;   0.20;   20;  0.47;      10; 300",   'Strong axis OK
     "Strip 40x4;   2.13;   40;   1.2;      20; 1600",
     "Strip 50x5;   5.20;   50;   1.96;     25; 3125",
     "Strip 60x6;   10.8;   60;   2.8;      30; 5400",
     "Strip 60x8;   14.4;   60;   3.7;      30; 7200",
     "Strip 60x10;  18.0;   60;   4.7;      30; 9000",
     "Strip 80x6;   25.6;   80;   3.7;      40; 9600",
     "Strip 80x8;   34.1;   80;     5;      40; 12800",
     "Strip 80x10;  42.7;   80;     6.2;    40; 16000",
     "Strip 80x12;  51.2;   80;     7.5;    40; 19200", 'New
     "Strip 80x15;  64.0;   80;     9.4;    40; 24000", 'New
     "Strip 80x20;  85.3;   80;    12.6;    40; 32000", 'New
     "Strip 100x6;    50;  100;     4.7;    50; 15000",
     "Strip 100x8;   66.7; 100;     6.2;    50; 20000",
     "Strip 100x10;  83.3; 100;     7.8;    50; 25000",
     "Strip 100x12; 100.0; 100;     9.4;    50; 30000", 'new
     "Strip 100x15; 125.0; 100;     11.8;   50; 37500", 'new
     "Strip 100x20; 166.7; 100;     15.7;   50; 33300", 'new
     "Strip 120x6;   86.4; 120;     5.6;    60; 21600",
     "Strip 120x8;  115.2; 120;     7.5;    60; 28800",
     "Strip 120x10; 144.0; 120;     9.4;    60; 36000",
     "Strip 120x12; 172.8; 120;     11.2;   60; 43200",
     "Strip 140x6;  137.2; 140;     6.6;    70; 29400",
     "Strip 140x8;  182.9; 140;      8.7;   70; 39200",
     "Strip 140x10; 228.7; 140;     10.9;   70; 49000",
     "Strip 140x12; 274.4; 140;     13.1;   70; 58800",
     "Strip 150x8;  225.0; 150;      9.4;   75; 45000", 'new
     "Strip 150x10; 281.1; 150;     11.8;   75; 56250", 'new
     "Strip 160x6;  204.8; 160;      7.5;   80; 38400",
     "Strip 160x8;  273.1; 160;       10;   80; 51200",
     "Strip 160x10; 341.3; 160;     12.5;   80; 64000",
     "Strip 160x12; 409.6; 160;     15.0;   80; 76800",
     "Strip 180x6;  291.6; 180;      8.4;   90; 48600",
     "Strip 180x8;  388.8; 180;     11.2;   90; 64800",
     "Strip 180x10; 486.0; 180;     14;     90; 81000",
     "Strip 180x12; 583.2; 180;     16.8;   90; 97200",
     "Strip 200x8;  533.3; 200;     12.5;   100; 80000",
     "Strip 200x10; 666.7; 200;     15.6;   100; 100000",
     "Strip 200x12; 800.0; 200;     18.7;   100; 120000",
     "Strip 200x15;  1000; 200;     23.4;   100; 150000",
     "Strip 250x15;  1953; 250;     29.2;   125; 234375",
     "Strip 250x20;  2604; 250;     39.0;   125; 312500",
     "Strip 300x20;  4500; 300;     46.8;   150; 450000",
     "UNP 40;   14.1; 40;   4.9;     20;     5832",
     "UNP 50;   26.4; 50;   5.6;     25;     11350",
     "UNP 65;   57.5; 65;   7.2;     32.5;   19212",
     "UNP 80;   106; 80;    8.6;     40;     29228",
     "UNP 100;  206; 100;  10.6;     50;     49221",    'OK
     "UNP 120;  364; 120;  13.4;     60;     73152",
     "UNP 140;  605; 140;  16.0;     70;     103200",
     "UNP 160;  925; 160;  18.8;     80;     138261",
     "UNP 180;  1350; 180; 22.0;     90;     180058",
     "UNP 200;  1910; 200; 25.3;     100;    229155",
     "UNP 220;  2690; 220; 29.4;     110;    284841",
     "UNP 240;  3600; 240; 33.2;     120;    359601",
     "UNP 260;  4820; 260; 37.9;     130;    444520",
     "UNP 280;  6280; 280; 41.8;     140;    533875",
     "UNP 300;  8030; 300; 46.2;     150;    633960",
     "UNP 320;  10870; 320; 59.5;    160;    813663",
     "UNP 350;  12840; 350; 60.6;    175;    888334",
     "UNP 380;  15760; 380; 62.6;    190;    1002770",
     "UNP 400;  20350; 400; 71.8;    200;    1132900"
     }

    Public Shared ρsteel As Double = 7850     'Density steel [kg/m3]
    Public Shared σ02 As Double               'Yield stress [N/mm2]
    Public Shared sf As Double                 'Safety factor [-]
    Public Shared σdesign As Double           'Design stress [N/mm2]
    Public Shared younggpa As Double          'Youngs modulus [Gpa]
    Public Shared youngNmm2 As Double         'Youngs modulus [N/mm2]
    Public Shared youngNm2 As Double          'Youngs modulus [N/m2]


    Private Sub Form1Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        For hh = 0 To UNP.Length - 1             'Fill combobox1
            words = UNP(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
            ComboBox2.Items.Add(words(0))   'Girders (short)
            ComboBox3.Items.Add(words(0))
            ComboBox4.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 49    'HEB 
        ComboBox2.SelectedIndex = 94    'UNP 160    Girders (short)
        ComboBox3.SelectedIndex = 89    'UNP 65
        ComboBox4.SelectedIndex = 63    'Strip 100x10

        For Each tg As TabPage In TabControl1.TabPages
            tg.BackColor = Color.Snow
        Next
        Getsteelproperties()

    End Sub
    Private Sub Getsteelproperties()
        σ02 = NumericUpDown10.Value               '[N/mm2]Yield strength
        sf = NumericUpDown43.Value                 'Safety value
        σdesign = σ02 / sf                     'Design stress [N/mm2]
        TextBox115.Text = σdesign.ToString("F0")  '[N/mm2] design stress
    End Sub
    Public Sub Calcyoung(tt As Double)
        'https://www.engineeringtoolbox.com/young-modulus-d773.html

        If tt < -200 Or tt > 650 Then MsgBox("Youngs modulus, temperature out of range")
        younggpa = -0.000000324 * tt ^ 3 + 0.000049951 * tt ^ 2 - 0.04930174 * tt + 203.386 '[GPa]
        youngNmm2 = younggpa * 10 ^ 3               '[N/mm2] 
        youngNm2 = younggpa * 10 ^ 9                '[N/m2] 

        TextBox116.Text = younggpa.ToString("F1")     '[GPa] 
        TextBox117.Text = youngNmm2.ToString("F0")    '[N/mm2] 
    End Sub

    Private Sub Button1Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged
        Calcinput()
    End Sub
    Private Sub Calcinput()
        'http://www.roymech.co.uk/UsefulTables/Mechanics/Plates.html
        'Rectangle simply supported
        Dim a, b, t, flex As Double
        Dim p, σm, yt As Double
        Dim weight, cost As Double

        Calcyoung(NumericUpDown5.Value)               'Get young modulus
        Getsteelproperties()                          'Allowed stress

        p = NumericUpDown1.Value * 100                  '[mbar]->[N/m2]]
        TextBox4.Text = p.ToString                      '[N/m2]
        TextBox43.Text = (p / 10 ^ 6).ToString          '[N/mm2]
        TextBox49.Text = NumericUpDown1.Value.ToString  '[mbar]

        a = NumericUpDown2.Value / 1000 '[m]        'Length
        b = NumericUpDown3.Value / 1000 '[m]        'Width
        t = NumericUpDown4.Value / 1000 '[m]        'Thickness
        weight = a * b * t * ρsteel               '[kg]
        cost = weight * NumericUpDown33.Value       '[euro]

        If a >= b Then
            Label40.Visible = False
            σm = 0.75 * p * b ^ 2
            σm /= t ^ 2 * (1.61 * (b / a) ^ 3 + 1)
            σm /= 10 ^ 6        '[N/mm2]

            yt = 0.142 * p * b ^ 4
            yt /= youngNm2 * t ^ 3 * (2.21 * (b / a) ^ 3 + 1)
            yt *= 1000          '[m]-->[mm]
            flex = a * 10 ^ 3 / yt
        Else
            σm = 0
            yt = 0
            Label40.Visible = True
            flex = 0
        End If

        '--- Roark's 8 edition, chapter 11.13, ref 43, page 468 ----
        Dim Wu As Double        'Ultimate load
        Dim β As Double
        Dim ratio As Double

        ratio = b / a
        Select Case True
            Case (ratio <= 1 And ratio > 0.9)
                β = 5.48
            Case (ratio <= 0.9 And ratio > 0.8)
                β = 5.5
            Case (ratio <= 0.8 And ratio > 0.7)
                β = 5.58
            Case (ratio <= 0.7 And ratio > 0.6)
                β = 5.64
            Case (ratio <= 0.6 And ratio > 0.5)
                β = 5.89
            Case (ratio <= 0.5 And ratio > 0.4)
                β = 6.15
            Case (ratio <= 0.4 And ratio > 0.3)
                β = 6.7
            Case (ratio <= 0.3 And ratio > 0.2)
                β = 7.68
            Case (ratio <= 0.2)
                β = 9.69
        End Select

        Wu = β * σ02 * (t * 1000) ^ 2     '[N] Ultimate load
        Wu /= (a * 1000) * (b * 1000)       '[N/mm2] Ultimate pressure
        Wu *= 10 ^ 4                        '[N/mm2] --> [mbar]

        '----- Present -----------------
        TextBox2.Text = σm.ToString("0")
        TextBox3.Text = yt.ToString("0.0")
        TextBox25.Text = flex.ToString("0")
        TextBox51.Text = weight.ToString("0")
        TextBox87.Text = cost.ToString("0")

        TextBox69.Text = β.ToString("0.0")
        TextBox70.Text = Wu.ToString("0")
        TextBox71.Text = ratio.ToString("0.00")
        TextBox72.Text = σ02.ToString("0")            'Yield stress
        '===== checks ================
        TextBox2.BackColor = CType(IIf(σm > σ02, Color.Red, Color.LightGreen), Color)
        TextBox70.BackColor = CType(IIf(Wu < NumericUpDown1.Value, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button2Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage3.Enter, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged
        'Rectangle clamped edges
        Dim a, b, t As Double
        Dim p, σm, yt, flex As Double
        Dim weight, cost As Double

        p = NumericUpDown1.Value * 100 '[mbar->[N/m2]]
        TextBox4.Text = p.ToString
        a = NumericUpDown8.Value / 1000         '[m]
        b = NumericUpDown7.Value / 1000         '[m]
        t = NumericUpDown6.Value / 1000         '[m]
        weight = a * b * t * ρsteel           '[kg]
        cost = weight * NumericUpDown33.Value   '[Euro]

        If a >= b Then
            Label39.Visible = False
            σm = p * b ^ 2
            σm /= t ^ 2 * (2.63 * (b / a) ^ 6 + 1)
            σm /= 10 ^ 6                        '[N/mm2]

            yt = 0.0284 * p * b ^ 4
            yt /= youngNm2 * t ^ 3 * (1.056 * (b / a) ^ 5 + 1)
            yt *= 1000                           '[m]-->[mm]
            flex = a * 10 ^ 3 / yt
        Else
            σm = 0
            yt = 0
            Label39.Visible = True
            flex = 0
        End If

        TextBox6.Text = σm.ToString("0")
        TextBox5.Text = yt.ToString("0.0")
        TextBox24.Text = flex.ToString("0")             '[mm]
        TextBox50.Text = NumericUpDown1.Value.ToString  '[mbar]
        TextBox52.Text = weight.ToString("0")           '[kg]
        TextBox86.Text = cost.ToString("0")             '[euro]
        '===== check ================
        TextBox6.BackColor = CType(IIf(σm > σdesign, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button4Click(sender As Object, e As EventArgs) Handles Button4.Click, TabPage4.Enter, NumericUpDown9.ValueChanged, NumericUpDown11.ValueChanged
        'Round plate simply supported
        Dim dia, r, t As Double
        Dim p, σm, yt As Double
        Dim weight, cost As Double

        p = NumericUpDown1.Value * 100              '[mbar->[N/m2]]
        TextBox4.Text = p.ToString
        dia = NumericUpDown11.Value / 1000          '[m]
        r = dia / 2 '[m]
        t = NumericUpDown9.Value / 1000             '[m]
        weight = Math.PI / 4 * dia ^ 2 * t * 7850
        cost = weight * NumericUpDown33.Value

        σm = (1.238 * p * r ^ 2) / t ^ 2
        σm /= 10 ^ 6                                '[N/mm2]

        yt = (0.696 * p * r ^ 4) / (youngNm2 * t ^ 3)
        yt *= 1000                                  '[m]--->[mm]

        TextBox113.Text = p.ToString("F0")          '[N/m2]
        TextBox114.Text = (p / 100).ToString("F0")  '[mbar]
        TextBox8.Text = σm.ToString("F0")           'σm [N/m2] bens stress
        TextBox7.Text = yt.ToString("F1")           '[mm2]
        TextBox78.Text = weight.ToString("F0")      '[kg]
        TextBox79.Text = cost.ToString("F0")
        TextBox118.Text = younggpa.ToString("F0")    '[Gpa] 
        TextBox120.Text = NumericUpDown5.Value.ToString '[c]
        '===== check ================
        TextBox8.BackColor = CType(IIf(σm > σdesign, Color.Red, Color.LightGreen), Color)
    End Sub
    Private Sub Button5Click(sender As Object, e As EventArgs) Handles Button5.Click, TabPage6.Enter, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged
        'Round with hole
        Dim dia, diahole As Double
        Dim a, b, t As Double
        Dim p, σm, ym As Double
        Dim x, k1, k2 As Double
        Dim wght, cost As Double

        p = NumericUpDown1.Value * 100          '[mbar->[N/m2]]
        dia = NumericUpDown17.Value / 1000      '[meter]
        diahole = NumericUpDown16.Value / 1000 '[meter]
        t = NumericUpDown15.Value / 1000        '[meter]
        a = dia / 2
        b = diahole / 2

        '============= determine k1, k2 =================
        x = a / b

        k1 = 0.0067 * x ^ 4 - 0.0584 * x ^ 3 + 0.0519 * x ^ 2 + 0.6132 * x - 0.4358
        k2 = 0.0127 * x ^ 4 - 0.131 * x ^ 3 + 0.3117 * x ^ 2 + 0.6069 * x - 0.219

        If x > 5 Then k1 = 0.815
        If x > 5 Then k2 = 2.2

        'MessageBox.Show("a= " & a.ToString & " b= " & b.ToString & " t=" & t.ToString & " e=" & ee.ToString)
        σm = k2 * p * a ^ 2 / t ^ 2
        σm /= 10 ^ 6                        '[N/mm2]

        ym = k1 * p * a ^ 4
        ym /= youngNm2 * t ^ 3
        ym *= 1000                          '[mm]
        wght = PI * dia ^ 2 * t * 7850      '[kg]
        cost = wght * NumericUpDown33.Value '[Euro]

        TextBox22.Text = x.ToString("0.0")
        TextBox20.Text = (a * 1000).ToString("0")
        TextBox21.Text = (b * 1000).ToString("0")
        TextBox12.Text = σm.ToString("0")
        TextBox17.Text = k1.ToString("0.000")
        TextBox16.Text = k2.ToString("0.000")
        TextBox11.Text = ym.ToString("0.0")
        TextBox80.Text = wght.ToString("0")
        TextBox85.Text = cost.ToString("0")
        '===== check ================
        TextBox12.BackColor = CType(IIf(σm > σdesign, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button3Click(sender As Object, e As EventArgs) Handles Button3.Click, TabPage5.Enter, NumericUpDown18.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown24.ValueChanged
        Calcbeam()
    End Sub
    Private Sub Calcbeam()
        'http://beamguru.com/online/beam-calculator/
        'https://www.amesweb.info/StructuralBeamDeflection/SimplySupportedBeamStressDeflectionAnalysis.aspx
        Dim l As Double                     'Beam length
        Dim Iy, w As Double                 '[mm]
        Dim flex, mommax, σbend As Double
        Dim p, loadwidth, area As Double
        Dim beamheight, beamsectionarea As Double
        Dim σtension, tension As Double
        Dim σcombi As Double
        Dim weight, cost As Double

        l = NumericUpDown18.Value * 10 ^ 3              '[m->mm]
        Double.TryParse(TextBox41.Text, Iy)
        Iy *= 10 ^ 4     '[mm4]

        '===== Distributed load ============
        p = NumericUpDown1.Value * 10 ^ 2 * 10 ^ 6      '[mbar->N/mm2]
        loadwidth = NumericUpDown19.Value * 10 ^ 3     '[mm]
        area = loadwidth                               '[mm2]
        w = p * area                                    '[N/mm]

        '===== Flex, moment and stress ============
        mommax = (w * l ^ 2) / 8
        flex = 5 * w * l ^ 4 / (384 * youngNmm2 * Iy)
        Double.TryParse(TextBox18.Text, beamheight)
        σbend = mommax * beamheight * 0.5 / Iy
        σbend /= 10 ^ 12                             '[N/mm2]

        '===== Tension σ ============================
        Double.TryParse(TextBox48.Text, beamsectionarea)
        tension = NumericUpDown24.Value * 10 ^ 3     '[N]
        σtension = tension / beamsectionarea      '[N/mm2]

        '========= Combines stress =============
        σcombi = σbend + σtension

        '========= Cost =============
        Double.TryParse(TextBox48.Text, weight)
        cost = weight * NumericUpDown33.Value

        '===== Present =
        TextBox15.Text = (w / 10 ^ 12).ToString("0.00")     '[kN/m]
        TextBox10.Text = (flex / 10 ^ 12).ToString("0.0")   '[mm] flex
        TextBox13.Text = (mommax / 10 ^ 18).ToString("0.0") '[Nm]
        TextBox14.Text = σbend.ToString("0") '[N/mm2]
        TextBox23.Text = NumericUpDown1.Value.ToString      '[mbar]
        TextBox44.Text = σbend.ToString("0.0")             '[N/mm2]
        TextBox45.Text = σtension.ToString("0.0")          '[N/mm2]
        TextBox46.Text = σcombi.ToString("0.0")            '[N/mm2]
        TextBox82.Text = cost.ToString("0")                 '[Euro]
        '===== check ================
        TextBox14.BackColor = CType(IIf(σbend > σdesign, Color.Red, Color.LightGreen), Color)
        TextBox46.BackColor = CType(IIf(σcombi > σdesign, Color.Red, Color.LightGreen), Color)

    End Sub
    Private Sub ComboBox1SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim area As Double
        Dim weightm As Double
        Dim weighttotal As Double
        Dim Beamlength As Double

        Try
            Beamlength = NumericUpDown19.Value
            Dim words() As String = UNP(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            TextBox41.Text = words(1)               'Inertia Iy [cm4]
            TextBox18.Text = words(2)               'Beam Height [mm]
            TextBox47.Text = words(3)               'Beam weight [mm]

            weightm = CDbl(words(3))
            area = Math.Round(weightm * 10 ^ 6 / ρsteel, 0)
            weighttotal = weightm * Beamlength
            TextBox81.Text = area.ToString              'Beam area
            TextBox48.Text = weighttotal.ToString("0") 'Beam weight
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
        Calcbeam()
    End Sub

    Private Sub NumericUpDown1ValueChanged(sender As Object, e As EventArgs)
        Calcinput()
    End Sub

    Private Sub Button6Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown20.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, RadioButton1.CheckedChanged, NumericUpDown21.ValueChanged
        Dim moi As Double           'Moment of Inertia
        Dim wi, le, th As Double    '[N/mm2]
        Dim maxd As Double         'Max deflection
        Dim ω As Double             'Uniformly distrib load
        Dim pload As Double        'Point load [N]
        Dim wght, cost As Double


        le = NumericUpDown12.Value      '[mm] length
        wi = NumericUpDown13.Value      '[mm] width
        th = NumericUpDown20.Value      '[mm] thchick
        pload = NumericUpDown21.Value  '[N]
        moi = wi * th ^ 3 / 12          'Moment of Inertia

        '---------- Select load type -------------
        If RadioButton1.Checked Then    'Point 
            PictureBox8.Visible = False 'Distr. load
            PictureBox9.Visible = True  'Point load
            GroupBox14.Visible = False
            GroupBox17.Visible = True
            maxd = pload * le ^ 3 / (3 * youngNmm2 * moi)

        Else                            'Distributed load
            PictureBox8.Visible = True  'Distr. load
            PictureBox9.Visible = False 'Point load
            GroupBox14.Visible = True
            GroupBox17.Visible = False

            ω = wi * th / 10 ^ 9 * 7800 * 10            '[N/mm]
            maxd = ω * le ^ 4 / (8 * youngNmm2 * moi)
        End If
        wght = le * wi * th * 7850 / 10 ^ 9     '[kg]
        cost = wght * NumericUpDown33.Value     '[Euro]

        TextBox26.Text = moi.ToString("0")      'Moment of Inertia
        TextBox27.Text = maxd.ToString("0.00") 'Max deflection
        TextBox28.Text = ω.ToString("0.00")     'Distrib. Load
        TextBox83.Text = wght.ToString("0")     'Weight
        TextBox88.Text = cost.ToString("0")     'Cost
    End Sub

    Private Sub Button7Click(sender As Object, e As EventArgs) Handles Button7.Click, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown22.ValueChanged, ComboBox3.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, TabPage8.Enter
        'Calculate the Girders and Stiffeners one time
        If ComboBox2.SelectedIndex > -1 And ComboBox3.SelectedIndex > -1 Then
            Calcgrill(ComboBox2.SelectedIndex, ComboBox3.SelectedIndex)
        End If
    End Sub

    Private Sub Button9Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Auto select the Girders and Stiffeners
        '--- NOTE GIRDER is vertical
        Dim weight As Double = 10 ^ 8       'Init value
        Dim calcweight As Double           'Girder+Beam weight
        Dim bestgirder As Integer = 999    'Init value
        Dim beststif As Integer = 999      'Init value
        Dim Count As Integer = 0

        Button9.BackColor = Color.Red
        Button9.Text = "Auto select WAIT..."
        ProgressBar1.Visible = True
        ProgressBar1.Value = 0
        Me.Update()
        TextBox112.Clear()

        Array.Clear(grillresults, 0, grillresults.Length)

        If CheckBox1.Checked Then 'Identical beams (horizontal and vertical)
            For profile = 1 To (UNP.Length - 1)

                ProgressBar1.Value += 1
                If ProgressBar1.Value = 9999 Then ProgressBar1.Value = 1

                Calcgrill(profile, profile)
                Double.TryParse(TextBox37.Text, calcweight)
                Me.Update()

                If TextBox29.BackColor <> Color.Red And calcweight > 0 Then
                    grillresults(Count).girder = profile
                    grillresults(Count).beam = profile
                    grillresults(Count).weight = calcweight
                    Count += 1
                End If
            Next
        Else        'Different beams (horizontal and vertical)
            For vertgirder = 1 To (UNP.Length - 1)
                For horbeam = 0 To (UNP.Length - 1)

                    ProgressBar1.Value += 1
                    If ProgressBar1.Value = 9999 Then ProgressBar1.Value = 1

                    Calcgrill(vertgirder, horbeam)
                    Double.TryParse(TextBox37.Text, calcweight)
                    Me.Update()

                    If TextBox29.BackColor <> Color.Red And calcweight > 0 Then
                        grillresults(Count).girder = vertgirder
                        grillresults(Count).beam = horbeam
                        grillresults(Count).weight = calcweight
                        Count += 1
                    End If
                Next
            Next
        End If

        '====  SORT THE CALC RESULTS ==========
        Button9.Text = "SORTING....."
        Button9.BackColor = Color.LightBlue
        Me.Update()
        'https://social.msdn.microsoft.com/Forums/en-US/9af03200-58d6-4035-84e8-2554347bc25b/vbnet-2005-structure-arraysort-how?forum=vblanguage
        'https://stackoverflow.com/questions/1751768/sort-an-array-of-structures-in-net
        grillresults = grillresults.OrderBy(Function(c) c.weight).ToArray

        '===== LOG results ==========
        For i = 0 To grillresults.Length - 1
            If grillresults(i).weight > 0 Then
                TextBox112.Text &= "Girder name= " & Getprofilename(grillresults(i).girder) & vbTab & "  "
                TextBox112.Text &= "Beam name= " & Getprofilename(grillresults(i).beam) & vbTab
                TextBox112.Text &= "Weight= " & grillresults(i).weight & vbCrLf
            End If
        Next

        For i = 0 To grillresults.Length - 1
            If grillresults(i).weight > 0 Then
                Calcgrill(grillresults(i).girder, grillresults(i).beam)
                ComboBox2.SelectedIndex = grillresults(i).girder   'Girders (short)
                ComboBox3.SelectedIndex = grillresults(i).beam
                Exit For
            End If
        Next

        Button9.Text = "Auto select"
        Button9.BackColor = Color.Transparent
        ProgressBar1.Visible = False
    End Sub

    'Design of Ship Hull Structures ISBN: 978-3-642-10009-3
    'Chapter Grillage Structure Page 254 
    'Girders (short) most important, vertical, 
    'Beams horizontal against buckling
    'Girder and Beam are selected from the UNP table
    Private Sub Calcgrill(Girdervert As Integer, Beamhor As Integer)
        Dim press As Double             'Uniform load
        Dim ahor As Double             'Longer horizontal edge
        Dim bvert As Double            'Shorter vertical edge
        Dim Ihorbeam As Double        'Second Moment of inertia
        Dim Ivertgirder As Double     'Second Moment of inertia
        Dim novertgirders As Double   'No Girders vertical
        Dim nohorbeams As Double      'No Beams horintal
        Dim eygirder As Double         'Distance COG to plate face
        Dim spacebeams As Double       'Beams space 
        Dim spacegirders As Double     'Girder space 
        Dim σy As Double                'Midpoint Stiffener Stress
        Dim δ As Double                 'Midpoint (Max) deflection
        Dim force As Double             'Pressure x area
        Dim weight As Double            'Total weight
        Dim cost As Double              'Total cost
        Dim girvertwht As Double      'Girdervert 
        Dim beamhorwht As Double      'Beam vertical
        Dim wordshor() As String
        Dim wordsvert() As String
        Dim lopti As Double
        Dim δ1, δ2 As Double            'Temp value for readability
        Dim girdername, beamname As String

        ComboBox2.SelectedIndex = Girdervert   'Girders number (short)
        ComboBox3.SelectedIndex = Beamhor      'Beam number

        '---------- get data -------------
        Double.TryParse(TextBox43.Text, press)          '[N/mm2]
        TextBox33.Text = press.ToString                 '[N/mm2]
        TextBox40.Text = NumericUpDown1.Value.ToString  '[mbar]

        '--- NOTE GIRDER is vertical , Stiff-BEAM is horizontal------
        novertgirders = NumericUpDown29.Value     'no Girder
        nohorbeams = NumericUpDown28.Value        'no Beams
        ahor = NumericUpDown22.Value               'Longer edge
        bvert = NumericUpDown23.Value              'Shortes edge

        '--- Girders (short) VERTICAL
        wordsvert = UNP(Girdervert).Split(CType(";", Char()))
        girdername = wordsvert(0)
        Ivertgirder = CDbl(wordsvert(1)) * (10 ^ 4)      'Inertia Iy [cm^4->mm^4]
        girvertwht = CDbl(wordsvert(3))                  '[kg]
        TextBox38.Text = girvertwht.ToString
        TextBox32.Text = (Ivertgirder / 10 ^ 4).ToString     '[cm4]
        eygirder = CDbl(wordsvert(2)) - CDbl(wordsvert(4))
        TextBox42.Text = eygirder.ToString("0.0")          'Distance to plate face [mm]

        '--- Beams Horizontal
        wordshor = UNP(Beamhor).Split(CType(";", Char()))
        beamname = wordshor(0)                            'Beam horizontal
        Ihorbeam = CDbl(wordshor(1)) * (10 ^ 4)          'Inertia Iy [cm^4->nogird^4]
        beamhorwht = CDbl(wordshor(3))                   '[kg]
        TextBox39.Text = beamhorwht.ToString
        TextBox34.Text = (Ihorbeam / 10 ^ 4).ToString     '[cm4]

        '--------- calc girder spacing -------------
        spacebeams = ahor / (nohorbeams + 1)            'Beam space
        spacegirders = bvert / (novertgirders + 1)      'Girder space

        '------ deflection and stress @ midpoint--------
        '------ formula 6.1.1---------------------------

        force = ahor * bvert * press

        δ1 = Ihorbeam * (nohorbeams + 1) / ahor ^ 3
        δ2 = Ivertgirder * (novertgirders + 1) / bvert ^ 3

        δ = force
        δ /= PI ^ 6 * youngNmm2 / 16
        δ /= (δ1 + δ2)

        '----Stress @ midpoint--------
        '------ formula 6.1.1---------------------------
        σy = PI ^ 2 * δ * youngNmm2 * eygirder / bvert ^ 2

        '------ Girder weight ---------------------
        weight = novertgirders * bvert / 1000 * girvertwht     'Girder vertical [kg]
        weight += nohorbeams * ahor / 1000 * beamhorwht        'Beam horizontal [kg]

        '------ Optimum girder distance (formula 6.2.6)-----------
        '------ Both ends of the girder is supported--------------
        lopti = 0.72 * spacebeams ^ (1 / 3) * bvert ^ (2 / 3)

        '------ cost ---------
        cost = weight * NumericUpDown33.Value

        '------ present ----------
        TextBox29.Text = σy.ToString("0")       '[N/mm2]
        TextBox31.Text = δ.ToString("0.0")      '[mm]

        TextBox35.Text = spacegirders.ToString("0")    'Girder distance
        TextBox36.Text = spacebeams.ToString("0")      'Beam distance
        TextBox37.Text = weight.ToString("0.0")         '[kg]
        TextBox84.Text = cost.ToString("0")             '[Euro]

        Label148.Text = "If girders and beam are identical the Optimum Girder space= " & lopti.ToString("0") & " [mm]"

        '------ check -shorter/longer edge---------
        NumericUpDown22.BackColor = CType(IIf(ahor > bvert, Color.Yellow, Color.Red), Color)
        TextBox29.BackColor = CType(IIf(σy < σdesign, Color.LightGreen, Color.Red), Color)

        '------ logging if possible solution ---------------
        'If TextBox29.BackColor <> Color.Red Then
        '    TextBox112.Text &= "girdername= " & girdername & ", beamname= " & beamname & vbCrLf
        '    TextBox112.Text &= "σy = " & σy.ToString("0") & ", Weight= " & weight.ToString("0") & vbCrLf
        'End If
    End Sub
    'Stress and strain table 8.13 page 260
    Private Sub Button8Click(sender As Object, e As EventArgs) Handles Button8.Click, TabPage10.Enter, NumericUpDown25.ValueChanged, NumericUpDown14.ValueChanged, ComboBox4.SelectedIndexChanged
        Dim words() As String
        Dim Mp As Double            'Plastic moment
        Dim σ As Double             'max allowed stress [n/mm2]
        Dim lbeam As Double        'Beam length [mm]
        Dim pwidth As Double       'Pressure width [mm]
        Dim press As Double         '[N/mm2]
        Dim areabeam, Zplastic As Double 'Plasic modul [mm3]
        Dim Wa As Double            'Actual Uniform load [N/mm]
        Dim Wac As Double           'Collapse Uniform load [N/mm]

        '------------ get data -----------
        σ = σ02                           'max allowed stress [N/mm2]
        lbeam = NumericUpDown25.Value      'beam length [mm]
        pwidth = NumericUpDown14.Value     'pressure width [mm]

        If ComboBox4.SelectedIndex > -1 Then
            '--- Beams Horizontal
            words = UNP(ComboBox4.SelectedIndex).Split(CType(";", Char()))
            TextBox55.Text = words(1)  'Inertia Iy [cm^4]
            areabeam = Math.Round(CDbl(words(3)) * 10 ^ 6 / ρsteel, 0)
            TextBox56.Text = areabeam.ToString     'Area [cm^2]
            TextBox57.Text = words(3)               '[kg]
            TextBox58.Text = words(2)               '[mm]
            TextBox61.Text = words(5)               '[mm3]
            Zplastic = CDbl(words(5))              '[mm3]
            press = NumericUpDown1.Value / 10 ^ 4   '[mbar]-->[N/mm2]
        End If

        '-------------- calc -------------
        Wa = press * pwidth                    '[N/mm]
        Mp = σ * Zplastic                      '[Nm]
        Wac = 16 * Mp * lbeam ^ 2 / lbeam ^ 4 '[Nm]

        '-------------- Present ----------
        TextBox53.Text = press.ToString("F3")               '[N/mm2] pressure
        TextBox73.Text = (press * 10 ^ 4).ToString("F0")    '[mbar]
        TextBox54.Text = (Mp / 1000).ToString("F1")         '[kN/mm2]
        TextBox62.Text = Wa.ToString("F1")                  '[N/mm] 
        TextBox63.Text = Wac.ToString("F1")                 '[N/mm]
        '-------------- Checks --------
        TextBox62.BackColor = CType(IIf(Wa < Wac, Color.LightGreen, Color.Red), Color)
    End Sub

    Private Sub Button10Click(sender As Object, e As EventArgs) Handles Button10.Click, NumericUpDown32.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown27.ValueChanged, NumericUpDown26.ValueChanged, TabPage12.Enter
        'https://nl.wikipedia.org/wiki/Weerstandsmoment
        Dim force As Double
        Dim torquebend As Double
        Dim plate As Double     'Plate thickness
        Dim wy As Double        'weerstand moment
        Dim H, B1, B2 As Double
        Dim σb As Double        'Bending stress
        Dim σt As Double        'Tensile stress
        Dim σc As Double        'Tensile stress

        force = NumericUpDown26.Value * 10 ^ 3  '[N] Force at tip
        plate = NumericUpDown27.Value           '[mm]
        H = NumericUpDown30.Value               '[mm]
        B1 = NumericUpDown31.Value              '[mm]
        B2 = NumericUpDown32.Value              '[mm]

        '----------- bending stress ----------
        torquebend = force * (B1 - B2 / 2)     '[N.mm]
        wy = plate * B2 ^ 2 / 6                 '[mm3]
        σb = torquebend / wy                   '[N/mm2]

        '----------- tensile stress -----------
        σt = force / (plate * B2)               '[N/mm2]

        '----------- combined stress -----------
        σc = σt + σb                            '[N/mm2]

        TextBox77.Text = wy.ToString("0")                   '[mm3] 
        TextBox75.Text = (torquebend / 1000).ToString("0") '[N.mm] 
        TextBox76.Text = σc.ToString("0")                   '[N/mm] 
        TextBox74.Text = (force / 10).ToString("0")         '[kg] 

        '-------------- Checks --------
        TextBox76.BackColor = CType(IIf(σb < σdesign, Color.LightGreen, Color.Red), Color)
    End Sub

    Private Sub Button11Click(sender As Object, e As EventArgs) Handles Button11.Click, TabPage13.Enter, NumericUpDown39.ValueChanged, NumericUpDown38.ValueChanged, NumericUpDown37.ValueChanged, NumericUpDown36.ValueChanged, NumericUpDown35.ValueChanged, NumericUpDown34.ValueChanged, NumericUpDown40.ValueChanged, NumericUpDown42.ValueChanged, NumericUpDown41.ValueChanged
        Dim press As Double
        Dim presswidth As Double
        Dim load As Double
        Dim platet As Double
        Dim mvert, mhor As Double
        Dim shvert, shhor As Double   'Shear forces
        Dim s1, s2, s3, s4 As Double    'width
        Dim i1, i2, i3, i4 As Double    'area moment of inertia
        Dim w1, w2, w3, w4 As Double    'Resistance Moment
        Dim b, h As Double
        Dim σ1, σ2, σ3, σ4 As Double    'Bending stress
        Dim τ1, τ2, τ3, τ4 As Double    'Shear stress
        Dim σt1, σt2, σt3, σt4 As Double 'Combined stress


        press = NumericUpDown34.Value * 10 ^ 2      '[mbar]->[Pa] 
        presswidth = NumericUpDown40.Value         '[mm] 

        s1 = NumericUpDown35.Value         '[mm] width
        s2 = NumericUpDown39.Value         '[mm] width
        s3 = NumericUpDown42.Value         '[mm] width
        s4 = NumericUpDown41.Value         '[mm] width
        b = NumericUpDown36.Value          '[mm] width
        h = NumericUpDown37.Value          '[mm] width

        platet = NumericUpDown38.Value    '[mm] plate thicknes

        '-------- evenly distributed load --------
        load = press * presswidth / 1000  '[N/mm] evenly distributed load

        '-------- Moments both ends clamped -----
        mhor = load * (h - s2 - s4) ^ 2 / 24       '[Nmm]
        mvert = load * (b - s1 - s3) ^ 2 / 24      '[Nmm]

        '-------- Shear forces ------------------
        shvert = load * (b - s1 - s3)         '[N]
        shhor = load * (h - s2 - s4)          '[N]

        '---- https://en.wikipedia.org/wiki/Bending ----
        '---- https://nl.wikipedia.org/wiki/Vergeet-mij-nietje(mechanica)
        '---- https://en.wikipedia.org/wiki/Listofsecondmomentsofarea --

        '--- area moment of inertia
        i1 = platet * s1 ^ 3 / 12
        i2 = platet * s2 ^ 3 / 12
        i3 = platet * s3 ^ 3 / 12
        i4 = platet * s4 ^ 3 / 12

        '---- 'Resistance Moment about the neutral axis ---
        w1 = i1 / (s1 * 0.5)
        w2 = i2 / (s2 * 0.5)
        w3 = i3 / (s3 * 0.5)
        w4 = i4 / (s4 * 0.5)

        '---- Bending stress -------------
        σ1 = mhor / w1
        σ3 = mhor / w3
        σ2 = mvert / w2
        σ4 = mvert / w4

        '---- Shear stress -------------
        τ1 = shhor * 0.5 / (platet * s1)
        τ3 = shhor * 0.5 / (platet * s3)
        τ2 = shvert * 0.5 / (platet * s2)
        τ4 = shvert * 0.5 / (platet * s4)

        '---- Combined stress -----------
        σt1 = σ1 + τ1
        σt3 = σ3 + τ3
        σt2 = σ2 + τ2
        σt4 = σ4 + τ4

        '---------- present -------------
        TextBox89.Text = load.ToString("0")         '[N/m2]

        TextBox91.Text = (mvert / 10 ^ 3).ToString("0")       '[N/m] 
        TextBox92.Text = (mhor / 10 ^ 3).ToString("0")        '[N/m] 
        TextBox93.Text = (mvert / 10 ^ 3).ToString("0")       '[N/m] 
        TextBox94.Text = (mhor / 10 ^ 3).ToString("0")        '[N/m] 

        TextBox90.Text = i1.ToString("0")           '[mm4] 
        TextBox99.Text = i2.ToString("0")           '[mm4] 
        TextBox100.Text = i3.ToString("0")          '[mm4] 
        TextBox101.Text = i4.ToString("0")          '[mm4] 

        TextBox95.Text = σ1.ToString("0")           '[N/mm2] 
        TextBox96.Text = σ2.ToString("0")           '[N/mm2] 
        TextBox97.Text = σ3.ToString("0")           '[N/mm2] 
        TextBox98.Text = σ4.ToString("0")           '[N/mm2] 

        TextBox102.Text = shhor.ToString("0")      '[N] 
        TextBox103.Text = shvert.ToString("0")     '[N] 

        TextBox107.Text = τ1.ToString("0")          '[N/mm2] 
        TextBox106.Text = τ2.ToString("0")          '[N/mm2] 
        TextBox105.Text = τ3.ToString("0")          '[N/mm2] 
        TextBox104.Text = τ4.ToString("0")          '[N/mm2] 

        TextBox111.Text = σt1.ToString("0")         '[N/mm2] 
        TextBox110.Text = σt2.ToString("0")         '[N/mm2] 
        TextBox109.Text = σt3.ToString("0")         '[N/mm2] 
        TextBox108.Text = σt4.ToString("0")         '[N/mm2] 

        '-------------- Checks --------
        TextBox111.BackColor = CType(IIf(σt1 < σdesign, Color.LightGreen, Color.Red), Color)
        TextBox110.BackColor = CType(IIf(σt2 < σdesign, Color.LightGreen, Color.Red), Color)
        TextBox109.BackColor = CType(IIf(σt3 < σdesign, Color.LightGreen, Color.Red), Color)
        TextBox108.BackColor = CType(IIf(σt4 < σdesign, Color.LightGreen, Color.Red), Color)
    End Sub

    Private Sub Button12Click(sender As Object, e As EventArgs) Handles Button12.Click, NumericUpDown5.ValueChanged, NumericUpDown33.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown43.ValueChanged
        Calcinput()
    End Sub

    Private Sub Button13Click(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox112.Clear()
    End Sub
    Private Function Getprofilename(no As Integer) As String
        Dim wwords() As String
        wwords = UNP(no).Split(CType(";", Char()))
        Return (wwords(0))
    End Function
End Class
