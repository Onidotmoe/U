Imports Microsoft.Xna.Framework

Public Module ColorManipulation
#Region "Color Converters"
    Public Function Invert(Color As System.Drawing.Color, Optional DoAlpha As Boolean = False) As System.Drawing.Color
        If (Not DoAlpha) Then
            Return System.Drawing.Color.FromArgb((255 - System.Convert.ToInt32(Color.R)), (255 - System.Convert.ToInt32(Color.G)), (255 - System.Convert.ToInt32(Color.B)))
        Else
            Return System.Drawing.Color.FromArgb((255 - System.Convert.ToInt32(Color.R)), (255 - System.Convert.ToInt32(Color.G)), (255 - System.Convert.ToInt32(Color.B)), (255 - System.Convert.ToInt32(Color.A)))
        End If
    End Function

    Public Function HSBToColor(HSB As HSB) As Color
        Return RGBToColor(HSBToRGB(HSB))
    End Function

    Public Function LABToColor(LAB As LAB) As Color
        Return RGBToColor(LABToRGB(LAB))
    End Function

    Public Function LABToRGB(LAB As LAB) As RGB
        Return XYZToRGB(LABtoXYZ(LAB))
    End Function

    Public Function LABToHSB(LAB As LAB) As HSB
        Return RGBToHSB(LABToRGB(LAB))
    End Function

    Public Function RGBToLAB(RGB As RGB) As LAB
        Return XYZToLAB(RGBToXYZ(RGB))
    End Function

    Public Function LABtoXYZ(LAB As LAB) As XYZ
        Dim X, Y, Z As New Double
        Y = ((LAB.L + 16.0) / 116.0)
        X = ((LAB.A / 500.0) + Y)
        Z = (Y - (LAB.B / 200.0))

        Dim Pow_X = Math.Pow(X, 3.0)
        Dim Pow_Y = Math.Pow(Y, 3.0)
        Dim Pow_Z = Math.Pow(Z, 3.0)

        Dim Less = (216 / 24389)

        If (Pow_X > Less) Then
            X = Pow_X
        Else
            X = ((X - (16.0 / 116.0)) / 7.787)
        End If
        If (Pow_Y > Less) Then
            Y = Pow_Y
        Else
            Y = ((Y - (16.0 / 116.0)) / 7.787)
        End If
        If (Pow_Z > Less) Then
            Z = Pow_Z
        Else
            Z = ((Z - (16.0 / 116.0)) / 7.787)
        End If

        Return New XYZ((X * 95.047), (Y * 100.0), (Z * 108.883))
    End Function

    Public Function XYZToRGB(XYZ As XYZ) As RGB
        Dim X, Y, Z As New Double
        Dim R, G, B As New Double
        Dim Pow As Double = (1.0 / 2.4)
        Dim Less As Double = 0.0031308

        X = (XYZ.X / 100)
        Y = (XYZ.Y / 100)
        Z = (XYZ.Z / 100)

        R = ((X * 3.24071) + (Y * -1.53726) + (Z * -0.498571))
        G = ((X * -0.969258) + (Y * 1.87599) + (Z * 0.0415557))
        B = ((X * 0.0556352) + (Y * -0.203996) + (Z * 1.05707))

        If (R > Less) Then
            R = ((1.055 * Math.Pow(R, Pow)) - 0.055)
        Else
            R *= 12.92
        End If
        If (G > Less) Then
            G = ((1.055 * Math.Pow(G, Pow)) - 0.055)
        Else
            G *= 12.92
        End If
        If (B > Less) Then
            B = ((1.055 * Math.Pow(B, Pow)) - 0.055)
        Else
            B *= 12.92
        End If

        Return New RGB((R * 255), (G * 255), (B * 255))
    End Function

    Public Function RGBToXYZ(RGB As RGB) As XYZ
        Dim X, Y, Z As New Double
        Dim R, G, B As New Double
        Dim Less As Double = 0.04045

        R = (RGB.R / 255)
        G = (RGB.G / 255)
        B = (RGB.B / 255)

        If (R > Less) Then
            R = Math.Pow(((R + 0.055) / 1.055), 2.4)
        Else
            R = (R / 12.92)
        End If
        If (G > Less) Then
            G = Math.Pow(((G + 0.055) / 1.055), 2.4)
        Else
            G = (G / 12.92)
        End If
        If (B > Less) Then
            B = Math.Pow(((B + 0.055) / 1.055), 2.4)
        Else
            B = (B / 12.92)
        End If

        X = ((R * 0.4124) + (G * 0.3576) + (B * 0.1805))
        Y = ((R * 0.2126) + (G * 0.7152) + (B * 0.0722))
        Z = ((R * 0.0193) + (G * 0.1192) + (B * 0.9505))

        Return New XYZ(X * 100, Y * 100, Z * 100)
    End Function

    Public Function XYZToLAB(XYZ As XYZ) As LAB
        Dim X, Y, Z As New Double
        Dim L, A, B As New Double
        Dim Less As Double = 0.008856
        Dim Pow As Double = (1.0 / 3.0)

        X = ((XYZ.X / 100) / 0.9505)
        Y = (XYZ.Y / 100)
        Z = ((XYZ.Z / 100) / 1.089)

        If (X > Less) Then
            X = Math.Pow(X, Pow)
        Else
            X = ((7.787 * X) + (16.0 / 116.0))
        End If
        If (Y > Less) Then
            Y = Math.Pow(Y, Pow)
        Else
            Y = ((7.787 * Y) + (16.0 / 116.0))
        End If
        If (Z > Less) Then
            Z = Math.Pow(Z, Pow)
        Else
            Z = ((7.787 * Z) + (16.0 / 116.0))
        End If

        L = ((116.0 * Y) - 16.0)
        A = (500.0 * (X - Y))
        B = (200.0 * (Y - Z))

        Return New LAB(CInt(L), CInt(A), CInt(B))
    End Function

    Public Function HSBToRGB(HSB As HSB) As RGB
        Select Case 0.0
            Case HSB.S
                Return New RGB((HSB.B * 255), (HSB.B * 255), (HSB.B * 255))
            Case HSB.B
                Return New RGB(0, 0, 0)
            Case Else
                Dim R, G, B As New Double
                Dim H2 As Double = (HSB.H * 6.0)
                Dim H2Floor As Integer = CInt(Math.Floor(H2))
                Dim [Mod] As Double = (H2 - H2Floor)
                Dim v1 As Double = (HSB.B * (1.0 - HSB.S))
                Dim v2 As Double = (HSB.B * (1.0 - (HSB.S * [Mod])))
                Dim v3 As Double = (HSB.B * (1.0 - (HSB.S * (1.0 - [Mod]))))

                Select Case (H2Floor + 1)
                    Case 0
                        R = HSB.B
                        G = v1
                        B = v2
                    Case 1
                        R = HSB.B
                        G = v3
                        B = v1
                    Case 2
                        R = v2
                        G = HSB.B
                        B = v1
                    Case 3
                        R = v1
                        G = HSB.B
                        B = v3
                    Case 4
                        R = v1
                        G = v2
                        B = HSB.B
                    Case 5
                        R = v3
                        G = v1
                        B = HSB.B
                    Case 6
                        R = HSB.B
                        G = v1
                        B = v2
                    Case 7
                        R = HSB.B
                        G = v3
                        B = v1
                End Select

                Return New RGB((R * 255), (G * 255), (B * 255))
        End Select
    End Function

    Public Function RGBToHSB(RGB As RGB) As HSB
        Dim R, G, B As New Double
        Dim H, S, V As New Double
        Dim Max, Min, Delta As New Double

        R = (RGB.R / 255)
        G = (RGB.G / 255)
        B = (RGB.B / 255)

        Max = Math.Max(R, Math.Max(G, B))
        Min = Math.Min(R, Math.Min(G, B))
        Delta = (Max - Min)
        V = Max

        If (Delta = 0) Then
            H = 0
            S = 0
        Else
            S = (Delta / Max)

            Dim Delta_R, Delta_G, Delta_B As New Double
            Delta_R = ((((Max - R) / 6) + (Max / 2)) / Delta)
            Delta_G = ((((Max - G) / 6) + (Max / 2)) / Delta)
            Delta_B = ((((Max - B) / 6) + (Max / 2)) / Delta)

            Select Case Max
                Case R
                    H = (Delta_B - Delta_G)
                Case G
                    H = ((1 / 3) + (Delta_R - Delta_B))
                Case B
                    H = ((2 / 3) + (Delta_G - Delta_R))
            End Select

            Select Case H
                Case < 0
                    H += 1
                Case > 1
                    H -= 1
            End Select
        End If

        Return New HSB(H, S, V)
    End Function

    Public Function RGBToCMYK(RGB As RGB) As CMYK
        Dim R As Double = (RGB.R / 255)
        Dim G As Double = (RGB.G / 255)
        Dim B As Double = (RGB.B / 255)

        Dim K As Double = (1 - Math.Max(Math.Max(R, G), B))
        Dim C As Double = (1 - R - K) / (1 - K)
        Dim M As Double = (1 - G - K) / (1 - K)
        Dim Y As Double = (1 - B - K) / (1 - K)

        Return New CMYK(C, M, Y, K)
    End Function

    Public Function CMYKToHSB(CMYK As CMYK) As HSB
        Return RGBToHSB(CMYKToRGB(CMYK))
    End Function

    Public Function CMYKToRGB(CMYK As CMYK) As RGB
        Dim R As Double = ((1 - CMYK.C) * (1 - CMYK.K))
        Dim G As Double = ((1 - CMYK.M) * (1 - CMYK.K))
        Dim B As Double = ((1 - CMYK.Y) * (1 - CMYK.K))

        Return New RGB(R, G, B)
    End Function

    Public Function CMYKToColor(CMYK As CMYK) As Color
        Return RGBToColor(CMYKToRGB(CMYK))
    End Function

    Public Function GetNearestWebSafeColor(Color As Color) As Color
        Return New Color(CInt(Math.Round((Convert.ToInt32(Color.R) / 255) * 5) * 51), CInt(Math.Round((Convert.ToInt32(Color.G) / 255) * 5) * 51), CInt(Math.Round((Convert.ToInt32(Color.B) / 255) * 5) * 51))
    End Function

    Public Function GetNearestWebSafeColor(RGB As RGB) As RGB
        Return New RGB(CInt(Math.Round((RGB.R / 255) * 5) * 51), CInt(Math.Round((RGB.G / 255) * 5) * 51), CInt((Math.Round(RGB.B / 255) * 5) * 51))
    End Function

    Public Function XNAColorToSystemColor(Color As Color) As System.Drawing.Color
        Return System.Drawing.Color.FromArgb(Convert.ToInt32(Color.A), Convert.ToInt32(Color.R), Convert.ToInt32(Color.G), Convert.ToInt32(Color.B))
    End Function

    Public Function RGBToColor(RGB As RGB) As Color
        Return New Color(CInt(RGB.R), CInt(RGB.G), CInt(RGB.B), CInt(RGB.A))
    End Function

    Function ColorToRGB(Color As Color) As RGB
        Return New RGB(Convert.ToInt32(Color.R), Convert.ToInt32(Color.G), Convert.ToInt32(Color.B), Convert.ToInt32(Color.A))
    End Function

    Function ColorToRGB(R As Double, G As Double, B As Double) As RGB
        Return New RGB(R, G, B)
    End Function

    Function ColorToRGB(R As Double, G As Double, B As Double, A As Double) As RGB
        Return New RGB(R, G, B, A)
    End Function

    Public Function ToStringHexData(RGB As RGB, Optional _Alpha As Integer = 255) As String
        Dim Color = RGBToColor(RGB)
        Dim Red As String = Convert.ToString(Color.R, 16)
        Dim Green As String = Convert.ToString(Color.G, 16)
        Dim Blue As String = Convert.ToString(Color.B, 16)
        Dim Alpha As String = Convert.ToString(Convert.ToByte(_Alpha), 16)

        If (Red.Length < 2) Then
            Red = ("0" & Red)
        End If
        If (Green.Length < 2) Then
            Green = ("0" & Green)
        End If
        If (Blue.Length < 2) Then
            Blue = ("0" & Blue)
        End If
        If (Alpha.Length < 2) Then
            Alpha = ("0" & Alpha)
        End If

        Return (Alpha.ToUpper() & Red.ToUpper() & Green.ToUpper() & Blue.ToUpper())
    End Function

    Public Function ParseStringHexData(Hex_Data As String, ByRef Alpha As Integer) As RGB
        Hex_Data = "00000000" & Hex_Data
        Hex_Data = Hex_Data.Remove(0, Hex_Data.Length - 8)
        Dim R_Text, G_Text, B_Text, A_Text As String
        Dim R, G, B, A As Integer
        A_Text = Hex_Data.Substring(0, 2)
        R_Text = Hex_Data.Substring(2, 2)
        G_Text = Hex_Data.Substring(4, 2)
        B_Text = Hex_Data.Substring(6, 2)
        R = Integer.Parse(R_Text, Globalization.NumberStyles.HexNumber)
        G = Integer.Parse(G_Text, Globalization.NumberStyles.HexNumber)
        B = Integer.Parse(B_Text, Globalization.NumberStyles.HexNumber)
        A = Integer.Parse(A_Text, Globalization.NumberStyles.HexNumber)

        Alpha = A
        Return New RGB(R, G, B)
    End Function
#End Region

#Region "Color Classes"

    ''' <summary>
    ''' <para>The Default XNA Color class has a problematic behavior with fractions being considered % of the max 255 color value.</para>
    ''' <para>Which means that any decimal number is considered a fraction even when it is larger than 1.</para>
    ''' <para>This class solves those problems and maintains precision that might otherwise be lost in conversion.</para>
    ''' </summary>
    Class RGB
        Public ReadOnly Min As Double = 0.0
        Public ReadOnly Max As Double = 255.0

        Public Sub New()
        End Sub

        Public Sub New(R As Double, G As Double, B As Double)
            Me.R = R
            Me.G = G
            Me.B = B
        End Sub

        Public Sub New(R As Double, G As Double, B As Double, A As Double)
            Me.R = R
            Me.G = G
            Me.B = B
            Me.A = A
        End Sub

        Public Sub New(Color As Color)
            Me.R = Convert.ToInt32(Color.R)
            Me.G = Convert.ToInt32(Color.G)
            Me.B = Convert.ToInt32(Color.B)
            Me.A = Convert.ToInt32(Color.A)
        End Sub

        Private _R As New Double
        Private _G As New Double
        Private _B As New Double
        Private _A As Double = Max

        Public Property R As Double
            Get
                Return _R
            End Get
            Set
                _R = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Public Property G As Double
            Get
                Return _G
            End Get
            Set
                _G = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Public Property B As Double
            Get
                Return _B
            End Get
            Set
                _B = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Public Property A As Double
            Get
                Return _A
            End Get
            Set
                _A = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Overrides Function ToString() As String
            Return (_R.ToString & ":"c & _G.ToString & ":"c & _B.ToString & ":"c & _A.ToString)
        End Function

        Public Shared Operator =(Left As RGB, Right As RGB) As Boolean
            If ((Left.R = Right.R) AndAlso (Left.G = Right.G) AndAlso (Left.B = Right.B) AndAlso (Left.A = Right.A)) Then
                Return True
            Else
                Return False
            End If
        End Operator

        Public Shared Operator <>(Left As RGB, Right As RGB) As Boolean
            Return (Not (Left = Right))
        End Operator

    End Class

    Public Class XYZ
        Public ReadOnly Min As Double = 0

        Public Sub New()
        End Sub

        Public Sub New(X As Double, Y As Double, Z As Double)
            Me.X = X
            Me.Y = Y
            Me.Z = Z
        End Sub

        Private _X As New Double
        Private _Y As New Double
        Private _Z As New Double

        Public Property X As Double
            Get
                Return _X
            End Get
            Set
                _X = LimitInRange(Value, Min, 95.05)
            End Set
        End Property

        Public Property Y As Double
            Get
                Return _Y
            End Get
            Set
                _Y = LimitInRange(Value, Min, 100)
            End Set
        End Property

        Public Property Z As Double
            Get
                Return _Z
            End Get
            Set
                _Z = LimitInRange(Value, Min, 108.9)
            End Set
        End Property

        Overrides Function ToString() As String
            Return (_X.ToString & ":"c & _Y.ToString & ":"c & _Z.ToString)
        End Function

    End Class

    Public Class LAB
        Public ReadOnly Min As Double = -128
        Public ReadOnly Max As Double = 127

        Sub New()
        End Sub

        Sub New(L As Double, A As Double, B As Double)
            Me.L = L
            Me.A = A
            Me.B = B
        End Sub

        Private _L As New Double
        Private _A As New Double
        Private _B As New Double

        Property L As Double
            Get
                Return _L
            End Get
            Set
                _L = LimitInRange(Value, 0, 100)
            End Set
        End Property

        Property A As Double
            Get
                Return _A
            End Get
            Set
                _A = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Property B As Double
            Get
                Return _B
            End Get
            Set
                _B = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Overrides Function ToString() As String
            Return (_L.ToString & ":"c & _A.ToString & ":"c & _B.ToString)
        End Function

    End Class

    Public Class HSB
        Public ReadOnly Min As Double = 0
        Public ReadOnly Max As Double = 1

        Public Sub New()
        End Sub

        Public Sub New(H As Double, S As Double, B As Double)
            Me.H = H
            Me.S = S
            Me.B = B
        End Sub

        Private _H As New Double
        Private _S As New Double
        Private _B As New Double

        Public Property H As Double
            Get
                Return _H
            End Get
            Set
                _H = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Public Property S As Double
            Get
                Return _S
            End Get
            Set
                _S = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Public Property B As Double
            Get
                Return _B
            End Get
            Set
                _B = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Overrides Function ToString() As String
            Return (_H.ToString & ":"c & _S.ToString & ":"c & _B.ToString)
        End Function

        Public Shared Operator =(Left As HSB, Right As HSB) As Boolean
            If ((Left.H = Right.H) AndAlso (Left.S = Right.S) AndAlso (Left.B = Right.B)) Then
                Return True
            Else
                Return False
            End If
        End Operator

        Public Shared Operator <>(Left As HSB, Right As HSB) As Boolean
            Return (Not (Left = Right))
        End Operator

    End Class

    Public Class CMYK
        Public ReadOnly Min As Double = 0
        Public ReadOnly Max As Double = 1

        Public Sub New()
        End Sub

        Public Sub New(C As Double, M As Double, Y As Double, K As Double)
            Me.C = C
            Me.M = M
            Me.Y = Y
            Me.K = K
        End Sub

        Private _C As New Double
        Private _M As New Double
        Private _Y As New Double
        Private _K As New Double

        Public Property C As Double
            Get
                Return _C
            End Get
            Set
                _C = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Public Property M As Double
            Get
                Return _M
            End Get
            Set
                _M = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Public Property Y As Double
            Get
                Return _Y
            End Get
            Set
                _Y = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Public Property K As Double
            Get
                Return _K
            End Get
            Set
                _K = LimitInRange(Value, Min, Max)
            End Set
        End Property

        Overrides Function ToString() As String
            Return (_C.ToString & ":"c & _M.ToString & ":"c & _Y.ToString & ":"c & _K.ToString)
        End Function

    End Class

#End Region

End Module