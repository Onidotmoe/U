Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.CompilerServices
Imports Microsoft.Xna.Framework
Imports Microsoft.Xna.Framework.Graphics

Public Module ImageProcessing

    ''' <summary>Crops the Bitmap to the specified width and height and returns a new bitmap.</summary>
    <Extension>
    Public Function Crop(Bitmap As Drawing.Bitmap, Width As Integer, Height As Integer) As Drawing.Bitmap
        Using Result As Drawing.Bitmap = CreateBitmap(Width, Height)
            Using Graphics = Drawing.Graphics.FromImage(Result)
                Graphics.DrawImage(Bitmap, 0, 0, New Drawing.Rectangle(0, 0, Width, Height), Drawing.GraphicsUnit.Pixel)
                Graphics.Save()
            End Using

            Return New Drawing.Bitmap(Result)
        End Using
    End Function

    ''' <summary>Creates a new System.Drawing.Bitmap with specified parameters.</summary>
    ''' <param name="Width">Actual Width of the bitmap.</param>
    ''' <param name="Height">Actual Height of the bitmap.</param>
    ''' <param name="xDPI">Dots per inch on X axes.</param>
    ''' <param name="yDPI">Dots per inch on Y axes.</param>
    Public Function CreateBitmap(Width As Integer, Height As Integer, Optional xDPI As Single = 256, Optional yDPI As Single = 256) As Drawing.Bitmap
        Dim Result = New Drawing.Bitmap(Width, Height)
        Result.SetResolution(xDPI, yDPI)
        Return Result
    End Function

    ''' <summary>Might not work.</summary>
    Public Sub DrawStringOnBitmap(ByRef Bitmap As Drawing.Bitmap, Text As String, Font As Drawing.Font, X As Single, Y As Single, Optional Foreground As Drawing.Brush = Nothing, Optional Clear As Boolean = False)
        If (Foreground Is Nothing) Then
            Foreground = Drawing.Brushes.Black
        End If

        Dim Rectangle As Drawing.RectangleF = New Drawing.RectangleF(X, Y, Bitmap.Width, Bitmap.Height)
        Using Graphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(Bitmap)
            With Graphics
                If Clear Then
                    .Clear(System.Drawing.Color.Transparent)
                End If
                .SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
                .InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                .PixelOffsetMode = Drawing.Drawing2D.PixelOffsetMode.HighQuality
                .TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit

                .DrawString(Text, Font, Foreground, Rectangle)
                .Save()
            End With
        End Using
    End Sub
    ''' <summary>Simple tester to check if image stream conversion is possible in all formats.</summary>
    Public Sub StreamConversionTester()
        Dim Bitmap As System.Drawing.Bitmap = New Drawing.Bitmap(10, 10)

        For Each IFormat In New ImageFormat() {ImageFormat.Bmp, ImageFormat.Emf, ImageFormat.Exif, ImageFormat.Gif, ImageFormat.Icon, ImageFormat.Jpeg, ImageFormat.MemoryBmp, ImageFormat.Png, ImageFormat.Tiff, ImageFormat.Wmf}
            Using MemoryStream As MemoryStream = New MemoryStream
                Console.Write("Testing - {0} :", IFormat)
                Dim Success As Boolean = True

                Try
                    Bitmap.Save(MemoryStream, IFormat)
                Catch Exception As Exception
                    Success = False
                End Try

                Console.WriteLine(Microsoft.VisualBasic.vbTab & "{0}", If(Success, "OK", "FAIL"))
            End Using
        Next
    End Sub

    ''' <summary>Replaces all pixels in the texture with the specified color.</summary>
    <Extension>
    Function Fill(Texture As Texture2D, Color As Color) As Texture2D
        Dim Data As Color() = New Color(Texture.Width * Texture.Height - 1) {}

        For Pixel As Integer = 0 To Data.Count() - 1
            Data(Pixel) = Color
        Next

        Texture.SetData(Data)
        Return Texture
    End Function

    Function CreateTexture(Device As GraphicsDevice, Optional Width As Integer = Nothing, Optional Height As Integer = Nothing, Optional Fill As Color = Nothing, Optional Palette As Dictionary(Of Integer, Color) = Nothing) As Texture2D
        Width = If(Width <> Nothing, Width, 50)
        Height = If(Height <> Nothing, Height, 50)

        Dim Texture As Texture2D = New Texture2D(Device, Width, Height)
        Dim Data As Color() = New Color(Width * Height - 1) {}

        If (Palette IsNot Nothing) Then
            For Pixel As Integer = 0 To Data.Count() - 1
                Data(Pixel) = Palette(Pixel)
            Next
        ElseIf (Not (Fill = Nothing)) Then
            For Pixel As Integer = 0 To Data.Count() - 1
                Data(Pixel) = Fill
            Next
        End If

        Texture.SetData(Data)
        Return Texture
    End Function
    Function BitmapToByteArrayLockBits(Bitmap As Drawing.Bitmap) As Byte()
        Dim Rectangle As New Drawing.Rectangle(0, 0, Bitmap.Width, Bitmap.Height)
        Dim Data As BitmapData = Bitmap.LockBits(Rectangle, ImageLockMode.ReadWrite, Bitmap.PixelFormat)
        Dim NumBytes As Integer = (Data.Stride * Bitmap.Height)
        Dim Bytes As Byte() = New Byte(NumBytes - 1) {}
        System.Runtime.InteropServices.Marshal.Copy(Data.Scan0, Bytes, 0, NumBytes)
        Bitmap.UnlockBits(Data)

        Return Bytes
    End Function
    ''' <summary>Significantly faster than SetData and MemoryStream.</summary>
    Function BitmapToTexture2DLockBits(Bitmap As Drawing.Bitmap, Device As GraphicsDevice) As Texture2D
        Dim Result As New Texture2D(Device, Bitmap.Size.Width, Bitmap.Size.Height)
        Dim Rectangle As New Drawing.Rectangle(0, 0, Bitmap.Width, Bitmap.Height)
        Dim Data As BitmapData = Bitmap.LockBits(Rectangle, ImageLockMode.ReadWrite, Bitmap.PixelFormat)
        Dim NumBytes As Integer = (Data.Stride * Bitmap.Height)
        Dim Bytes As Byte() = New Byte(NumBytes - 1) {}
        System.Runtime.InteropServices.Marshal.Copy(Data.Scan0, Bytes, 0, NumBytes)
        Bitmap.UnlockBits(Data)
        Result.SetData(Bytes)

        Return Result
    End Function
    ''' <summary>Significantly Slower than LockBits and MemoryStream.</summary>
    Function BitmapToTexture2DSetData(Bitmap As Drawing.Bitmap, Device As GraphicsDevice) As Texture2D
        Dim Result As New Texture2D(Device, Bitmap.Size.Width, Bitmap.Size.Height)
        Result.SetData(BitmapToColorArrayMonogame(Bitmap))

        Return Result
    End Function
    ''' <summary>Significantly slower than LockBits but faster than SetData.</summary>
    Function BitmapToTexture2DMemoryStream(Bitmap As Drawing.Bitmap, Device As GraphicsDevice) As Texture2D
        Dim Result As Texture2D = Nothing
        Using MemoryStream As MemoryStream = New MemoryStream
            Bitmap.Save(MemoryStream, ImageFormat.Bmp)
            MemoryStream.Seek(0, SeekOrigin.Begin)
            Result = Texture2D.FromStream(Device, MemoryStream)
        End Using

        Return Result
    End Function

    Function BitmapToColorArray(Bitmap As Drawing.Bitmap) As Drawing.Color()
        Dim Color() As Drawing.Color = New Drawing.Color(Bitmap.Size.Width * Bitmap.Size.Height - 1) {}

        For X As Integer = 0 To (Bitmap.Size.Width - 1)
            For Y As Integer = 0 To (Bitmap.Size.Height - 1)
                Color((Bitmap.Size.Width * Y) + X) = Bitmap.GetPixel(X, Y)
            Next
        Next
        Return Color
    End Function
    Function BitmapToColorArrayMonogame(Bitmap As Drawing.Bitmap) As Microsoft.Xna.Framework.Color()
        Dim Color() As Color = New Color(Bitmap.Size.Width * Bitmap.Size.Height - 1) {}

        For X As Integer = 0 To (Bitmap.Size.Width - 1)
            For Y As Integer = 0 To (Bitmap.Size.Height - 1)
                Color((Bitmap.Size.Width * Y) + X) = SystemColorToMonogame(Bitmap.GetPixel(X, Y))
            Next
        Next
        Return Color
    End Function
    Function SystemColorToMonogame(Color As System.Drawing.Color) As Microsoft.Xna.Framework.Color
        Return New Microsoft.Xna.Framework.Color(Color.R, Color.G, Color.B, Color.A)
    End Function
    Function Texture2DToBitmapMemoryStream(Texture As Texture2D) As Drawing.Bitmap
        Dim Result As Drawing.Bitmap = Nothing

        Using MemoryStream As MemoryStream = New MemoryStream()
            Texture.SaveAsPng(MemoryStream, Texture.Width, Texture.Height)
            Result = New System.Drawing.Bitmap(MemoryStream)
        End Using

        Return Result
    End Function
    Function Texture2DToBitmapLockBits(Texture As Texture2D) As Drawing.Bitmap
        Dim Rectangle As New Drawing.Rectangle(0, 0, Texture.Width, Texture.Height)
        Dim Bitmap As New Drawing.Bitmap(Texture.Height, Texture.Width)
        Dim Data As BitmapData = Bitmap.LockBits(Rectangle, ImageLockMode.ReadWrite, Bitmap.PixelFormat)
        Dim NumBytes As Integer = (Data.Stride * Bitmap.Height)
        Dim Bytes As Byte() = New Byte(NumBytes - 1) {}
        System.Runtime.InteropServices.Marshal.Copy(Data.Scan0, Bytes, 0, NumBytes)
        Bitmap.UnlockBits(Data)

        Return Bitmap
    End Function
End Module