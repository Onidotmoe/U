Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices

Public Module FontManipulation
    <StructLayout(LayoutKind.Sequential)>
    Structure KerningPair
        Public wFirst As Short
        Public wSecond As Short
        Public iKernelAmount As Integer
    End Structure

    <DllImport("gdi32.dll", SetLastError:=True, CallingConvention:=CallingConvention.Winapi)>
    Public Function GetKerningPairs(hdc As IntPtr, nPairs As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> <Out()> pairs() As KerningPair) As Integer
    End Function

    <DllImport("gdi32.dll")>
    Private Function SelectObject(hdc As IntPtr, hObject As IntPtr) As IntPtr
    End Function

    Public Function GetKerningPairs(Font As Drawing.Font) As IList(Of KerningPair)
        Using Graphics As Drawing.Graphics = Drawing.Graphics.FromHwnd(IntPtr.Zero)
            Graphics.PageUnit = Drawing.GraphicsUnit.Pixel

            Dim Pairs() As KerningPair
            Dim Hdc As IntPtr = Graphics.GetHdc
            Dim HFont As IntPtr = Font.ToHfont
            Dim Old As IntPtr = SelectObject(Hdc, HFont)
            Try
                Dim NumPairs As Integer = GetKerningPairs(Hdc, 0, Nothing)
                If (NumPairs > 0) Then
                    Pairs = New KerningPair(NumPairs - 1) {}
                    NumPairs = GetKerningPairs(Hdc, NumPairs, Pairs)
                    Return Pairs
                Else
                    Return Nothing
                End If
            Finally
                Old = SelectObject(Hdc, Old)
            End Try
        End Using
    End Function

    Sub ExaminePairs(FontFamily As Drawing.FontFamily)
        Try
            Using Font = New Drawing.Font(FontFamily, 25)
                Dim Pairs = GetKerningPairs(Font)
                If Pairs IsNot Nothing Then
                    Debug.WriteLine("#Pairs: {0}", Pairs.Count)
                Else
                    Debug.WriteLine("No pairs found")
                End If
            End Using
        Catch Ex As Exception
            Debug.WriteLine("Error: {0} for: {1}", Ex.Message, FontFamily.Name)
        End Try
    End Sub

    Public Function GetAllInstalledFonts() As List(Of Drawing.FontFamily)
        Return New List(Of Drawing.FontFamily)((New Drawing.Text.InstalledFontCollection).Families)
    End Function
    Public Function GetCharactersFromUnicodeRange(Start As String, [End] As String) As List(Of Char)
        Dim CharBag As New ConcurrentBag(Of Char)
        Parallel.For(CInt("&H" & Start.Remove(0, 2)), CInt("&H" & [End].Remove(0, 2)), Sub(i) CharBag.Add(Microsoft.VisualBasic.ChrW(i)))
        Dim CharList = CharBag.ToList
        CharList.Sort()
        Return CharList
    End Function
End Module