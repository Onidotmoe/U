Public Module StringManipulation

    Function Box(ParamArray Strings() As String) As String
        Dim Output As String = Nothing

        For i As Integer = 0 To Strings.Count - 1
            If (i <> (Strings.Count - 1)) Then
                Output += Strings(i) + vbCrLf
            Else
                Output += Strings(i)
            End If
        Next

        Return Output
    End Function

    Public Sub FindString(ByVal Source As String, ByVal Search As String, ByRef Output As String, Optional ByVal EndChar As String = """", Optional Overwrite As Boolean = False)
        If (Not String.IsNullOrWhiteSpace(Source)) AndAlso (String.IsNullOrWhiteSpace(Output) OrElse (Overwrite = True)) Then
            Dim Index_Start As New Integer
            If (Not String.IsNullOrWhiteSpace(Search)) Then
                Index_Start = Source.IndexOf(Search, StringComparison.InvariantCultureIgnoreCase)
            End If

            If (Index_Start > -1) Then
                Dim Index_End As New Integer
                Dim SearchLength As Integer = If(String.IsNullOrEmpty(Search), 0, Search.Length)

                If (Not String.IsNullOrWhiteSpace(EndChar)) Then
                    Index_End = Source.IndexOf(EndChar, Index_Start + SearchLength, StringComparison.InvariantCultureIgnoreCase)
                Else
                    Index_End = Source.Length
                End If

                If (Index_End > -1) Then
                    Dim Result As String = Source.Substring(Index_Start + SearchLength, Index_End - (Index_Start + SearchLength))

                    If (Not String.IsNullOrWhiteSpace(Result)) Then
                        Output = Result
                    End If
                End If
            End If
        End If
    End Sub

    Public Function ReplaceFirst(text As String, search As String, replace As String) As String
        Dim pos As Integer = text.IndexOf(search)
        If pos < 0 Then
            Return text
        End If
        Return (text.Substring(0, pos) & replace) + text.Substring(pos + search.Length)
    End Function

    Public Function RemoveWhiteSpace(ByVal Input As String) As String
        Return Input.Where(Function(fItem) Not Char.IsWhiteSpace(fItem)).ToArray()
    End Function

End Module