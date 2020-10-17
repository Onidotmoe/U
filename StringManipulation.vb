Imports System.Runtime.CompilerServices

Public Module StringManipulation
    ''' <summary>
    ''' Adds a New Line After Each String
    ''' </summary>
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

    ''' <summary>
    ''' Searches a String for the First Occurrence of a Substring and Replaces it
    ''' </summary>
    ''' <param name="Text">Parent String</param>
    ''' <param name="Substring">Substring to Replace</param>
    ''' <param name="Replace">String to Replace Substring With</param>
    Public Function ReplaceFirst(Text As String, Substring As String, Replace As String) As String
        Dim Position As Integer = Text.IndexOf(Substring)

        If (Position < 0) Then
            Return Text
        End If

        Return (Text.Substring(0, Position) & Replace) + Text.Substring(Position + Substring.Length)
    End Function
    ''' <summary>
    ''' Searches a String for the Last Occurrence of a Substring and Replaces it
    ''' </summary>
    ''' <param name="Text">Parent String</param>
    ''' <param name="Substring">Substring to Replace</param>
    ''' <param name="Replace">String to Replace Substring With</param>
    Public Function ReplaceLast(Text As String, Substring As String, Replace As String) As String
        Dim Position As Integer = Text.LastIndexOf(Substring)

        If (Position < 0) Then
            Return Text
        End If

        Return (Text.Substring(0, Position) & Replace) + Text.Substring(Position + Substring.Length)
    End Function
    ''' <summary>
    ''' Removes All WhiteSpace Characters from the String
    ''' </summary>
    Public Function RemoveWhiteSpace(Input As String) As String
        Return Input.Where(Function(F) Not Char.IsWhiteSpace(F)).ToArray()
    End Function

    ''' <summary>
    ''' Removes the First Occurrence of the Specified Substring in the String
    ''' </summary>
    <Extension>
    Public Function RemoveFirst([String] As String, Remove As String) As String
        Return ReplaceFirst([String], Remove, String.Empty)
    End Function
    ''' <summary>
    ''' Removes the First Occurrence of Each of the Specified Substrings in the String
    ''' </summary>
    <Extension>
    Public Function RemoveFirsts([String] As String, ParamArray Substrings As String()) As String
        For Each Substring As String In Substrings
            [String] = [String].RemoveFirst(Substring)
        Next

        Return [String]
    End Function
    ''' <summary>
    ''' Removes the Last Occurrence of the Specified Substring in the String
    ''' </summary>
    <Extension>
    Public Function RemoveLast([String] As String, Remove As String) As String
        Return ReplaceLast([String], Remove, String.Empty)
    End Function

    ''' <summary>
    ''' Removes All Occurrence of the Specified Substrings in the String
    ''' </summary>
    <Extension>
    Public Function RemoveAll([String] As String, ParamArray Substrings As String()) As String
        For Each Remove As String In Substrings
            [String] = [String].Replace(Remove, String.Empty)
        Next

        Return [String]
    End Function
    ''' <summary>
    ''' Returns the First Substring between the Start and End Strings exclusively.
    ''' </summary>
    <Extension>
    Public Function Substring([String] As String, Start As String, [End] As String) As String
        Dim Index_Start = ([String].IndexOf(Start) + 1)
        Dim Index_End = ([String].IndexOf([End]) - Index_Start)

        Return [String].Substring(Index_Start, Index_End)
    End Function
    ''' <summary>
    ''' Returns the Inner Substring between the First Occurrence of the Start String and the First Occurrence of the End String after the Start String.
    ''' </summary>
    <Extension>
    Public Function SubstringInner([String] As String, Start As String, [End] As String) As String
        Dim Index_Start = ([String].IndexOf(Start) + Start.Length)
        Dim Index_End = ([String].IndexOf([End], Index_Start) - Index_Start)

        Return [String].Substring(Index_Start, Index_End)
    End Function
End Module