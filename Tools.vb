Imports System.IO
Imports System.Reflection
Imports System.Runtime.Serialization.Formatters.Binary

Public Module Tools

    ''' <summary>
    ''' Goes through an Array of objects until one is not Nothing, returns Nothing if All objects are Nothing.
    ''' </summary>
    ''' <param name="Objects"></param>
    Function IfElse(Of T)(ParamArray Objects() As T) As T
        For Each Obj In Objects
            If (Obj IsNot Nothing) Then
                Return Obj
            End If
        Next

        Return Nothing
    End Function

    '''' <summary>
    '''' Independent of the first If statement, this one will run if true alone.
    '''' </summary>
    ''IfAlso

    ''' <summary>
    ''' If Object IsNot Nothing then commit Action.
    ''' </summary>
    Sub IfNot([Object] As Object, [Action] As Action)
        If ([Object] IsNot Nothing) Then
            [Action]()
        End If
    End Sub

    ''' <summary>
    ''' If Object IsNot Nothing then return result, otherwise return Nothing.
    ''' </summary>
    Function IfNot(Of T)([Object] As Object, Result As T) As T
        If ([Object] IsNot Nothing) Then
            Return Result
        Else
            Return Nothing
        End If
    End Function

    Class Pair(Of T1, T2)

        Public Sub New()
        End Sub

        Public Sub New(First As T1, Second As T2)
            Me.First = First
            Me.Second = Second
        End Sub

        Public Property First() As T1
        Public Property Second() As T2
    End Class

    Class Debug

        Public Shared Sub WriteLine(ParamArray Objects() As Object)
            Dim Line As String = Nothing

            For i = 0 To Objects.Count - 1
                If (i > 0) Then
                    Line += " _ "
                End If
                If (Objects(i) IsNot Nothing) Then
                    Line += Objects(i).ToString
                Else
                    Line += " "
                End If
            Next

            Diagnostics.Debug.WriteLine(Line)
        End Sub

    End Class

    Public Enum SizeUnits
        [Byte]
        KB
        MB
        GB
        TB
        PB
        EB
        ZB
        YB
    End Enum

    Public Function ToByteArray(Of T)(Input As T) As Byte()
        If (Input Is Nothing) Then
            Return Nothing
        Else
            Dim BinaryFormatter As New BinaryFormatter
            Using MemoryStream As MemoryStream = New MemoryStream
                BinaryFormatter.Serialize(MemoryStream, Input)
                Return MemoryStream.ToArray()
            End Using
        End If
    End Function

    Public Function FromByteArray(Of T)(Input As T) As T
        If (Input Is Nothing) Then
            Return CType(Nothing, T)
        Else
            Dim BinaryFormatter As New BinaryFormatter
            Using MemoryStream As MemoryStream = New MemoryStream
                Dim Result As Object = BinaryFormatter.Deserialize(MemoryStream)
                Return CType(Result, T)
            End Using
        End If
    End Function

    Public Async Function DelayTask(Time As Double) As Threading.Tasks.Task
        Await Threading.Tasks.Task.Delay(TimeSpan.FromSeconds(Time))
    End Function

    '''' <summary>Delays action to allow userinput to finish. Checks if start-input differs from end-input, cancel further actions.</summary>
    'Async Function IfNotChanged(Text As String, TextInput As GeonBit.UI.Entities.TextInput) As Task(Of Boolean)
    '    Await DelayTask(1.5)

    '    If (Not Equals(Text, TextInput.Value)) Then
    '        Return False
    '    Else
    '        Return True
    '    End If
    'End Function
    Public Function IsNullOrDefaultValue([Object] As Object) As Boolean
        Return ([Object] Is Nothing) OrElse ([Object].GetType.IsValueType AndAlso ([Object] Is Nothing))
    End Function

    ''' <summary>Returns a the number that's the Nearest multiplier of specified Number</summary>
    Public Function RoundToNearest(Number As Double, Nearest As Double) As Double
        Return (Math.Ceiling(Number / Nearest) * Nearest)
    End Function
    Public Function DescendingComparer() As Comparer(Of Integer)
        Return Comparer(Of Integer).Create(Function(X, Y) Y.CompareTo(X))
    End Function



End Module