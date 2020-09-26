Imports System.Collections.Concurrent
Imports System.Runtime.CompilerServices
Imports Microsoft.Xna.Framework

Public Module Extensions

    <Extension>
    Function None(Of TSource)(source As IEnumerable(Of TSource)) As Boolean
        Return (Not source.Any())
    End Function

    <Extension>
    Function None(Of TSource)(source As IEnumerable(Of TSource), predicate As Func(Of TSource, Boolean)) As Boolean
        Return (Not source.Any(predicate))
    End Function

    <Extension>
    Function TryRemove(Of TKey, TValue)(Self As ConcurrentDictionary(Of TKey, TValue), key As TKey) As Boolean
        Dim Ignored As TValue
        Return Self.TryRemove(key, Ignored)
    End Function

    <Extension>
    Function Remove(Of TKey, TValue)(Self As ConcurrentDictionary(Of TKey, TValue), key As TKey) As Boolean
        Return (CType(Self, IDictionary(Of TKey, TValue))).Remove(key)
    End Function

    <Extension()>
    Function ToSize(Value As Int64, Unit As SizeUnits) As String
        Return (Value / Math.Pow(1024, CType(Unit, Long))).ToString("0.00" + Unit.ToString)
    End Function

    <Extension()>
    Sub ForEach(Of T)(Seq As IEnumerable(Of T), Action As Action(Of T))
        For Each Item In Seq
            Action(Item)
        Next
    End Sub

    <Extension()>
    Iterator Function ForEachLazy(Of T)(Seq As IEnumerable(Of T), Action As Action(Of T)) As IEnumerable(Of T)
        For Each Item In Seq
            Action(Item)
            Yield Item
        Next
    End Function

    <Extension()>
    Function ToVector(Point As Point) As Vector2
        Return New Vector2(Point.X, Point.Y)
    End Function
    <Extension>
    Public Sub AddRange(Of T)(Bag As ConcurrentBag(Of T), Range As IEnumerable(Of T))
        For Each Item In Range
            Bag.Add(Item)
        Next
    End Sub

End Module