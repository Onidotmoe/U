Imports System.Reflection

Public Module PropertyManipulation

    Sub SetPropertyAllOverwrite(OldOjb As Object, NewObj As Object)
        For Each OldProp In OldOjb.[GetType].GetProperties(BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Instance)
            Dim NewProp As PropertyInfo = NewObj.GetType().GetProperty(OldProp.Name, BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Instance)
            If ((NewProp IsNot Nothing) AndAlso (NewProp.CanWrite)) Then
                Dim Value = Convert.ChangeType(OldProp.GetValue(OldOjb, Nothing), NewProp.PropertyType)
                If (Value IsNot Nothing) Then
                    NewProp.SetValue(NewObj, Value)
                End If
            End If
        Next
    End Sub

    Sub SetPropertyValueByName(Obj As Object, Name As String, Value As Object)
        If Name.Contains("."c) Then
            SetNestedProperty(Name, Obj, Value)
        Else
            Dim Prop As PropertyInfo = Obj.GetType().GetProperty(Name, BindingFlags.[Public] Or BindingFlags.Instance)
            If ((Prop IsNot Nothing) AndAlso (Prop.CanWrite)) Then
                Value = Convert.ChangeType(Value, Prop.PropertyType)
                Prop.SetValue(Obj, Value)
            End If
        End If
    End Sub

    Sub SetNestedProperty(Name As String, Obj As Object, Value As Object)
        Dim Bits As String() = Name.Split("."c)

        For i As Integer = 0 To Bits.Length - 2
            Dim PropToGet As PropertyInfo = Obj.[GetType]().GetProperty(Bits(i))
            Obj = PropToGet.GetValue(Obj, Nothing)
        Next

        Dim PropToSet As PropertyInfo = Obj.[GetType]().GetProperty(Bits.Last, BindingFlags.SetProperty Or BindingFlags.IgnoreCase Or BindingFlags.Public Or BindingFlags.Instance)
        Value = Convert.ChangeType(Value, PropToSet.PropertyType)
        PropToSet.SetValue(Obj, Value, Nothing)
    End Sub

    Function GetNestedProperty(Name As String, Obj As Object) As Object
        Dim Bits As String() = Name.Split("."c)

        For i As Integer = 0 To Bits.Length - 1
            Dim PropToGet As PropertyInfo = Obj.[GetType]().GetProperty(Bits(i))
            Obj = PropToGet.GetValue(Obj, Nothing)
        Next

        Return Obj
    End Function

End Module
