Imports System.ComponentModel
Imports System.Reflection

Public Module PropertyManipulation

    Sub SetPropertyAllOverwrite(Old_Object As Object, New_Object As Object)
        For Each Old_Property In Old_Object.[GetType].GetProperties(BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Instance)
            Dim New_Property As PropertyInfo = New_Object.GetType().GetProperty(Old_Property.Name, BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Instance)

            If ((New_Property IsNot Nothing) AndAlso New_Property.CanWrite) Then
                Dim Value = Convert.ChangeType(Old_Property.GetValue(Old_Object, Nothing), New_Property.PropertyType)

                If (Value IsNot Nothing) Then
                    New_Property.SetValue(New_Object, Value)
                End If
            End If
        Next
    End Sub
    ''' <summary>
    ''' Gets the value on a Property on the given Object.
    ''' </summary>
    ''' <param name="Obj">Object containing the property.</param>
    ''' <param name="Name">Name or Full path to the property.</param>
    ''' <returns>The Property Value Object. Returns nothing if it couldn't be found or read.</returns>
    Function GetPropertyValueByName(Obj As Object, Name As String) As Object
        If Name.Contains("."c) Then
            Return GetNestedProperty(Obj, Name)
        Else
            Dim Prop As PropertyInfo = Obj.GetType().GetProperty(Name, BindingFlags.[Public] Or BindingFlags.Instance)

            If ((Prop IsNot Nothing) AndAlso (Prop.CanRead)) Then
                Return Prop.GetValue(Obj)
            Else
                Return Nothing
            End If
        End If
    End Function
    ''' <summary>
    ''' Returns the Value of the specified Property from its full path.
    ''' </summary>
    ''' <param name="Obj">Object containing the property value.</param>
    ''' <param name="Name">Full path to the property.</param>
    ''' <returns>Property value on the given Object.</returns>
    Function GetNestedProperty(Obj As Object, Name As String) As Object
        Dim Bits As String() = Name.Split("."c)

        For i As Integer = 0 To Bits.Length - 1
            Dim PropToGet As PropertyInfo = Obj.[GetType]().GetProperty(Bits(i))
            Obj = PropToGet.GetValue(Obj, Nothing)
        Next

        Return Obj
    End Function


    ''' <summary>
    ''' Sets the value on a Property on the given Object.
    ''' </summary>
    ''' <param name="Obj">Object containing the property.</param>
    ''' <param name="Name">Name or Full path to the property.</param>
    ''' <param name="Value">Value to set.</param>
    Sub SetPropertyValueByName(Obj As Object, Name As String, Value As Object)
        If Name.Contains("."c) Then
            SetNestedProperty(Obj, Name, Value)
        Else
            Dim Prop As PropertyInfo = Obj.GetType().GetProperty(Name, BindingFlags.[Public] Or BindingFlags.Instance)

            If ((Prop IsNot Nothing) AndAlso (Prop.CanWrite)) Then
                Value = Convert.ChangeType(Value, Prop.PropertyType)
                Prop.SetValue(Obj, Value)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Sets the Nested Property on the given Object.
    ''' </summary>
    ''' <param name="Name">Full path to the property.</param>
    ''' <param name="Obj">Object containing the property value.</param>
    ''' <param name="Value">Value to set.</param>
    Sub SetNestedProperty(Obj As Object, Name As String, Value As Object)
        Dim Bits As String() = Name.Split("."c)

        For i As Integer = 0 To Bits.Length - 2
            Dim PropToGet As PropertyInfo = Obj.[GetType]().GetProperty(Bits(i))
            Obj = PropToGet.GetValue(Obj, Nothing)
        Next

        Dim PropToSet As PropertyInfo = Obj.[GetType]().GetProperty(Bits.Last, BindingFlags.SetProperty Or BindingFlags.IgnoreCase Or BindingFlags.Public Or BindingFlags.Instance)
        Value = Convert.ChangeType(Value, PropToSet.PropertyType)
        PropToSet.SetValue(Obj, Value, Nothing)
    End Sub
    ''' <summary>
    ''' Sets a Value on a Field in a Structure and handles the Boxing/Unboxing Process.
    ''' </summary>
    ''' <param name="Struct">Reference to the Structure that Holds the Field</param>
    ''' <param name="FieldName">Name of the Field to Set the Value On</param>
    ''' <param name="Value">The Value to Set on the Field</param>
    ''' <returns></returns>
    Function StructureSetValue(ByRef Struct As Object, FieldName As String, Value As Object) As Object
        Dim StructValueType As ValueType = CType(Struct, ValueType)
        Dim Field As FieldInfo = StructValueType.[GetType]().GetField(FieldName)

        Field.SetValue(StructValueType, Value)
        Return StructValueType
    End Function
    ''' <summary>
    ''' Sets a Field's Value on the given Object.
    ''' </summary>
    ''' <param name="Obj">Object containing the field.</param>
    ''' <param name="Name">Name of the field.</param>
    ''' <param name="Value">Value to set.</param>
    Sub SetFieldValueByName(Obj As Object, Name As String, Value As Object)
        Dim Field As FieldInfo = Obj.GetType().GetField(Name, BindingFlags.[Public] Or BindingFlags.Instance)

        If (Field IsNot Nothing) Then
            Value = Convert.ChangeType(Value, Field.FieldType)
            Field.SetValue(Obj, Value)
        End If
    End Sub

    ''' <summary>
    ''' Converts the String to the Property Type from the Object
    ''' </summary>
    Function ToObjectProperty([Object] As Object, Name As String, Value As String) As Object
        Return TypeDescriptor.GetConverter([Object].GetType.GetProperty(Name).PropertyType).ConvertFromInvariantString(Value)
    End Function
    ''' <summary>
    ''' Converts the String to the Field Type from the Object
    ''' </summary>
    Function ToObjectField([Object] As Object, Name As String, Value As String) As Object
        Return TypeDescriptor.GetConverter([Object].GetType.GetField(Name).FieldType).ConvertFromInvariantString(Value)
    End Function

End Module
