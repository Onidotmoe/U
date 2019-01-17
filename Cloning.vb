Imports System.Reflection
Imports System.Runtime.CompilerServices

Public Module Cloning

    <Extension()>
    Function GetUnderlyingType(ByVal member As MemberInfo) As Type
        Dim type As Type
        Select Case member.MemberType
            Case MemberTypes.Field
                type = (CType(member, FieldInfo)).FieldType
            Case MemberTypes.[Property]
                type = (CType(member, PropertyInfo)).PropertyType
            Case MemberTypes.[Event]
                type = (CType(member, EventInfo)).EventHandlerType
            Case Else
                Throw New ArgumentException("member must be if type FieldInfo, PropertyInfo or EventInfo", "member")
        End Select

        Return If(Nullable.GetUnderlyingType(type), type)
    End Function

    <Extension()>
    Function GetFieldsAndProperties(ByVal type As Type) As MemberInfo()
        Dim fps As List(Of MemberInfo) = New List(Of MemberInfo)()
        fps.AddRange(type.GetFields())
        fps.AddRange(type.GetProperties())
        fps = fps.OrderBy(Function(x) x.MetadataToken).ToList()
        Return fps.ToArray()
    End Function

    <Extension()>
    Function GetValue(ByVal member As MemberInfo, ByVal target As Object) As Object
        If TypeOf member Is PropertyInfo Then
            Return (DirectCast(member, PropertyInfo)).GetValue(target, Nothing)
        ElseIf TypeOf member Is FieldInfo Then
            Return (DirectCast(member, FieldInfo)).GetValue(target)
        Else
            Throw New Exception("member must be either PropertyInfo or FieldInfo")
        End If
    End Function

    <Extension()>
    Sub SetValue(ByVal member As MemberInfo, ByVal target As Object, ByVal value As Object)
        If TypeOf member Is PropertyInfo Then
            DirectCast(member, PropertyInfo).SetValue(target, value, Nothing)
        ElseIf TypeOf member Is FieldInfo Then
            DirectCast(member, FieldInfo).SetValue(target, value)
        Else
            Throw New Exception("destinationMember must be either PropertyInfo or FieldInfo")
        End If
    End Sub

    <Extension()>
    Function DeepClone(ByVal obj As Object) As Object
        If obj Is Nothing Then Return Nothing
        Dim type As Type = obj.[GetType]()
        If TypeOf obj Is IList Then
            Dim list As IList = (CType(obj, IList))
            Dim newlist As IList = CType(Activator.CreateInstance(obj.[GetType](), list.Count), IList)
            For Each elem As Object In list
                newlist.Add(DeepClone(elem))
            Next

            Return newlist
        End If

        If type.IsValueType OrElse type = GetType(String) Then
            Return obj
        ElseIf type.IsArray Then
            Dim elementType As Type = Type.[GetType](type.FullName.Replace("[]", String.Empty))
            Dim array = TryCast(obj, Array)
            Dim copied As Array = Array.CreateInstance(elementType, array.Length)
            For i As Integer = 0 To array.Length - 1
                copied.SetValue(DeepClone(array.GetValue(i)), i)
            Next

            Return Convert.ChangeType(copied, obj.[GetType]())
        ElseIf type.IsClass Then
            Dim toret As Object = Activator.CreateInstance(obj.[GetType]())
            Dim fields As MemberInfo() = type.GetFieldsAndProperties()
            For Each field As MemberInfo In fields
                If field.Name = "parent" Then
                    Continue For
                End If

                Dim fieldValue As Object = field.GetValue(obj)
                If fieldValue Is Nothing Then Continue For
                field.SetValue(toret, DeepClone(fieldValue))
            Next

            Return toret
        Else
            If Debugger.IsAttached Then Debugger.Break()
            Return Nothing
        End If
    End Function

    <Extension>
    Function Clone(Of T)(Obj As T) As T
        Return DirectCast(Obj.DeepClone(), T)
    End Function

End Module