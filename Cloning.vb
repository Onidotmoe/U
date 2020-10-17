Imports System.Reflection
Imports System.Runtime.CompilerServices

Public Module Cloning

    <Extension()>
    Function GetUnderlyingType(Member As MemberInfo) As Type
        Dim Type As Type

        Select Case Member.MemberType
            Case MemberTypes.Field
                Type = (CType(Member, FieldInfo)).FieldType
            Case MemberTypes.[Property]
                Type = (CType(Member, PropertyInfo)).PropertyType
            Case MemberTypes.[Event]
                Type = (CType(Member, EventInfo)).EventHandlerType
            Case Else
                Throw New ArgumentException("Member must be if Type FieldInfo, PropertyInfo or EventInfo", "Member")
        End Select

        Return If(Nullable.GetUnderlyingType(Type), Type)
    End Function

    <Extension()>
    Function GetFieldsAndProperties(Type As Type) As MemberInfo()
        Dim Members As List(Of MemberInfo) = New List(Of MemberInfo)()

        Members.AddRange(Type.GetFields())
        Members.AddRange(Type.GetProperties())
        Members = Members.OrderBy(Function(F) F.MetadataToken).ToList()

        Return Members.ToArray()
    End Function

    <Extension()>
    Function GetValue(Member As MemberInfo, Target As Object) As Object
        If (TypeOf Member Is PropertyInfo) Then
            Return (DirectCast(Member, PropertyInfo)).GetValue(Target, Nothing)

        ElseIf (TypeOf Member Is FieldInfo) Then
            Return (DirectCast(Member, FieldInfo)).GetValue(Target)
        Else
            Throw New Exception("Member must be either PropertyInfo or FieldInfo")
        End If
    End Function

    <Extension()>
    Sub SetValue(Member As MemberInfo, Target As Object, Value As Object)
        If (TypeOf Member Is PropertyInfo) Then
            DirectCast(Member, PropertyInfo).SetValue(Target, Value, Nothing)

        ElseIf (TypeOf Member Is FieldInfo) Then
            DirectCast(Member, FieldInfo).SetValue(Target, Value)
        Else
            Throw New Exception("destinationMember must be either PropertyInfo or FieldInfo")
        End If
    End Sub

    <Extension()>
    Function DeepClone(Obj As Object) As Object
        If Obj Is Nothing Then
            Return Nothing
        End If

        Dim Type As Type = Obj.[GetType]()

        If (TypeOf Obj Is IList) Then
            Dim List As IList = (CType(Obj, IList))
            Dim Newlist As IList = CType(Activator.CreateInstance(Obj.[GetType](), List.Count), IList)

            For Each Element As Object In List
                Newlist.Add(DeepClone(Element))
            Next

            Return Newlist
        End If

        If (Type.IsValueType OrElse (Type = GetType(String))) Then
            Return Obj

        ElseIf Type.IsArray Then
            Dim ElementType As Type = Type.[GetType](Type.FullName.Replace("[]", String.Empty))
            Dim Array = TryCast(Obj, Array)
            Dim Copied As Array = Array.CreateInstance(ElementType, Array.Length)

            For i As Integer = 0 To array.Length - 1
                copied.SetValue(DeepClone(array.GetValue(i)), i)
            Next

            Return Convert.ChangeType(copied, Obj.[GetType]())

        ElseIf Type.IsClass Then
            Dim Result As Object = Activator.CreateInstance(Obj.[GetType]())
            Dim Fields As MemberInfo() = Type.GetFieldsAndProperties()

            For Each Field As MemberInfo In Fields
                If (Field.Name = "parent") Then
                    Continue For
                End If

                Dim FieldValue As Object = Field.GetValue(Obj)

                If (FieldValue Is Nothing) Then
                    Continue For
                End If

                Field.SetValue(Result, DeepClone(FieldValue))
            Next

            Return Result
        Else
            If Debugger.IsAttached Then
                Debugger.Break()
            End If

            Return Nothing
        End If
    End Function

    <Extension>
    Function Clone(Of T)(Obj As T) As T
        Return DirectCast(Obj.DeepClone(), T)
    End Function

End Module