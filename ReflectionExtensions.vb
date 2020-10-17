Imports System.Reflection
Imports System.Runtime.CompilerServices

Public Module ReflectionExtensions
    ''' <summary>
    ''' Get all Extension Methods of this Type in the Specified Assembly
    ''' </summary>
    ''' <param name="Type">The Type that has been Extended with the Methods</param>
    ''' <param name="ExtensionsAssembly">The Assembly where the Extension Methods are located</param>
    <Extension()>
    Public Function GetExtensionMethods(Type As Type, ExtensionsAssembly As Assembly) As IEnumerable(Of MethodInfo)
        Dim Query As New List(Of MethodInfo)

        For Each TypeItem In ExtensionsAssembly.GetTypes
            If (Not TypeItem.IsGenericType AndAlso Not TypeItem.IsNested) Then
                For Each Method In TypeItem.GetMethods(BindingFlags.[Static] Or BindingFlags.[Public] Or BindingFlags.NonPublic)
                    If Method.IsDefined(GetType(ExtensionAttribute), False) Then
                        If (Method.GetParameters()(0).ParameterType = Type) Then
                            Query.Add(Method)
                        End If
                    End If
                Next
            End If
        Next

        Return Query
    End Function
    ''' <summary>
    ''' Searches the Assembly for a Method with the given Name in a Type
    ''' </summary>
    ''' <param name="Type">The Type that has been Extended with the Method</param>
    ''' <param name="ExtensionsAssembly">The Assembly where the Extension Method is located</param>
    ''' <param name="Name">Name of the Method to search for</param>
    <Extension()>
    Public Function GetExtensionMethod(Type As Type, ExtensionsAssembly As Assembly, Name As String) As MethodInfo
        Return Type.GetExtensionMethods(ExtensionsAssembly).FirstOrDefault(Function(F) F.Name = Name)
    End Function

    ''' <summary>
    ''' Searches the Assembly for a Method with the given Name in a Type
    ''' </summary>
    ''' <param name="Type">The Type that has been Extended with the Method</param>
    ''' <param name="ExtensionsAssembly">The Assembly where the Extension Method is located</param>
    ''' <param name="Name">Name of the Method to search for</param>
    ''' <param name="Types">TODO: What is this for???</param>
    <Extension()>
    Public Function GetExtensionMethod(Type As Type, ExtensionsAssembly As Assembly, Name As String, Types As Type()) As MethodInfo
        Dim Methods = (From F In Type.GetExtensionMethods(ExtensionsAssembly) Where (F.Name = Name) AndAlso (F.GetParameters().Count() = (Types.Length + 1)) Select F).ToList()

        If Methods.None() Then
            Return Nothing
        End If

        If (Methods.Count() = 1) Then
            Return Methods.First()
        End If

        For Each MethodInfo In Methods
            Dim Parameters = MethodInfo.GetParameters()
            Dim Found As Boolean = True

            For B As Byte = 0 To CByte(Types.Length - 1)
                Found = True

                If (Parameters(B).[GetType]() <> Types(B)) Then
                    Found = False
                End If
            Next

            If Found Then
                Return MethodInfo
            End If
        Next

        Return Nothing
    End Function
    ''' <summary>
    ''' Searches for a Type in all assemblies associated with the execution application
    ''' </summary>
    ''' <param name="Name">Name of the Type</param>
    Public Function GetAssemblyType(Name As String) As Type
        For Each Assembly As Assembly In AppDomain.CurrentDomain.GetAssemblies
            For Each Type As Type In Assembly.GetTypes
                If Type.Name.Equals(Name, StringComparison.OrdinalIgnoreCase) Then
                    Return Type
                End If
            Next
        Next

        Return Nothing
    End Function
    ''' <summary>
    ''' Searches for a Type in all assemblies associated with the execution application
    ''' </summary>
    ''' <param name="FullName">Full Name of the Type</param>
    Public Function GetAssemblyTypeFullName(FullName As String) As Type
        For Each Assembly As Assembly In AppDomain.CurrentDomain.GetAssemblies
            For Each Type As Type In Assembly.GetTypes
                If Type.FullName.Equals(FullName, StringComparison.OrdinalIgnoreCase) Then
                    Return Type
                End If
            Next
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Searches for a Type in all assemblies associated with the execution application, additionally specify which Assembly to look in, used when ambiguity is a concern
    ''' </summary>
    ''' <param name="Name">Name of the Type</param>
    ''' <param name="Assembly">Name of specific Assembly to search in</param>
    Public Function GetAssemblyType(Name As String, Assembly As String) As Type
        For Each Type As Type In AppDomain.CurrentDomain.GetAssemblies.First(Function(F) F.GetName().Name = Assembly).GetTypes
            If Type.Name.Equals(Name, StringComparison.OrdinalIgnoreCase) Then
                Return Type
            End If
        Next

        Return Nothing
    End Function
    ''' <summary>
    ''' Searches for a Type in all assemblies associated with the execution application, additionally specify which Assembly to look in, used when ambiguity is a concern
    ''' </summary>
    ''' <param name="Name">Name of the Type</param>
    ''' <param name="Assembly">Specific Assembly to search in</param>
    Public Function GetAssemblyType(Name As String, Assembly As Assembly) As Type
        For Each Type As Type In Assembly.GetTypes
            If Type.Name.Equals(Name, StringComparison.OrdinalIgnoreCase) Then
                Return Type
            End If
        Next

        Return Nothing
    End Function
    ''' <summary>
    ''' Creates an Instance of a type from the specified Name using its parameterless constructor.
    ''' </summary>
    ''' <param name="Name">Name of the Type.</param>
    ''' <returns>Instance of the Object.</returns>
    Public Function CreateClass(Name As String) As Object
        Dim Type = GetAssemblyType(Name)
        Return IfNot(Type, Activator.CreateInstance(Type))
    End Function
    ''' <summary>
    ''' Creates an Instance of a type from the specified Name using the constructor that fits the parameters best.
    ''' </summary>
    ''' <param name="Name">Name of the Type.</param>
    ''' <param name="Objects">Parameter objects in their correct order.</param>
    ''' <returns>Instance of the Object.</returns>
    Public Function CreateClass(Name As String, ParamArray Objects() As Object) As Object
        Dim Type = GetAssemblyType(Name)
        Return IfNot(Type, Activator.CreateInstance(Type, Objects))
    End Function
End Module
