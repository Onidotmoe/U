Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.Xsl

Public Module XML

    Function Deserialize(Of T)(Input As String) As T
        Dim XMLSerializer As XmlSerializer = XmlSerializer.FromTypes(New System.Type() {GetType(T)})(0)
        Dim Result As T

        Using Reader As XmlTextReader = New XmlTextReader(New StringReader(Input))
            Result = CType(XMLSerializer.Deserialize(Reader), T)
        End Using

        Return Result
    End Function

    Function Serialize(Path As String, Obj As Object) As Boolean
        Dim XMLSerializer As XmlSerializer = New XmlSerializer(Obj.GetType)

        Using Writer As IO.TextWriter = New IO.StreamWriter(Path)
            XMLSerializer.Serialize(Writer, Obj)
        End Using

        Return IO.File.Exists(Path)
    End Function

    Function Write(Path As String, Obj As Object) As Boolean
        Return Serialize(Path, Obj)
    End Function

    Function Read(Of T)(Path As String) As T
        If IO.File.Exists(Path) Then
            Return Deserialize(Of T)(File.ReadAllText(Path))
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Reads a XML file with the specified XSL scheme applied.
    ''' </summary>
    Function ReadWith(Of T)(Path As Uri, Scheme As String) As T
        If (IO.File.Exists(Path.ToString) AndAlso IO.File.Exists(Scheme)) Then
            Dim XSLT As XslCompiledTransform = New XslCompiledTransform()
            Dim Result As T
            XSLT.Load(XmlReader.Create(New StringReader(Scheme)))

            Using sw As StringWriter = New StringWriter()
                XSLT.Transform(XmlReader.Create(Path.ToString), XmlWriter.Create(sw, XSLT.OutputSettings))

                Dim XMLSerializer As XmlSerializer = XmlSerializer.FromTypes(New System.Type() {GetType(T)})(0)

                Using Reader As XmlTextReader = New XmlTextReader(New StringReader(sw.ToString()))
                    Result = CType(XMLSerializer.Deserialize(Reader), T)
                End Using
            End Using

            Return Result
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>Reads a XML string with the specified XSL scheme applied.</summary>
    Function ReadWith(Of T)(XML As String, Scheme As String) As T
        Dim XSLT As XslCompiledTransform = New XslCompiledTransform()
        Dim Result As T
        XSLT.Load(XmlReader.Create(New StringReader(Scheme)))

        Using sw As StringWriter = New StringWriter()
            XSLT.Transform(XmlReader.Create(New StringReader(XML)), XmlWriter.Create(sw, XSLT.OutputSettings))

            Dim XMLSerializer As XmlSerializer = XmlSerializer.FromTypes(New System.Type() {GetType(T)})(0)

            Using Reader As XmlTextReader = New XmlTextReader(New StringReader(sw.ToString()))
                Result = CType(XMLSerializer.Deserialize(Reader), T)
            End Using
        End Using

        Return Result
    End Function

    Function ReadArray(Of T)(Input As String, Root As String) As List(Of T)
        Dim XMLSerializer As XmlSerializer = New XmlSerializer(GetType(List(Of T)), New XmlRootAttribute(Root))
        Dim Result As List(Of T)

        Using Reader As XmlTextReader = New XmlTextReader(New StringReader(Input))
            Result = CType(XMLSerializer.Deserialize(Reader), List(Of T))
        End Using

        Return Result
    End Function

    Function TestLoad(Path As String) As Boolean
        Try
            Dim Xdoc As New XDocument()
            Xdoc = XDocument.Load(Path)
            Return True
        Catch exception As XmlException
            Return False
        End Try
    End Function

    <AttributeUsage(AttributeTargets.[Property], AllowMultiple:=False)>
    Class XmlCommentAttribute
        Inherits Attribute
        Private _Comment As String

        Public Property Comment() As String
            Get
                Return _Comment
            End Get
            Set
                _Comment = Comment
            End Set
        End Property

    End Class

    Public Function DataSerialize(Of T)(List As T) As String
        Dim StringWriter As StringWriter = New StringWriter()
        Dim [String] As New XmlSerializer(List.GetType())
        [String].Serialize(StringWriter, List)
        Return StringWriter.ToString()
    End Function

    Public Function DataDeserialize(Data As String) As List(Of Object)
        Dim XMLSerializer As New XmlSerializer(GetType(List(Of Object)))
        Dim List As List(Of Object) = CType(XMLSerializer.Deserialize(New StringReader(Data)), List(Of Object))
        Return List
    End Function

End Module