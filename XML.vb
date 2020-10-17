Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.Xsl

Public Module XML
    ''' <summary>
    ''' Deserializes a XML String into an Object.
    ''' </summary>
    ''' <typeparam name="T">Type of Object represented by the String.</typeparam>
    ''' <param name="Input">Input XML String to Deserialize.</param>
    ''' <returns>Instance of the deserialized Object.</returns>
    Function Deserialize(Of T)(Input As String) As T
        Dim XMLSerializer As XmlSerializer = XmlSerializer.FromTypes(New System.Type() {GetType(T)})(0)
        Dim Result As T

        Using Reader As XmlTextReader = New XmlTextReader(New StringReader(Input))
            Result = CType(XMLSerializer.Deserialize(Reader), T)
        End Using

        Return Result
    End Function
    ''' <summary>
    ''' Serializes the Object to its XML representation and writes it to the given Path as a XML document.
    ''' </summary>
    ''' <param name="Path">Full path; including file name and extension.</param>
    ''' <param name="Obj">Object to Serialize.</param>
    ''' <returns>True if file was created.</returns>
    Function Serialize(Path As String, Obj As Object) As Boolean
        Dim XMLSerializer As XmlSerializer = New XmlSerializer(Obj.GetType)

        Using Writer As IO.TextWriter = New IO.StreamWriter(Path)
            XMLSerializer.Serialize(Writer, Obj)
        End Using

        Return IO.File.Exists(Path)
    End Function
    ''' <summary>
    ''' Serializes the Object to its XML representation and writes it to the given Path as a XML document.
    ''' </summary>
    ''' <param name="Path">Full path; including file name and extension.</param>
    ''' <param name="Obj">Object to Serialize.</param>
    ''' <returns>True if file was created.</returns>
    Function Write(Path As String, Obj As Object) As Boolean
        Return Serialize(Path, Obj)
    End Function
    ''' <summary>
    ''' Reads a XML document from the given Path and deserializes it into an Object.
    ''' </summary>
    ''' <typeparam name="T">Type of Object to deserialize into.</typeparam>
    ''' <param name="Path">Full path; including file name and extension.</param>
    ''' <returns>Instance of the deserialized Object. Nothing if the specified file does not exist.</returns>
    Function Read(Of T)(Path As String) As T
        If IO.File.Exists(Path) Then
            Return Deserialize(Of T)(File.ReadAllText(Path))
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Reads a XML file with the specified XSL scheme applied, then deserializes it into an Object.
    ''' </summary>
    ''' <typeparam name="T">Type of Object to deserialize into.</typeparam>
    ''' <param name="Path">Full path to the XML file; including file name and extension.</param>
    ''' <param name="Scheme">XSL scheme as a String.</param>
    ''' <returns>Instance of the deserialized file.</returns>
    Function ReadWith(Of T)(Path As Uri, Scheme As String) As T
        If (IO.File.Exists(Path.ToString) AndAlso IO.File.Exists(Scheme)) Then
            Dim XSLT As XslCompiledTransform = New XslCompiledTransform()
            Dim Result As T
            XSLT.Load(XmlReader.Create(New StringReader(Scheme)))

            Using StringWriter As StringWriter = New StringWriter()
                XSLT.Transform(XmlReader.Create(Path.ToString), XmlWriter.Create(StringWriter, XSLT.OutputSettings))

                Dim XMLSerializer As XmlSerializer = XmlSerializer.FromTypes(New System.Type() {GetType(T)})(0)

                Using Reader As XmlTextReader = New XmlTextReader(New StringReader(StringWriter.ToString()))
                    Result = CType(XMLSerializer.Deserialize(Reader), T)
                End Using
            End Using

            Return Result
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Reads a XML document String with the specified XSL scheme applied, then deserializes it into an Object.
    ''' </summary>
    ''' <typeparam name="T">Type of Object to deserialize into.</typeparam>
    ''' <param name="XML">XML document as a String.</param>
    ''' <param name="Scheme">XSL scheme as a String.</param>
    ''' <returns>Instance of the deserialized file.</returns>
    Function ReadWith(Of T)(XML As String, Scheme As String) As T
        Dim XSLT As XslCompiledTransform = New XslCompiledTransform()
        Dim Result As T
        XSLT.Load(XmlReader.Create(New StringReader(Scheme)))

        Using StringWriter As StringWriter = New StringWriter()
            XSLT.Transform(XmlReader.Create(New StringReader(XML)), XmlWriter.Create(StringWriter, XSLT.OutputSettings))

            Dim XMLSerializer As XmlSerializer = XmlSerializer.FromTypes(New System.Type() {GetType(T)})(0)

            Using Reader As XmlTextReader = New XmlTextReader(New StringReader(StringWriter.ToString()))
                Result = CType(XMLSerializer.Deserialize(Reader), T)
            End Using
        End Using

        Return Result
    End Function
    ''' <summary>
    ''' Reads a XML document String with an array of Elements starting from <paramref name="Root"/> into a List of <typeparamref name="T"/>
    ''' </summary>
    ''' <typeparam name="T">Type of Object that list contains.</typeparam>
    ''' <param name="Input">XML document as a String.</param>
    ''' <param name="Root">Root Element of the Array.</param>
    ''' <returns>A List of <typeparamref name="T"/></returns>
    Function ReadArray(Of T)(Input As String, Root As String) As List(Of T)
        Dim XMLSerializer As XmlSerializer = New XmlSerializer(GetType(List(Of T)), New XmlRootAttribute(Root))
        Dim Result As List(Of T)

        Using Reader As XmlTextReader = New XmlTextReader(New StringReader(Input))
            Result = CType(XMLSerializer.Deserialize(Reader), List(Of T))
        End Using

        Return Result
    End Function
    ''' <summary>
    ''' Tries to Load a XML Document from a file.
    ''' </summary>
    ''' <param name="Path">Full path to the XML file; including file name and extension.</param>
    ''' <returns>True if load was successful.</returns>
    Function TestLoad(Path As String) As Boolean
        Try
            Dim XDocument As New XDocument()
            XDocument = XDocument.Load(Path)
            Return True
        Catch Exception As XmlException
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Serialize an Object into a XML String
    ''' </summary>
    Public Function DataSerialize(Of T)([Object] As T) As String
        Dim StringWriter As StringWriter = New StringWriter()
        Dim [String] As New XmlSerializer([Object].GetType())
        [String].Serialize(StringWriter, [Object])
        Return StringWriter.ToString()
    End Function
    ''' <summary>
    ''' Deserialize a XML String into a List of Objects
    ''' </summary>
    Public Function DataDeserialize(Data As String) As List(Of Object)
        Dim XMLSerializer As New XmlSerializer(GetType(List(Of Object)))
        Return CType(XMLSerializer.Deserialize(New StringReader(Data)), List(Of Object))
    End Function
    ''' <summary>
    ''' Deserialize a XML String using the given Type into an Object
    ''' </summary>
    Public Function DataDeserialize(Data As String, Type As Type) As Object
        Dim XMLSerializer As New XmlSerializer(Type)
        Return XMLSerializer.Deserialize(New StringReader(Data))
    End Function
    ''' <summary>
    ''' Get the XsdElement Name for the given Type
    ''' </summary>
    ''' <param name="Type">Type to request name for</param>
    ''' <remarks>Useful for Reflection of XML Elements.</remarks>
    Public Function GetXsdMappingName(Type As Type) As String
        Return (New SoapReflectionImporter().ImportTypeMapping(Type)).XsdElementName
    End Function
    ''' <summary>
    ''' Removes all Namespaces from the XML Document String.
    ''' </summary>
    ''' <param name="XmlDocument">A Xml Document as a String.</param>
    ''' <returns>A Xml Document String with Namespaces stripped.</returns>
    Function RemoveAllNamespaces(XmlDocument As String) As String
        Return RemoveAllNamespaces(XElement.Parse(XmlDocument)).ToString()
    End Function
    ''' <summary>
    ''' Removes all Namespaces from the XElement.
    ''' </summary>
    ''' <param name="XmlDocument">XElement containing namespaces.</param>
    ''' <returns>A XElement with Namespaces stripped.</returns>
    Function RemoveAllNamespaces(XmlDocument As XElement) As XElement
        If (Not XmlDocument.HasElements) Then
            Dim Element As XElement = New XElement(XmlDocument.Name.LocalName)
            Element.Value = XmlDocument.Value

            For Each Attribute As XAttribute In XmlDocument.Attributes()
                Element.Add(Attribute)
            Next

            Return Element
        End If

        Return New XElement(XmlDocument.Name.LocalName, XmlDocument.Elements().[Select](Function(F) RemoveAllNamespaces(F)))
    End Function

End Module