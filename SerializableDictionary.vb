Imports System.Xml.Serialization

<XmlRoot("Dictionary")>
Public Class Dictionary(Of TKey, TValue)
    Inherits Collections.Generic.Dictionary(Of TKey, TValue)
    Implements IXmlSerializable

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements IXmlSerializable.GetSchema
        Return Nothing
    End Function

    <XmlRoot("Item")>
    Public Structure KeyValuePair
        <XmlElement>
        Public Key As TKey
        <XmlElement>
        Public Value As TValue
        Sub New(Key As TKey, Value As TValue)
            Me.Key = Key
            Me.Value = Value
        End Sub
    End Structure

    Public Sub ReadXml(Reader As System.Xml.XmlReader) Implements IXmlSerializable.ReadXml
        Dim ItemSerializer As XmlSerializer = New XmlSerializer(GetType(KeyValuePair))
        Dim WasEmpty As Boolean = Reader.IsEmptyElement
        Reader.Read()

        If WasEmpty Then
            Return
        End If

        While (Reader.NodeType <> System.Xml.XmlNodeType.EndElement)
            Reader.ReadInnerXml()
            Reader.ReadOuterXml()
            Dim KeyValuePair As KeyValuePair = CType(ItemSerializer.Deserialize(Reader), KeyValuePair)
            Me.Add(KeyValuePair.Key, KeyValuePair.Value)
        End While

        Reader.ReadEndElement()
    End Sub

    Public Sub WriteXml(Writer As System.Xml.XmlWriter) Implements IXmlSerializable.WriteXml
        Dim ItemSerializer As XmlSerializer = New XmlSerializer(GetType(KeyValuePair))

        Dim BlankNameSpace As XmlSerializerNamespaces = New XmlSerializerNamespaces()
        BlankNameSpace.Add("", "")

        For Each Key As TKey In Me.Keys
            Dim Value As TValue = Me(Key)
            ItemSerializer.Serialize(Writer, New KeyValuePair(Key, Value), BlankNameSpace)
        Next
    End Sub

    Public Shared Function RemoveAllNamespaces(ByVal xmlDocument As String) As String
        Dim xmlDocumentWithoutNs As XElement = RemoveAllNamespaces(XElement.Parse(xmlDocument))
        Return xmlDocumentWithoutNs.ToString()
    End Function

    Private Shared Function RemoveAllNamespaces(ByVal xmlDocument As XElement) As XElement
        If Not xmlDocument.HasElements Then
            Dim xElement As XElement = New XElement(xmlDocument.Name.LocalName)
            xElement.Value = xmlDocument.Value

            For Each attribute As XAttribute In xmlDocument.Attributes()
                xElement.Add(attribute)
            Next

            Return xElement
        End If

        Return New XElement(xmlDocument.Name.LocalName, xmlDocument.Elements().[Select](Function(el) RemoveAllNamespaces(el)))
    End Function

End Class