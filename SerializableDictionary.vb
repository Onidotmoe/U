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
        'Dim KeySerializer As XmlSerializer = New XmlSerializer(GetType(TKey))
        'Dim ValueSerializer As XmlSerializer = New XmlSerializer(GetType(TValue))
        Dim WasEmpty As Boolean = Reader.IsEmptyElement
        Reader.Read()
        'Debug.WriteLine(Reader.ReadInnerXml)
        ' Debug.WriteLine(Reader.ReadOuterXml)
        ' Debug.WriteLine(Reader.LocalName)

        If WasEmpty Then
            Return
        End If

        'RemoveAllNamespaces(Reader.)

        While (Reader.NodeType <> System.Xml.XmlNodeType.EndElement)
            ' Debug.WriteLine("inner", Reader.ReadInnerXml())
            '   Debug.WriteLine("outer", Reader.ReadOuterXml())
            Reader.ReadInnerXml()
            Reader.ReadOuterXml()
            Dim KeyValuePair As KeyValuePair = CType(ItemSerializer.Deserialize(Reader), KeyValuePair)
            Me.Add(KeyValuePair.Key, KeyValuePair.Value)

            'Reader.ReadStartElement("Item")
            'Reader.ReadStartElement("Key")
            'Dim Key As TKey = CType(KeySerializer.Deserialize(Reader), TKey)
            'Reader.ReadEndElement()
            'Reader.ReadStartElement("Value")
            'Dim Value As TValue = CType(ValueSerializer.Deserialize(Reader), TValue)
            'Reader.ReadEndElement()
            'Me.Add(Key, Value)
            'Reader.ReadEndElement()
            'Reader.MoveToContent()

            'Reader.ReadStartElement("Item")
            'Reader.ReadStartElement("Key")
            'Dim Key As TKey = CType(KeySerializer.Deserialize(Reader), TKey)
            'Reader.ReadEndElement()
            'Reader.ReadStartElement("Value")
            'Dim Value As TValue = CType(ValueSerializer.Deserialize(Reader), TValue)
            'Reader.ReadEndElement()
            'Me.Add(Key, Value)
            'Reader.ReadEndElement()
            'Reader.MoveToContent()
        End While

        Reader.ReadEndElement()
    End Sub

    Public Sub WriteXml(Writer As System.Xml.XmlWriter) Implements IXmlSerializable.WriteXml
        ' Dim KeySerializer As XmlSerializer = New XmlSerializer(GetType(TKey), "")
        ' Dim ValueSerializer As XmlSerializer = New XmlSerializer(GetType(TValue), "")
        Dim ItemSerializer As XmlSerializer = New XmlSerializer(GetType(KeyValuePair))

        Dim BlankNameSpace As XmlSerializerNamespaces = New XmlSerializerNamespaces()
        BlankNameSpace.Add("", "")

        For Each Key As TKey In Me.Keys
            Dim value As TValue = Me(Key)
            ItemSerializer.Serialize(Writer, New KeyValuePair(Key, value), BlankNameSpace)

            'Writer.WriteStartElement("Item")
            'Writer.WriteStartElement("Key")
            'KeySerializer.Serialize(Writer, Key, BlankNameSpace)
            'Writer.WriteEndElement()
            'Writer.WriteStartElement("Value")
            'Dim value As TValue = Me(Key)
            'ValueSerializer.Serialize(Writer, value, BlankNameSpace)
            'Writer.WriteEndElement()
            'Writer.WriteEndElement()

            'Writer.WriteStartElement("Item")
            'Writer.WriteStartElement("Key")
            'KeySerializer.Serialize(Writer, Key)
            'Writer.WriteEndElement()
            'Writer.WriteStartElement("Value")
            'Dim value As TValue = Me(Key)
            'ValueSerializer.Serialize(Writer, value)
            'Writer.WriteEndElement()
            'Writer.WriteEndElement()
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

'<XmlRoot("Dictionary")>
'Public Class Dictionary(Of TKey, TValue)
'    Inherits Collections.Generic.Dictionary(Of TKey, TValue)
'    Implements IXmlSerializable

'    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements IXmlSerializable.GetSchema
'        Return Nothing
'    End Function

'    Public Sub ReadXml(Reader As System.Xml.XmlReader) Implements IXmlSerializable.ReadXml
'        Dim KeySerializer As XmlSerializer = New XmlSerializer(GetType(TKey))
'        Dim ValueSerializer As XmlSerializer = New XmlSerializer(GetType(TValue))
'        Dim WasEmpty As Boolean = Reader.IsEmptyElement
'        Reader.Read()

'        If WasEmpty Then
'            Return
'        End If

'        While (Reader.NodeType <> System.Xml.XmlNodeType.EndElement)
'            Reader.ReadStartElement("Item")
'            Reader.ReadStartElement("Key")
'            Dim Key As TKey = CType(KeySerializer.Deserialize(Reader), TKey)
'            Reader.ReadEndElement()
'            Reader.ReadStartElement("Value")
'            '  Debug.WriteLine("inner", Reader.ReadInnerXml())
'            'Debug.WriteLine("outer", Reader.ReadOuterXml())
'            'Debug.WriteLine("type", Reader.NodeType)

'            ' Debug.WriteLine(Reader.XmlSpace)

'            Dim Value As TValue = CType(ValueSerializer.Deserialize(Reader), TValue)
'            Reader.ReadEndElement()
'            Me.Add(Key, Value)
'            Reader.ReadEndElement()
'            Reader.MoveToContent()
'        End While

'        Reader.ReadEndElement()
'    End Sub

'    Public Sub WriteXml(Writer As System.Xml.XmlWriter) Implements IXmlSerializable.WriteXml
'        Dim KeySerializer As XmlSerializer = New XmlSerializer(GetType(TKey), "")
'        Dim ValueSerializer As XmlSerializer = New XmlSerializer(GetType(TValue), "")

'        'Dim BlankNameSpace As XmlSerializerNamespaces = New XmlSerializerNamespaces()
'        'BlankNameSpace.Add("", "")

'        For Each Key As TKey In Me.Keys
'            Writer.WriteStartElement("Item")
'            Writer.WriteStartElement("Key")
'            KeySerializer.Serialize(Writer, Key)
'            Writer.WriteEndElement()
'            Writer.WriteStartElement("Value")
'            Dim value As TValue = Me(Key)
'            ValueSerializer.Serialize(Writer, value)
'            Writer.WriteEndElement()
'            Writer.WriteEndElement()
'        Next
'    End Sub
'End Class