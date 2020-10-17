Imports System.Runtime.CompilerServices
Imports System.Xml.Serialization

<XmlRoot("Dictionary")>
Public Class Dictionary(Of TKey, TValue)
    Inherits Collections.Generic.Dictionary(Of TKey, TValue)
    Implements IXmlSerializable

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

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements IXmlSerializable.GetSchema
        Return Nothing
    End Function

    Public Sub ReadXml(Reader As System.Xml.XmlReader) Implements IXmlSerializable.ReadXml
        If (Not Reader.IsEmptyElement) Then
            Dim ItemSerializer As XmlSerializer = New XmlSerializer(GetType(KeyValuePair))

            Reader.Read()

            While (Reader.NodeType <> System.Xml.XmlNodeType.EndElement)
                Reader.ReadInnerXml()
                Reader.ReadOuterXml()
                Dim KeyValuePair As KeyValuePair = CType(ItemSerializer.Deserialize(Reader), KeyValuePair)

                Me.Add(KeyValuePair.Key, KeyValuePair.Value)
            End While

            Reader.ReadEndElement()
        End If
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

End Class
