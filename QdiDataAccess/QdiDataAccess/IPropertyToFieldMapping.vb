Option Strict On
Option Explicit On

''' <summary>
'''     Interface for PropertyToFieldMappingList.vb
''' </summary>
Public Interface IPropertyToFieldMappingList
    ReadOnly Property PropertyToFieldMapping() As Dictionary(Of String, String)
    ReadOnly Property FieldList() As System.Collections.Generic.Dictionary(Of String, String).KeyCollection
    ReadOnly Property PropertyList() As System.Collections.Generic.Dictionary(Of String, String).KeyCollection
    ReadOnly Property FieldToPropertyMapping() As Dictionary(Of String, String)

    Function ReturnFieldName(ByVal inProperty As String) As String
    Function ReturnProperty(ByVal inField As String) As String

End Interface


