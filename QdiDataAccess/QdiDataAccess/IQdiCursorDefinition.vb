Option Strict On
Option Explicit On

Imports ESRI.ArcGIS.Geodatabase

''' <summary>
'''     Interface of mostly properties used byrsor definition classes.
''' </summary>
Public Interface IQdiCursorDefinition
    'ReadOnly Property SQLstring() As String
    Property ObjectId() As Decimal

    WriteOnly Property RelateId() As Qdi.BusinessLogic.IRelateIdValidator
    Property Township() As Nullable(Of Decimal)
    Property Range() As Nullable(Of Decimal)
    Property Section() As Nullable(Of Decimal)
    Property DataSource() As String
    Property FromDate() As Nullable(Of Decimal)
    Property ToDate() As Nullable(Of Decimal)
    Property FromDepth() As Nullable(Of Decimal)
    Property ToDepth() As Nullable(Of Decimal)
    Property Counties() As List(Of String)
    Property Quadrangles() As List(Of String)
    Property FirstBedrockUnits() As List(Of String)
    Property LastBedrockUnits() As List(Of String)
    Property FirstStratUnits() As List(Of String)
    Property LastStratUnits() As List(Of String)
    Property Aquifers() As List(Of String)
    Property RelateIDs() As List(Of String)

    '** A Cursor of Qdix
    Function QdixICursor() As ICursor

    Sub Refresh()

    ReadOnly Property ValidationErrors() As List(Of String)
End Interface
