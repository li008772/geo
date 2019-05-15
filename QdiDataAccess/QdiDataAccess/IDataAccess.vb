Option Strict On
Option Explicit On

Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports mgs.Domain
Imports Qdi.BusinessLogic
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto

''' <summary>
'''     Interface for DataAccess.vb
''' </summary>
Public Interface IDataAccess

    ReadOnly Property CurrentLoginStatus() As Qdi.DataAccess.ConnectionStatus
    ReadOnly Property QdiCursor() As IQdiCursorDefinition
    ReadOnly Property PropertyToFieldMapping() As IPropertyToFieldMappingList
    'ReadOnly Property 44QdiFeatureWorkspace() As IFeatureWorkspace
    ReadOnly Property QdiWorkspace() As IWorkspace
    ReadOnly Property ReferenceMapFieldLookup() As Dictionary(Of String, String)
    ReadOnly Property RequiredCodeVersion() As String

    Sub LoadDataIntoView(ByVal activeView As ESRI.ArcGIS.Carto.IActiveView)

    'Sub Connect(ByRef pCodeVersion As String)
    'Sub Disconnect()
    Sub UpdateConnection(ByVal SkipDisconnect As Boolean, ByVal CodeVersion As String)

    Sub Add(ByRef qdiRecord As Qdi.BusinessLogic.IQdiRecord)
    Sub Update(ByRef qdiRecord As Qdi.BusinessLogic.IQdiRecord)
    Sub Delete(ByRef qdiRecord As Qdi.BusinessLogic.IQdiRecord)
    Function SetEditable(ByRef pWorkspaceEdit As IWorkspaceEdit) As Boolean
    Sub SaveEdits(ByRef pWorkspaceEdit As IWorkspaceEdit, Optional ByVal pInDataTable As String = "")
    Sub SelectRelateId(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal pRelateId As String, Optional ByVal SelectionType As esriSelectionResultEnum = esriSelectionResultEnum.esriSelectionResultNew)

    Function ReadbyRelateId(ByRef relateID As String) As IQdiRecord
    Function ReadByOID(ByRef ObjectID As Double) As IQdiRecord

    Function SpatialSelectOneRecord(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal polygon As IGeometry, Optional ByVal addSelect As Boolean = False) As IQdiRecord
    Sub SpatialSelectManyRecords(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal polygon As IGeometry, Optional ByVal addSelect As Boolean = False)
    Sub ZoomToSelectedQDIXRecords(ByVal ActiveView As ESRI.ArcGIS.Carto.IActiveView)


    Function RelateIdExists(ByVal relateId As String) As Boolean
    Function ObjectdIdExists(ByVal ObjectId As Double) As Boolean
    Function CanUpdateObjectIdtoRelateId(ByVal ObjectId As Double, ByVal relateId As String) As Boolean

    Function NextQseriesID() As String
    Function NextUniqueWellId() As String

    Function GetFeatureLayer(ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As ESRI.ArcGIS.Carto.IFeatureLayer
    Function GetFeatureTable(ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As ESRI.ArcGIS.Geodatabase.ITable
    Function GetFeatureTable(ByVal featureClassName As Qdi.BusinessLogic.NamedTables) As ESRI.ArcGIS.Geodatabase.ITable
    Function GetFeatureLayerInView(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As ESRI.ArcGIS.Geodatabase.IDataset

    Function ConvertFeatureToQdiRecord(ByVal feature As IFeature) As Qdi.BusinessLogic.IQdiRecord

End Interface
