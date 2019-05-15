Option Strict On
Option Explicit On


Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
'Imports ESRI.ArcGIS.Editor
'Imports ESRI.ArcGIS.arcmapui
Imports mgs.CursorBuilder
Imports mgs.Domain
Imports Qdi.BusinessLogic
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Reflection

Public MustInherit Class DataAccessBaseClass

    Implements IDataAccess

    Private m_Workspace As IWorkspace
    Private m_GUID As Guid = Guid.NewGuid
    Protected m_QdiCursor As Qdi.DataAccess.IQdiCursorDefinition
    'Protected m_QdixFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_ConnectionStatus As Qdi.DataAccess.ConnectionStatus  'requiredCodeVersion
    'Const m_UserName As String = System.Environment.UserName.ToString
    Protected m_DataSetDict As Dictionary(Of String, ESRI.ArcGIS.Geodatabase.IDataset) = New Dictionary(Of String, ESRI.ArcGIS.Geodatabase.IDataset)
    Private m_RequiredCodeVersion As String = Nothing



#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "a40ca571-a1c9-4279-835e-58a215dab935"
    Public Const InterfaceId As String = "eb3677e8-f388-4243-a0a2-4173d813139b"
    Public Const EventsId As String = "af29e429-be9e-4466-8228-13ec13b0dea1"
#End Region

#Region "Creation Subs"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.

    'Protected Sub New()
    'End Sub
    ''' <summary>
    ''' A subroutine for creating a new object
    ''' </summary>
    Protected Sub New()
        MyBase.New()
        ConnectionStatus = ConnectionStatus.NotConnected
    End Sub

#End Region

#Region "Connection Subs"
    ''' <summary>
    ''' The ConnectionStatus property
    ''' </summary>
    ''' <returns>Returns the property</returns>
    Friend Property ConnectionStatus() As Qdi.DataAccess.ConnectionStatus
        Get
            Return m_ConnectionStatus
        End Get
        Set(ByVal value As ConnectionStatus)
            m_ConnectionStatus = value
        End Set
    End Property
    ''' <summary>
    ''' A subroutine for connecting code
    ''' </summary>
    ''' <param name="pCodeVersion">The code version</param>
    Private Sub Connect(ByRef pCodeVersion As String) 'Implements IDataAccess.Connect
        If (Workspace Is Nothing) Or (ConnectionStatus = ConnectionStatus.ConnectedWithOldCode) Then
            'Dim pUserName As String
            Dim pLoginForm As New QdiLoginForm()
            Dim pResult As System.Windows.Forms.DialogResult

            pResult = pLoginForm.ShowDialog()

            If (pResult <> Windows.Forms.DialogResult.OK) Then
                Exit Sub
            End If

            'pUserName = pLoginForm.UsernameTextBox.Text.Trim()

            Try
                Dim pWorkSpace As IWorkspace = GetWorkspace()
                Me.Workspace = pWorkSpace
                ConnectionStatus = ConnectionStatus.ConnectedViewer

                If (requiredCodeVersion.ToLower = pCodeVersion.ToLower) Then
                    SwitchToEditorWorkspace()
                    'm_UserName = pUserName
                Else

                    ConnectionStatus = ConnectionStatus.ConnectedWithOldCode
                End If

            Catch ex As Exception
                Throw ex
            End Try
        End If
    End Sub

    'Protected MustOverride Sub SwitchToEditorWorkspace(ByVal userName As String) 'Implements IDataAccess.SwitchToEditorWorkspace
    Protected MustOverride Sub SwitchToEditorWorkspace()
    Protected MustOverride Function GetWorkspace() As IWorkspace
    ''' <summary>
    ''' A subroutine for disconnecting code
    ''' </summary>
    Friend Sub Disconnect() 'Implements IDataAccess.Disconnect
        Try
            'If (Me.CurrentLoginStatus > ConnectionStatus.ConnectedViewer) Then
            '    'm_Workspace.ExecuteSQL("delete from gis.qdix where relateid in (select relateid from gis.qdix where type = 'SB' limit 10)")
            'End If
        Catch ex As Exception
            Throw ex
        Finally
            Me.Workspace = Nothing
        End Try
    End Sub
    ''' <summary>
    ''' A subroutine for updating a connection
    ''' </summary>
    ''' <param name="SkipDisconnect">boolean for if disconnection should be skipped or not</param>
    ''' <param name="CodeVersion">The code version</param>
    Public Sub UpdateConnection(ByVal SkipDisconnect As Boolean, ByVal CodeVersion As String) Implements IDataAccess.UpdateConnection

        If (Me.CurrentLoginStatus < Qdi.DataAccess.ConnectionStatus.ConnectedViewer) Then
            Try
                Me.Connect(CodeVersion)
            Catch ex As Exception
                Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Unable to Connect to Database", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Try
        Else
            If (SkipDisconnect <> True) Then
                Dim pDisconnectAnswer As MsgBoxResult
                Windows.Forms.MessageBox.Show("Are you sure you want to disconnect?", "Disconnect?", Windows.Forms.MessageBoxButtons.OKCancel, Windows.Forms.MessageBoxIcon.Question)
                If (pDisconnectAnswer <> MsgBoxResult.Ok) Then
                    Exit Sub
                Else
                    Try
                        Me.Disconnect()
                    Catch ex As Exception
                        Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Unable to Connect to Database", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                    End Try

                End If
            End If
        End If
    End Sub

#End Region

#Region "Version Control"
    ''' <summary>
    ''' The required code version property
    ''' </summary>
    ''' <returns>The required code version</returns>
    Protected ReadOnly Property requiredCodeVersion() As String Implements IDataAccess.RequiredCodeVersion
        Get
            If (m_RequiredCodeVersion = Nothing) Then
                updateCodeVersions()
            End If
            Return m_RequiredCodeVersion
        End Get
    End Property
    ''' <summary>
    ''' A subroutine for updating code versions
    ''' </summary>
    Private Sub updateCodeVersions()

        Dim pCursor As ICursor = getTableCursor(BusinessLogic.NamedTables.qdvr, " ")

        Dim pRow As IRow
        pRow = pCursor.NextRow

        Dim pHighRequired As Double = -1
        Dim pHighNameRequired As String = Nothing

        Dim pName As String
        Dim pOrder As Double
        Dim pVersionNumIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, My.Resources.r_qdversion_versionnum_field)
        Dim pVersionNameIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, My.Resources.r_qdversion_versionname_field)

        While Not pRow Is Nothing

            pOrder = CType(pRow.Value(pVersionNumIndex), Double)
            pName = CType(pRow.Value(pVersionNameIndex), String)

            '*** Check to see if this is a required update & if it is newer
            If (pHighRequired < 0) Then
                pHighRequired = pOrder
                pHighNameRequired = pName
            ElseIf (pHighRequired < pOrder) Then
                pHighRequired = pOrder
                pHighNameRequired = pName
            End If

            pRow = pCursor.NextRow
        End While

        If Not (pHighNameRequired Is Nothing) Then
            m_RequiredCodeVersion = pHighNameRequired
        End If

    End Sub
    ''' <summary>
    ''' A subroutine for getting the table cursor
    ''' </summary>
    ''' <param name="pNamedTableType">The type of named table</param>
    ''' <param name="pWhereClause">The 'where' cause as a string</param>
    ''' <returns></returns>
    Friend Function getTableCursor(ByVal pNamedTableType As Qdi.BusinessLogic.NamedTables, ByVal pWhereClause As String) As ICursor
        Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
        pTable = GetTable(DatabaseSpecificTableName(pNamedTableType))

        Dim pQueryFilter As IQueryFilter
        pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = pWhereClause

        Dim pCursor As ICursor = pTable.Search(pQueryFilter, True)

        Return pCursor

    End Function

#End Region

#Region "Records Selections Subs"
    ''' <summary>
    ''' A function that allows for spatial selection
    ''' </summary>
    ''' <param name="geometry">The geometric type</param>
    ''' <param name="pOneFeature">An optional feature boolean that seems to always be passed in as false</param>
    ''' <returns></returns>
    Private Function SpatialSelect(ByVal geometry As IGeometry, Optional ByVal pOneFeature As Boolean = False) As Qdi.BusinessLogic.IQdiRecord 'Implements IDataAccess.SpatialSelect

        Dim pSpatialIdentifier As ISpatialIdentifier = New SpatialIdentifier
        Dim pFeature As IFeature

        Dim featureSelection As ESRI.ArcGIS.Carto.IFeatureSelection = TryCast(Me.qdixFeatureLayer, ESRI.ArcGIS.Carto.IFeatureSelection)

        If Not (TypeOf geometry Is IEnvelope) Then
            pOneFeature = True
        End If
        pFeature = pSpatialIdentifier.SpatialSelectFeature(geometry, Me.qdixFeatureLayer)

        If (pOneFeature = True) Then
            If Not (pFeature Is Nothing) Then
                Return ConvertFeatureToQdiRecord(pFeature)
            End If
        End If

        Return Nothing
    End Function
    ''' <summary>
    ''' A function for spatially selecting one record
    ''' </summary>
    ''' <param name="pActiveView">The active view</param>
    ''' <param name="geometry">The geometric type</param>
    ''' <param name="addSelect">A boolean for addSelect. This is an optional param.</param>
    ''' <returns>Returns the QdiRecord</returns>
    Public Function SpatialSelectOneRecord(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal geometry As IGeometry, Optional ByVal addSelect As Boolean = False) As Qdi.BusinessLogic.IQdiRecord Implements IDataAccess.SpatialSelectOneRecord
        If (addSelect = False) Then
            ClearSelectedMapFeatures(pActiveView, Me.qdixFeatureLayer)
        End If

        Dim pQdiRecord As Qdi.BusinessLogic.IQdiRecord = SpatialSelect(geometry)
        If Not pQdiRecord Is Nothing Then
            SelectRelateId(pActiveView, pQdiRecord.RelateId)
        End If

        Return pQdiRecord
    End Function
    ''' <summary>
    ''' A function for spatially selecting multiple records
    ''' </summary>
    ''' <param name="pActiveView">The active view</param>
    ''' <param name="pGeometry">The geometric type</param>
    ''' <param name="addSelect">A boolean for addSelect. This is an optional param.</param>
    Public Sub SpatialSelectManyRecords(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal pGeometry As ESRI.ArcGIS.Geometry.IGeometry, Optional ByVal addSelect As Boolean = False) Implements IDataAccess.SpatialSelectManyRecords
        Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
        pDataset = GetFeatureLayerInView(pActiveView, BusinessLogic.NamedFeatureClass.qdix)

        Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        pFeatureLayer = CType(pDataset, ESRI.ArcGIS.Carto.IFeatureLayer)

        'Dim pFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
        'pFeatureClass =  CType(pFeatureLayer., ESRI.ArcGIS.Geodatabase.IFeatureClass)
        ' create a spatial query filter
        Dim spatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilterClass()

        ' specify the geometry to query with
        spatialFilter.Geometry = pGeometry

        ' specify what the geometry field is called on the Feature Class that we will be querying against
        Dim nameOfShapeField As System.String = pFeatureLayer.FeatureClass.ShapeFieldName
        spatialFilter.GeometryField = nameOfShapeField

        ' specify the type of spatial operation to use
        spatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

        ' create the where statement
        'spatialFilter.WhereClause = whereClause

        ' perform the query and use a cursor to hold the results
        Dim queryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilterClass()
        queryFilter = CType(spatialFilter, ESRI.ArcGIS.Geodatabase.IQueryFilter)
        'Dim featureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor = FeatureClass.Search(queryFilter, False)

        Dim featureSelection As ESRI.ArcGIS.Carto.IFeatureSelection = TryCast(pFeatureLayer, ESRI.ArcGIS.Carto.IFeatureSelection)
        If (addSelect = True) Then
            featureSelection.SelectFeatures(queryFilter, esriSelectionResultEnum.esriSelectionResultAdd, False)
        Else
            featureSelection.SelectFeatures(queryFilter, esriSelectionResultEnum.esriSelectionResultNew, False)
        End If

        pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)

    End Sub
    ''' <summary>
    ''' A subroutine for zooming to selected qdi records within the active view
    ''' </summary>
    ''' <param name="pActiveView">The active view object</param>
    Sub ZoomToSelectedQDIXRecords(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView) Implements IDataAccess.ZoomToSelectedQDIXRecords

        Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
        pDataset = GetFeatureLayerInView(pActiveView, BusinessLogic.NamedFeatureClass.qdix)

        Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        pFeatureLayer = CType(pDataset, ESRI.ArcGIS.Carto.IFeatureLayer)

        'Dim pLayer As IFeatureLayer
        Dim pFSel As IFeatureSelection
        pFSel = CType(pFeatureLayer, IFeatureSelection)

        'Get the selected features
        Dim pSelSet As ISelectionSet
        pSelSet = pFSel.SelectionSet

        Dim pEnumGeom As IEnumGeometry
        Dim pEnumGeomBind As IEnumGeometryBind

        pEnumGeom = New EnumFeatureGeometry
        pEnumGeomBind = CType(pEnumGeom, IEnumGeometryBind)

        pEnumGeomBind.BindGeometrySource(Nothing, pSelSet)

        Dim pGeomFactory As IGeometryFactory
        pGeomFactory = CType(New GeometryEnvironment, IGeometryFactory)

        Dim pGeom As IGeometry
        pGeom = pGeomFactory.CreateGeometryFromEnumerator(pEnumGeom)

        If Not pGeom.IsEmpty Then
            Dim pEnvelope As IEnvelope2 = CType(pGeom.Envelope, IEnvelope2)

            pEnvelope.Expand(1000, 1000, False)

            pActiveView.Extent = pEnvelope
            pActiveView.Refresh()
        End If

    End Sub

    ''' <summary>
    ''' A subroutine for selecting a related id.
    ''' </summary>
    ''' <param name="pActiveView">The active view object</param>
    ''' <param name="pRelateId">The id to select</param>
    ''' <param name="SelectionType">The type of selection</param>
    Public Sub SelectRelateId(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal pRelateId As String, Optional ByVal SelectionType As esriSelectionResultEnum = esriSelectionResultEnum.esriSelectionResultNew) Implements IDataAccess.SelectRelateId
        Dim pDataset As ESRI.ArcGIS.Geodatabase.IDataset
        pDataset = GetFeatureLayerInView(pActiveView, BusinessLogic.NamedFeatureClass.qdix)

        Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        pFeatureLayer = CType(pDataset, ESRI.ArcGIS.Carto.IFeatureLayer)

        Dim queryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilterClass
        queryFilter.WhereClause = "relateid = '" + pRelateId + "'"

        Dim featureSelection As ESRI.ArcGIS.Carto.IFeatureSelection = TryCast(pFeatureLayer, ESRI.ArcGIS.Carto.IFeatureSelection)
        featureSelection.SelectFeatures(queryFilter, SelectionType, False)

        pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
    End Sub
    ''' <summary>
    ''' A subroutine that clears the selected map features within the active view's feature layer
    ''' </summary>
    ''' <param name="activeView">The active view object</param>
    ''' <param name="featureLayer">The current feature layer</param>
    Public Sub ClearSelectedMapFeatures(ByVal activeView As ESRI.ArcGIS.Carto.IActiveView, ByVal featureLayer As ESRI.ArcGIS.Carto.IFeatureLayer)

        If activeView Is Nothing OrElse featureLayer Is Nothing Then
            Return
        End If

        Dim featureSelection As ESRI.ArcGIS.Carto.IFeatureSelection = TryCast(featureLayer, ESRI.ArcGIS.Carto.IFeatureSelection) ' Dynamic Cast

        ' Invalidate only the selection cache. Flag the original selection
        activeView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)

        ' Clear the selection
        featureSelection.Clear()

        ' Flag the new selection
        activeView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)

    End Sub


#End Region

#Region "Record Manipulation Subs"
    ''' <summary>
    ''' A subroutine that adds a record
    ''' </summary>
    ''' <param name="qdiRecord">The record to add</param>
    Public Sub Add(ByRef qdiRecord As BusinessLogic.IQdiRecord) Implements IDataAccess.Add

        If (qdiRecord.IsAdd = False) Then
            Exit Sub
        End If

        If (Me.CurrentLoginStatus < ConnectionStatus.ConnectedEditor) Then
            Exit Sub
        End If

        MakeEdit(qdiRecord, New Row())

    End Sub
    ''' <summary>
    ''' A subroutine that deletes a record
    ''' </summary>
    ''' <param name="qdiRecord">The record to delete</param>
    Public Sub Delete(ByRef qdiRecord As BusinessLogic.IQdiRecord) Implements IDataAccess.Delete
        If (qdiRecord.IsUpdate() = False) Then
            Exit Sub
        End If

        If (Me.CurrentLoginStatus < ConnectionStatus.ConnectedEditor) Then
            Exit Sub
        End If

        Try
            qdiRecord.RelateId = "00Q9999999"
            Update(qdiRecord)

            Dim pFeature As IFeature
            pFeature = ReadFeature(qdiRecord.ObjectId)
            Dim pWorkspaceEdit As IWorkspaceEdit
            pWorkspaceEdit = CType(Me.Workspace, IWorkspaceEdit)
            Dim pUneditAtEnd As Boolean = False



            If (Not pFeature Is Nothing) Then
                Dim pRow As IRow
                pRow = CType(pFeature, IRow)

                If (Not pRow Is Nothing) Then
                    Try
                        If Not (pWorkspaceEdit.IsBeingEdited) Then
                            pWorkspaceEdit.StartEditing(True)
                            pWorkspaceEdit.StartEditOperation()
                            pUneditAtEnd = True
                        End If
                        pRow.Delete()

                    Catch ex As Exception
                        Throw (New CustomException("Error, unable to delete original record with [ObjectId] = " + qdiRecord.ObjectId.ToString))
                    Finally
                        If (pUneditAtEnd = True) Then
                            SaveEdits(pWorkspaceEdit)
                            'pWorkspaceEdit.StopEditOperation()
                            'pWorkspaceEdit.StopEditing(True)
                        End If
                    End Try
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Subroutine for updating a record
    ''' </summary>
    ''' <param name="qdiRecord">The qdiRecord to update</param>
    Public Sub Update(ByRef qdiRecord As BusinessLogic.IQdiRecord) Implements IDataAccess.Update

        If (qdiRecord.IsUpdate() = False) Then
            Exit Sub
        End If

        If (Me.CurrentLoginStatus < ConnectionStatus.ConnectedEditor) Then
            Exit Sub
        End If
        Try
            Dim pFeature As IFeature
            pFeature = ReadFeature(qdiRecord.ObjectId)

            If (Not pFeature Is Nothing) Then
                Dim pRow As IRow
                pRow = CType(pFeature, IRow)
                MakeEdit(qdiRecord, pRow)
            Else
                Throw (New CustomException("Error, unable to locate original record with [ObjectId] = " + qdiRecord.ObjectId.ToString))
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    ''' <summary>
    ''' A subroutine for editing a qdiRecord within a specific row
    ''' </summary>
    ''' <param name="qdiRecord">The record to edit</param>
    ''' <param name="row">The row of the desired record</param>
    Public Sub MakeEdit(ByRef qdiRecord As BusinessLogic.IQdiRecord, ByRef row As IRow)

        If (Me.CurrentLoginStatus < ConnectionStatus.ConnectedEditor) Then
            Exit Sub
        End If

        Dim pRelateId As String = qdiRecord.RelateId
        Dim pIRow As ESRI.ArcGIS.Geodatabase.IRow
        Dim pUneditAtEnd As Boolean = False
        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = CType(Me.Workspace, IWorkspaceEdit)

        Try
            If Not (pWorkspaceEdit.IsBeingEdited) Then
                pWorkspaceEdit.StartEditing(True)
                pWorkspaceEdit.StartEditOperation()
                pUneditAtEnd = True
            End If

            If (qdiRecord.IsAdd) Then
                Dim p_HookLayer As ESRI.ArcGIS.Geodatabase.IDataset
                Dim p_HookTable As ESRI.ArcGIS.Geodatabase.ITable
                p_HookLayer = CType(Me.qdixFeatureLayer, ESRI.ArcGIS.Geodatabase.IDataset)
                p_HookTable = CType(p_HookLayer, ESRI.ArcGIS.Geodatabase.ITable)
                pIRow = p_HookTable.CreateRow()
            Else
                pIRow = row
            End If

            '**20110214 Dim OriginalPropertyList As PropertyInfo() = GetType(Qdi.BusinessLogic.IQdiRecord).GetProperties()
            Dim OriginalPropertyList As System.Reflection.PropertyInfo() = GetType(Qdi.BusinessLogic.IQdiRecord).GetProperties()
            Dim PropertyToFieldMapping As Dictionary(Of String, String) = Me.PropertyToFieldMapping.PropertyToFieldMapping
            Dim pPropertyName As String
            Dim pFieldName As String
            Dim pFieldIndex As Integer
            Dim pPropertyValue As Object = Nothing
            Dim pPropertyInfo As System.Reflection.PropertyInfo

            For iPropertyIndex As Integer = 0 To OriginalPropertyList.Length - 1
                pPropertyInfo = OriginalPropertyList(iPropertyIndex)
                pPropertyName = pPropertyInfo.Name

                If (PropertyToFieldMapping.ContainsKey(pPropertyName)) Then
                    pFieldName = PropertyToFieldMapping(pPropertyName)

                    If (pFieldIndex > -1) Then
                        If pPropertyInfo.GetIndexParameters().Length = 0 Then
                            If (pPropertyInfo.CanRead) Then
                                Try
                                    pPropertyValue = pPropertyInfo.GetValue(qdiRecord, Nothing)
                                    If (pPropertyValue Is Nothing) Then
                                        'pPropertyValue = DBNull.Value
                                    End If
                                    WriteValueToField(pFieldName, pIRow, pPropertyValue)
                                Catch ex As Exception
                                    Windows.Forms.MessageBox.Show("Write Error:" + pFieldName + ", " + pPropertyValue.ToString, "Error Writing Data", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                                End Try
                            End If
                        End If
                    End If
                End If

            Next iPropertyIndex

            WriteValueDictionary(qdiRecord, pIRow)

            pIRow.Store()


        Catch ex As Exception
            Throw ex
        Finally

            If (pUneditAtEnd = True) Then
                SaveEdits(pWorkspaceEdit)
                'pWorkspaceEdit.StopEditOperation()
                'pWorkspaceEdit.StopEditing(True)
            End If

        End Try

    End Sub
    ''' <summary>
    ''' A subroutine for reconciling versions with each other
    ''' </summary>
    ''' <param name="pWorkspaceEdit">A workspace edit object</param>
    Private Sub ReconcileVersion(ByRef pWorkspaceEdit As IWorkspaceEdit)
        Dim version As IVersion = DirectCast(pWorkspaceEdit, IVersion)
        Dim versionName As [String] = version.VersionName

        ' Reconcile the version. Modify this code to reconcile and handle conflicts
        ' in a manner appropriate for the specific application.
        Dim versionEdit4 As IVersionEdit4 = DirectCast(pWorkspaceEdit, IVersionEdit4)
        Dim attempt As Integer = 1

        Try
            versionEdit4.Reconcile4(versionName, True, False, True, False)
        Catch ex2 As Exception
            MsgBox("Error Reconciling Edits" + vbNewLine + vbNewLine + ex2.Message)
        End Try
    End Sub
    'Private Function GetEditor() As IEditor
    '    'Dim uid As UID = New UIDClass()
    '    'uid.Value = "esriEditor.Editor"

    '    'Dim PTempApplication As ESRI.ArcGIS.Framework.IApplication
    '    'PTempApplication = TryCast(My.Application, ESRI.ArcGIS.Framework.IApplication)
    '    Dim editorUID As ESRI.ArcGIS.esriSystem.UID
    '    editorUID = New ESRI.ArcGIS.esriSystem.UID
    '    editorUID.Value = "esriEditor.Editor"

    '    Dim editor As IEditor3
    '    Dim a As ESRI.ArcGIS.Framework.IApplication = CType(My.Application, ESRI.ArcGIS.Framework.IApplication)
    '    Dim b As ESRI.ArcGIS.esriSystem.IExtension = a.FindExtensionByCLSID(editorUID)
    '    editor = CType(b, IEditor3)

    '    'Dim editorExtension As IExtension
    '    'editorExtension = app.FindExtensionByName("ESRI Object Editor")

    '    'ESRI.ArcGIS.Framework.IApplication.FindExtensionByCLSID(ESRI.ArcGIS.esriSystem.UID) As ESRI.ArcGIS.esriSystem.IExtension

    '    'a.FindExtensionByCLSID(editorUID)
    '    'editor = Applic

    '    'Dim editor As IEditor = CType(PTempApplication.FindExtensionByCLSID(uid), IEditor)

    '    Return editor
    'End Function
    ''' <summary>
    ''' A function that returns if a workspaceEdit is locked.
    ''' </summary>
    ''' <param name="pWorkspaceEdit">The workspace to edit</param>
    ''' <returns>A boolean for if the workspace is locked</returns>
    Private Function haveLock(ByRef pWorkspaceEdit As IWorkspaceEdit) As Boolean
        Dim pHaveLock As Boolean = False
        If (isLocked(pWorkspaceEdit, "", pHaveLock)) Then
            Return pHaveLock
        End If
        Return False
    End Function
    ''' <summary>
    ''' A function for if a workspace is locked
    ''' </summary>
    ''' <param name="pWorkspaceEdit">The workspace to edit</param>
    ''' <param name="pText">THe text being passed in</param>
    ''' <param name="bySelf">An optional boolean for if the lock is implemented by the workspace itself</param>
    ''' <returns>Whatever not pRow is nothing means. Extremely confusing, probably needs to be cleaned up.</returns>
    Private Function isLocked(ByRef pWorkspaceEdit As IWorkspaceEdit, ByRef pText As String, Optional ByRef bySelf As Boolean = False) As Boolean

        Dim pCursor As ICursor = getTableCursor(BusinessLogic.NamedTables.qdlk, " ")

        Dim pRow As IRow = Nothing
        Dim pResult As Boolean = False

        Try
            pRow = pCursor.NextRow
            pResult = True
            pText = "User: " + pRow.Value(pRow.Fields.FindField("who")).ToString
            pText += " on Machine: " + pRow.Value(pRow.Fields.FindField("machine")).ToString() + vbNewLine
            pText += " Since: " + pRow.Value(pRow.Fields.FindField("date")).ToString() + vbNewLine

            If Debugger.IsAttached Then
                pText += " GUID: " + pRow.Value(pRow.Fields.FindField("guid")).ToString() + vbNewLine
                pText += " My GUID: " + myGUID.ToString()
            End If

            If (pRow.Value(pRow.Fields.FindField("who")).ToString = System.Environment.UserName.ToString()) Then
                If (pRow.Value(pRow.Fields.FindField("machine")).ToString = System.Environment.MachineName.ToString()) Then
                    If (pRow.Value(pRow.Fields.FindField("guid")).ToString = myGUID.ToString) Then
                        bySelf = True
                    End If
                End If
            End If


            System.Runtime.InteropServices.Marshal.ReleaseComObject(pCursor)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pRow)
        Catch ex As Exception

        End Try

        Return Not (pRow Is Nothing)

    End Function

    ''' <summary>
    ''' A function that gets a lock for a workspace
    ''' </summary>
    ''' <param name="pWorkspaceEdit">the workspace to lock</param>
    ''' <returns></returns>
    Private Function GetLock(ByRef pWorkspaceEdit As IWorkspaceEdit) As Boolean
        Dim pWaitForm As Qdi.DataAccess.frmWaitingForm = New Qdi.DataAccess.frmWaitingForm
        Dim pFormText As String = ""
        Dim pForceLock As Boolean = False
        Dim pRowBuffer As IRowBuffer
        Dim pNewObject As Object

        While isLocked(pWorkspaceEdit, pFormText) And (pForceLock = False)
            System.Threading.Thread.Sleep(1000)
            pWaitForm.lbl_From.Text = pFormText
            Dim forceLockResults As System.Windows.Forms.DialogResult = pWaitForm.ShowDialog
            If (forceLockResults = Windows.Forms.DialogResult.Abort) Then
                pForceLock = True
            ElseIf (forceLockResults = Windows.forms.DialogResult.Cancel) Then
                Return False
            End If
        End While

        Try

            Dim pMWorkspaceEdit As IMultiuserWorkspaceEdit = CType(pWorkspaceEdit, IMultiuserWorkspaceEdit)

            If (pMWorkspaceEdit.SupportsMultiuserEditSessionMode(esriMultiuserEditSessionMode.esriMESMVersioned)) Then
                pMWorkspaceEdit.StartMultiuserEditing(esriMultiuserEditSessionMode.esriMESMVersioned)
            Else
                pWorkspaceEdit.StartEditing(False)
            End If

            pWorkspaceEdit.StartEditOperation()

            Dim pITable As ESRI.ArcGIS.Geodatabase.ITable = GetFeatureTable(BusinessLogic.NamedTables.qdlk)
            Dim pInsertCursor As ESRI.ArcGIS.Geodatabase.ICursor = pITable.Insert(True)


            pRowBuffer = pITable.CreateRowBuffer
            pRowBuffer.Value(pRowBuffer.Fields.FindField("who")) = System.Environment.UserName.ToString
            pRowBuffer.Value(pRowBuffer.Fields.FindField("machine")) = System.Environment.MachineName.ToString
            Dim ptemp As String = DateTime.Now.ToString
            pRowBuffer.Value(pRowBuffer.Fields.FindField("date")) = System.DateTime.Now.ToString("HH:mm:ss (MM/dd/yyyy)")
            pRowBuffer.Value(pRowBuffer.Fields.FindField("guid")) = myGUID.ToString


            'Goingto remove exising locks
            RemoveLock(pWorkspaceEdit)

            pNewObject = pInsertCursor.InsertRow(pRowBuffer)

            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pInsertCursor)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pRowBuffer)
            Catch ex As Exception

            End Try
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function
    ''' <summary>
    ''' A subroutine for removing a lock from a workspace
    ''' </summary>
    ''' <param name="pWorkspaceEdit">the workspace to unlock</param>
    Private Sub RemoveLock(ByRef pWorkspaceEdit As IWorkspaceEdit)

        Dim pCursorBuilder As mgs.CursorBuilder.IAttributeCursorBuilder
        pCursorBuilder = New mgs.CursorBuilder.AttributeCursorBuilder

        Dim pITable As ESRI.ArcGIS.Geodatabase.ITable = GetFeatureTable(BusinessLogic.NamedTables.qdlk)
        Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor = pCursorBuilder.ReturnUpdateCursor(pITable, "")
        Dim pRow As IRow = pCursor.NextRow

        While Not pRow Is Nothing
            pRow.Delete()
            pRow = pCursor.NextRow
        End While

        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pCursor)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pRow)
        Catch ex As Exception

        End Try

    End Sub


    'Public Sub SetEditable(ByRef pWorkspaceEdit As IWorkspaceEdit) Implements IDataAccess.SetEditable
    '    'Dim version As IVersion = DirectCast(pWorkspaceEdit, IVersion)
    '    'version.RefreshVersion()

    '    If Not (pWorkspaceEdit.IsBeingEdited) Then 'Not pObjClassInfo2.CanBypassEditSession Then
    '        Dim pMWorkspaceEdit As IMultiuserWorkspaceEdit = CType(pWorkspaceEdit, IMultiuserWorkspaceEdit)

    '        Dim editor As IEditor = GetEditor()

    '        If editor.EditState = esriEditState.esriStateNotEditing Then
    '            'If (pMWorkspaceEdit.SupportsMultiuserEditSessionMode(esriMultiuserEditSessionMode.esriMESMVersioned)) Then
    '            '    pMWorkspaceEdit.StartMultiuserEditing(esriMultiuserEditSessionMode.esriMESMVersioned)
    '            'Else
    '            Dim pWorkspace As IWorkspace = CType(pWorkspaceEdit, IWorkspace)
    '            editor.StartEditing(pWorkspace)
    '            'End If
    '        End If
    '        editor.StartOperation()
    '    End If

    'End Sub
    ''' <summary>
    ''' A function for setting a workspace as editable
    ''' </summary>
    ''' <param name="pWorkspaceEdit">the workspace to set as editable</param>
    ''' <returns></returns>
    Public Function SetEditable(ByRef pWorkspaceEdit As IWorkspaceEdit) As Boolean Implements IDataAccess.SetEditable

        While haveLock(pWorkspaceEdit) <> True
            If GetLock(pWorkspaceEdit) = False Then
                Return False
            End If
        End While

        If Not (pWorkspaceEdit.IsBeingEdited) Then
            Dim pMWorkspaceEdit As IMultiuserWorkspaceEdit = CType(pWorkspaceEdit, IMultiuserWorkspaceEdit)

            If (pMWorkspaceEdit.SupportsMultiuserEditSessionMode(esriMultiuserEditSessionMode.esriMESMVersioned)) Then
                pMWorkspaceEdit.StartMultiuserEditing(esriMultiuserEditSessionMode.esriMESMVersioned)
            Else
                pWorkspaceEdit.StartEditing(False)
            End If

            pWorkspaceEdit.StartEditOperation()
        End If

        Return True
    End Function
    ''' <summary>
    ''' A subroutine that saves edits
    ''' </summary>
    ''' <param name="pWorkspaceEdit">The workspace</param>
    ''' <param name="pInDataTable">An optional string to pass in</param>
    Public Sub SaveEdits(ByRef pWorkspaceEdit As IWorkspaceEdit, Optional ByVal pInDataTable As String = "") Implements IDataAccess.SaveEdits
        'Dim pEditor As IEditor = GetEditor()
        Try
            RemoveLock(pWorkspaceEdit)
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)
        Catch ex As Runtime.InteropServices.COMException

            MsgBox("Error Saving " + pInDataTable + " Edits, Will Try Again, You May Want To Restart ArcMap and Verify Edits Were Saved." + vbNewLine + ex.ErrorCode.ToString + vbNewLine + ex.Message, MsgBoxStyle.Critical)

            If ex.ErrorCode = CInt(fdoError.FDO_E_VERSION_REDEFINED) Then
                Try
                    ReconcileVersion(pWorkspaceEdit)
                    ' Stop the edit session.
                    pWorkspaceEdit.StopEditing(True)
                Catch ex2 As Exception
                    MsgBox("Error Saving Edits--Please Exit ArcMap and retry." + vbNewLine + vbNewLine + ex2.Message)
                    pWorkspaceEdit.StopEditing(False)
                End Try

            Else
                ' A different error has occurred. Handle in an appropriate way for the application.\
                MsgBox("Error Saving Edits--Probably Need to Exit ArcMap and retry." + ex.ErrorCode.ToString + vbNewLine + vbNewLine + ex.Message)
                pWorkspaceEdit.StopEditing(False)
            End If
        End Try


    End Sub

    'Public Sub SaveEditsIEditor(ByRef pWorkspaceEdit As IWorkspaceEdit, Optional ByVal pInDataTable As String = "") Implements IDataAccess.SaveEdits
    '    'Dim pEditor As IEditor = GetEditor()
    '    Try
    '        pEditor.StopOperation("edit")
    '        pEditor.StopEditing(True)
    '    Catch ex As Runtime.InteropServices.COMException

    '        MsgBox("Error Saving " + pInDataTable + " Edits, Will Try Again, You May Want To Restart ArcMap and Verify Edits Were Saved." + vbNewLine + ex.ErrorCode.ToString + vbNewLine + ex.Message, MsgBoxStyle.Critical)
    '        ' End If

    '        If ex.ErrorCode = CInt(fdoError.FDO_E_VERSION_REDEFINED) Then
    '            Try
    '                ReconcileVersion(pWorkspaceEdit)
    '                ' Stop the edit session.
    '                pEditor.StopEditing(True)
    '            Catch ex2 As Exception
    '                MsgBox("Error Saving Edits--Please Exit ArcMap and retry." + vbNewLine + vbNewLine + ex2.Message)
    '                pEditor.StopEditing(False)
    '            End Try

    '        Else
    '            ' A different error has occurred. Handle in an appropriate way for the application.\
    '            MsgBox("Error Saving Edits--Probably Need to Exit ArcMap and retry." + ex.ErrorCode.ToString + vbNewLine + vbNewLine + ex.Message)
    '            pEditor.StopEditing(False)
    '        End If
    '    End Try

    'End Sub
    ''' <summary>
    ''' A subroutine for writing a value to a field, with the field name taken in as a string
    ''' </summary>
    ''' <param name="pFieldName">The field's name (as a string)</param>
    ''' <param name="pIRow">The row</param>
    ''' <param name="pValue">The value to write</param>
    Private Sub WriteValueToField(ByVal pFieldName As String, ByVal pIRow As IRow, ByVal pValue As Object)
        Dim pFieldIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pIRow, pFieldName)

        'If (pFieldIndex = -1) Then
        '    pFieldIndex = Find FieldCaseInsensitive(pIRow, pFieldName)
        'End If

        If (pFieldIndex > -1) Then
            WriteValueToField(pFieldIndex, pIRow, pValue)
        End If
    End Sub
    ''' <summary>
    ''' Another subroutine for writing a value to a field, but with the field index taken in as an integer
    ''' </summary>
    ''' <param name="pFieldIndex">The field index as an integer</param>
    ''' <param name="pIRow">The row</param>
    ''' <param name="pValue">The value to write</param>
    Private Sub WriteValueToField(ByVal pFieldIndex As Integer, ByVal pIRow As IRow, ByVal pValue As Object)
        If (pFieldIndex > -1) Then

            Dim pField As IField
            pField = pIRow.Fields.Field(pFieldIndex)

            If Not (mgs.Domain.DomainHandler.GetCodedValueDomainFromField(pField) Is Nothing) And (Not pValue Is Nothing) Then
                Dim pValueTemp As Object = mgs.Domain.DomainHandler.ReturnCVDValue(mgs.Domain.DomainHandler.GetCodedValueDomainFromField(pField), pValue.ToString)
                If Not (pValueTemp Is Nothing) Then
                    pValue = pValueTemp
                End If
            End If

            Select Case pField.Type
                Case esriFieldType.esriFieldTypeGeometry
                    Dim pGeometry As ESRI.ArcGIS.Geometry.Point
                    Try
                        pGeometry = CType(pValue, ESRI.ArcGIS.Geometry.Point)
                        pIRow.Value(pFieldIndex) = pValue
                    Catch ex As Exception

                    End Try

                Case esriFieldType.esriFieldTypeString
                    If (pValue Is Nothing) Then
                        pValue = ""
                    End If
                    pIRow.Value(pFieldIndex) = pValue
                Case esriFieldType.esriFieldTypeInteger, esriFieldType.esriFieldTypeSmallInteger
                    Dim pInteger As Integer
                    If (Integer.TryParse(pValue.ToString, pInteger)) Then
                        pIRow.Value(pFieldIndex) = pInteger
                    Else
                        If (pValue Is Nothing) Then
                            'Dim a As DBNull = New DBNull
                            If (pField.IsNullable) Then
                                pIRow.Value(pFieldIndex) = DBNull.Value
                            End If

                        Else
                            Dim pDouble As Double
                            If (Double.TryParse(pValue.ToString, pDouble)) Then
                                Try
                                    pInteger = CType((pDouble - 0.5), Integer) '** I want to truncate so I subtract 0.5
                                    pIRow.Value(pFieldIndex) = pInteger
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                        End If
                Case esriFieldType.esriFieldTypeDouble
                    If (pValue Is Nothing) Then
                        'Dim a As DBNull = New DBNull
                        If (pField.IsNullable) Then
                            pIRow.Value(pFieldIndex) = DBNull.Value
                        End If

                    Else
                        Dim pDouble As Double
                        If (Double.TryParse(pValue.ToString, pDouble)) Then
                            pIRow.Value(pFieldIndex) = pDouble
                        End If
                    End If
                Case Else
                        Windows.Forms.MessageBox.Show(pField.Name + " Found Type: " + pField.Type.ToString, " (DBG 87)", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Select
        End If
    End Sub

    ''' <summary>
    ''' A subroutine that writes a value dictionary
    ''' </summary>
    ''' <param name="qdiRecord">The current qdi record</param>
    ''' <param name="pIRow">The row to write to a dictionary</param>
    Protected Sub WriteValueDictionary(ByRef qdiRecord As BusinessLogic.IQdiRecord, ByRef pIRow As IRow)
        Dim pValueDictionary As Dictionary(Of String, String)
        pValueDictionary = qdiRecord.ValueDictionary

        If Not (pValueDictionary Is Nothing) Then
            Dim pValue As String
            Dim pKey As String

            For Each pKey In pValueDictionary.Keys
                pValue = pValueDictionary(pKey)

                Try
                    WriteValueToField(pKey, pIRow, pValue)
                Catch ex As Exception
                    Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                End Try
            Next
        End If
    End Sub

#End Region

#Region "Read Data"
    ''' <summary>
    ''' A function for reading by related id
    ''' </summary>
    ''' <param name="relateId">The id as a string</param>
    ''' <returns></returns>
    Public Function ReadbyRelateId(ByRef relateId As String) As BusinessLogic.IQdiRecord Implements IDataAccess.ReadbyRelateId
        Return ConvertFeatureToQdiRecord(ReadFeature(relateId))
    End Function
    ''' <summary>
    ''' A function for reading by object id
    ''' </summary>
    ''' <param name="objectId">The id as a double</param>
    ''' <returns></returns>
    Public Function ReadbyObjectId(ByRef objectId As Double) As BusinessLogic.IQdiRecord Implements IDataAccess.ReadByOID
        Return ConvertFeatureToQdiRecord(ReadFeature(objectId))
    End Function
    ''' <summary>
    ''' A function for reading a feature, which is searched for by its object id
    ''' </summary>
    ''' <param name="ObjectId">The id of the object to search for</param>
    ''' <returns>The feature that was searched for, if it was found.</returns>
    Protected Function ReadFeature(ByVal ObjectId As Double) As IFeature
        Dim pFCursor As IFeatureCursor
        pFCursor = SearchByObjectId(ObjectId)

        Try
            Return ReadFeature(pFCursor)
        Catch ex As Exception
            Throw (New CustomException("Error, " + ex.ToString + " records with [ObjectId] = " + ObjectId.ToString + ") exist"))
        End Try

    End Function
    ''' <summary>
    ''' A function for reading a feature, which is searched for by its related Id
    ''' </summary>
    ''' <param name="relateId">The related id as a string</param>
    ''' <returns>The feature that was searched for, if it was found.</returns>
    Protected Function ReadFeature(ByVal relateId As String) As IFeature
        Dim pFCursor As IFeatureCursor
        pFCursor = SearchByRelateId(relateId)

        Try
            Return ReadFeature(pFCursor)
        Catch ex As Exception
            Throw (New CustomException("Error, " + ex.ToString + " records with [RelateId] = " + relateId + ") exist"))
        End Try

    End Function
    ''' <summary>
    ''' A function for reading a feature, which is searched for by its Feature cursor
    ''' </summary>
    ''' <param name="pFCursor">The feature cursor to search by</param>
    ''' <returns>The feature if it was found</returns>
    Protected Function ReadFeature(ByVal pFCursor As IFeatureCursor) As IFeature
        Dim returnFeature As IFeature = Nothing
        Dim pFeature As IFeature
        pFeature = pFCursor.NextFeature

        Dim int As Integer = 0

        While Not pFeature Is Nothing
            int += 1
            If Not (pFeature Is Nothing) Then
                returnFeature = pFeature
            End If

            pFeature = pFCursor.NextFeature
        End While

        If (int <> 1) Then
            Throw (New CustomException(int.ToString))
        End If
        Return returnFeature
    End Function
#End Region

#Region "Database Validation Operations"
    ''' <summary>
    ''' A function that returns if a related id exists
    ''' </summary>
    ''' <param name="relateId">The id to check to see if it exists</param>
    ''' <returns>true or false, depending on if the id exists or not</returns>
    Public Function RelateIdExists(ByVal relateId As String) As Boolean Implements IDataAccess.RelateIdExists

        Dim pFCursor As IFeatureCursor = SearchByRelateId(relateId)

        Dim pFeature As IFeature
        pFeature = pFCursor.NextFeature

        If pFeature Is Nothing Then
            Return False
        Else
            Return True
        End If

    End Function
    ''' <summary>
    ''' A function that returns a boolean based on if an object id can be updated to a related id
    ''' </summary>
    ''' <param name="objectId">The object id to check</param>
    ''' <param name="relateId">The related id to check</param>
    ''' <returns>If the object id can be updated to a related id (as a boolean, true or false)</returns>
    Public Function CanUpdateObjectIdtoRelateId(ByVal objectId As Double, ByVal relateId As String) As Boolean Implements IDataAccess.CanUpdateObjectIdtoRelateId
        Dim pFCursor As IFeatureCursor = SearchByObjectId(objectId)
        Dim pFeature As IFeature
        pFeature = pFCursor.NextFeature

        Dim pCurrentRelateId As String = ReturnStandardString(pFeature, "RelateId")

        Dim PCurrentRelateIdProper As String
        Dim pNewRelateIdProper As String

        Dim pRelateIdValidator As Qdi.BusinessLogic.IRelateIdValidator = New Qdi.BusinessLogic.RelateIdValidator(pCurrentRelateId)
        PCurrentRelateIdProper = pRelateIdValidator.BestGuessRelateId

        pRelateIdValidator.RelateId = relateId

        pNewRelateIdProper = pRelateIdValidator.BestGuessRelateId

        If (pNewRelateIdProper = PCurrentRelateIdProper) Then
            Return True
        Else
            If Me.RelateIdExists(pNewRelateIdProper) Then
                Return False
            End If
        End If

        Return True

    End Function
    ''' <summary>
    ''' A function for checking if an object id exists
    ''' </summary>
    ''' <param name="objectId">The object to check</param>
    ''' <returns>True or false, if the object id exists or not.</returns>
    Public Function ObjectIdExists(ByVal objectId As Double) As Boolean Implements IDataAccess.ObjectdIdExists

        Dim pFCursor As IFeatureCursor = SearchByObjectId(objectId)

        Dim pFeature As IFeature
        pFeature = pFCursor.NextFeature

        If pFeature Is Nothing Then
            Return False
        Else
            Return True
        End If

    End Function
    ''' <summary>
    ''' A function for searching by related id
    ''' </summary>
    ''' <param name="relateId">The id to search by</param>
    ''' <returns>The Feature cursor if found</returns>
    Public Function SearchByRelateId(ByVal relateId As String) As IFeatureCursor
        Dim pRelateIdValidator As Qdi.BusinessLogic.IRelateIdValidator
        pRelateIdValidator = New Qdi.BusinessLogic.RelateIdValidator(relateId)
        Dim pRelateId As String

        Try
            pRelateId = pRelateIdValidator.BestGuessRelateId
        Catch ex As Exception
            Throw ex
        End Try

        Dim sqlString As String
        sqlString = "relateid = '" & pRelateId & "'"

        Dim pFCursor As IFeatureCursor = ReturnIFeatureCursor(sqlString, Me.GetFeatureLayer(NamedFeatureClass.qdix))

        Return pFCursor
    End Function
    ''' <summary>
    ''' A function for returning an IFeatureCursor
    ''' </summary>
    ''' <param name="sqlString">The sql string to use in the query</param>
    ''' <param name="mFeatureLayer">The feature layer</param>
    ''' <returns>A feature cursor</returns>
    Private Function ReturnIFeatureCursor(ByVal sqlString As String, ByVal mFeatureLayer As IFeatureLayer) As IFeatureCursor
        Dim pFCursor As IFeatureCursor
        pFCursor = QueryFeatureClass(sqlString, mFeatureLayer)

        Return pFCursor
    End Function
    ''' <summary>
    ''' A function for searching by object id
    ''' </summary>
    ''' <param name="ObjectId">The id to search by</param>
    ''' <returns>A feature cursor</returns>
    Public Function SearchByObjectId(ByVal ObjectId As Double) As IFeatureCursor
        Dim sqlString As String = SelectByObjectID(ObjectId)

        Dim pFCursor As IFeatureCursor = ReturnIFeatureCursor(sqlString, Me.GetFeatureLayer(NamedFeatureClass.qdix))

        Return pFCursor
    End Function
    ''' <summary>
    ''' A function that gets the next related id
    ''' </summary>
    ''' <param name="sqlString">The sql string used in the query</param>
    ''' <returns>The next related id</returns>
    Friend Function NextRelateId(ByVal sqlString As String) As String
        Dim pCursorBuilder As mgs.CursorBuilder.IAttributeCursorBuilder
        pCursorBuilder = New mgs.CursorBuilder.AttributeCursorBuilder

        Dim pFCursor As IFeatureCursor = pCursorBuilder.ReturnFeatureCursor(Me.GetFeatureLayer(NamedFeatureClass.qdix), sqlString, "ORDER BY RELATEID", True)
        Dim pFeature As IFeature = pFCursor.NextFeature

        Dim pRelateIdFieldIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pFeature, "relateid") '2Do: XMLThis

        Dim pMaxRelateId As String = CType(pFeature.Value(pRelateIdFieldIndex), String)
        Dim iRelateId As String

        Dim pRelateIdValidator As IRelateIdValidator
        pRelateIdValidator = New RelateIdValidator("")

        Do Until pFeature Is Nothing

            iRelateId = CType(pFeature.Value(pRelateIdFieldIndex), String)
            If Not (iRelateId Is Nothing) Then
                pRelateIdValidator.RelateId = iRelateId
                If Not (pRelateIdValidator.IsInvalid) Then
                    If (pMaxRelateId < iRelateId) Then
                        pMaxRelateId = iRelateId
                    End If
                End If
            End If

            pFeature = pFCursor.NextFeature
        Loop

        pRelateIdValidator.RelateId = pMaxRelateId
        Dim pRelateId As String
        pRelateId = pRelateIdValidator.NextRelateId

        Return pRelateId
    End Function
    ''' <summary>
    ''' A function for returning an FCursor by query feature class
    ''' </summary>
    ''' <param name="queryString">The query string</param>
    ''' <param name="mFeatureLayer">The feature layer</param>
    ''' <returns>The FCursor</returns>
    Protected Function QueryFeatureClass(ByVal queryString As String, ByVal mFeatureLayer As IFeatureLayer) As IFeatureCursor
        Dim pQueryFilter As IQueryFilter
        pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = queryString

        Dim tempFeatureClass As IFeatureLayer
        tempFeatureClass = CType(Me.qdixFeatureLayer, IFeatureLayer)

        Dim pFCursor As IFeatureCursor
        pFCursor = mFeatureLayer.Search(pQueryFilter, False)

        Return pFCursor
    End Function
    ''' <summary>
    ''' A function that returns a query cursor
    ''' </summary>
    ''' <param name="whereString">a string for where to search</param>
    ''' <param name="pTableName">The table name</param>
    ''' <returns>A query cursor</returns>
    Protected Function QueryCursor(ByVal whereString As String, ByVal pTableName As Qdi.BusinessLogic.NamedTables) As ICursor
        Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
        pTable = GetTable(DatabaseSpecificTableName(pTableName))

        Return QueryCursor(whereString, pTable)
    End Function
    ''' <summary>
    ''' A function that returns a pCursor
    ''' </summary>
    ''' <param name="whereString">A string for where to search</param>
    ''' <param name="pTable">The table</param>
    ''' <returns></returns>
    Protected Function QueryCursor(ByVal whereString As String, ByVal pTable As ESRI.ArcGIS.Geodatabase.ITable) As ICursor

        Dim pQueryFilter As IQueryFilter
        pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = whereString

        Dim pCursor As ICursor
        pCursor = pTable.Search(pQueryFilter, True)

        Return pCursor
    End Function

#End Region

#Region "Properties"
    Protected MustOverride ReadOnly Property ViewerConnectionString() As String
    Protected MustOverride ReadOnly Property EditorConnectionString() As String
    Protected MustOverride ReadOnly Property QdiCursor() As IQdiCursorDefinition Implements IDataAccess.QdiCursor
    Protected MustOverride ReadOnly Property PropertyToFieldMapping() As IPropertyToFieldMappingList Implements IDataAccess.PropertyToFieldMapping

    ''' <summary>
    ''' The current login status property
    ''' </summary>
    ''' <returns>this property</returns>
    Public ReadOnly Property CurrentLoginStatus() As Qdi.DataAccess.ConnectionStatus Implements IDataAccess.CurrentLoginStatus
        Get
            If Not (Workspace Is Nothing) Then
                Return ConnectionStatus
            End If

            Return Qdi.DataAccess.ConnectionStatus.NotConnected
        End Get
    End Property

    '*** This was protected
    'Public ReadOnly Property tempIFeatureWorkspace() As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace Implements IDataAccess.QdiFeatureWorkspace
    '    Get
    '        Dim pFeatureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = Nothing

    '        Try
    '            pFeatureWorkspace = CType(Workspace, ESRI.ArcGIS.Geodatabase.IFeatureWorkspace)
    '        Catch ex As Exception
    '            Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
    '        End Try


    '        Return pFeatureWorkspace
    '    End Get
    'End Property
    ''' <summary>
    ''' The QdiWorkspace property
    ''' </summary>
    ''' <returns>this property</returns>
    Public ReadOnly Property QdiWorkspace() As ESRI.ArcGIS.Geodatabase.IWorkspace Implements IDataAccess.QdiWorkspace
        Get
            Return Workspace
        End Get
    End Property
    ''' <summary>
    ''' The workspace property
    ''' </summary>
    ''' <returns>this property</returns>
    Public Property Workspace() As IWorkspace
        Set(ByVal value As IWorkspace)
            m_Workspace = value
        End Set
        Get
            Return m_Workspace
        End Get
    End Property
    ''' <summary>
    ''' The OutdatedCode property
    ''' </summary>
    ''' <returns>this property as a boolean</returns>
    Protected ReadOnly Property OutdatedCode() As Boolean
        Get
            Return ConnectionStatus = ConnectionStatus.ConnectedWithOldCode
        End Get
    End Property
    ''' <summary>
    ''' The myGUID property
    ''' </summary>
    ''' <returns>this property as a Guid</returns>
    Protected ReadOnly Property myGUID() As Guid
        Get
            Return m_GUID
        End Get
    End Property
    ''' <summary>
    ''' The Editor property
    ''' </summary>
    ''' <returns>this property as a boolean</returns>
    Protected ReadOnly Property Editor() As Boolean
        Get
            Return ConnectionStatus > ConnectionStatus.ConnectedViewer
        End Get
    End Property
    ''' <summary>
    ''' The Administrator property, except they spelled it as "Administator" which is incorrect, but hopefully whenever this property is called, it is also spelled wrong.
    ''' </summary>
    ''' <returns>this property as a boolean</returns>
    Protected ReadOnly Property Administator() As Boolean
        Get
            Return ConnectionStatus = ConnectionStatus.ConnectedAdmin
        End Get
    End Property
    ''' <summary>
    ''' The qdixFeature Layer property
    ''' </summary>
    ''' <returns>this property</returns>
    Protected ReadOnly Property qdixFeatureLayer() As ESRI.ArcGIS.Carto.IFeatureLayer
        Get
            'If (m_QdixFeatureLayer Is Nothing) Then
            '    m_QdixFeatureLayer = GetFeatureLayer(BusinessLogic.NamedFeatureClass.qdix)
            'End If

            Return GetFeatureLayer(BusinessLogic.NamedFeatureClass.qdix) 'm_QdixFeatureLayer
        End Get

    End Property
    ''' <summary>
    ''' A subroutine for loading data into the active view
    ''' </summary>
    ''' <param name="activeView">The active view</param>
    Public Sub LoadDataIntoView(ByVal activeView As ESRI.ArcGIS.Carto.IActiveView) Implements IDataAccess.LoadDataIntoView

        If (CurrentLoginStatus > ConnectionStatus.NotConnected) Then


            LoadData(activeView, BusinessLogic.NamedFeatureClass.county, True)
            LoadData(activeView, BusinessLogic.NamedFeatureClass.Quad_24K, False)
            LoadData(activeView, BusinessLogic.NamedFeatureClass.sections, False)

            Dim pEnumLayer As ESRI.ArcGIS.Carto.IEnumLayer
            Dim pLayer As ILayer
            pEnumLayer = activeView.FocusMap.Layers

            pEnumLayer.Reset()
            pLayer = pEnumLayer.Next

            LoadData(activeView, BusinessLogic.NamedFeatureClass.qdix, True)
        End If
    End Sub
    ''' <summary>
    ''' A subroutine for loading data
    ''' </summary>
    ''' <param name="activeView">The active view</param>
    ''' <param name="layerName">The layer name</param>
    ''' <param name="setVisible">An optional boolean for setting the data visible</param>
    Protected Sub LoadData(ByVal activeView As ESRI.ArcGIS.Carto.IActiveView, ByVal layerName As Qdi.BusinessLogic.NamedFeatureClass, Optional ByVal setVisible As Boolean = True)
        If DataExists(activeView, layerName) Then
            Exit Sub
        End If

        Dim featureLayer As ESRI.ArcGIS.Carto.IFeatureLayer = New ESRI.ArcGIS.Carto.FeatureLayerClass
        featureLayer = GetFeatureLayer(layerName)
 
        If Not (featureLayer Is Nothing) Then
            featureLayer.Visible = setVisible
            activeView.FocusMap.AddLayer(featureLayer)
        End If

    End Sub

    ''' <summary>
    ''' A function for converting a feature to a QdiRecord
    ''' </summary>
    ''' <param name="feature">The feature to convert</param>
    ''' <returns>The converted feature as a QdiRecord</returns>
    Public Function ConvertFeatureToQdiRecord(ByVal feature As IFeature) As Qdi.BusinessLogic.IQdiRecord Implements IDataAccess.ConvertFeatureToQdiRecord
        Dim mCursorToQdiRecord As ICursorToQdiRecord
        mCursorToQdiRecord = New CursorToQdiRecord(Nothing, Me.PropertyToFieldMapping)

        Dim pQdiRecord As Qdi.BusinessLogic.IQdiRecord = mCursorToQdiRecord.ConvertFeatureIntoQdiRecord(feature)
        mCursorToQdiRecord = Nothing

        Return pQdiRecord
    End Function
    ''' <summary>
    ''' A function for returning a value
    ''' </summary>
    ''' <param name="pFeature">The feature</param>
    ''' <param name="fieldname">The fieldname</param>
    ''' <returns></returns>
    Protected Function ReturnValue(ByRef pFeature As IFeature, ByVal fieldname As String) As Object

        Dim pFieldIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pFeature, fieldname)


        Return pFeature.Value(pFieldIndex)
    End Function
    ''' <summary>
    ''' A function that returns a standard string
    ''' </summary>
    ''' <param name="feature">The feature</param>
    ''' <param name="fieldname">The field name</param>
    ''' <param name="NullValue">An optional null value which is a string. Very confusing.</param>
    ''' <returns></returns>
    Protected Function ReturnStandardString(ByRef feature As IFeature, ByVal fieldname As String, Optional ByVal NullValue As String = "") As String

        Dim value As Object = ReturnValue(feature, fieldname)
        Dim pString As String

        If (value Is System.DBNull.Value) Then
            pString = Nothing
        Else
            Try
                pString = value.ToString()
            Catch ex As Exception
                Throw ex
            End Try

        End If

        Return pString
    End Function

    'Protected Function Find FieldCaseInsensitive(ByVal row As IRow, ByVal fieldname As String) As Integer
    '    Dim iIndex As Integer = -1

    '    Dim pField As IField

    '    Dim iCounter As Integer

    '    For iCounter = 0 To row.Fields.FieldCount - 1
    '        pField = row.Fields.Field(iCounter)

    '        If (UCase(pField.Name) = UCase(fieldname)) Or (UCase(pField.AliasName) = UCase(fieldname)) Then
    '            Return iCounter
    '        End If
    '    Next

    '    If (iIndex = -1) Then
    '        Throw (New CustomException("Field (" + fieldname + ") Not Found"))
    '    End If
    '    Return iIndex
    'End Function

    'Protected Function Find FieldCaseInsensitive(ByVal feature As IFeature, ByVal fieldname As String) As Integer
    '    Dim pRow As IRow
    '    pRow = feature

    '    Return Find FieldCaseInsensitive(pRow, fieldname)

    'End Function
#End Region

#Region "Functions"
    Public MustOverride ReadOnly Property ReferenceMapFieldLookup() As Dictionary(Of String, String) Implements IDataAccess.ReferenceMapFieldLookup
    ''' <summary>
    ''' A function for getting a feature table
    ''' </summary>
    ''' <param name="featureClassName">the class name of the feature</param>
    ''' <returns>the ITable</returns>
    Public Function GetFTable(ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As ESRI.ArcGIS.Geodatabase.ITable Implements IDataAccess.GetFeatureTable
        Dim pFeatureClass As IFeatureClass
        pFeatureClass = GetFeatureClass(featureClassName)

        Dim pITable As ITable
        pITable = CType(pFeatureClass, ITable)

        Return pITable

    End Function
    ''' <summary>
    ''' A function for getting a feature table
    ''' </summary>
    ''' <param name="pTableName">the table's name</param>
    ''' <returns>the pTable</returns>
    Function GetFeatureTable(ByVal pTableName As Qdi.BusinessLogic.NamedTables) As ESRI.ArcGIS.Geodatabase.ITable Implements IDataAccess.GetFeatureTable
        Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
        pTable = GetTable(DatabaseSpecificTableName(pTableName))

        Return pTable
    End Function
    ''' <summary>
    ''' A function for getting a table
    ''' </summary>
    ''' <param name="tableName">the name of the table</param>
    ''' <returns>The pUserTable</returns>
    Protected Function GetTable(ByVal tableName As String) As ESRI.ArcGIS.Geodatabase.ITable

        Dim pUserTable As ESRI.ArcGIS.Geodatabase.ITable = Nothing
        Dim pDataSet As ESRI.ArcGIS.Geodatabase.IDataset = GetDataSet(tableName, esriDatasetType.esriDTTable)

        pUserTable = CType(pDataSet, ITable)


        Return pUserTable
    End Function
    ''' <summary>
    ''' Returns an ArcSde Workspace from a string passed in
    ''' </summary>
    ''' <param name="connectionString">the string for connecting the workspaces</param>
    ''' <returns></returns>
    Protected Function ArcSdeWorkspaceFromString(ByVal connectionString As String) As IWorkspace
        Dim workspaceFactory As IWorkspaceFactory2 = New SdeWorkspaceFactoryClass()
        Return workspaceFactory.OpenFromString(connectionString, 0)
    End Function
    ''' <summary>
    ''' A function for getting a feature class
    ''' </summary>
    ''' <param name="featureClassName">the name of the class</param>
    ''' <returns>the pFeature class</returns>
    Public Function GetFeatureClass(ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim pFeatureClass As IFeatureClass
        pFeatureClass = GetFeatureClass(DatabaseSpecificClassName(featureClassName))

        Return pFeatureClass
    End Function
    ''' <summary>
    ''' A function for getting a feature class
    ''' </summary>
    ''' <param name="featureLayerName">the feature's layer</param>
    ''' <returns>the pFeature class</returns>
    Protected Function GetFeatureClass(ByVal featureLayerName As String) As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim pDataSet As ESRI.ArcGIS.Geodatabase.IDataset = GetDataSet(featureLayerName, esriDatasetType.esriDTFeatureClass)

        Dim pFeatureClass As IFeatureClass
        pFeatureClass = CType(pDataSet, IFeatureClass)

        Return pFeatureClass
    End Function
    ''' <summary>
    ''' A function for getting a feature layer
    ''' </summary>
    ''' <param name="featureClassName">the class name of the feature</param>
    ''' <returns>the feature layer</returns>
    Protected Function GetFeatureLayer(ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As ESRI.ArcGIS.Carto.IFeatureLayer Implements IDataAccess.GetFeatureLayer
        Dim pFeatureClass As IFeatureClass = GetFeatureClass(featureClassName)
        Dim pFeatureLayer As IFeatureLayer

        If Not (pFeatureClass Is Nothing) Then
            pFeatureLayer = New FeatureLayer
            pFeatureLayer.FeatureClass = pFeatureClass
            pFeatureLayer.Name = pFeatureClass.AliasName
        Else
            pFeatureLayer = Nothing
        End If

        Return pFeatureLayer
    End Function
    ''' <summary>
    ''' A function for getting a feature layer
    ''' </summary>
    ''' <param name="featureLayerName">the feature layer's name</param>
    ''' <returns>the feature layer</returns>
    Protected Function GetFeatureLayer(ByVal featureLayerName As String) As ESRI.ArcGIS.Carto.IFeatureLayer
        Dim pFeatureLayer As IFeatureLayer
        pFeatureLayer = Nothing

        Dim pFeatureClass As IFeatureClass = GetFeatureClass(featureLayerName)

        pFeatureLayer = New FeatureLayer
        pFeatureLayer.FeatureClass = pFeatureClass
        pFeatureLayer.Name = pFeatureClass.AliasName

        Return pFeatureLayer
    End Function

    ''' <summary>
    ''' A function for getting a feature layer into the active view
    ''' </summary>
    ''' <param name="pActiveView">the active view</param>
    ''' <param name="featureClassName">the name of the feature class</param>
    ''' <returns>a Data Set</returns>
    Protected Function GetFeatureLayerInView(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As ESRI.ArcGIS.Geodatabase.IDataset Implements IDataAccess.GetFeatureLayerInView
        Dim pDataSet As IDataset = PickData(pActiveView, featureClassName)

        Return pDataSet

    End Function
    ''' <summary>
    ''' The DataSetDict property
    ''' </summary>
    ''' <returns>this property as a dictionary</returns>
    Private ReadOnly Property DataSetDict() As Dictionary(Of String, ESRI.ArcGIS.Geodatabase.IDataset)
        Get
            Return m_DataSetDict
        End Get
    End Property
    ''' <summary>
    ''' A subroutine for setting a key data set dict
    ''' </summary>
    ''' <param name="inKey">the key</param>
    ''' <param name="inDataset">the datasate</param>
    Private Sub SetKeyDataSetDict(ByVal inKey As String, ByVal inDataset As ESRI.ArcGIS.Geodatabase.IDataset)
        If Not (DataSetDict Is Nothing) Then
            If (Not DataSetDict.ContainsKey(inKey)) And (Not inDataset Is Nothing) Then
                DataSetDict.Add(inKey, inDataset)
            End If
        End If
    End Sub
    ''' <summary>
    ''' A function that gets a data set
    ''' </summary>
    ''' <param name="datasetName">The name of the data set</param>
    ''' <param name="datasetType">The type of the data set</param>
    ''' <returns>The data set</returns>
    Protected Function GetDataSet(ByVal datasetName As String, ByVal datasetType As esriDatasetType) As ESRI.ArcGIS.Geodatabase.IDataset
        If Not (DataSetDict Is Nothing) Then
            If (DataSetDict.ContainsKey(datasetName)) Then
                Return DataSetDict(datasetName)
            End If
        End If

        Dim pFinalDataSet As IDataset = Nothing
        Dim pEnumDataset As IEnumDataset
        pEnumDataset = Workspace.Datasets(esriDatasetType.esriDTAny)

        Dim pDataset As IDataset
        pDataset = pEnumDataset.Next

        While Not pDataset Is Nothing
            If (pDataset.Type = datasetType) Then
                If (isSame(datasetName, pDataset)) Then
                    pFinalDataSet = pDataset
                    Exit While
                End If
            ElseIf (pDataset.Type = esriDatasetType.esriDTFeatureDataset) Then
                Dim pEnumDatasetSubset As IEnumDataset
                pEnumDatasetSubset = pDataset.Subsets

                Dim pDatasubset As IDataset
                pDatasubset = pEnumDatasetSubset.Next

                While Not pDatasubset Is Nothing
                    If (isSame(datasetName, pDatasubset)) Then
                        pFinalDataSet = pDatasubset
                        Exit While
                    End If
                    pDatasubset = pEnumDatasetSubset.Next
                End While
            End If

            pDataset = pEnumDataset.Next
        End While

        If (pFinalDataSet Is Nothing) Then
            Throw (New CustomException("Error, unable to locate dataset: " + datasetName))
        End If

        SetKeyDataSetDict(datasetName, pFinalDataSet)
        'DataSetDict.Add(datasetName, pFinalDataSet)
        Return pFinalDataSet
    End Function
    ''' <summary>
    ''' A function that checks to see if a data set is the same
    ''' </summary>
    ''' <param name="targetName">the target's name</param>
    ''' <param name="dataset">the data set</param>
    ''' <returns>true or false depending on the result</returns>
    Private Function isSame(ByVal targetName As String, ByVal dataset As IDataset) As Boolean
        If (LCase(dataset.Name) = LCase(targetName)) Then
            Return True
        End If
        If (dataset.Type = esriDatasetType.esriDTFeatureClass) Then
            Try

                Dim featureClass As IFeatureClass = CType(dataset, IFeatureClass)

                If (LCase(featureClass.AliasName) = LCase(targetName)) Then
                    Return True
                End If


            Catch ex As Exception

            End Try
        End If

        Return False
    End Function
    ''' <summary>
    ''' A function for picking data to show in the active view
    ''' </summary>
    ''' <param name="pActiveView">the active view</param>
    ''' <param name="featureClassName">the feature class's name</param>
    ''' <returns></returns>
    Protected Function PickData(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As IDataset
        Dim iIndex As Integer
        Dim pMap As ESRI.ArcGIS.Carto.IMap = pActiveView.FocusMap
        Dim pDataset As IDataset = Nothing

        If (pMap.LayerCount > 0) Then
            Dim iMapLayer As ESRI.ArcGIS.Carto.ILayer

            ' Dim pFeatureLayer As IFeatureLayer


            Dim pPostgresLayerName As String = DatabaseSpecificClassName(featureClassName)

            For iIndex = 0 To (pMap.LayerCount - 1)
                iMapLayer = pMap.Layer(iIndex)

                If (LCase(iMapLayer.Name) = LCase(pPostgresLayerName)) Then
                    pDataset = CType(iMapLayer, IDataset)
                    'Return pDataset
                End If

                Try
                    Dim pDatasetTest As IDataset
                    pDatasetTest = CType(iMapLayer, IDataset)
                    If Not (pDatasetTest Is Nothing) Then
                        If (LCase(pDatasetTest.BrowseName) = LCase(pPostgresLayerName)) Then
                            pDataset = pDatasetTest
                        End If
                    End If

                Catch ex As Exception

                End Try

            Next
        End If


        If Not (DataSetDict Is Nothing) Then
            SetKeyDataSetDict(featureClassName.ToString, pDataset)
        End If

        Return pDataset
    End Function
    ''' <summary>
    ''' A function for checking if data exists or not
    ''' </summary>
    ''' <param name="pActiveView">the active view</param>
    ''' <param name="featureClassName">the feature class's name</param>
    ''' <returns>true or false depending on the result</returns>
    Protected Function DataExists(ByVal pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As Boolean
        Return (Not (PickData(pActiveView, featureClassName) Is Nothing))
    End Function

    Protected MustOverride Function DatabaseSpecificClassName(ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As String
    Protected MustOverride Function DatabaseSpecificTableName(ByVal tableName As Qdi.BusinessLogic.NamedTables) As String
    Protected MustOverride Function SelectByObjectID(ByVal tableName As Double) As String
    Public MustOverride Function NextQseriesId() As String Implements IDataAccess.NextQseriesID
    Public MustOverride Function NextUniqueWellId() As String Implements IDataAccess.NextUniqueWellId
#End Region




End Class




