Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem

'<ComClass(PersonalGDBDataAccess.ClassId, PersonalGDBDataAccess.InterfaceId, PersonalGDBDataAccess.EventsId), _
' ProgId("QdiDataEntry.QdiDataAccess")> _
Public NotInheritable Class PersonalGDBDataAccess

    Inherits DataAccessBaseClass

    Private m_ConnectionString As String
    Private m_FieldMapping As IPropertyToFieldMappingList = Nothing


#Region "Creation Subs"

    '    ' A creatable COM class must have a Public Sub New() 
    '    ' with no parameters, otherwise, the class will not be 
    '    ' registered in the COM registry and cannot be created 
    '    ' via CreateObject.

    '    Private Sub New()
    '    End Sub

    Friend Sub New()
        MyBase.New()
    End Sub

    Friend Sub New(ByVal pConnectionString As String)
        MyBase.New()
        m_ConnectionString = pConnectionString
    End Sub
#End Region

    ''' <summary>
    '''     Several accessor and modifier methods, and other helper methods.
    ''' </summary>

    Protected Overrides Sub SwitchToEditorWorkspace()
        MsgBox("Not Implemented!")
    End Sub
    'Protected Overrides Sub SwitchToEditorWorkspace(ByVal userName As String)

    '    Dim pCursor As ICursor = getTableCursor(BusinessLogic.NamedTables.qduf, "userid = '" + userName + "'")
    '    'pCursor = pTable.Search(pQueryFilter, True)

    '    Dim pRow As IRow
    '    pRow = pCursor.NextRow

    '    Dim pUserRight As String
    '    Dim pUser As String

    '    Dim pThisUserRight As String = Nothing

    '    If (pRow Is Nothing) Then
    '        Windows.Forms.MessageBox.Show("User Name: " + userName + " is Invalid, Can Not Connect", "Can Not Connect", Windows.Forms.MessageBoxButtons.OK)
    '    Else
    '        Dim pUserFieldIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, My.Resources.r_qduserfile_userid_field) '2Done: XML This
    '        Dim pRightsFieldIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, My.Resources.r_qduserfile_userright_field) '2Done: XML This

    '        While Not pRow Is Nothing

    '            pUser = CType(pRow.Value(pUserFieldIndex), String)

    '            If (UCase(pUser) = UCase(userName)) Then
    '                pUserRight = CType(pRow.Value(pRightsFieldIndex), String)
    '                pThisUserRight = pUserRight
    '            End If

    '            pRow = pCursor.NextRow
    '        End While
    '    End If

    '    If (pThisUserRight = "E") Then
    '        Me.ConnectionStatus = ConnectionStatus.ConnectedEditor
    '    ElseIf (pThisUserRight = "A") Then
    '        Me.ConnectionStatus = ConnectionStatus.ConnectedAdmin
    '    ElseIf (pThisUserRight Is Nothing) Then
    '        Disconnect()
    '    End If

    '    'If (m_ConnectionStatus > ConnectionStatus.ConnectedViewer) Then
    '    '    Disconnect()
    '    '    GetWorkspace()
    '    '    Dim pWorkSpace As IWorkspace
    '    '    pWorkSpace = ArcSdeWorkspaceFromString(Me.EditorConnectionString)

    '    '    Me.Workspace = pWorkSpace
    '    'End If
    'End Sub

    Protected Overrides ReadOnly Property ViewerConnectionString() As String
        Get
            Return m_ConnectionString
        End Get
    End Property

    Protected Overrides ReadOnly Property EditorConnectionString() As String
        Get
            Return m_ConnectionString
        End Get
    End Property

    Protected Overrides ReadOnly Property QdiCursor() As IQdiCursorDefinition
        Get
            'If (m_QdiCursor Is Nothing) Then
            '    m_QdiCursor = New PersonalQdiCursorDefinition(Me)
            'End If
            Return New PersonalQdiCursorDefinition(Me) 'MyBase.m_QdiCursor
        End Get
    End Property

    Protected Overrides ReadOnly Property PropertyToFieldMapping() As IPropertyToFieldMappingList
        Get
            If (m_FieldMapping Is Nothing) Then
                Dim pDataAccessFactory As New Qdi.DataAccess.PropertyToFieldMappingFactory
                m_FieldMapping = pDataAccessFactory.Create(Qdi.DataAccess.PropertyToFieldMappingFactory.DatabaseType.PostGres)
            End If
            Return m_FieldMapping
        End Get
    End Property

    Public Overrides ReadOnly Property ReferenceMapFieldLookup() As Dictionary(Of String, String)
        Get
            Dim pReferenceMapFieldLookup As Dictionary(Of String, String) = New Dictionary(Of String, String)
            pReferenceMapFieldLookup.Add("county_c", "county_id")
            pReferenceMapFieldLookup.Add("township", "twp")
            pReferenceMapFieldLookup.Add("range", "rng")
            pReferenceMapFieldLookup.Add("range_dir", "rngdir")
            pReferenceMapFieldLookup.Add("section", "sec")
            pReferenceMapFieldLookup.Add("mgsquad_c", "mgsquad")
            Return pReferenceMapFieldLookup

        End Get
    End Property

    Public Overrides Function NextQseriesId() As String
        Dim sqlString As String = "relateid Like ""00Q*"""

        Return NextRelateId(sqlString)
    End Function

    Public Overrides Function NextUniqueWellId() As String
        Dim sqlString As String = "relateid not like ""00Q*"""

        Return NextRelateId(sqlString)
    End Function

    Protected Overrides Function DatabaseSpecificClassName(ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As String
        Return [Enum].GetName(GetType(Qdi.BusinessLogic.NamedFeatureClass), featureClassName)
    End Function

    Protected Overrides Function DatabaseSpecificTableName(ByVal tableName As Qdi.BusinessLogic.NamedTables) As String
        Return [Enum].GetName(GetType(Qdi.BusinessLogic.NamedTables), tableName)
    End Function

    Protected Overrides Function SelectByObjectID(ByVal ObjectId As Double) As String
        Return "ObjectId = " + ObjectId.ToString
    End Function

    ''' <summary>
    '''     Creates a Workspace object from the connection string (path of database).
    ''' </summary>
    ''' <returns>Information about the workspace container of datasets.</returns>
    Protected Overrides Function GetWorkspace() As IWorkspace
        Dim pWorkSpace As IWorkspace
        Try
            Dim pWorkspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactoryClass()
            pWorkSpace = pWorkspaceFactory.OpenFromFile(m_ConnectionString, 0)

        Catch ex As Exception
            Throw ex
        End Try


        Return pWorkSpace

    End Function


End Class




