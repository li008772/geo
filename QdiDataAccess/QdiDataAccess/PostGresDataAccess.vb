Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem

Public NotInheritable Class PostGresDataAccess

    Inherits DataAccessBaseClass

    Private m_ConnectionString As String = "SERVER=*SERVER*;DATABASE=*DATABASE*;INSTANCE=*INSTANCE*;USER=*USER*;PASSWORD=*PASSWORD*;VERSION=*VERSION*"

    Private m_Server As String = My.Resources.r_Server

    Private m_Version As String = My.Resources.r_Version
    Private m_FieldMapping As IPropertyToFieldMappingList = Nothing

    Private m_dataEditorUserName As String = My.Resources.r_dataEditorUserName
    Private m_dataEditorPassword As String = My.Resources.r_dataEditorPassword
    Private m_dataViewerUserName As String = My.Resources.r_dataViewerUserName
    Private m_dataViewerPassword As String = My.Resources.r_dataViewerPassword



#Region "Creation Subs"
    Friend Sub New()
        MyBase.New()
    End Sub
#End Region

    ''' <summary>
    ''' Returns the Arc Sde workspace
    ''' </summary>
    ''' <returns></returns>
    Protected Overrides Function GetWorkspace() As IWorkspace
        Dim pWorkSpace As IWorkspace
        pWorkSpace = ArcSdeWorkspaceFromString(Me.ViewerConnectionString)

        Return pWorkSpace

    End Function

    ''' <summary>
    ''' Switch to Acr SDE database
    ''' </summary>
    Protected Overrides Sub SwitchToEditorWorkspace()
        'Dim pCursor As ICursor = getTableCursor(BusinessLogic.NamedTables.qduf, "userid = '" + userName + "'")

        'Dim pRow As IRow
        'pRow = pCursor.NextRow

        'Dim pUserRight As String
        'Dim pUser As String

        'Dim pThisUserRight As String = Nothing

        'If (pRow Is Nothing) Then
        '    Windows.Forms.MessageBox.Show("User Name: " + userName + " is Invalid, Can Not Connect", "Can Not Connect", Windows.Forms.MessageBoxButtons.OK)
        'Else

        '    Dim pUserFieldIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, My.Resources.r_qduserfile_userid_field) '"userid") '2Done: XML This
        '    Dim pRightsFieldIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, My.Resources.r_qduserfile_userright_field) '2Done: XML This

        '    While Not pRow Is Nothing

        '        pUser = CType(pRow.Value(pUserFieldIndex), String)

        '        If (UCase(pUser) = UCase(userName)) Then
        '            pUserRight = CType(pRow.Value(pRightsFieldIndex), String)
        '            pThisUserRight = pUserRight
        '        End If

        '        pRow = pCursor.NextRow
        '    End While
        'End If

        'If (pThisUserRight = "E") Then
        ConnectionStatus = ConnectionStatus.ConnectedEditor
        'ElseIf (pThisUserRight = "A") Then
        'ConnectionStatus = ConnectionStatus.ConnectedAdmin
        'ElseIf (pThisUserRight Is Nothing) Then
        'Disconnect()
        'End If

        If (ConnectionStatus > ConnectionStatus.ConnectedViewer) Then
            Disconnect()
            Dim pWorkSpace As IWorkspace
            pWorkSpace = ArcSdeWorkspaceFromString(Me.EditorConnectionString)

            Me.Workspace = pWorkSpace
        End If
    End Sub
    'Protected Overrides Sub SwitchToEditorWorkspace(ByVal userName As String)

    '    Dim pCursor As ICursor = getTableCursor(BusinessLogic.NamedTables.qduf, "userid = '" + userName + "'")

    '    Dim pRow As IRow
    '    pRow = pCursor.NextRow

    '    Dim pUserRight As String
    '    Dim pUser As String

    '    Dim pThisUserRight As String = Nothing

    '    If (pRow Is Nothing) Then
    '        Windows.Forms.MessageBox.Show("User Name: " + userName + " is Invalid, Can Not Connect", "Can Not Connect", Windows.Forms.MessageBoxButtons.OK)
    '    Else

    '        Dim pUserFieldIndex As Integer = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, My.Resources.r_qduserfile_userid_field) '"userid") '2Done: XML This
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
    '        ConnectionStatus = ConnectionStatus.ConnectedEditor
    '    ElseIf (pThisUserRight = "A") Then
    '        ConnectionStatus = ConnectionStatus.ConnectedAdmin
    '    ElseIf (pThisUserRight Is Nothing) Then
    '        Disconnect()
    '    End If

    '    If (ConnectionStatus > ConnectionStatus.ConnectedViewer) Then
    '        Disconnect()
    '        Dim pWorkSpace As IWorkspace
    '        pWorkSpace = ArcSdeWorkspaceFromString(Me.EditorConnectionString)

    '        Me.Workspace = pWorkSpace
    '    End If
    'End Sub

    ''' <summary>
    ''' SDE database conneciton string
    ''' </summary>
    ''' <returns></returns>
    Protected Overrides ReadOnly Property ViewerConnectionString() As String
        Get
            Dim pConnectionString As String

            pConnectionString = m_ConnectionString.Replace("*USER*", m_dataViewerUserName).Replace("*PASSWORD*", m_dataViewerPassword).Replace("*INSTANCE*", QDIInstance())
            pConnectionString = pConnectionString.Replace("*SERVER*", m_Server).Replace("*DATABASE*", QDIDataBase()).Replace("*VERSION*", m_Version)

            Return pConnectionString
        End Get
    End Property

    ''' <summary>
    ''' SDE database editor connection string
    ''' </summary>
    ''' <returns></returns>
    Protected Overrides ReadOnly Property EditorConnectionString() As String
        Get
            Dim pConnectionString As String

            pConnectionString = m_ConnectionString.Replace("*USER*", m_dataEditorUserName).Replace("*PASSWORD*", m_dataEditorPassword).Replace("*INSTANCE*", QDIInstance())
            pConnectionString = pConnectionString.Replace("*SERVER*", m_Server).Replace("*DATABASE*", QDIDataBase()).Replace("*VERSION*", m_Version)

            Return pConnectionString
        End Get
    End Property

    ''' <summary>
    ''' Gets the current postgres cursor 
    ''' </summary>
    ''' <returns></returns>
    Protected Overrides ReadOnly Property QdiCursor() As IQdiCursorDefinition
        Get
            If (m_QdiCursor Is Nothing) Then
                m_QdiCursor = New PostGresQdiCursorDefinition(Me)
            End If
            Return MyBase.m_QdiCursor
        End Get
    End Property

    ''' <summary>
    ''' Get the property to field
    ''' </summary>
    ''' <returns></returns>
    Protected Overrides ReadOnly Property PropertyToFieldMapping() As IPropertyToFieldMappingList
        Get
            If (m_FieldMapping Is Nothing) Then
                Dim pDataAccessFactory As New Qdi.DataAccess.PropertyToFieldMappingFactory
                m_FieldMapping = pDataAccessFactory.Create(Qdi.DataAccess.PropertyToFieldMappingFactory.DatabaseType.PostGres)
            End If
            Return m_FieldMapping
        End Get
    End Property

    ''' <summary>
    ''' Map referenced field
    ''' </summary>
    ''' <returns></returns>
    Public Overrides ReadOnly Property ReferenceMapFieldLookup() As Dictionary(Of String, String)
        Get
            Dim pReferenceMapFieldLookup As Dictionary(Of String, String) = New Dictionary(Of String, String)
            pReferenceMapFieldLookup.Add("county_c", "countycode")
            pReferenceMapFieldLookup.Add("township", "town")
            pReferenceMapFieldLookup.Add("range", "rang")
            pReferenceMapFieldLookup.Add("section", "sect")
            pReferenceMapFieldLookup.Add("range_dir", "rngdir")
            pReferenceMapFieldLookup.Add("mgsquad_c", "mgsquad")
            'pReferenceMapFieldLookup.Add("CountyInt", "coun")
            'pReferenceMapFieldLookup.Add("Township", "town")
            'pReferenceMapFieldLookup.Add("Range", "rang")
            'pReferenceMapFieldLookup.Add("Section", "sect")
            'pReferenceMapFieldLookup.Add("range_dir", "rngdir")
            'pReferenceMapFieldLookup.Add("MgsQuadCode", "mgsquad")
            Return pReferenceMapFieldLookup

        End Get
    End Property

    ''' <summary>
    ''' Gets next relate id
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function NextQseriesId() As String
        Dim sqlString As String = "relateid like '00Q%'"

        Return NextRelateId(sqlString)
    End Function

    ''' <summary>
    ''' Gets next relate id not a q series
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function NextUniqueWellId() As String
        Dim sqlString As String = "relateid not like '00Q%'"

        Return NextRelateId(sqlString)
    End Function

    ''' <summary>
    ''' Gets the specific database name with fields
    ''' </summary>
    ''' <param name="featureClassName"></param>
    ''' <returns></returns>
    Protected Overrides Function DatabaseSpecificClassName(ByVal featureClassName As Qdi.BusinessLogic.NamedFeatureClass) As String
        Select Case featureClassName
            Case BusinessLogic.NamedFeatureClass.county, BusinessLogic.NamedFeatureClass.sections, BusinessLogic.NamedFeatureClass.Quad_24K
                Return QDIOwner() + [Enum].GetName(GetType(Qdi.BusinessLogic.NamedFeatureClass), featureClassName)
        End Select

        Return QDIOwner() + [Enum].GetName(GetType(Qdi.BusinessLogic.NamedFeatureClass), featureClassName)
    End Function

    ''' <summary>
    ''' Gets the qdi name and owner
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <returns></returns>
    Protected Overrides Function DatabaseSpecificTableName(ByVal tableName As Qdi.BusinessLogic.NamedTables) As String
        Return QDIOwner() + [Enum].GetName(GetType(Qdi.BusinessLogic.NamedTables), tableName)
    End Function

    ''' <summary>
    ''' Get object id
    ''' </summary>
    ''' <param name="ObjectId"></param>
    ''' <returns></returns>
    Protected Overrides Function SelectByObjectID(ByVal ObjectId As Double) As String
        Return "ObjectId = '" & ObjectId.ToString & "'"
    End Function

    ''' <summary>
    ''' Use enviorment test 
    ''' </summary>
    ''' <returns></returns>
    Private Function useTestData() As Boolean
        If Debugger.IsAttached And (My.Resources.r_testMode <> "True") Then
            Return True
        ElseIf (System.Environment.UserName.ToUpper = "MJRANTAL") And (My.Resources.r_forceDebugMode.ToUpper = "TRUE") Then
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Gets the QDI resource owner
    ''' </summary>
    ''' <returns></returns>
    Private Function QDIOwner() As String
        If useTestData() Then
            Return My.Resources.r_QDI_Owner_debug
        End If
        Return My.Resources.r_QDI_Owner
    End Function

    ''' <summary>
    ''' Gets the QDI database debug
    ''' </summary>
    ''' <returns></returns>
    Private Function QDIDataBase() As String
        If useTestData() Then
            Return My.Resources.r_QDI_database_debug
        End If
        Return My.Resources.r_QDI_database
    End Function

    ''' <summary>
    ''' Gets the QDI debug instance
    ''' </summary>
    ''' <returns></returns>
    Private Function QDIInstance() As String
        If useTestData() Then
            Return My.Resources.r_QDI_instance_debug
        End If
        Return My.Resources.r_QDI_instance
    End Function
End Class




