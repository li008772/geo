Option Strict On
Option Explicit On

Imports ESRI.ArcGIS.Geodatabase

''' <summary>
''' Data fields of the qdi cursor
''' </summary>
Public MustInherit Class QdiCursorBaseClass
    Implements IQdiCursorDefinition

    Private m_ObjectId As Decimal
    Private m_DataAccess As IDataAccess
    Private m_RelateIdValidator As Qdi.BusinessLogic.IRelateIdValidator
    Private m_Township As Nullable(Of Decimal)
    Private m_Range As Nullable(Of Decimal)
    Private m_Section As Nullable(Of Decimal)
    Private m_DataSource As String
    Private m_FromDate As Nullable(Of Decimal)
    Private m_ToDate As Nullable(Of Decimal)
    Private m_FromDepth As Nullable(Of Decimal)
    Private m_ToDepth As Nullable(Of Decimal)
    Private m_Counties As List(Of String)
    Private m_Quadrangles As List(Of String)
    Private m_FirstBedrocks As List(Of String)
    Private m_LastBedrocks As List(Of String)
    Private m_FirstStratUnits As List(Of String)
    Private m_LastStratUnits As List(Of String)
    Private m_Aquifers As List(Of String)
    Private m_RelateIDs As List(Of String)

#Region "Creation"

    Protected Sub New()
    End Sub

    Protected Sub New(ByRef pDataAcess As IDataAccess)
        m_DataAccess = pDataAcess
    End Sub
#End Region

#Region "Properties"

    'Public MustOverride ReadOnly Property SqlString() As String Implements IQdiCursor.SQLstring

#Region "IQdiCursor Properties"

    ''' <summary>
    ''' m_ObjectId getter/setter. 
    ''' </summary>
    ''' <returns>m_ObjectId valuse</returns>
    Public Property ObjectId() As Decimal Implements IQdiCursorDefinition.ObjectId
        Get
            Return m_ObjectId
        End Get
        Set(ByVal value As Decimal)
            m_ObjectId = value
        End Set
    End Property

    ''' <summary>
    ''' Sets the m_RelateIdValidator
    ''' </summary>
    Public WriteOnly Property RelateId() As Qdi.BusinessLogic.IRelateIdValidator Implements IQdiCursorDefinition.RelateId
        Set(ByVal value As Qdi.BusinessLogic.IRelateIdValidator)

            If Not (value Is Nothing) Then
                If (value.IsInvalid) Then
                    Throw (New CustomException("Invalid [RelateId] in Query"))
                Else
                    m_RelateIdValidator = value
                End If
            Else
                m_RelateIdValidator = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' Refreshes all values, zeros and sets new
    ''' </summary>
    Public Sub Refresh() Implements IQdiCursorDefinition.Refresh
        Township = Nothing
        Range = Nothing
        Section = Nothing
        DataSource = Nothing
        FromDate = Nothing
        ToDate = Nothing
        FromDepth = Nothing
        ToDepth = Nothing
        Counties = New List(Of String)
        Quadrangles = New List(Of String)
        FirstBedrockUnits = New List(Of String)
        LastBedrockUnits = New List(Of String)
        FirstStratUnits = New List(Of String)
        LastStratUnits = New List(Of String)
        Aquifers = New List(Of String)
        RelateIDs = New List(Of String)
    End Sub

    ''' <summary>
    ''' m_Township getter/setter
    ''' </summary>
    ''' <returns>m_Township value</returns>
    Public Property Township() As Nullable(Of Decimal) Implements IQdiCursorDefinition.Township
        Set(ByVal value As Nullable(Of Decimal))
            m_Township = value
        End Set
        Get
            Return m_Township
        End Get
    End Property

    ''' <summary>
    ''' m_Range getter/setter
    ''' </summary>
    ''' <returns>m_Range value</returns>
    Public Property Range() As Nullable(Of Decimal) Implements IQdiCursorDefinition.Range
        Set(ByVal value As Nullable(Of Decimal))
            m_Range = value
        End Set
        Get
            Return m_Range
        End Get
    End Property

    ''' <summary>
    ''' m_Section getter/setter
    ''' </summary>
    ''' <returns>m_Section value</returns>
    Public Property Section() As Nullable(Of Decimal) Implements IQdiCursorDefinition.Section
        Set(ByVal value As Nullable(Of Decimal))
            m_Section = value
        End Set
        Get
            Return m_Section
        End Get
    End Property

    ''' <summary>
    ''' m_DataSource getter/setter
    ''' </summary>
    ''' <returns>m_DataSource value</returns>
    Public Property DataSource() As String Implements IQdiCursorDefinition.DataSource
        Set(ByVal value As String)
            m_DataSource = value
        End Set
        Get
            Return m_DataSource
        End Get
    End Property


    ''' <summary>
    ''' m_FromDate getter/setter
    ''' </summary>
    ''' <returns></returns>
    Public Property FromDate() As Nullable(Of Decimal) Implements IQdiCursorDefinition.FromDate
        Set(ByVal value As Nullable(Of Decimal))
            m_FromDate = value
        End Set
        Get
            Return m_FromDate
        End Get
    End Property

    ''' <summary>
    '''m_ToDate getter/setter
    ''' </summary>
    ''' <returns>m_ToDate value</returns>
    Public Property ToDate() As Nullable(Of Decimal) Implements IQdiCursorDefinition.ToDate
        Set(ByVal value As Nullable(Of Decimal))
            m_ToDate = value
        End Set
        Get
            Return m_ToDate
        End Get
    End Property

    ''' <summary>
    ''' m_FromDepth getter/setter
    ''' </summary>
    ''' <returns>m_FromDepth value</returns>
    Public Property FromDepth() As Nullable(Of Decimal) Implements IQdiCursorDefinition.FromDepth
        Set(ByVal value As Nullable(Of Decimal))
            m_FromDepth = value
        End Set
        Get
            Return m_FromDepth
        End Get
    End Property

    ''' <summary>
    ''' m_ToDepth getter/setter
    ''' </summary>
    ''' <returns>m_ToDepth value</returns>
    Public Property ToDepth() As Nullable(Of Decimal) Implements IQdiCursorDefinition.ToDepth
        Set(ByVal value As Nullable(Of Decimal))
            m_ToDepth = value
        End Set
        Get
            Return m_ToDepth
        End Get
    End Property

    ''' <summary>
    ''' m_Counties getter/setter
    ''' </summary>
    ''' <returns>m_Counties value</returns>
    Public Property Counties() As List(Of String) Implements IQdiCursorDefinition.Counties
        Set(ByVal value As List(Of String))
            m_Counties = value
        End Set
        Get
            Return m_Counties
        End Get
    End Property

    ''' <summary>
    ''' m_Quadrangles getter/setter
    ''' </summary>
    ''' <returns>m_Quadrangles value</returns>
    Public Property Quadrangles() As List(Of String) Implements IQdiCursorDefinition.Quadrangles
        Set(ByVal value As List(Of String))
            m_Quadrangles = value
        End Set
        Get
            Return m_Quadrangles
        End Get
    End Property

    ''' <summary>
    ''' m_FirstBedrocks getter/setter
    ''' </summary>
    ''' <returns>m_FirstBedrocks value</returns>
    Public Property FirstBedrockUnits() As List(Of String) Implements IQdiCursorDefinition.FirstBedrockUnits
        Set(ByVal value As List(Of String))
            m_FirstBedrocks = value
        End Set
        Get
            Return m_FirstBedrocks
        End Get
    End Property

    ''' <summary>
    ''' m_LastBedrocks getter/setter
    ''' </summary>
    ''' <returns>m_LastBedrocks value</returns>
    Public Property LastBedrockUnits() As List(Of String) Implements IQdiCursorDefinition.LastBedrockUnits
        Set(ByVal value As List(Of String))
            m_LastBedrocks = value
        End Set
        Get
            Return m_LastBedrocks
        End Get
    End Property

    ''' <summary>
    ''' m_FirstStratUnits getter/setter
    ''' </summary>
    ''' <returns>m_FirstStratUnits value</returns>
    Public Property FirstStratUnits() As List(Of String) Implements IQdiCursorDefinition.FirstStratUnits
        Set(ByVal value As List(Of String))
            m_FirstStratUnits = value
        End Set
        Get
            Return m_FirstStratUnits
        End Get
    End Property

    ''' <summary>
    ''' m_LastStratUnits getter/seetter
    ''' </summary>
    ''' <returns>m_LastStratUnits value</returns>
    Public Property LastStratUnits() As List(Of String) Implements IQdiCursorDefinition.LastStratUnits
        Set(ByVal value As List(Of String))
            m_LastStratUnits = value
        End Set
        Get
            Return m_LastStratUnits
        End Get
    End Property

    ''' <summary>
    ''' m_Aquifers getter/setter
    ''' </summary>
    ''' <returns>m_Aquifers value</returns>
    Public Property Aquifers() As List(Of String) Implements IQdiCursorDefinition.Aquifers
        Set(ByVal value As List(Of String))
            m_Aquifers = value
        End Set
        Get
            Return m_Aquifers
        End Get
    End Property

    ''' <summary>
    ''' m_RelateIDs getter/setter
    ''' </summary>
    ''' <returns>m_RelateIDs value</returns>
    Public Property RelateIDs() As List(Of String) Implements IQdiCursorDefinition.RelateIDs
        Set(ByVal value As List(Of String))
            m_RelateIDs = value
        End Set
        Get
            Return m_RelateIDs
        End Get
    End Property


    ''' <summary>
    ''' pValidationErrors getter/setter
    ''' </summary>
    ''' <returns>pValidationErrors value</returns>
    Public ReadOnly Property ValidationErrors() As List(Of String) Implements IQdiCursorDefinition.ValidationErrors
        Get
            Dim pValidationErrors As List(Of String)
            pValidationErrors = New List(Of String)

            If Not (FromDepth Is Nothing) Then

            End If
            Return pValidationErrors
        End Get

    End Property

#End Region

    ''' <summary>
    ''' m_RelateIdValidator value
    ''' </summary>
    ''' <returns>m_RelateIdValidator value</returns>
    Protected ReadOnly Property RelateIdValidator() As Qdi.BusinessLogic.IRelateIdValidator
        'Set(ByVal value As Qdi.BusinessLogic.IRelateIdValidator)
        '    m_RelateIdValidator = value
        'End Set
        Get
            Return m_RelateIdValidator
        End Get
    End Property

    ''' <summary>
    ''' m_DataAccess getter/setter
    ''' </summary>
    ''' <returns>m_DataAccess value</returns>
    Protected Property DataAccess() As IDataAccess
        Set(ByVal value As IDataAccess)
            m_DataAccess = value
        End Set
        Get
            Return m_DataAccess
        End Get
    End Property

#End Region

#Region "Functions"
    Public MustOverride Function QdixICursor() As ICursor Implements IQdiCursorDefinition.QdixICursor
#End Region


End Class
