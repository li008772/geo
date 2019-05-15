Option Strict On
Option Explicit On

''' <summary>
''' This maps the properties Of a QDIRecord To the actual fields In qdix.qdi
''' </summary>
Public Class PropertyToFieldMappingPostGres
    Implements Qdi.DataAccess.IPropertyToFieldMappingList

    Dim m_propertyToFieldMapping As Dictionary(Of String, String) = Nothing
    Dim m_fieldToPropertyMapping As Dictionary(Of String, String) = Nothing

    ''' <summary>
    ''' Map a property to a field for the database
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property PropertyToFieldMapping() As System.Collections.Generic.Dictionary(Of String, String) Implements IPropertyToFieldMappingList.PropertyToFieldMapping
        Get
            If (m_propertyToFieldMapping Is Nothing) Then
                BuildFieldMapping()
            End If
            Return m_propertyToFieldMapping
        End Get
    End Property

    ''' <summary>
    ''' Map field to a property 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property FieldToPropertyMapping() As System.Collections.Generic.Dictionary(Of String, String) Implements IPropertyToFieldMappingList.FieldToPropertyMapping
        Get
            If (m_fieldToPropertyMapping Is Nothing) Then
                m_fieldToPropertyMapping = New Dictionary(Of String, String)
                Dim pPropertyList As System.Collections.Generic.Dictionary(Of String, String).KeyCollection = Me.PropertyList
                Dim pField As String

                For Each pProperty As String In pPropertyList
                    pField = ReturnField(pProperty)
                    m_fieldToPropertyMapping.Add(pField, pProperty)
                Next
            End If
            Return m_fieldToPropertyMapping
        End Get
    End Property

    ''' <summary>
    ''' Returns the mapping field in the property
    ''' </summary>
    ''' <param name="inProperty"></param>
    ''' <returns></returns>
    Public Function ReturnField(ByVal inProperty As String) As String Implements IPropertyToFieldMappingList.ReturnFieldName
        If (PropertyToFieldMapping.ContainsKey(inProperty)) Then
            Return PropertyToFieldMapping(inProperty)
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Returns mappint property from dictionary
    ''' </summary>
    ''' <param name="inField"></param>
    ''' <returns></returns>
    Public Function ReturnProperty(ByVal inField As String) As String Implements IPropertyToFieldMappingList.ReturnProperty
        If (FieldToPropertyMapping.ContainsKey(inField)) Then
            Return FieldToPropertyMapping(inField)
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Returnd the property list keys
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property PropertyList() As System.Collections.Generic.Dictionary(Of String, String).KeyCollection Implements IPropertyToFieldMappingList.PropertyList
        Get
            Return PropertyToFieldMapping.Keys
        End Get
    End Property

    ''' <summary>
    ''' Returnd the field list
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property FieldList() As System.Collections.Generic.Dictionary(Of String, String).KeyCollection Implements IPropertyToFieldMappingList.FieldList
        Get
            Return Me.FieldToPropertyMapping.Keys
        End Get
    End Property

    ''' <summary>
    ''' Build a field mapping dictionary
    ''' </summary>
    Private Sub BuildFieldMapping()
        '*AP2QR* Add Property to Field
        '2DO: XML
        m_propertyToFieldMapping = New Dictionary(Of String, String)

        m_propertyToFieldMapping.Add("Point", "Shape")
        m_propertyToFieldMapping.Add("RelateId", "RelateId")
        m_propertyToFieldMapping.Add("WellName", "wellname")
        m_propertyToFieldMapping.Add("Elevation", "Elevation")
        m_propertyToFieldMapping.Add("UTMN", "UTMN")
        m_propertyToFieldMapping.Add("UTME", "UTME")
        m_propertyToFieldMapping.Add("Subsection", "Subsection")
        m_propertyToFieldMapping.Add("DataSource", "data_src")
        m_propertyToFieldMapping.Add("Aquifer", "Aquifer")
        m_propertyToFieldMapping.Add("County", "County_c")
        m_propertyToFieldMapping.Add("Township", "Township")
        m_propertyToFieldMapping.Add("Range", "Range")
        m_propertyToFieldMapping.Add("Section", "Section")
        m_propertyToFieldMapping.Add("DepthDrill", "Depth_drll")
        m_propertyToFieldMapping.Add("DepthToBedrock", "depth2bdrk")
        m_propertyToFieldMapping.Add("MgsQuadCode", "mgsquad_c")
        m_propertyToFieldMapping.Add("FirstStrat", "First_Strat")
        m_propertyToFieldMapping.Add("LastStrat", "Last_Strat")
        m_propertyToFieldMapping.Add("FirstBedrock", "First_Bdrk")
        m_propertyToFieldMapping.Add("LastBedrock", "Last_Bdrk")
        m_propertyToFieldMapping.Add("RangeDir", "range_dir")
        m_propertyToFieldMapping.Add("EntryDateInteger", "entry_date")
        m_propertyToFieldMapping.Add("UpdateDateInteger", "updt_date")
        m_propertyToFieldMapping.Add("DepthCompleted", "depth_comp")
        m_propertyToFieldMapping.Add("HoleDateInteger", "date_drll")
        m_propertyToFieldMapping.Add("StratDateInteger", "strat_date")
        m_propertyToFieldMapping.Add("StratUpDateInteger", "strat_upd")
        m_propertyToFieldMapping.Add("LocationMethod", "loc_mc")
        m_propertyToFieldMapping.Add("LocationSource", "loc_src")
        m_propertyToFieldMapping.Add("InputSource", "input_src")
        m_propertyToFieldMapping.Add("StratSource", "strat_src")
        m_propertyToFieldMapping.Add("StratGeologist", "strat_geol")
        m_propertyToFieldMapping.Add("StratMethod", "strat_mc")
        m_propertyToFieldMapping.Add("Type", "type")
        m_propertyToFieldMapping.Add("SoilClass", "soilclass")
        m_propertyToFieldMapping.Add("UniqueNo", "unique_no")
        m_propertyToFieldMapping.Add("HasCuttings", "cuttings")
        m_propertyToFieldMapping.Add("HasCore", "core")
        m_propertyToFieldMapping.Add("HasBHGeophys", "bhgeophys")
        m_propertyToFieldMapping.Add("HasGeochem", "geochem")
        m_propertyToFieldMapping.Add("HasWaterChem", "waterchem")
        m_propertyToFieldMapping.Add("HasAgeDate", "agedate")
        m_propertyToFieldMapping.Add("HasPaleoMag", "paleomag")
        m_propertyToFieldMapping.Add("HasPollen", "pollen")


    End Sub

End Class
