Option Strict On
Option Explicit On

Public Class TexturePropertyToFieldMapping
    Implements Qdi.DataAccess.IPropertyToFieldMappingList

    Dim m_propertyToFieldMapping As Dictionary(Of String, String) = Nothing
    Dim m_fieldToPropertyMapping As Dictionary(Of String, String) = Nothing

    ''' <summary>
    ''' repersents the private variable m_propertyToFieldMapping
    ''' </summary>
    ''' <returns>m_propertyToFieldMapping</returns>
    Public ReadOnly Property PropertyToFieldMapping() As System.Collections.Generic.Dictionary(Of String, String) Implements IPropertyToFieldMappingList.PropertyToFieldMapping
        Get
            If (m_propertyToFieldMapping Is Nothing) Then
                BuildFieldMapping()
            End If
            Return m_propertyToFieldMapping
        End Get
    End Property

    ''' <summary>
    ''' repersents the private variable m_fieldToPropertyMapping
    ''' </summary>
    ''' <returns>m_fieldToPropertyMapping</returns>
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
    ''' returns the value of the input string (key) in PropertyToFieldMapping, returns nothing if there is no such key
    ''' </summary>
    ''' <param name="inProperty">a string repersents the key to search in PropertyToFieldMapping</param>
    ''' <returns>returns the value of the input string (key) in PropertyToFieldMapping</returns>
    Public Function ReturnField(ByVal inProperty As String) As String Implements IPropertyToFieldMappingList.ReturnFieldName
        If (PropertyToFieldMapping.ContainsKey(inProperty)) Then
            Return PropertyToFieldMapping(inProperty)
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    '''  returns the value of the input string (key) in FieldToPropertyMapping, returns nothing if there is no such key
    ''' </summary>
    ''' <param name="inField">a string repersents the key to search in FieldToPropertyMapping</param>
    ''' <returns>returns the value of the input string (key) in FieldToPropertyMapping</returns>
    Public Function ReturnProperty(ByVal inField As String) As String Implements IPropertyToFieldMappingList.ReturnProperty
        If (FieldToPropertyMapping.ContainsKey(inField)) Then
            Return FieldToPropertyMapping(inField)
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' repersents all keys in PropertyToFieldMapping
    ''' </summary>
    ''' <returns>a KeyCollection of all keys in PropertyToFieldMapping</returns>
    Public ReadOnly Property PropertyList() As System.Collections.Generic.Dictionary(Of String, String).KeyCollection Implements IPropertyToFieldMappingList.PropertyList
        Get
            Return PropertyToFieldMapping.Keys
        End Get
    End Property

    ''' <summary>
    ''' repersents all keys in FieldToPropertyMapping
    ''' </summary>
    ''' <returns>a KeyCollection of all keys in FieldToPropertyMapping</returns>
    Public ReadOnly Property FieldList() As System.Collections.Generic.Dictionary(Of String, String).KeyCollection Implements IPropertyToFieldMappingList.FieldList
        Get
            Return Me.FieldToPropertyMapping.Keys
        End Get
    End Property

    ''' <summary>
    ''' return the value of key "date_modify" in PropertyToFieldMapping
    ''' </summary>
    ''' <returns>the value of key "date_modify" in PropertyToFieldMapping</returns>
    Public ReadOnly Property UpdateDateFieldName() As String
        Get
            Return ReturnField("date_modify")
        End Get
    End Property

    ''' <summary>
    ''' populates the mapping for private variable m_propertyToFieldMapping
    ''' </summary>
    Private Sub BuildFieldMapping()
        '*AP2QR* Add Property to Field
        '2DO: XML
        m_propertyToFieldMapping = New Dictionary(Of String, String)

        '*** Property, Field Name
        m_propertyToFieldMapping.Add("RelateId", "RelateId") 'SAME
        m_propertyToFieldMapping.Add("Depth", "sampledepth") ' "depth_of_s")
        m_propertyToFieldMapping.Add("Scientist", "scientist") 'Same
        m_propertyToFieldMapping.Add("CountedBy", "counted_by") 'Same
        m_propertyToFieldMapping.Add("Project", "project") 'Same
        m_propertyToFieldMapping.Add("Depth_Top", "depth_top") ' NEW
        m_propertyToFieldMapping.Add("Depth_Bottom", "depth_bottom") ' NEW
        m_propertyToFieldMapping.Add("Depth_Units", "depth_unit") ' NEW
        m_propertyToFieldMapping.Add("Correlation", "correlation") ' NEW
        m_propertyToFieldMapping.Add("Reference", "reference") ' NEW
        m_propertyToFieldMapping.Add("UpdateDate", "date_modify")
        m_propertyToFieldMapping.Add("isSummaryRecordCode", "incomplete")
        m_propertyToFieldMapping.Add("SourceFile", "sourcefile")
        m_propertyToFieldMapping.Add("ImportErrors", "importerrors")

        '**** Texture
        'm_propertyToFieldMapping.Add("AbnormalUnit", "unit") 'Same
        'm_propertyToFieldMapping.Add("StandardizedUnit", "strat") 'Same
        m_propertyToFieldMapping.Add("PreviousInterptetation", "previnterpt") 'AbnormalUnit", "unit") 'Same
        m_propertyToFieldMapping.Add("Leached", "leached") 'AbnormalUnit", "unit") 'Same
        m_propertyToFieldMapping.Add("WorkingInterpretation", "strat_work") 'AbnormalUnit", "unit") 'Same
        m_propertyToFieldMapping.Add("DepositType", "deposittype") 'Same
        m_propertyToFieldMapping.Add("StratUnit", "stratunit") 'Same

        m_propertyToFieldMapping.Add("DryColor", "dry_color") 'Same
        m_propertyToFieldMapping.Add("WetColor", "wet_color") 'Same
        m_propertyToFieldMapping.Add("SampleName", "sample_num") 'Same
        m_propertyToFieldMapping.Add("Split", "split") 'SAME
        m_propertyToFieldMapping.Add("MainNote", "notes") '"notes_")
        m_propertyToFieldMapping.Add("Notes", "grain_note") '"grain_note") Was [notes_]
        '*** Measurements
        m_propertyToFieldMapping.Add("TotalWeight", "totalweight") '"totalweigh")
        m_propertyToFieldMapping.Add("GravelWeight", "gravel") 'gravelweig")
        m_propertyToFieldMapping.Add("Sand12mmWeight", "sand_12") 'weight_of_")
        m_propertyToFieldMapping.Add("SandCoarsetoMediumWeight", "sand_a") '1/5th to 1 mm
        m_propertyToFieldMapping.Add("SandCoarse_Percent", "coarse_pct") 'less than 1/5 mm
        m_propertyToFieldMapping.Add("SandFineWeight", "sand_b") 'less than 1/5 mm
        m_propertyToFieldMapping.Add("SandAB", "other_split") '"weight_o_1")
        m_propertyToFieldMapping.Add("SandFine_Percent", "fine_pct") 'less than 1/5 mm
        m_propertyToFieldMapping.Add("SandOtherWeight", "other_weight") 'weight_oth")
        m_propertyToFieldMapping.Add("Hydrometer", "hydrometer") 'Same
        m_propertyToFieldMapping.Add("HydrometerControl", "hydrometer_control") 'control_hy")

        'DerivedSubsection
        m_propertyToFieldMapping.Add("AdjustedWeight", "adjustweight") 'adjustedwe")
        m_propertyToFieldMapping.Add("MoistureCorrected", "moist_correct") 'adjustedwe")
        m_propertyToFieldMapping.Add("AdjustedDryWeight", "adjdryweight") 'adjustedwe")
        m_propertyToFieldMapping.Add("GravelPercent", "gravel_pct") '"percent_gr")
        m_propertyToFieldMapping.Add("SandTotalWeight", "sand_total") '"total_weig")
        m_propertyToFieldMapping.Add("SandPercent", "sand_pct") 'percent_sa")
        m_propertyToFieldMapping.Add("ClayWeight", "clay") 'weight_cla")
        m_propertyToFieldMapping.Add("ClayPercent", "clay_pct") 'percent_cl")
        m_propertyToFieldMapping.Add("SandAndClayWeight", "sand_clay") 'weight_san")
        m_propertyToFieldMapping.Add("SiltWeight", "silt") 'weight_sil")
        m_propertyToFieldMapping.Add("SiltPercent", "silt_pct") 'percent_si")
        m_propertyToFieldMapping.Add("Sand12mmPercent", "sand_12_pct") 'percent_1_")
        'DELETED! m_propertyToFieldMapping.Add("Sand12mmPercent_Duplicate", "percent_ve") '*** 2Do: Deletethis, duplicate of Sand12mmPercent
        'DELETED! m_propertyToFieldMapping.Add("SandCoarsetoMediumWeight_Duplicate", "coarse_to_") '*** 2Do: Deletethis, duplicate of SandCoarsetoMediumWeight
        'DELETED! m_propertyToFieldMapping.Add("SandCoarsetoMediumPercent", "percent_co")
        'DELETED! m_propertyToFieldMapping.Add("SandFineWeight_Duplicate", "fine_sand") '*** 2Do: Deletethis, duplicate of SandFineWeight
        'DELETED! m_propertyToFieldMapping.Add("SandFinePercent", "percent_fi")

        'm_propertyToFieldMapping.Add("SandCoarsetoMed", "coarse_to_")
        'm_propertyToFieldMapping.Add("SandFineintoVeryFine", "fine_sand")

        'Light Crystalline Precambrian
        m_propertyToFieldMapping.Add("Felsic", "felsic") ' _") '46
        m_propertyToFieldMapping.Add("Felsic_ClassPercent", "felic_cls") '"granite_cl")
        m_propertyToFieldMapping.Add("Quartzite", "quartzite") 'SAME
        m_propertyToFieldMapping.Add("Quartzite_ClassPercent", "qtz_cls") '"qtz_class")
        m_propertyToFieldMapping.Add("ClearQuartz", "clear_quartz") '")
        m_propertyToFieldMapping.Add("ClearQuartz_ClassPercent", "clear_quartz_cls") '"clear_qu_1")
        m_propertyToFieldMapping.Add("Light_GroupPercent", "light_cls") '"light_clas")

        'Dark Crystalline Precambrian
        m_propertyToFieldMapping.Add("Mafic", "mafic_igneous") 'mafic_igne")
        m_propertyToFieldMapping.Add("Mafic_ClassPercent", "mafic_cls") '"mafic_clas")
        m_propertyToFieldMapping.Add("MetaSedVol", "metasedvol") 'Same
        m_propertyToFieldMapping.Add("MetaSedVol_ClassPercent", "meta_cls") '"meta_class")
        m_propertyToFieldMapping.Add("Dark_GroupPercent", "dark_cls") '"dark_class")

        'Red  Crystalline Precambrian
        m_propertyToFieldMapping.Add("IronFormation", "iron") '_forma")
        m_propertyToFieldMapping.Add("IronFormation_ClassPercent", "iron_cls") '"iron_class")
        m_propertyToFieldMapping.Add("RedVolcanic", "red_volcanic") '"red_volcan") 
        m_propertyToFieldMapping.Add("RedVolcanic_ClassPercent", "rhyolite_cls") '"rhyolite_c")
        m_propertyToFieldMapping.Add("PCarkosic", "pc_arkosic") 'Same
        m_propertyToFieldMapping.Add("PCarkosic_ClassPercent", "pc_arkosic_cls") '"pc_arkos_1")
        m_propertyToFieldMapping.Add("PCQuartzAreniteSST", "pc_qtz_arenite") '"pc_qtz_are")
        m_propertyToFieldMapping.Add("PCQuartzAreniteSST_ClassPercent", "pc_arenite_cls") '"pc_arenite")
        m_propertyToFieldMapping.Add("Red_GroupPercent", "red_cls") '"x_red_clas")

        'Precambrian Totals
        m_propertyToFieldMapping.Add("PrecambrianOther", "pc_other") 'precambria")
        m_propertyToFieldMapping.Add("PrecambrianOther_ClassPercent", "pcother_cls") '"pcother_cl")
        m_propertyToFieldMapping.Add("Precambrian_Total", "precambrian") 'precambr_1")
        m_propertyToFieldMapping.Add("Precambrian_BulkPercent", "precambrian_blk") '"precambr_2")
        m_propertyToFieldMapping.Add("Crystalline_BulkPercent", "crystalline_blk") '"crystall_1")

        'Carbonate Fields
        m_propertyToFieldMapping.Add("Carbonate", "carbonate") ' "carbonate1")
        m_propertyToFieldMapping.Add("CarbonateTotal", "carbonate_total") '"carbonate_")
        m_propertyToFieldMapping.Add("Carbonate_ClassPercent", "carbchert_cls") '"carbchert_")
        m_propertyToFieldMapping.Add("Carbonate_GroupPercent", "chert_grp") 'chert_grou")
        m_propertyToFieldMapping.Add("Carbonate_BulkPercent", "carbonate_blk") '"carbonate")
        m_propertyToFieldMapping.Add("PaleoSandStone", "paleo_ss") 'SAME
        m_propertyToFieldMapping.Add("PaleoSandStone_ClassPercent", "ss_cls") '"ss_class") 
        m_propertyToFieldMapping.Add("PaleoSandStone_GroupPercent", "sst_grp") '"sst_group")
        m_propertyToFieldMapping.Add("PaleoShale", "paleo_shale") '"paleo_shal")
        m_propertyToFieldMapping.Add("PaleoShale_ClassPercent", "shale_cls") '"shale_clas")
        m_propertyToFieldMapping.Add("PaleoShale_GroupPercent", "psh_grp") '"psh_group")
        m_propertyToFieldMapping.Add("PaleoOther", "paleo_other") '"paleo_othe")
        m_propertyToFieldMapping.Add("PaleoOther_ClassPercent", "pother_cls") '"pother_cla")
        m_propertyToFieldMapping.Add("PaleozoicTotal", "paleozoic") ', "paleozoic_")
        m_propertyToFieldMapping.Add("Paleozoic_BulkPercent", "paleozoic_blk") '"paleozoic")

        'Cretceous-Other
        m_propertyToFieldMapping.Add("GrayShale", "gray_shale") 'SAME
        m_propertyToFieldMapping.Add("GrayShale_ClassPercent", "gray_cls") '"gray_class")
        m_propertyToFieldMapping.Add("GSHGroup", "gsh_grp") '"gsh_group")
        m_propertyToFieldMapping.Add("SpeckledShale", "speckled_shale") '"speckled_s")
        m_propertyToFieldMapping.Add("SpeckledShale_ClassPercent", "speck_cls") '"speck_clas")
        m_propertyToFieldMapping.Add("ShaleTotal", "shale")   ' "shale_tota")
        m_propertyToFieldMapping.Add("Shale_BulkPercent", "shale_blk") '"shale")

        m_propertyToFieldMapping.Add("Limestone", "limestone") 'SAME
        m_propertyToFieldMapping.Add("Limestone_ClassPercent", "lms_cls") '"lms_class")
        m_propertyToFieldMapping.Add("InoceramusShells", "inoceramus") 'SAME
        m_propertyToFieldMapping.Add("InoceramusShells_ClassPercent", "shell_cls") '"shell_clas")
        m_propertyToFieldMapping.Add("Pyrite", "pyrite") 'SAME
        m_propertyToFieldMapping.Add("Pyrite_ClassPercent", "pyrite_cls") '"pyrite_cla")
        m_propertyToFieldMapping.Add("Lignite", "lignite") 'SAME
        m_propertyToFieldMapping.Add("Lignite_ClassPercent", "lignite_cls") '"lignite_cl")
        m_propertyToFieldMapping.Add("OstranderSand", "ostrandersand") '"ostrander_")
        m_propertyToFieldMapping.Add("OstranderSand_ClassPercent", "ostrander_cls") '"ostrander1")
        m_propertyToFieldMapping.Add("CretaceousOther", "cret_other") 'SAME
        m_propertyToFieldMapping.Add("CretaceousOther_ClassPercent", "cothers_cls") '"cothers_cl")
        m_propertyToFieldMapping.Add("CretaceousTotal", "cret_total") 'SAME
        m_propertyToFieldMapping.Add("Cretaceous_BulkPercent", "cretaceous_blk") '"cretaceous")
        m_propertyToFieldMapping.Add("Chert", "chert") 'SAME
        m_propertyToFieldMapping.Add("Unknown", "unknown") '"unknown_")
        m_propertyToFieldMapping.Add("Secondary", "secondary") 'SAME
        m_propertyToFieldMapping.Add("OtherOther", "other_other") '"other_othe")
        m_propertyToFieldMapping.Add("OtherTotal", "other_total") '"other_tota")
        m_propertyToFieldMapping.Add("Other_GroupPercent", "others_pct") '"others")
        m_propertyToFieldMapping.Add("CCSOther", "ccs_other") '"ccs_other_")
        m_propertyToFieldMapping.Add("CCSOthers_GroupPercent", "ccs_others_pct") '"ccs_others")
        m_propertyToFieldMapping.Add("GrandTotal", "grand_total") '"grand_tota")
        m_propertyToFieldMapping.Add("BulkTotal", "bulk_total") 'SAME
        m_propertyToFieldMapping.Add("MiscK_GroupPercent", "misc_k_grp") '"misc_k_gro")
        m_propertyToFieldMapping.Add("CCSBulk", "ccs_blk") 'SAME
        m_propertyToFieldMapping.Add("NormalSilt", "silt_norm") '"normal_sil")
        m_propertyToFieldMapping.Add("NormalXlline", "xlline_norm") '"normal_xll")
        'DELETED!  m_propertyToFieldMapping.Add("Crystalline_Total_Duplicate", "crystalline") '"crystallin")


        '*m_propertyToFieldMapping.Add("NegativeNine", "paleozoic_")
        'm_propertyToFieldMapping.Add("NegativeNine", "paleozoic")
        'm_propertyToFieldMapping.Add("NegativeNine", "others")
        'm_propertyToFieldMapping.Add("NegativeNine", "carbonate_")

        'm_propertyToFieldMapping.Add("NegativeNine", "prefix_cop")
        'm_propertyToFieldMapping.Add("NegativeNine", "depth_top")
        'm_propertyToFieldMapping.Add("NegativeNine", "depth_bot")
        'm_propertyToFieldMapping.Add("NegativeNine", "depth_unit")
        'm_propertyToFieldMapping.Add("NegativeNine", "correlatio")

        Dim ptemplist As List(Of String) = New List(Of String)
        For Each ivalue In m_propertyToFieldMapping.Values
            If (ptemplist.Contains(ivalue)) Then
                MsgBox(ivalue.ToString + " ALREADY EXISTS")
            End If
            ptemplist.Add(ivalue)
        Next

    End Sub

End Class
