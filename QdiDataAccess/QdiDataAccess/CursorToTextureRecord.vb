Option Explicit On
Option Strict On

Imports System.Reflection

Public Class CursorToTextureRecord

    Private m_RecordList As List(Of Qdi.BusinessLogic.ITextureRecord)
    Private m_IQdiCursor As ESRI.ArcGIS.Geodatabase.ICursor
    Private m_FieldMapping As IPropertyToFieldMappingList = New Qdi.DataAccess.TexturePropertyToFieldMapping
    Private m_FieldDictionary As Dictionary(Of String, Integer) = Nothing
    Private m_PropertyToFieldMappingList As Qdi.DataAccess.IPropertyToFieldMappingList = New Qdi.DataAccess.TexturePropertyToFieldMapping
    ''' <summary>
    ''' This subroutine appears to do nothing
    ''' </summary>
    Private Sub New()
    End Sub
    ''' <summary>
    ''' A subroutine for creating a new IQdiCursor.
    ''' </summary>
    ''' <param name="pCursorDef"></param>
    Public Sub New(ByVal pCursorDef As ESRI.ArcGIS.Geodatabase.ICursor)
        m_IQdiCursor = pCursorDef
    End Sub

    ''' <summary>
    '''The RecordList property. If there is no list, then the list must be built by calling the BuildRecordList() function. Whether this is necessary or not, a RecordList property will be returned when this property is called.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property RecordList() As System.Collections.Generic.List(Of BusinessLogic.ITextureRecord)
        Get
            If (m_RecordList Is Nothing) Then
                BuildRecordList()
            End If
            Return m_RecordList
        End Get
    End Property
    ''' <summary>
    ''' The PropertyToFieldMapping property, which is returned when this is called.
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property PropertyToFieldMapping() As IPropertyToFieldMappingList
        Get
            Return m_FieldMapping
        End Get
    End Property
    ''' <summary>
    ''' A function that converts features into the texture record. As long as no exceptions are caught, a texture record is returned.
    ''' </summary>
    ''' <param name="pRow">The current row in the DB</param>
    ''' <returns></returns>
    Public Function ConvertFeatureIntoTextureRecord(ByVal pRow As ESRI.ArcGIS.Geodatabase.IRow) As Qdi.BusinessLogic.ITextureRecord

        Dim pNewTextureRecord As Qdi.BusinessLogic.ITextureRecord = Nothing

        Try
            Dim ptemp As Dictionary(Of String, Integer) = FieldIndexDictionary(pRow)
            If (ptemp.ContainsKey("RelateId")) Then

                Dim pIndexRelatetId As Integer = ptemp.Item("RelateId")
                Dim pIndexObjectId As Integer = ptemp.Item("objectid")

                pNewTextureRecord = New Qdi.BusinessLogic.TextureRecordList.TextureRecord(pRow.Value(pIndexRelatetId).ToString, CType(pRow.Value(pIndexObjectId), Integer))

                copyFeatureValuesIntoProperties(pRow, pNewTextureRecord)
            End If
        Catch ex As Exception
            Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            pNewTextureRecord = Nothing
        End Try



        Return pNewTextureRecord

    End Function
    ''' <summary>
    ''' A subroutine that copies the feature values into the properties. This subroutine continues as long as there are values to copy over.
    ''' </summary>
    ''' <param name="pFeature">The feature to be copied</param>
    ''' <param name="pTextureRecord">The current texture record</param>
    Private Sub copyFeatureValuesIntoProperties(ByVal pFeature As ESRI.ArcGIS.Geodatabase.IRow, ByRef pTextureRecord As Qdi.BusinessLogic.ITextureRecord)

        Dim pFieldDictionary As Dictionary(Of String, Integer) = FieldIndexDictionary(pFeature)
        Dim IQdiRecordPropertyList As PropertyInfo() = GetType(Qdi.BusinessLogic.ITextureRecord).GetProperties()
        Dim pPropertyValue As Object
        Dim pPropertyInfo As System.Reflection.PropertyInfo
        Dim pFieldName As String

        For iPropertyIndex As Integer = 0 To IQdiRecordPropertyList.Length - 1
            pPropertyInfo = IQdiRecordPropertyList(iPropertyIndex)
            If pPropertyInfo.GetIndexParameters().Length = 0 Then

                If (pPropertyInfo.CanWrite) Then
                    pFieldName = PropertyToFieldMapping.ReturnFieldName(pPropertyInfo.Name)

                    If Not (pFieldName Is Nothing) Then
                        If Not (pFieldDictionary.ContainsKey(pFieldName)) Then
                            Continue For
                        End If
                        If (pFieldDictionary(pFieldName) > -1) Then
                            pPropertyValue = ReadFeatureValue(pFeature, pFieldDictionary(pFieldName))

                            Try
                                'MsgBox(pPropertyInfo.Name.ToString + "   =   " + pPropertyValue.ToString) 
                                If Not (pTextureRecord Is Nothing) Then
                                    If Not (pPropertyValue Is Nothing) Then
                                        SetObjectProperty(pTextureRecord, pPropertyInfo, pPropertyValue)
                                    End If
                                End If

                            Catch ex As Exception
                                Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                            End Try
                        Else
                            Windows.Forms.MessageBox.Show("Did not find:" + pFieldName + vbNewLine + pFieldDictionary(pFieldName).ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                        End If
                    End If

                End If
            End If

        Next iPropertyIndex
        pTextureRecord.Preserve()

    End Sub
    ''' <summary>
    ''' A function that reads the row's value and returns it.
    ''' </summary>
    ''' <param name="pRow">The row to read from</param>
    ''' <param name="pFieldIndex">The field index to read from</param>
    ''' <returns></returns>
    Private Function ReadRowValue(ByVal pRow As ESRI.ArcGIS.Geodatabase.IRow, ByVal pFieldIndex As Integer) As Object
        Dim pPropertyValue As Object

        If Not ((mgs.Domain.DomainHandler.ReturnCVDNameAsStringFromFieldIndex(pRow, pFieldIndex)) Is Nothing) Then
            pPropertyValue = mgs.Domain.DomainHandler.ReturnCVDNameAsStringFromFieldIndex(pRow, pFieldIndex)
        Else
            pPropertyValue = pRow.Value(pFieldIndex)
            If (pPropertyValue.GetType.ToString = "System.DBNull") Then
                pPropertyValue = CType(Nothing, Object)
            End If
        End If

        Return pPropertyValue
    End Function
    ''' <summary>
    ''' A subroutine that sets an object's property.
    ''' </summary>
    ''' <param name="pQdiRecord">The qdiRecord to set the value in</param>
    ''' <param name="pPropertyInfo">The property's info parameter</param>
    ''' <param name="pPropertyValue">The value to be set to the object's property value</param>
    Private Sub SetObjectProperty(ByVal pQdiRecord As Qdi.BusinessLogic.ITextureRecord, ByVal pPropertyInfo As System.Reflection.PropertyInfo, ByVal pPropertyValue As Object)
        Try

            pPropertyInfo.SetValue(pQdiRecord, pPropertyValue, Nothing)

        Catch ex As Exception
            MsgBox(pPropertyInfo.Name + vbNewLine + pPropertyValue.ToString)

            Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' The FieldIndexDictionary property. This simply returns the property when called.
    ''' </summary>
    ''' <param name="pRow">The current row</param>
    ''' <returns></returns>
    Private ReadOnly Property FieldIndexDictionary(ByVal pRow As ESRI.ArcGIS.Geodatabase.IRow) As Dictionary(Of String, Integer)
        Get
            If (m_FieldDictionary Is Nothing) Then

                m_FieldDictionary = New Dictionary(Of String, Integer)
                Dim pFieldList As System.Collections.Generic.Dictionary(Of String, String).KeyCollection = PropertyToFieldMapping.FieldList
                Dim pFieldIndex As Integer
                pFieldIndex = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, "ObjectID")
                m_FieldDictionary.Add("objectid", pFieldIndex)

                For Each pField As String In pFieldList
                    pFieldIndex = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pRow, pField)
                    If (pFieldIndex > -1) Then
                        m_FieldDictionary.Add(pField, pFieldIndex)
                    Else
                        If (Debugger.IsAttached) Then
                            MsgBox("Unable to find this field, look in CursorToTextureRecords: " + pField.ToLower())
                        End If

                    End If
                Next
            End If

            Return m_FieldDictionary
        End Get
    End Property
    ''' <summary>
    ''' A subroutine that builds a record list. This function is called when there isn't already a record list.
    ''' </summary>
    Private Sub BuildRecordList()
        Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor = m_IQdiCursor

        Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
        pRow = pCursor.NextRow

        If pRow Is Nothing Then
            Windows.Forms.MessageBox.Show("No Records Selected", "No Records Selected", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim pResultList As List(Of Qdi.BusinessLogic.ITextureRecord) = New List(Of Qdi.BusinessLogic.ITextureRecord)
        Dim pTempQdiRecord As Qdi.BusinessLogic.ITextureRecord

        Dim iCounter As Integer = 0

        While Not (pRow Is Nothing)
            iCounter += 1
            Try
                pTempQdiRecord = ConvertFeatureIntoTextureRecord(pRow)
                pResultList.Add(pTempQdiRecord)
            Catch ex As Exception
                Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Finally
                pRow = pCursor.NextRow
            End Try
        End While

        m_IQdiCursor = Nothing
        m_RecordList = pResultList

    End Sub
    ''' <summary>
    ''' Converts an object to an integer and returns an integer. This may not be a necessary function, since there is likely a built in function that performs the same task.
    ''' </summary>
    ''' <param name="pObject">The object to be converted to an integer</param>
    ''' <returns></returns>
    Private Function ObjToInteger(ByVal pObject As Object) As Integer
        Dim pInteger As Integer
        pInteger = Nothing

        If (Integer.TryParse(pObject.ToString, pInteger)) Then
            Return pInteger
        End If

        Return pInteger
    End Function
    ''' <summary>
    ''' Converts an object to an double and returns an double. This may not be a necessary function, since there is likely a built in function that performs the same task.
    ''' </summary>
    ''' <param name="pObject">The object to be converted to a double</param>
    ''' <returns></returns>
    Private Function ObjToDouble(ByVal pObject As Object) As Double
        Dim pDouble As Double
        pDouble = Nothing

        If (Double.TryParse(pObject.ToString, pDouble)) Then
            Return pDouble
        End If

        Return pDouble
    End Function
    ''' <summary>
    ''' A function that updates a texture record. Not only does it update the record, but records the date when each item was updated.
    ''' </summary>
    ''' <param name="pRow">The current row</param>
    ''' <param name="pTextureRecord">The texture record to be updated</param>
    ''' <returns></returns>
    Public Function UpdateRecord(ByRef pRow As ESRI.ArcGIS.Geodatabase.IRow, ByVal pTextureRecord As Qdi.BusinessLogic.ITextureRecord) As Boolean
        Dim pMadeChange As Boolean = False

        Dim pFieldDictionary As Dictionary(Of String, Integer) = FieldIndexDictionary(pRow)
        Dim IQdiRecordPropertyList As PropertyInfo() = GetType(Qdi.BusinessLogic.ITextureRecord).GetProperties()
        Dim pOldPropertyValue As Object
        Dim pNewPropertyValue As Object
        Dim pPropertyInfo As System.Reflection.PropertyInfo
        Dim pFieldName As String
        Dim pFieldDictionaryIndex As Integer

        Dim DebugOn As Boolean = True

        For iPropertyIndex As Integer = 0 To IQdiRecordPropertyList.Length - 1
            pPropertyInfo = IQdiRecordPropertyList(iPropertyIndex)
            If pPropertyInfo.GetIndexParameters().Length = 0 Then

                pFieldName = PropertyToFieldMapping.ReturnFieldName(pPropertyInfo.Name)

                If Not (pFieldName Is Nothing) Then
                    If Not (pFieldDictionary.ContainsKey(pFieldName)) Then
                        Continue For
                    End If
                    If (pFieldDictionary(pFieldName) > -1) Then
                        ' End If
                        pOldPropertyValue = ReadFeatureValue(pRow, pFieldDictionary(pFieldName))
                        pNewPropertyValue = pPropertyInfo.GetValue(pTextureRecord, Nothing)
                        Try
                            Dim pFieldType As ESRI.ArcGIS.Geodatabase.IField = pRow.Fields.Field(pFieldDictionaryIndex)
                            Dim pMakeAChange As Boolean = False
                            If ((pOldPropertyValue Is Nothing) Or (pNewPropertyValue Is Nothing)) Then
                                If (pOldPropertyValue Is Nothing) And (pNewPropertyValue Is Nothing) Then
                                    Continue For
                                End If
                                pMakeAChange = True
                            ElseIf (pOldPropertyValue.ToString <> pNewPropertyValue.ToString) Then
                                pMakeAChange = True
                            End If


                            If (pMakeAChange) Then
                                pMadeChange = True
                                pFieldDictionaryIndex = pFieldDictionary(pFieldName)
                                If (pFieldDictionaryIndex > -1) Then

                                    Select Case pRow.Fields.Field(pFieldDictionaryIndex).Type
                                        Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeDouble
                                            Dim pNewDouble As Double
                                            If (Double.TryParse(pNewPropertyValue.ToString, pNewDouble) = True) Then
                                                pRow.Value(pFieldDictionaryIndex) = pNewDouble
                                            End If
                                        Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeString
                                            Dim pString As String

                                            If (pNewPropertyValue Is Nothing) Then
                                                pString = ""
                                            Else
                                                pString = pNewPropertyValue.ToString
                                            End If

                                            pRow.Value(pFieldDictionaryIndex) = pString.ToString
                                        Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeSmallInteger
                                            Dim pInteger As Integer
                                            If (Integer.TryParse(pNewPropertyValue.ToString, pInteger)) Then
                                                pRow.Value(pFieldDictionaryIndex) = pInteger
                                            End If
                                        Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeOID
                                            Dim pLongInt As Long
                                            If (Long.TryParse(pNewPropertyValue.ToString, pLongInt)) Then
                                                pRow.Value(pFieldDictionaryIndex) = pLongInt
                                            End If

                                        Case Else
                                            'MsgBox("Unable to write value to Field, " + pFieldName + ", type: " + pRow.Fields.Field(pFieldDictionaryIndex).Type.ToString + vbNewLine + "Old Value: " + pOldPropertyValue.ToString + vbNewLine + "New Value: " + pNewPropertyValue.ToString)

                                    End Select
                                End If

                            End If
                        Catch ex As Exception
                            Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                        End Try
                    Else
                        Windows.Forms.MessageBox.Show("Did not find:" + pFieldName + vbNewLine + pFieldDictionary(pFieldName).ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                    End If
                End If

                'End If
            End If


        Next iPropertyIndex

        If (pMadeChange = True) Then
            Dim a As String = m_PropertyToFieldMappingList.ReturnFieldName("UpdateDate")
            If Not (a Is Nothing) Then
                Dim pIndex As Integer = pFieldDictionary(a)
                If (pIndex > -1) Then
                    pRow.Value(pIndex) = Date.Today
                End If
            End If

        End If

        Return pMadeChange
    End Function

    ''' <summary>
    ''' A function for reading feature values. It returns a property value.
    ''' </summary>
    ''' <param name="pRow">The current row</param>
    ''' <param name="pFieldIndex">The current field index.</param>
    ''' <returns></returns>
    Private Function ReadFeatureValue(ByVal pRow As ESRI.ArcGIS.Geodatabase.IRow, ByVal pFieldIndex As Integer) As Object
        Dim pPropertyValue As Object

        If Not ((mgs.Domain.DomainHandler.ReturnCVDNameAsStringFromFieldIndex(pRow, pFieldIndex)) Is Nothing) Then
            pPropertyValue = mgs.Domain.DomainHandler.ReturnCVDNameAsStringFromFieldIndex(pRow, pFieldIndex)
        Else
            pPropertyValue = pRow.Value(pFieldIndex)
            If (pPropertyValue.GetType.ToString = "System.DBNull") Then
                pPropertyValue = CType(Nothing, Object)
            End If
        End If

        Return pPropertyValue
    End Function
End Class


