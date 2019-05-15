Option Explicit On
Option Strict On

Imports System.Reflection

Public Class CursorToQdiRecord
    Implements Qdi.DataAccess.ICursorToQdiRecord

    Private m_RecordList As List(Of Qdi.BusinessLogic.IQdiRecord)
    Private m_IQdiCursor As Qdi.DataAccess.IQdiCursorDefinition
    Private m_FieldMapping As IPropertyToFieldMappingList
    Private m_FieldDictionary As Dictionary(Of String, Integer) = Nothing
    ''' <summary>
    ''' A sub that appears to do nothing
    ''' </summary>
    Private Sub New()
    End Sub
    ''' <summary>
    ''' A subroutine for creating a new IQdiCursor.
    ''' </summary>
    ''' <param name="pCursorDef"></param>
    Public Sub New(ByVal pCursorDef As Qdi.DataAccess.IQdiCursorDefinition, ByVal pFieldMapping As IPropertyToFieldMappingList)
        m_IQdiCursor = pCursorDef
        m_FieldMapping = pFieldMapping
    End Sub


    ''' <summary>
    '''The RecordList property. If there is no list, then the list must be built by calling the BuildRecordList() function. Whether this is necessary or not, a RecordList property will be returned when this property is called.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property RecordList() As System.Collections.Generic.List(Of BusinessLogic.IQdiRecord) Implements ICursorToQdiRecord.RecordList
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
    ''' A function that converts features into the qdi record. As long as no exceptions are caught, a qdi record is returned.
    ''' </summary>
    ''' <param name="pFeature">The current feature</param>
    ''' <returns></returns>
    Public Function ConvertFeatureIntoQdiRecord(ByVal pFeature As ESRI.ArcGIS.Geodatabase.IFeature) As Qdi.BusinessLogic.IQdiRecord Implements ICursorToQdiRecord.ConvertFeatureIntoQdiRecord
        Dim pPoint As ESRI.ArcGIS.Geometry.IPoint

        pPoint = CType(pFeature.ShapeCopy, ESRI.ArcGIS.Geometry.IPoint)

        Dim pObjectId As Double = ObjToDouble(pFeature.Value(FieldIndexDictionary(pFeature).Item("objectid")))

        Dim pNewQdiRecord As Qdi.BusinessLogic.IQdiRecord
        pNewQdiRecord = New Qdi.BusinessLogic.QdiRecord(pPoint, pObjectId)

        Try
            pNewQdiRecord = copyFeatureValuesIntoProperties(pFeature, pNewQdiRecord)

        Catch ex As Exception
            Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            pNewQdiRecord = Nothing
        End Try


        Return pNewQdiRecord

    End Function
    ''' <summary>
    ''' A subroutine that copies the feature values into the properties. This subroutine continues as long as there are values to copy over.
    ''' </summary>
    ''' <param name="pFeature">The feature to be copied</param>
    ''' <param name="pQdiRecord">The current qdi record</param>
    Private Function copyFeatureValuesIntoProperties(ByVal pFeature As ESRI.ArcGIS.Geodatabase.IFeature, ByVal pQdiRecord As Qdi.BusinessLogic.IQdiRecord) As Qdi.BusinessLogic.IQdiRecord

        Dim pFieldDictionary As Dictionary(Of String, Integer) = FieldIndexDictionary(pFeature)
        Dim IQdiRecordPropertyList As PropertyInfo() = GetType(Qdi.BusinessLogic.IQdiRecord).GetProperties()
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
                                SetObjectProperty(pQdiRecord, pPropertyInfo, pPropertyValue)
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

        Dim pKeyCollectionList As Dictionary(Of String, String).KeyCollection = pQdiRecord.ValueDictionary.Keys
        Dim iTemp As List(Of String) = New List(Of String)
        For Each iFieldName In pKeyCollectionList
            iTemp.Add(iFieldName)
        Next

        Dim pFieldIndex As Integer

        For Each iFieldName In iTemp
            pFieldIndex = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pFeature, iFieldName)
            If (pFieldIndex > -1) Then
                pPropertyValue = ReadFeatureValue(pFeature, pFieldIndex)
                If Not (pPropertyValue Is Nothing) Then
                    pQdiRecord.ValueDictionarySetValue(iFieldName, pPropertyValue.ToString)
                End If
            End If
        Next
        Return pQdiRecord
    End Function
    ''' <summary>
    ''' A function that reads feature values
    ''' </summary>
    ''' <param name="pFeature">the current feature</param>
    ''' <param name="pFieldIndex">the field index</param>
    ''' <returns>the property value</returns>
    Private Function ReadFeatureValue(ByVal pFeature As ESRI.ArcGIS.Geodatabase.IFeature, ByVal pFieldIndex As Integer) As Object
        Dim pPropertyValue As Object

        If Not ((mgs.Domain.DomainHandler.ReturnCVDNameAsStringFromFieldIndex(pFeature, pFieldIndex)) Is Nothing) Then
            pPropertyValue = mgs.Domain.DomainHandler.ReturnCVDNameAsStringFromFieldIndex(pFeature, pFieldIndex)
        Else
            pPropertyValue = pFeature.Value(pFieldIndex)
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
    Private Sub SetObjectProperty(ByVal pQdiRecord As Qdi.BusinessLogic.IQdiRecord, ByVal pPropertyInfo As System.Reflection.PropertyInfo, ByVal pPropertyValue As Object)
        Try
            pPropertyInfo.SetValue(pQdiRecord, pPropertyValue, Nothing)
        Catch ex As Exception
            Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub
    ''' <summary>
    ''' The FieldIndexDictionary property. This simply returns the property when called.
    ''' </summary>
    ''' <param name="pFeature>The current feature</param>
    ''' <returns></returns>
    Private ReadOnly Property FieldIndexDictionary(ByVal pFeature As ESRI.ArcGIS.Geodatabase.IFeature) As Dictionary(Of String, Integer)
        Get
            If (m_FieldDictionary Is Nothing) Then

                m_FieldDictionary = New Dictionary(Of String, Integer)
                Dim pFieldList As System.Collections.Generic.Dictionary(Of String, String).KeyCollection = PropertyToFieldMapping.FieldList
                Dim pFieldIndex As Integer
                pFieldIndex = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pFeature, "ObjectID")
                m_FieldDictionary.Add("objectid", pFieldIndex)

                For Each pField As String In pFieldList
                    pFieldIndex = mgs.Domain.FindFieldByNameOrAlias.FindFieldIndexByNameOrAlias(pFeature, pField)
                    If (pFieldIndex > -1) Then
                        m_FieldDictionary.Add(pField, pFieldIndex)
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
        Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor = m_IQdiCursor.QdixICursor
        Dim pFeatureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
        pFeatureCursor = CType(pCursor, ESRI.ArcGIS.Geodatabase.IFeatureCursor)

        Dim pFeature As ESRI.ArcGIS.Geodatabase.IFeature
        pFeature = pFeatureCursor.NextFeature

        If pFeature Is Nothing Then
            Windows.Forms.MessageBox.Show("No Records Selected", "No Records Selected", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim pResultList As List(Of Qdi.BusinessLogic.IQdiRecord) = New List(Of Qdi.BusinessLogic.IQdiRecord)
        Dim pTempQdiRecord As Qdi.BusinessLogic.IQdiRecord

        Dim iCounter As Integer = 0

        While Not (pFeature Is Nothing)
            iCounter += 1
            'Debug.WriteLine(iCounter.ToString, "Count ")

            Try
                pTempQdiRecord = ConvertFeatureIntoQdiRecord(pFeature)
                pResultList.Add(pTempQdiRecord)

            Catch ex As Exception
                Windows.Forms.MessageBox.Show(ex.ToString + vbNewLine + vbNewLine + ex.Message.ToString, "Cancelling...", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Finally
                pFeature = pFeatureCursor.NextFeature
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
End Class
