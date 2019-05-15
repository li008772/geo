Option Explicit On
Option Strict On
Imports System.Reflection

Public Class QdiRecordStandardizer
    Implements Qdi.DataAccess.IQdiRecordStandardizer

    ''' <summary>
    ''' default constructor for class QdiRecordStandardizer
    ''' </summary>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' format and standardize the input QDI record
    ''' </summary>
    ''' <param name="pQdiRecord">an instance of BusinessLogic.IQdiRecord</param>
    ''' <param name="pDataAccess">an instance of IDataAccess</param>
    ''' <returns>a new QDI record</returns>
    Public Function Standardize(ByVal pQdiRecord As BusinessLogic.IQdiRecord, ByVal pDataAccess As IDataAccess) As BusinessLogic.IQdiRecord Implements IQdiRecordStandardizer.Standardize
        Dim newQdiRecord As Qdi.BusinessLogic.IQdiRecord = pQdiRecord.ReturnClone

        Dim IQdiRecordPropertyList As PropertyInfo() = GetType(Qdi.BusinessLogic.IQdiRecord).GetProperties()
        Dim pPropertyValue As Object
        Dim pPropertyValue2 As Object
        Dim pPropertyInfo As System.Reflection.PropertyInfo
        Dim pFieldName As String

        Dim pPropertyToFieldMapping As IPropertyToFieldMappingList = pDataAccess.PropertyToFieldMapping

        For iPropertyIndex As Integer = 0 To IQdiRecordPropertyList.Length - 1
            pPropertyInfo = IQdiRecordPropertyList(iPropertyIndex)
            If pPropertyInfo.GetIndexParameters().Length = 0 Then

                If ((pPropertyInfo.CanWrite) And (pPropertyInfo.CanRead)) Then

                    pFieldName = pPropertyToFieldMapping.ReturnFieldName(pPropertyInfo.Name)

                    If (Not pFieldName Is Nothing) Then
                        Dim pTable As ESRI.ArcGIS.Geodatabase.ITable = pDataAccess.GetFeatureTable(BusinessLogic.NamedFeatureClass.qdix)
                        Dim pCVDDoman As ESRI.ArcGIS.Geodatabase.ICodedValueDomain = mgs.Domain.DomainHandler.GetCodedValueDomainFromField(pTable, pFieldName)
                        If (Not pCVDDoman Is Nothing) Then
                            pPropertyValue = pPropertyInfo.GetValue(pQdiRecord, Nothing)
                            pPropertyValue2 = pPropertyInfo.GetValue(newQdiRecord, Nothing)

                            If (pPropertyValue2 Is Nothing) Then
                                If Not (pPropertyValue Is Nothing) Then
                                    pPropertyInfo.SetValue(newQdiRecord, pPropertyValue, Nothing)
                                End If
                            Else
                                Dim pNewValue As Object = mgs.Domain.DomainHandler.ReturnCVDValue(pCVDDoman, pPropertyValue.ToString)

                                If (Not (pNewValue Is Nothing)) Then
                                    If (pNewValue.ToString <> pPropertyValue.ToString) Then
                                        pPropertyInfo.SetValue(newQdiRecord, pNewValue, Nothing)
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If
            End If

        Next iPropertyIndex

        Return newQdiRecord
    End Function

End Class
