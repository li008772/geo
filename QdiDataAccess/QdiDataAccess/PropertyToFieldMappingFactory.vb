Option Strict On
Option Explicit On

''' <summary>
''' Mapping to a property field in the database dictionary
''' </summary>
Public Class PropertyToFieldMappingFactory

    Enum DatabaseType
        PostGres
    End Enum

    Public Function Create(ByVal databaseType As DatabaseType) As IPropertyToFieldMappingList
        Dim pPropertyToFieldMapping As IPropertyToFieldMappingList

        If (databaseType = PropertyToFieldMappingFactory.DatabaseType.PostGres) Then
            pPropertyToFieldMapping = New PropertyToFieldMappingPostGres()
        Else
            Throw (New CustomException("Invalid Database Specification"))
        End If
        Return pPropertyToFieldMapping

    End Function

    'Dim pDataAccessFactory As New Qdi.DataAccess.PropertyToFieldMappingFactory
    '            m_FieldMapping = pDataAccessFactory.Create(Qdi.DataAccess.PropertyToFieldMappingFactory.DatabaseType.PostGres)

End Class
