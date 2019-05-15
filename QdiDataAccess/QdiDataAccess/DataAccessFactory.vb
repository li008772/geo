Option Strict On
Option Explicit On


Public Class DataAccessFactory

    ''' <summary>
    ''' An enum for which database type is being used
    ''' </summary>
    Enum DatabaseType
        PostGres
        PersonalGeodataBase
    End Enum
    ''' <summary>
    ''' A function for creating a new personal database access string.
    ''' </summary>
    ''' <param name="pConnectionString">The string to be used in creating the new database access</param>
    ''' <returns></returns>
    Public Function Create(ByVal pConnectionString As String) As IDataAccess '2D0: This is temp
        If (pConnectionString = "") Then
            Return Create(DataAccessFactory.DatabaseType.PostGres)
        End If

        Return New PersonalGDBDataAccess(pConnectionString)
    End Function
    ''' <summary>
    ''' A function for creating an IDataAccess object based off of the database type passed in.
    ''' </summary>
    ''' <param name="databaseType">The database type passed in (has to be one of the specified types in the DatabaseType enum).</param>
    ''' <returns></returns>
    Public Function Create(ByVal databaseType As DatabaseType) As IDataAccess
        Dim pIQdiDataAccess As IDataAccess = Nothing

        Select Case databaseType
            Case DataAccessFactory.DatabaseType.PostGres
                pIQdiDataAccess = New PostGresDataAccess()
            Case DataAccessFactory.DatabaseType.PersonalGeodataBase
                pIQdiDataAccess = New PersonalGDBDataAccess()
            Case Else
                Throw (New CustomException("Invalid Database Specification"))
        End Select

        Return pIQdiDataAccess

    End Function
End Class
