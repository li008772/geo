Option Strict On
Option Explicit On

Imports ESRI.ArcGIS.Geodatabase
Imports mgs.CursorBuilder

''' <summary>
''' Cursor for the postgres database
''' </summary>
Public Class PostGresQdiCursorDefinition
    Inherits QdiCursorBaseClass

    Friend Sub New(ByVal pPostgresDataAcess As PostGresDataAccess)
        Me.DataAccess = pPostgresDataAcess
    End Sub

    ''' <summary>
    ''' Builds the sqlstring for the database
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property SqlString() As String
        Get
            Dim returnString As String = ""

            addClause(returnString, clauseRelateId)
            addClause(returnString, StandardWhereClause("township", MyBase.Township))
            addClause(returnString, StandardWhereClause("section", MyBase.Section))
            addClause(returnString, StandardWhereClause("data_src", MyBase.DataSource))
            addClause(returnString, StandardWhereClause("range", MyBase.Range))
            addClause(returnString, StandardWhereClause("date_drll", MyBase.FromDate, pOperator:=">="))
            addClause(returnString, StandardWhereClause("date_drll", MyBase.ToDate, pOperator:="<="))
            addClause(returnString, StandardWhereClause("depth2bdrk", MyBase.FromDepth, pOperator:=">="))
            addClause(returnString, StandardWhereClause("depth2bdrk", MyBase.ToDepth, pOperator:="<="))
            addClause(returnString, StandardWhereClause("county_c", MyBase.Counties, pOperator:="in"))
            addClause(returnString, StandardWhereClause("mgsquad_c", MyBase.Quadrangles, pOperator:="in"))
            addClause(returnString, StandardWhereClause("first_strat", MyBase.FirstStratUnits, pOperator:="in"))
            addClause(returnString, StandardWhereClause("last_strat", MyBase.LastStratUnits, pOperator:="in"))
            addClause(returnString, StandardWhereClause("first_bdrk", MyBase.FirstBedrockUnits, pOperator:="in"))
            addClause(returnString, StandardWhereClause("last_bdrk", MyBase.LastBedrockUnits, pOperator:="in"))
            addClause(returnString, StandardWhereClause("aquifer", MyBase.Aquifers, pOperator:="in"))
            addClause(returnString, StandardWhereClause("relateid", MyBase.RelateIDs, pOperator:="in"))

            If (returnString.Trim = "") Then
                returnString = "Not (Relateid IS NULL)" '#, StandardWhereClause("relateid", "'NULL'", pOperator:="<>"))
            End If

            Debug.WriteLine(returnString)
            Return returnString
        End Get
    End Property

    ''' <summary>
    ''' Adds a clause for the database
    ''' </summary>
    ''' <param name="pSqlString"></param>
    ''' <param name="pThisClause"></param>
    Protected Sub addClause(ByRef pSqlString As String, ByVal pThisClause As String)

        If Not (pThisClause Is Nothing) Then
            If (pThisClause.Trim.Length > 0) Then
                pSqlString = pThisClause
            End If
        End If

    End Sub


    ''' <summary>
    ''' Caluse for the relate id
    ''' </summary>
    ''' <returns></returns>
    Protected ReadOnly Property clauseRelateId() As String
        Get
            Dim pThisString As String = ""

            If Not (MyBase.RelateIdValidator Is Nothing) Then
                If Not (MyBase.RelateIdValidator.IsInvalid) Then
                    pThisString = String.Concat(" (relateid = '", RelateIdValidator.BestGuessRelateId, "') ")
                End If
            End If

            Debug.WriteLine(pThisString)
            Return pThisString
        End Get
    End Property

    ''' <summary>
    ''' Sets up the standard where clause
    ''' </summary>
    ''' <param name="fieldName">Field to update</param>
    ''' <param name="value">Value</param>
    ''' <param name="pOperator">Operator for value and field</param>
    ''' <returns></returns>
    Protected Function StandardWhereClause(ByRef fieldName As String, ByRef value As Nullable(Of Decimal), Optional ByRef pOperator As String = "=") As String
        Dim pStandardWithQuote As String = StandardWhereClause(fieldName, value.ToString, pOperator)

        If (pStandardWithQuote Is Nothing) Then
            Return pStandardWithQuote
        End If
        Return pStandardWithQuote.Replace("'", "")
    End Function

    ''' <summary>
    ''' Standard where clause/ builds string
    ''' </summary>
    ''' <param name="fieldName">Field to update</param>
    ''' <param name="value">Value</param>
    ''' <param name="pOperator">Operator for value and field</param>
    ''' <returns></returns>
    Protected Function StandardWhereClause(ByRef fieldName As String, ByRef value As String, Optional ByRef pOperator As String = "=") As String
        StandardWhereClause = Nothing

        If Not (value Is Nothing) Then
            If (value.Trim.Length > 0) Then
                If (fieldName.Trim.Length > 0) Then
                    StandardWhereClause = String.Concat(" (" + fieldName + " " + pOperator + " '", value, "') ")
                End If
            End If
        End If
    End Function

    ''' <summary>
    ''' Builds the standard clause for the database
    ''' </summary>
    ''' <param name="fieldName">Field to update</param>
    ''' <param name="value">Value</param>
    ''' <param name="pOperator">Operator for value and field</param>
    ''' <returns></returns>
    Protected Function StandardWhereClause(ByRef fieldName As String, ByRef value As List(Of String), Optional ByRef pOperator As String = "in") As String
        Dim pStandardWithQuote As String '= StandardWhereClause(fieldName, value.ToString, pOperator)

        If (value Is Nothing) Then
            Return Nothing
        ElseIf (value.Count = 0) Then
            Return Nothing
        End If
        pStandardWithQuote = " (" + fieldName + " " + pOperator + " ("

        Dim pString As String

        For Each pString In value
            pStandardWithQuote += "'" + pString + "', "
        Next

        pStandardWithQuote = RTrim(pStandardWithQuote)
        pStandardWithQuote = pStandardWithQuote.Remove(pStandardWithQuote.Length - 1)
        pStandardWithQuote += "))"
        Return pStandardWithQuote
    End Function

    ''' <summary>
    ''' Gets the qdix cursor
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function QdixICursor() As ICursor

        Dim pCursorBuilder As mgs.CursorBuilder.IAttributeCursorBuilder
        pCursorBuilder = New mgs.CursorBuilder.AttributeCursorBuilder

        Dim pCursor As ICursor
        pCursor = pCursorBuilder.ReturnCursor(MyBase.DataAccess.GetFeatureLayer(BusinessLogic.NamedFeatureClass.qdix), Me.SqlString, "ORDER BY RELATEID", True)

        Return pCursor
    End Function



End Class
