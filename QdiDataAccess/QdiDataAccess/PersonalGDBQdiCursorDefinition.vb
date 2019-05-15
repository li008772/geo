Option Strict On
Option Explicit On

Imports ESRI.ArcGIS.Geodatabase
Imports mgs.CursorBuilder


Public Class PersonalQdiCursorDefinition
    Inherits QdiCursorBaseClass

    Friend Sub New(ByVal pPersonalGDBDataAccess As PersonalGDBDataAccess)
        Me.DataAccess = pPersonalGDBDataAccess
    End Sub

    ''' <summary>
    '''     Builds and returns a long String of SQL statements.
    ''' </summary>
    ''' <returns>A String with multiple SQL statements.</returns>
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

            If (returnString.Trim = "") Then
                addClause(returnString, StandardWhereClause("relateid", "NULL", pOperator:="<>"))
            End If

            Debug.WriteLine(returnString)
            Return returnString
        End Get
    End Property

    ''' <summary>
    '''     If there's a valid clause, it's concatenated to a long String of SQL statements.
    ''' </summary>
    ''' <param name="pSqlString">String to store the built SQL statements</param>
    ''' <param name="pThisClause">Clause to be added to the long String of SQL statements.</param>
    Protected Sub addClause(ByRef pSqlString As String, ByVal pThisClause As String)

        If Not (pThisClause Is Nothing) Then
            If (pThisClause.Trim.Length > 0) Then
                If (pSqlString.Trim.Length > 0) Then
                    pSqlString += (String.Concat(" And ", pThisClause))
                Else
                    pSqlString = pThisClause
                End If
            End If
        End If

    End Sub

    ''' <summary>
    '''     Checks for valid RelateID, then builds the clause.
    ''' </summary>
    ''' <returns>A validly built RelateID clause.</returns>
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
    '''     Builds a where clause by error checking and then adding quotes.
    ''' </summary>
    ''' <param name="fieldName">The field to search.</param>
    ''' <param name="value">The value to search for, which gets cast as String by the function.</param>
    ''' <param name="pOperator">Optional argument for operator, it's assumed to be "=" by default.</param>
    ''' <returns>A where clause with quotes added in.</returns>
    Protected Function StandardWhereClause(ByRef fieldName As String, ByRef value As Nullable(Of Decimal), Optional ByRef pOperator As String = "=") As String
        Dim pStandardWithQuote As String = StandardWhereClause(fieldName, value.ToString, pOperator)

        If (pStandardWithQuote Is Nothing) Then
            Return pStandardWithQuote
        End If
        Return pStandardWithQuote.Replace("'", "")
    End Function

    ''' <summary>
    '''     Builds a where clause by error checking and then concatenating each term.
    ''' </summary>
    ''' <param name="fieldName">The field to search.</param>
    ''' <param name="value">The value to search for.</param>
    ''' <param name="pOperator">Optional argument for operator, it's assumed to be "=" by default.</param>
    ''' <returns>A where clause with quotes added in.</returns>
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
    '''     Builds a where clause by error checking and adding the appropriate spaces, parentheses, and quotes.
    ''' </summary>
    ''' <param name="fieldName">The field to search.</param>
    ''' <param name="value">The values to search for.</param>
    ''' <param name="pOperator">Optional argument for operator, it's assumed to be "in" by default.</param>
    ''' <returns></returns>
    Protected Function StandardWhereClause(ByRef fieldName As String, ByRef value As List(Of String), Optional ByRef pOperator As String = "in") As String
        Dim pStandardWithQuote As String '= StandardWhereClause(fieldName, value.ToString, pOperator)

        If (value.Count = 0) Then
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
    '''     Creates a cursor for the QDIX layer.
    ''' </summary>
    ''' <returns>A cursor based off of the contents of SqlString</returns>
    Public Overrides Function QdixICursor() As ICursor

        Dim pCursorBuilder As mgs.CursorBuilder.IAttributeCursorBuilder
        pCursorBuilder = New mgs.CursorBuilder.AttributeCursorBuilder

        Dim pCursor As ICursor
        pCursor = pCursorBuilder.ReturnCursor(MyBase.DataAccess.GetFeatureLayer(BusinessLogic.NamedFeatureClass.qdix), Me.SqlString, "ORDER BY RELATEID", True)

        Return pCursor
    End Function



End Class
