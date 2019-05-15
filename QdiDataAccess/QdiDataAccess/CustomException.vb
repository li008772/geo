Option Strict On
Option Explicit On

'Throw (New CustomException("Invalid Database Specification"))
''' <summary>
''' A custom exception class that allows for more specific exceptions than the standard system exception.
''' </summary>
Friend Class CustomException
    Inherits System.Exception
    Sub New(ByVal ErrMsg As String)
        MyBase.New(ErrMsg)
    End Sub
End Class
