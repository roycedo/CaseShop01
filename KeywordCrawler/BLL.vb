Public Class BLL

    Public Shared Function testDBConnection() As Integer
        Return DAL.testDBConnection
    End Function

    Public Shared Function getKeyword() As DataSet
        Return DAL.getKeyword()
    End Function

    Public Shared Function insKeyword(ByVal intSearchEngine As Integer, ByVal strResult As String, ByVal strQueryKeyword As String, ByVal intKeywordID As Integer, ByVal intLevel As Integer) As String
        Return DAL.insKeyword(intSearchEngine, strResult, strQueryKeyword, intKeywordID, intLevel)
    End Function

    Public Shared Sub insErrorMessageKeyword(ByVal intType As Integer, ByVal intKeywordID As Integer, ByVal strErrorMsg As String, ByVal strUserAgent As String)
        DAL.insErrorMessageKeyword(intType, intKeywordID, strErrorMsg, strUserAgent)
    End Sub

End Class
