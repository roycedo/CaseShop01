Public Class BLL

    Public Shared Sub updIsParser(ByVal intID As Integer)
        DAL.updIsParser(intID)
    End Sub

    Public Shared Sub updIsParserContent(ByVal intID As Integer)
        DAL.updIsParserContent(intID)
    End Sub

    Public Shared Sub insCategory(ByVal intParentID As Integer, strName As String, ByVal strWebUrl As String, ByVal intLevel As Integer, ByVal intIsParser As Integer, intIsParserContent As Integer, ByVal intType As Integer, ByVal intStatus As Integer)
        If DAL.getCategoryID(intParentID, strName, strWebUrl, intLevel, intType, intStatus) = 0 Then
            DAL.insCategory(intParentID, strName, strWebUrl, intLevel, intIsParser, intIsParserContent, intType, intStatus)
        End If
    End Sub

    Public Shared Function getCategoryID(ByVal intParentID As Integer, ByVal strName As String, ByVal strWebUrl As String, ByVal intLevel As Integer, ByVal intType As Integer, ByVal intStatus As Integer) As Integer
        Return DAL.getCategoryID(intParentID, strName, strWebUrl, intLevel, intType, intStatus)
    End Function

    Public Shared Function getCatogoryContent(ByVal intLevel As Integer) As DataSet
        Return DAL.getCatogoryContent(intLevel)
    End Function



    Public Shared Function getCategoryLevel3() As DataSet
        Return DAL.getCategoryLevel3()
    End Function
End Class
