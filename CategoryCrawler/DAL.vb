Imports MySql.Data.MySqlClient

Public Class DAL

    Public Shared Sub updIsParser(ByVal intID As Integer)
        Dim strSql As String = "update category set is_parser = 1 where id = ?id"
        MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql, _
                                    New MySqlParameter("id", intID))
    End Sub

    Public Shared Sub updIsParserContent(ByVal intID As Integer)
        Dim strSql As String = "update category set is_parser_content = 1 where id = ?id"
        MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql, _
                                    New MySqlParameter("id", intID))
    End Sub

    Public Shared Sub insCategory(ByVal intParentID As Integer, strName As String, ByVal strWebUrl As String, ByVal intLevel As Integer, ByVal intIsParser As Integer, intIsParserContent As Integer, ByVal intType As Integer, ByVal intStatus As Integer)
        Dim strSql As String = String.Empty
        strSql += "insert into category (parent_id, name, url, level, is_parser, is_parser_content, type, status, cre_date, cre_time) " & vbCrLf
        strSql += "value (?parent_id, ?name, ?url, ?level, ?is_parser, ?is_parser_content, ?type, ?status, ?cre_date, ?cre_time) " & vbCrLf
        MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql, _
                                           New MySqlParameter("?parent_id", intParentID), _
                                           New MySqlParameter("?name", strName), _
                                           New MySqlParameter("?url", strWebUrl), _
                                           New MySqlParameter("?level", intLevel), _
                                           New MySqlParameter("?is_parser", intIsParser), _
                                           New MySqlParameter("?is_parser_content", intIsParserContent), _
                                           New MySqlParameter("?type", intType), _
                                           New MySqlParameter("?status", intStatus), _
                                           New MySqlParameter("?cre_date", Now.ToShortDateString), _
                                           New MySqlParameter("?cre_time", Now.ToString("yyyy/MM/dd HH:mm:ss").Substring(11, 8)))

    End Sub

    Public Shared Function getCategoryID(ByVal intParentID As Integer, ByVal strName As String, ByVal strWebUrl As String, ByVal intLevel As Integer, ByVal intType As Integer, ByVal intStatus As Integer) As Integer
        Dim strValue As String = "0"
        Dim strSql As String = String.Empty
        strSql = "select id from category where parent_id = ?parent_id and name = ?name and url = ?url and level = ?level and type = ?type and status = ?status"
        Dim ds As New DataSet
        ds = MySqlHelper.ExecuteDataset(ComConfig.DB.connectString, strSql, _
                                           New MySqlParameter("?parent_id", intParentID), _
                                           New MySqlParameter("?name", strName), _
                                           New MySqlParameter("?url", strWebUrl), _
                                           New MySqlParameter("?level", intLevel), _
                                           New MySqlParameter("?type", intType), _
                                           New MySqlParameter("?status", intStatus))
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            strValue = ds.Tables(0).Rows(0).Item("id").ToString
        End If
        ds.Dispose()
        Return Integer.Parse(strValue)
    End Function

    Public Shared Function getCatogoryContent(ByVal intLevel As Integer) As DataSet
        Dim strSql As String = String.Empty
        strSql = "select id, level, url from category where type = 4 and status = 1 and level = ?level"
        Return MySqlHelper.ExecuteDataset(ComConfig.DB.connectString, strSql, _
                                           New MySqlParameter("?level", intLevel))
    End Function



    Public Shared Function getCategoryLevel3() As DataSet
        Dim strSql As String = String.Empty
        strSql += "select min(id) as minNum, min(id)+2000 as maxNum from category " & vbCrLf
        strSql += "where type = 4 and status = 1 and level = 3 " & vbCrLf
        strSql += "	and url <> 'https://tw.buy.yahoo.com' and name <> '更多'  " & vbCrLf
        strSql += "	and is_parser = 0 " & vbCrLf

        Dim intStart As Integer = 0
        Dim intEnd As Integer = 0
        Dim ds As New DataSet
        ds = MySqlHelper.ExecuteDataset(ComConfig.DB.connectString, strSql)
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            intStart = ds.Tables(0).Rows(0).Item("minNum").ToString
            intEnd = ds.Tables(0).Rows(0).Item("maxNum").ToString
        Else
            strSql = String.Empty
            strSql += "update category set " & vbCrLf
            strSql += "is_parser = 0 " & vbCrLf
            strSql += "where type = 4 and status = 1 and level = 3 " & vbCrLf
            strSql += "	and url <> 'https://tw.buy.yahoo.com' and name <> '更多'  " & vbCrLf
            strSql += "	and is_parser = 1 " & vbCrLf
            MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql)
        End If
        ds.Dispose()

        strSql = String.Empty
        strSql += "select id, url from category " & vbCrLf
        strSql += "where type = 4 and status = 1 and level = 3 " & vbCrLf
        strSql += "	and url <> 'https://tw.buy.yahoo.com' and name <> '更多' " & vbCrLf
        strSql += "	and is_parser = 0 " & vbCrLf
        strSql += "	and id between ?intStart and ?intEnd " & vbCrLf
        strSql += "order by id " & vbCrLf
        Return MySqlHelper.ExecuteDataset(ComConfig.DB.connectString, strSql, _
                                           New MySqlParameter("?intStart", intStart), _
                                           New MySqlParameter("?intEnd", intEnd))
    End Function
End Class
