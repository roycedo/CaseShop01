Imports MySql.Data.MySqlClient

Public Class DAL
   
    Public Shared Function testDBConnection() As Integer
        Dim strSql As String = String.Empty
        strSql = "select count(*) from keyword where id >=1 and id <= 150"
        Return MySqlHelper.ExecuteScalar(ComConfig.DB.connectString, strSql)
    End Function

    Public Shared Function getKeyword() As DataSet
        Dim strSql As String = String.Empty
        strSql += "select min(id) as minNum, min(id)+800 as maxNum from keyword where is_parser = 0 " & vbCrLf
        
        Dim intStart As Integer = 0
        Dim intEnd As Integer = 0
        Dim ds As New DataSet
        ds = MySqlHelper.ExecuteDataset(ComConfig.DB.connectString, strSql)
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            intStart = ds.Tables(0).Rows(0).Item("minNum").ToString
            intEnd = ds.Tables(0).Rows(0).Item("maxNum").ToString
        Else
            strSql = String.Empty
            strSql += "update keyword set " & vbCrLf
            strSql += "is_parser = 0 " & vbCrLf
            strSql += "where is_parser = 1 " & vbCrLf
            MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql)
        End If
        ds.Dispose()

        strSql = String.Empty
        strSql += "select id, name, level from keyword " & vbCrLf
        strSql += "where is_parser = 0 and id between ?intStart and ?intEnd " & vbCrLf
        strSql += "order by id " & vbCrLf
        Return MySqlHelper.ExecuteDataset(ComConfig.DB.connectString, strSql, _
                                           New MySqlParameter("?intStart", intStart), _
                                           New MySqlParameter("?intEnd", intEnd))
    End Function

    Public Shared Function insKeyword(ByVal intSearchEngine As Integer, ByVal strResult As String, ByVal strQueryKeyword As String, ByVal intKeywordID As Integer, ByVal intLevel As Integer) As String
        Dim strSql As String = String.Empty
        Dim arrResult As Array
        strResult = strResult.Replace("@@@@@", "＃")
        arrResult = strResult.Split("＃")

        ''Keyword
        Dim arrKeywordCollection As Array = arrResult(0).ToString.Split("♣")
        For i As Integer = 0 To arrKeywordCollection.Length - 1
            Dim arrKeyword As Array = arrKeywordCollection(i).ToString.Split("♠")
            Dim strTitle As String = arrKeyword(0).ToString
            If strTitle.Length >= 100 Then
                strTitle = strTitle.Substring(0, 98)
            End If
            Dim strWebUrl As String = arrKeyword(1).ToString
            Dim strDesc As String = arrKeyword(2).ToString
            If intSearchEngine = ComConfig.clConfig.SearchEngine.Google Then
                strDesc = ComUtility.StringPlus.ConvertBig5(strDesc)
            End If
            If strDesc.Length >= 255 Then
                strDesc = strDesc.Substring(0, 253)
            End If
            Dim strDomain As String = ComUtility.StringPlus.ExtractDomainFromURL(strWebUrl)
            If strDomain.Length >= 30 Then
                strDomain = strDomain.Substring(0, 28)
            End If

            If strTitle <> "" And strWebUrl <> "" And strDesc <> "" Then
                strSql = String.Empty

                strSql += "insert into parser_keyword_nature1 (keyword_id, title, description, url, domain_name, status, cre_date, cre_time) " & vbCrLf
                strSql += "values (?keyword_id, ?title, ?description, ?url, ?domain_name, ?status, ?cre_date, ?cre_time) " & vbCrLf
                MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql, _
                                            New MySqlParameter("?keyword_id", intKeywordID), _
                                            New MySqlParameter("?title", strTitle), _
                                            New MySqlParameter("?description", strDesc), _
                                            New MySqlParameter("?url", strWebUrl), _
                                            New MySqlParameter("?domain_name", strDomain), _
                                            New MySqlParameter("?status", IIf(strWebUrl = "" And strDesc = "", 0, 1)), _
                                            New MySqlParameter("?cre_date", Now.ToShortDateString), _
                                            New MySqlParameter("?cre_time", Now.ToString("yyyy/MM/dd HH:mm:ss").Substring(11, 8)))
            End If
        Next

        ''KeywordRelative
        Dim arrKeywordRelativeCollection As Array = arrResult(1).ToString.Split(",")
        For i As Integer = 0 To arrKeywordRelativeCollection.Length - 1
            Dim strName As String = arrKeywordRelativeCollection(i).ToString.Trim
            If isNameUsed(strName) = False Then
                strSql = String.Empty
                strSql += "insert into keyword1 (parent_id, name, name_length, level, type, status, cre_date, cre_time) " & vbCrLf
                strSql += "values (?parent_id, ?name, ?name_length, ?level, ?type, ?status, ?cre_date, ?cre_time) " & vbCrLf
                Dim intExe As Integer = MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql, _
                                            New MySqlParameter("?parent_id", intKeywordID), _
                                            New MySqlParameter("?name", strName), _
                                            New MySqlParameter("?name_length", strName.Length), _
                                            New MySqlParameter("?level", intLevel + 1), _
                                            New MySqlParameter("?type", intSearchEngine), _
                                            New MySqlParameter("?status", 1), _
                                            New MySqlParameter("?cre_date", Now.ToShortDateString), _
                                            New MySqlParameter("?cre_time", Now.ToString("yyyy/MM/dd HH:mm:ss").Substring(11, 8)))

                If intExe > 0 Then
                    strSql = String.Empty
                    strSql = "select max(id) as maxid from keyword1"
                    Dim intNewKeywordID As Integer = MySqlHelper.ExecuteScalar(ComConfig.DB.connectString, strSql)

                    strSql = String.Empty
                    strSql += "insert into keyword_relate1 (keyword_id, map_keyword_id, status, cre_date, cre_time) " & vbCrLf
                    strSql += "values (?keyword_id, ?map_keyword_id, ?status, ?cre_date, ?cre_time) " & vbCrLf
                    MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql, _
                                            New MySqlParameter("?keyword_id", intKeywordID), _
                                            New MySqlParameter("?map_keyword_id", intNewKeywordID), _
                                            New MySqlParameter("?status", 1), _
                                            New MySqlParameter("?cre_date", Now.ToShortDateString), _
                                            New MySqlParameter("?cre_time", Now.ToString("yyyy/MM/dd HH:mm:ss").Substring(11, 8)))

                    Dim intIsParserContent As Integer = IIf(arrKeywordCollection.Length > 0, 1, 0)
                    strSql = String.Empty
                    strSql += "update keyword1 set " & vbCrLf
                    strSql += "is_parser = ?is_parser, is_parser_content = ?is_parser_content, is_parser_relate = ?is_parser_relate " & vbCrLf
                    strSql += "where id = ?id" & vbCrLf
                    MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql, _
                                            New MySqlParameter("?is_parser", 1), _
                                            New MySqlParameter("?is_parser_content", intIsParserContent), _
                                            New MySqlParameter("?is_parser_relate", 1), _
                                            New MySqlParameter("?id", intKeywordID))
                End If
            End If
        Next

        Return strQueryKeyword & "OK"
    End Function

    Private Shared Function isNameUsed(ByVal strName) As Boolean
        Dim strSql As String = String.Empty
        strSql += "select count(*) from keyword1 where name = ?name"
        Dim intCount As Integer = Integer.Parse(MySqlHelper.ExecuteScalar(ComConfig.DB.connectString, strSql, New MySqlParameter("?name", strName)))
        Return IIf(intCount > 0, True, False)
    End Function

    Public Shared Sub insErrorMessageKeyword(ByVal intType As Integer, ByVal intKeywordID As Integer, ByVal strErrorMsg As String, ByVal strUserAgent As String)
        Dim strSql As String = String.Empty
        strSql += "insert into errmsg_keyword (type, keyword_id, errormsg, cre_date, cre_time, useragent) " & vbCrLf
        strSql += "values (?type, ?keyword_id, ?errormsg, ?cre_date, ?cre_time, ?useragent) " & vbCrLf
        MySqlHelper.ExecuteNonQuery(ComConfig.DB.connectString, strSql, _
                                            New MySqlParameter("?keyword_id", intKeywordID), _
                                            New MySqlParameter("?type", intType), _
                                            New MySqlParameter("?errormsg", strErrorMsg), _
                                            New MySqlParameter("?useragent", strUserAgent), _
                                            New MySqlParameter("?cre_date", Now.ToShortDateString), _
                                            New MySqlParameter("?cre_time", Now.ToString("yyyy/MM/dd HH:mm:ss").Substring(11, 8)))
    End Sub

End Class
