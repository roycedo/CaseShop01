Module Module1

    Sub Main()
        Process()
    End Sub

    Public Sub Process()
        Dim strMsg As String = String.Empty
        Dim ds As New DataSet
        ds = BLL.getKeyword
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                strMsg += CrawlKeyword(ds.Tables(0).Rows(i).Item("name").ToString, Integer.Parse(ds.Tables(0).Rows(i).Item("id").ToString), Integer.Parse(ds.Tables(0).Rows(i).Item("level").ToString))
            Next
        End If
        ds.Dispose()

        Console.WriteLine(strMsg)
        Console.Read()
    End Sub

    Private Function CrawlKeyword(ByVal strQueryKeyword As String, ByVal intKeywordID As Integer, ByVal intLevel As Integer) As String
        Dim lstSearchEngine As New List(Of Integer) From {ComConfig.clConfig.SearchEngine.Yahoo, ComConfig.clConfig.SearchEngine.Google}
        Dim strResult As String = String.Empty
        Dim strMsg As String = String.Empty

        For Each intSearchEngine As Integer In lstSearchEngine
            strResult = ComUtility.Crawler.CrawlSearchEngine(intSearchEngine, strQueryKeyword)

            If strResult.Length > 0 AndAlso strResult.Substring(0, 5) <> "Error" Then
                strMsg += saveData(intSearchEngine, strResult, strQueryKeyword, intKeywordID, intLevel)
            Else
                If strResult.Substring(0, 5) = "Error" Then
                    Dim arrTmp As Array = strResult.Split(":")
                    BLL.insErrorMessageKeyword(intSearchEngine, intKeywordID, arrTmp(1).ToString, arrTmp(2).ToString)
                    strResult = intSearchEngine & "Error.<br /><br />"
                End If
                strMsg += strResult
            End If

            System.Threading.Thread.Sleep(4000)
        Next

        Return strMsg
    End Function

    Private Function saveData(ByVal intSearchEngine As Integer, ByVal strResult As String, ByVal strQueryKeyword As String, ByVal intKeywordID As Integer, ByVal intLevel As Integer) As String
        Return BLL.insKeyword(intSearchEngine, strResult, strQueryKeyword, intKeywordID, intLevel)
    End Function

    Private Sub testDBConnection()
        Console.WriteLine("test Connection:" & BLL.testDBConnection)
        Console.Read()
    End Sub

End Module
