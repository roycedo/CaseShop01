Imports System.Net
Imports System.IO
Imports System.Text
Imports HtmlAgilityPack

Partial Public Class Crawler

#Region "Keyword & KeywordRelative"

    ''' <summary>
    ''' 搜尋引擎關鍵字搜尋
    ''' </summary>
    ''' <param name="strQueryKeyword"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CrawlSearchEngine(ByVal intSearchEngine As String, ByVal strQueryKeyword As String) As String
        '指定來源網頁
        Dim WebClient As New WebClient()
        Dim strUserAgent As String = ComConfig.clConfig.UserAgent(0, True).ToUpper
        WebClient.Headers(HttpRequestHeader.UserAgent) = strUserAgent

        '將網頁來源資料暫存到記憶體內
        Dim strSearchUrl As String = String.Empty
        Select Case intSearchEngine
            Case ComConfig.clConfig.SearchEngine.Google
                strSearchUrl = "https://www.google.com.tw/search?filter=0&pws=0&hl=zh-TW&op=" & strQueryKeyword & "&q=" & strQueryKeyword
            Case ComConfig.clConfig.SearchEngine.Yahoo
                strSearchUrl = "http://tw.search.yahoo.com/search?p=" & strQueryKeyword & "&fr=yfp&ei=utf-8&v=0&pstart=1&b=1"
        End Select
        Dim ms As New MemoryStream(WebClient.DownloadData(strSearchUrl))

        '使用預設編碼讀入 HTML 
        Dim doc As New HtmlDocument
        Select Case intSearchEngine
            Case ComConfig.clConfig.SearchEngine.Google
                doc.Load(ms, Encoding.[Default])
            Case ComConfig.clConfig.SearchEngine.Yahoo
                doc.Load(ms, Encoding.UTF8)
        End Select

        Dim hdc As New HtmlDocument

        '關鍵字結果
        Dim strErrKeyword As String = String.Empty
        Dim strKeyword As String = String.Empty
        Try
            Select Case intSearchEngine
                Case ComConfig.clConfig.SearchEngine.Google
                    strKeyword = CrawlGoogleKeyword(doc, hdc)
                Case ComConfig.clConfig.SearchEngine.Yahoo
                    strKeyword = CrawlYahooKeyword(doc, hdc)
            End Select
        Catch ex As Exception
            strErrKeyword = intSearchEngine & " Keyword-" & strQueryKeyword & "<br /><br />Error-" & ex.Message & "<br /><br />"
        End Try

        '關鍵字相關結果
        Dim strErrKeywordRelative As String = String.Empty
        Dim strKeywordRelative As String = String.Empty
        Try
            Select Case intSearchEngine
                Case ComConfig.clConfig.SearchEngine.Google
                    strKeywordRelative = CrawlGoogleKeywordRelative(doc, hdc)
                Case ComConfig.clConfig.SearchEngine.Yahoo
                    strKeywordRelative = CrawlYahooKeywordRelative(doc, hdc)
            End Select
        Catch ex As Exception
            strErrKeywordRelative = intSearchEngine & " KeywordRelative-" & strQueryKeyword & "<br /><br />Error-" & ex.Message & "<br /><br />"
        End Try

        '結果
        If (strErrKeyword & strErrKeywordRelative).Length > 0 Then
            Return ("Error: 1-" & strErrKeyword & strErrKeywordRelative & ":" & strUserAgent)
        Else
            If strKeyword.Length = 0 And strKeywordRelative.Length = 0 Then
                Return "Error: 2-No Data.:" & strUserAgent 'CrawlGoogleKeywordSimilar(doc, hdc)
            Else
                Return strKeyword & "@@@@@" & strKeywordRelative
            End If
        End If
    End Function

    ''' <summary>
    ''' Google 關鍵字結果
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CrawlGoogleKeyword(ByVal doc As HtmlDocument, ByVal hdc As HtmlDocument) As String
        'XPath 來解讀它
        'Dim strXPath As String = "/html/body/table[1]/tbody/tr[1]/td[2]/div/div[2]/div[2]/div[1]/*"
        Dim strXPath As String = "/html/body/table[@id='mn']/tbody//*[@id='center_col']/div[@id='res']/div[@id='search']/div[@id='ires']/*"
        Dim docHtmlNode As HtmlNode = doc.DocumentNode.SelectSingleNode(strXPath)

        Dim strContent As String = String.Empty
        If Not IsNothing(docHtmlNode) Then
            hdc.LoadHtml(docHtmlNode.InnerHtml)

            Dim LiHtmlNodeCollection As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./li[@class='g']")

            For Each HtmlNode As HtmlNode In LiHtmlNodeCollection
                Dim nodeTitle As HtmlNode = HtmlNode.SelectSingleNode("./h3")
                Dim strTitle As String = String.Empty
                If Not IsNothing(nodeTitle) Then strTitle = nodeTitle.InnerText.Trim()

                Dim nodeWebUrl As HtmlNode = HtmlNode.SelectSingleNode("./div/div/cite")
                Dim strWebUrl As String = String.Empty
                If Not IsNothing(nodeWebUrl) Then strWebUrl = nodeWebUrl.InnerText.Trim()

                Dim nodeDesc As HtmlNode = HtmlNode.SelectSingleNode("./div/span")
                Dim strDesc As String = String.Empty
                If Not IsNothing(nodeDesc) Then strDesc = WebUtility.HtmlDecode(nodeDesc.InnerHtml.Trim())

                If strTitle.Length > 0 And strWebUrl.Length > 0 And strDesc.Length > 0 Then
                    strContent += strTitle & "♠" & strWebUrl & "♠" & strDesc & "♣"
                End If
            Next
            strContent = IIf(strContent.Length >= 1, strContent.Substring(0, strContent.Length - 1), strContent)
        End If

        Return strContent
    End Function

    ''' <summary>
    ''' Google 關鍵字相關結果
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CrawlGoogleKeywordRelative(ByVal doc As HtmlDocument, ByVal hdc As HtmlDocument) As String
        'XPath 來解讀它
        'Dim strXPath As String = "/html/body/table[1]/tbody/tr[1]/td[2]/div/div/table"
        Dim strXPath As String = "/html/body/table[@id='mn']/tbody//*[@id='center_col']/div/table"
        Dim docHtmlNode As HtmlNode = doc.DocumentNode.SelectSingleNode(strXPath)

        Dim strContent As String = String.Empty
        If Not IsNothing(docHtmlNode) Then
            hdc.LoadHtml(docHtmlNode.InnerHtml)

            Dim trHtmlNodeCollection As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./tr")
            Dim tdHtmlNodeCollection As HtmlNodeCollection

            For intTr As Integer = 1 To trHtmlNodeCollection.Count
                tdHtmlNodeCollection = hdc.DocumentNode.SelectNodes("./tr[" & intTr & "]/td")
                For intTd As Integer = 1 To tdHtmlNodeCollection.Count
                    Dim strKeywordRelative As String = hdc.DocumentNode.SelectSingleNode("./tr[" & intTr & "]/td[" & intTd & "]").InnerText.Trim()
                    strContent += strKeywordRelative & ","
                Next
            Next
            strContent = IIf(strContent.Length >= 1, strContent.Substring(0, strContent.Length - 1), strContent)
        End If

        Return strContent
    End Function

    ''' <summary>
    ''' Yahoo 關鍵字結果
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CrawlYahooKeyword(ByVal doc As HtmlDocument, ByVal hdc As HtmlDocument) As String
        'XPath 來解讀它
        Dim strXPath As String = "/html/body/div/div[3]/div[2]//*[@id='web']/ol" ''result
        Dim docHtmlNode As HtmlNode = doc.DocumentNode.SelectSingleNode(strXPath)

        Dim strContent As String = String.Empty
        If Not IsNothing(docHtmlNode) Then
            hdc.LoadHtml(docHtmlNode.InnerHtml)

            Dim LiHtmlNodeCollection As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./li/div[@class='res']")

            For Each HtmlNode As HtmlNode In LiHtmlNodeCollection
                Dim nodeTitle As HtmlNode = HtmlNode.SelectSingleNode("./div/h3")
                Dim strTitle As String = String.Empty
                If Not IsNothing(nodeTitle) Then strTitle = nodeTitle.InnerText.Trim()

                Dim nodeWebUrl As HtmlNode = HtmlNode.SelectSingleNode("./span[@class='url']")
                Dim strWebUrl As String = String.Empty
                If Not IsNothing(nodeWebUrl) Then strWebUrl = nodeWebUrl.InnerText.Trim()

                Dim nodeDesc As HtmlNode = HtmlNode.SelectSingleNode("./div[@class='abstr']")
                Dim strDesc As String = String.Empty
                If Not IsNothing(nodeDesc) Then strDesc = WebUtility.HtmlDecode(nodeDesc.InnerHtml.Trim())

                If strTitle.Length > 0 And strWebUrl.Length > 0 And strDesc.Length > 0 Then
                    strContent += strTitle & "♠" & strWebUrl & "♠" & strDesc & "♣"
                End If
            Next
            strContent = IIf(strContent.Length >= 1, strContent.Substring(0, strContent.Length - 1), strContent)
        End If

        Return strContent
    End Function

    ''' <summary>
    ''' Yahoo 關鍵字相關結果
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CrawlYahooKeywordRelative(ByVal doc As HtmlDocument, ByVal hdc As HtmlDocument) As String
        'XPath 來解讀它
        Dim strXPath As String = "/html/body/div/div[3]/div[2]//*[@id='satat']/table/tbody" ''result
        Dim docHtmlNode As HtmlNode = doc.DocumentNode.SelectSingleNode(strXPath)

        Dim strContent As String = String.Empty
        If Not IsNothing(docHtmlNode) Then
            hdc.LoadHtml(docHtmlNode.InnerHtml)

            Dim trHtmlNodeCollection As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./tr")
            Dim tdHtmlNodeCollection As HtmlNodeCollection
            For intTr As Integer = 1 To trHtmlNodeCollection.Count
                tdHtmlNodeCollection = hdc.DocumentNode.SelectNodes("./tr[" & intTr & "]/td")
                For intTd As Integer = 1 To tdHtmlNodeCollection.Count
                    Dim strKeywordRelative As String = hdc.DocumentNode.SelectSingleNode("./tr[" & intTr & "]/td[" & intTd & "]").InnerText.Trim()
                    strContent += strKeywordRelative & ","
                Next
            Next
            strContent = IIf(strContent.Length >= 1, strContent.Substring(0, strContent.Length - 1), strContent)
        End If

        Return strContent
    End Function

#End Region

End Class
