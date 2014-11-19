Imports System.Net
Imports System.IO
Imports System.Text
Imports HtmlAgilityPack

Module Module1
    Dim strLinkUrl As String = "https://tw.buy.yahoo.com"

    Sub Main()
        ''取得所有類別
        Process()
        ''取得所以商品連結
        CrawYahooCategoryLevel3()
    End Sub

    Public Sub Process()
        Dim lstSearchEngine As New List(Of Integer) From {ComConfig.clConfig.SearchEngine.Yahoo}
        Dim strResult As String = String.Empty
        Dim strMsg As String = String.Empty
        Dim strSiteMapUrl As String = String.Empty

        For Each intSearchEngine As Integer In lstSearchEngine
            Select Case intSearchEngine
                Case ComConfig.clConfig.SearchEngine.Yahoo
                    strSiteMapUrl = "https://tw.buy.yahoo.com/help/helper.asp?p=sitemap"
                Case ComConfig.clConfig.SearchEngine.PCHome
                    strSiteMapUrl = "http://24h.pchome.com.tw/?mod=sitemap"
            End Select

            CrawlSiteMap(intSearchEngine, strSiteMapUrl)
        Next
    End Sub

    Private Sub CrawlSiteMap(ByVal intSearchEngine As Integer, ByVal strSiteMapUrl As String)
        '指定來源網頁
        Dim WebClient As New WebClient()
        Dim strUserAgent As String = ComConfig.clConfig.UserAgent(0, True).ToUpper
        WebClient.Headers(HttpRequestHeader.UserAgent) = strUserAgent

        ''將網頁來源資料暫存到記憶體內
        Dim ms As New MemoryStream(WebClient.DownloadData(strSiteMapUrl))

        ''使用預設編碼讀入 HTML 
        Dim doc As New HtmlDocument
        'doc.Load(ms, Encoding.Default)

        Dim hdc As New HtmlDocument

        Select Case intSearchEngine
            Case ComConfig.clConfig.SearchEngine.Yahoo
                doc.Load(ms, Encoding.Default)
                CrawlYahooSiteMap(doc, hdc)
            Case ComConfig.clConfig.SearchEngine.PCHome
                doc.Load(ms, Encoding.UTF8)
                'CrawlPCHomeSiteMap(doc, hdc)
        End Select
    End Sub

    Private Sub CrawlYahooSiteMap(ByVal doc As HtmlDocument, ByVal hdc As HtmlDocument)
        'XPath 來解讀它
        Dim strXPath As String = "/html/body//*[@id='cl-sitemap']/div/div" ''result
        Dim docHtmlNode As HtmlNode = doc.DocumentNode.SelectSingleNode(strXPath)

        Dim strContent As String = String.Empty
        If Not IsNothing(docHtmlNode) Then
            hdc.LoadHtml(docHtmlNode.InnerHtml)

            Dim UlHtmlNodeCollection As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./ul")
            Dim LiHtmlNodeCollection As HtmlNodeCollection

            For Each UlHtmlNode As HtmlNode In UlHtmlNodeCollection
                LiHtmlNodeCollection = UlHtmlNode.SelectNodes("./li")
                For Each LiHtmlNode As HtmlNode In LiHtmlNodeCollection
                    ''First Level
                    Dim strName As String = String.Empty
                    If Not IsNothing(LiHtmlNode) Then strName = LiHtmlNode.SelectSingleNode("./h3").InnerText.ToString
                    Dim strWebUrl As String = String.Empty
                    If Not IsNothing(LiHtmlNode) Then strWebUrl = ComUtility.StringPlus.getCssAhref(LiHtmlNode.SelectSingleNode("./h3").InnerHtml.ToString, strName)
                    BLL.insCategory(0, strName, strLinkUrl & strWebUrl, 0, 1, 1, ComConfig.clConfig.SearchEngine.Yahoo, 1)
                    Dim intParentID As Integer = BLL.getCategoryID(0, strName, strLinkUrl & strWebUrl, 0, ComConfig.clConfig.SearchEngine.Yahoo, 1)

                    ''Second Level
                    Dim SecondUlHtmlNodeCollection As HtmlNodeCollection = LiHtmlNode.SelectNodes("./ul")
                    Dim SecondLiHtmlNodeCollection As HtmlNodeCollection

                    For Each SecondUlHtmlNode As HtmlNode In SecondUlHtmlNodeCollection
                        SecondLiHtmlNodeCollection = SecondUlHtmlNode.SelectNodes("./li")
                        For Each SecondLiHtmlNode As HtmlNode In SecondLiHtmlNodeCollection
                            Dim strSecondName As String = String.Empty
                            If Not IsNothing(SecondLiHtmlNode) Then strSecondName = SecondLiHtmlNode.InnerText.ToString
                            Dim strSecondWebUrl As String = String.Empty
                            If Not IsNothing(SecondLiHtmlNode) Then strSecondWebUrl = ComUtility.StringPlus.getCssAhref(SecondLiHtmlNode.InnerHtml.ToString, strSecondName)

                            BLL.insCategory(intParentID, strSecondName, strLinkUrl & strSecondWebUrl, 1, 1, 0, ComConfig.clConfig.SearchEngine.Yahoo, 1)
                        Next
                        If SecondLiHtmlNodeCollection.Count > 0 Then
                            BLL.updIsParserContent(intParentID)
                        End If
                    Next
                Next
            Next

            ''Third, Fourth Level
            CrawlYahooSiteMapNextCategory(1)
        End If
    End Sub

    Private Sub CrawlYahooSiteMapNextCategory(ByVal intLevel As Integer)
        Dim dsThird As New DataSet
        dsThird = BLL.getCatogoryContent(intLevel)
        If dsThird.Tables.Count > 0 AndAlso dsThird.Tables(0).Rows.Count > 0 Then
            Dim intFirstID As Integer = 0
            Dim intFirstLevel As Integer = 0
            Dim strFirstUrl As String = String.Empty

            For i As Integer = 0 To dsThird.Tables(0).Rows.Count - 1
                intFirstID = dsThird.Tables(0).Rows(i).Item("id").ToString
                intFirstLevel = dsThird.Tables(0).Rows(i).Item("level").ToString
                strFirstUrl = dsThird.Tables(0).Rows(i).Item("url").ToString

                CrawlYahooSiteMapNextContent(intFirstID, intFirstLevel, strFirstUrl)
            Next
        End If
        dsThird.Dispose()
    End Sub

    Private Sub CrawlYahooSiteMapNextContent(ByVal intFirstID As Integer, ByVal intFirstLevel As Integer, ByVal strSiteMapUrl As String)
        '指定來源網頁
        Dim WebClient As New WebClient()
        Dim strUserAgent As String = ComConfig.clConfig.UserAgent(0, True).ToUpper
        WebClient.Headers(HttpRequestHeader.UserAgent) = strUserAgent

        ''將網頁來源資料暫存到記憶體內
        Dim ms As New MemoryStream(WebClient.DownloadData(strSiteMapUrl))

        ''使用預設編碼讀入 HTML 
        Dim doc As New HtmlDocument
        doc.Load(ms, Encoding.UTF8)

        Dim hdc As New HtmlDocument

        'XPath 來解讀它
        Dim strXPath As String = "/html/body/div[3]/div/div[@class='yui3-u-1-6']/div[@id='cl-menucate']/ul" ''result
        Dim docHtmlNode As HtmlNode = doc.DocumentNode.SelectSingleNode(strXPath)

        Dim strContent As String = String.Empty
        If Not IsNothing(docHtmlNode) Then
            hdc.LoadHtml(docHtmlNode.InnerHtml)

            Dim LiHtmlNodeCollection As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./li")

            Dim i As Integer = 1
            For Each HtmlNode As HtmlNode In LiHtmlNodeCollection
                Dim nodeTitle As HtmlNode
                Dim strSpan As String = String.Empty

                If intFirstLevel <= 1 Then
                    nodeTitle = HtmlNode.SelectSingleNode("./h3")
                    strSpan = HtmlNode.SelectSingleNode("./h3/span").Attributes("class").Value
                Else
                    ''intLevel 2 & 3
                    If i = 1 Then
                        nodeTitle = HtmlNode.SelectSingleNode("./h1")
                        strSpan = HtmlNode.SelectSingleNode("./h1/span").Attributes("class").Value
                    Else
                        nodeTitle = HtmlNode.SelectSingleNode("./h3")
                        strSpan = HtmlNode.SelectSingleNode("./h3/span").Attributes("class").Value
                    End If
                End If

                ''Third Level=2
                Dim strName As String = String.Empty
                If Not IsNothing(nodeTitle) Then strName = nodeTitle.InnerText.Trim
                Dim strWebUrl As String = String.Empty
                If Not IsNothing(nodeTitle) Then strWebUrl = ComUtility.StringPlus.getCssAhref(nodeTitle.InnerHtml.ToString, strName)

                BLL.insCategory(intFirstID, strName, strLinkUrl & strWebUrl, intFirstLevel + 1, 1, 0, ComConfig.clConfig.SearchEngine.Yahoo, 1)
                Dim intParentID As Integer = BLL.getCategoryID(intFirstID, strName, strLinkUrl & strWebUrl, intFirstLevel + 1, ComConfig.clConfig.SearchEngine.Yahoo, 1)
                BLL.updIsParserContent(intFirstID)

                ''Fourth Level=3
                Dim UlHtmlNodeCollection As HtmlNodeCollection = HtmlNode.SelectNodes("./ul")
                Dim SecondLiHtmlNodeCollection As HtmlNodeCollection

                If Not IsNothing(UlHtmlNodeCollection) Then
                    For Each SecondUlHtmlNode As HtmlNode In UlHtmlNodeCollection
                        SecondLiHtmlNodeCollection = SecondUlHtmlNode.SelectNodes("./li")
                        For Each SecondLiHtmlNode As HtmlNode In SecondLiHtmlNodeCollection
                            Dim strSecondName As String = String.Empty
                            If Not IsNothing(SecondLiHtmlNode) Then strSecondName = SecondLiHtmlNode.InnerText.ToString
                            Dim strSecondWebUrl As String = String.Empty
                            If Not IsNothing(SecondLiHtmlNode) Then strSecondWebUrl = ComUtility.StringPlus.getCssAhref(SecondLiHtmlNode.InnerHtml.ToString, strSecondName)

                            BLL.insCategory(intParentID, strSecondName, strLinkUrl & strSecondWebUrl, intFirstLevel + 2, 0, 0, ComConfig.clConfig.SearchEngine.Yahoo, 1)
                        Next
                        If SecondLiHtmlNodeCollection.Count > 0 Then
                            BLL.updIsParserContent(intParentID)
                        End If
                    Next
                End If

                i += 1
            Next

        End If
    End Sub

    Private Sub CrawYahooCategoryLevel3()
        Dim ds As New DataSet
        ds = BLL.getCategoryLevel3
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                Dim intID As Integer = Integer.Parse(ds.Tables(0).Rows(i).Item("id").ToString)
                Dim strWebUrl As String = ds.Tables(0).Rows(i).Item("url").ToString

                CrawYahooProductLink(intID, strWebUrl)
            Next
        End If
        ds.Dispose()
    End Sub

    Private Sub CrawYahooProductLink(ByVal intID As Integer, ByVal strSiteMapUrl As String)
        '指定來源網頁
        Dim WebClient As New WebClient()
        Dim strUserAgent As String = ComConfig.clConfig.UserAgent(0, True).ToUpper
        WebClient.Headers(HttpRequestHeader.UserAgent) = strUserAgent

        ''將網頁來源資料暫存到記憶體內
        Dim ms As New MemoryStream(WebClient.DownloadData(strSiteMapUrl))

        ''使用預設編碼讀入 HTML 
        Dim doc As New HtmlDocument
        doc.Load(ms, Encoding.UTF8)

        Dim hdc As New HtmlDocument

        'XPath 來解讀它
        Dim strXPath As String = "/html/body/div[3]/div[@id='bd']/div[2]/div[@class='yui3-g']/div[2]/div[@class='hero-content']" ''result
        Dim docHtmlNode As HtmlNode = doc.DocumentNode.SelectSingleNode(strXPath)

        Dim strContent As String = String.Empty
        If Not IsNothing(docHtmlNode) Then
            hdc.LoadHtml(docHtmlNode.InnerHtml)

            ''館長推薦
            Dim divHtmlNodeCollection1 As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./div[@id='cl-recproduct']/div")
            If Not IsNothing(divHtmlNodeCollection1) Then
                For Each HtmlNode As HtmlNode In divHtmlNodeCollection1
                    Dim nodeProduct As HtmlNode = HtmlNode.SelectSingleNode("./div")
                    Dim strName As String = String.Empty
                    If Not IsNothing(nodeProduct) Then strName = nodeProduct.SelectSingleNode("./ul/li[@class='yui3-u-1 name']").InnerText.Trim
                    Dim strWebUrl As String = String.Empty
                    If Not IsNothing(nodeProduct) Then strWebUrl = ComUtility.StringPlus.getCssAhref(nodeProduct.SelectSingleNode("./ul/li[@class='yui3-u-1 name']").InnerHtml, strName)

                    BLL.insCategory(intID, strName, strLinkUrl & strWebUrl, 4, 0, 0, ComConfig.clConfig.SearchEngine.Yahoo, 1)
                Next
            End If

            ''黃金促銷區
            Dim i As Integer = 1
            Dim divHtmlNodeCollection2 As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./div[@id='cl-gproduct']/div")
            If Not IsNothing(divHtmlNodeCollection1) Then
                For Each HtmlNode As HtmlNode In divHtmlNodeCollection2
                    If i >= 3 Then
                        Dim nodeProduct As HtmlNode = HtmlNode.SelectSingleNode("./div")
                        Dim strName As String = String.Empty
                        If Not IsNothing(nodeProduct) Then strName = nodeProduct.SelectSingleNode("./ul/li[@class='yui3-u-1 name']").InnerText.Trim
                        Dim strWebUrl As String = String.Empty
                        If Not IsNothing(nodeProduct) Then strWebUrl = ComUtility.StringPlus.getCssAhref(nodeProduct.SelectSingleNode("./ul/li[@class='yui3-u-1 name']").InnerHtml, strName)

                        BLL.insCategory(intID, strName, strLinkUrl & strWebUrl, 4, 0, 0, ComConfig.clConfig.SearchEngine.Yahoo, 1)
                    End If
                    i += 1
                Next
            End If

            ''推薦
            i = 1
            Dim divHtmlNodeCollection3 As HtmlNodeCollection = hdc.DocumentNode.SelectNodes("./div[@id='catit-recmd']/div")
            If Not IsNothing(divHtmlNodeCollection3) Then
                Dim UlHtmlNodeCollection As HtmlNodeCollection
                For Each HtmlNode As HtmlNode In divHtmlNodeCollection3
                    If i = 3 Then
                        Dim nodeList As HtmlNode = HtmlNode.SelectSingleNode(".//ul[@class='page-list']")
                        UlHtmlNodeCollection = nodeList.SelectNodes("./li/ul[@class='recom-list']")
                        For Each LiHtmlNode As HtmlNode In UlHtmlNodeCollection
                            Dim nodeProduct As HtmlNode = LiHtmlNode.SelectSingleNode("./li/div[@class='desc']")
                            Dim strName As String = String.Empty
                            If Not IsNothing(nodeProduct) Then strName = nodeProduct.InnerText.Trim
                            Dim strWebUrl As String = String.Empty
                            If Not IsNothing(nodeProduct) Then strWebUrl = ComUtility.StringPlus.getCssAhrefLevel4(nodeProduct.InnerHtml, "hpp")

                            BLL.insCategory(intID, strName, strLinkUrl & strWebUrl, 4, 0, 0, ComConfig.clConfig.SearchEngine.Yahoo, 1)
                        Next
                    End If
                    i += 1
                Next
            End If

            BLL.updIsParser(intID)

            Dim intCnt As Integer = 0
            If Not IsNothing(divHtmlNodeCollection1) AndAlso divHtmlNodeCollection1.Count Then intCnt += 1
            If Not IsNothing(divHtmlNodeCollection2) AndAlso divHtmlNodeCollection2.Count Then intCnt += 1
            If Not IsNothing(divHtmlNodeCollection3) AndAlso divHtmlNodeCollection3.Count Then intCnt += 1
            If intCnt > 0 Then BLL.updIsParserContent(intID)
        End If

    End Sub

End Module
