Imports System.Net
Imports System.IO
Imports System.Text
Imports HtmlAgilityPack

Module Module1

    Sub Main()
        Process()
    End Sub

    Public Sub Process()
        Dim strProductUrl As String = "https://tw.buy.yahoo.com/gdsale/gdsale.asp?gdid=5465709&hpp=p4a2"
        CrawlYahooProduct(strProductUrl)
    End Sub

    Private Sub CrawlYahooProduct(ByVal strSiteMapUrl As String)
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

        CrawlYahooProduct(doc, hdc)
    End Sub

    Private Function CrawlYahooProduct(ByVal doc As HtmlDocument, ByVal hdc As HtmlDocument) As String
        'XPath 來解讀它
        Dim strXPath As String = "/html/body/div[3]/div//*[@id='bd']/div[@class='yui3-u-5-6']/div[@class='yui3-g']/div[2]" ''result 商品主檔
        Dim docHtmlNode As HtmlNode = doc.DocumentNode.SelectSingleNode(strXPath)

        Dim strContent As String = String.Empty
        If Not IsNothing(docHtmlNode) Then
            hdc.LoadHtml(docHtmlNode.InnerHtml)

            Dim strPid As String = String.Empty
            Dim strMarketId As String = String.Empty
            Dim divProd1HtmlNode As HtmlNode = hdc.DocumentNode.SelectSingleNode("./div/div[1]")
            If Not IsNothing(divProd1HtmlNode) Then
                Dim divProdIdCollection As HtmlNode = divProd1HtmlNode.SelectSingleNode("./div[@class='itemgdid']")
                If Not IsNothing(divProdIdCollection) Then
                    Dim PidHtmlNode As HtmlNode = divProdIdCollection.SelectSingleNode("./div[@class='erpgdid']/span[@class='number']")
                    If Not IsNothing(PidHtmlNode) Then strPid = PidHtmlNode.InnerText.Trim ''商品編號
                    Dim MidHtmlNode As HtmlNode = divProdIdCollection.SelectSingleNode("./div[@class='gdid']/span[@class='number']")
                    If Not IsNothing(MidHtmlNode) Then strMarketId = MidHtmlNode.InnerText.Trim ''賣場編號
                End If
            End If

            Dim strPromoteTxt As String = String.Empty
            Dim strPname As String = String.Empty
            Dim strMarketPrice As String = String.Empty
            Dim strWebPrice As String = String.Empty
            Dim strInstallPeriod As String = String.Empty
            Dim strInstallPrice As String = String.Empty
            Dim strPfunc As String = String.Empty
            Dim divProd2HtmlNode As HtmlNode = hdc.DocumentNode.SelectSingleNode("./div/div[2]")
            If Not IsNothing(divProd2HtmlNode) Then
                Dim PromoteHtmlNode As HtmlNode = divProd2HtmlNode.SelectSingleNode("./div[@class='subtitle']")
                If Not IsNothing(PromoteHtmlNode) Then strPromoteTxt = PromoteHtmlNode.InnerText.ToString ''促銷文
                Dim PnameHtmlNode As HtmlNode = divProd2HtmlNode.SelectSingleNode("./div[@class='title']")
                If Not IsNothing(PnameHtmlNode) Then strPname = PnameHtmlNode.InnerText.ToString.Trim ''品名
                Dim MarketPriceHtmlNode As HtmlNode = divProd2HtmlNode.SelectSingleNode("./div[@class='suggest']/span[@class='price']")
                If Not IsNothing(MarketPriceHtmlNode) Then strMarketPrice = MarketPriceHtmlNode.InnerText.Trim.Replace(",", "").Replace("$", "") ''建議售價
                Dim WebPriceHtmlNode As HtmlNode = divProd2HtmlNode.SelectSingleNode("./div[@class='priceinfo']/span[@class='price']")
                If Not IsNothing(WebPriceHtmlNode) Then strWebPrice = WebPriceHtmlNode.InnerText.Trim.Replace(",", "").Replace("$", "") ''網路售價
                Dim InstallHtmlNode As HtmlNode = divProd2HtmlNode.SelectSingleNode("./div[@class='rate']")
                If Not IsNothing(InstallHtmlNode) Then
                    strInstallPeriod = InstallHtmlNode.SelectSingleNode("./div[@class='ratemax']/span[@class='num']").InnerText.ToString.Trim.Replace(",", "").Replace("$", "") ''最大分期數
                    strInstallPrice = InstallHtmlNode.SelectSingleNode("./div[@class='ratemax']/span[3]").InnerText.ToString.Trim.Replace(",", "").Replace("$", "")             ''分期金額
                End If
                Dim FuncHtmlNode As HtmlNode = divProd2HtmlNode.SelectSingleNode("./ul")
                If Not IsNothing(FuncHtmlNode) Then strPfunc = FuncHtmlNode.OuterHtml.ToString ''商品特色
            End If

            Dim strPdesc As String = String.Empty
            Dim strPspec As String = String.Empty
            Dim strValidDesc As String = String.Empty
            Dim strCarriageDesc As String = String.Empty
            Dim strPurchaseInfo As String = String.Empty
            Dim divProd3HtmlNode As HtmlNode = hdc.DocumentNode.SelectSingleNode("./div[@id='cl-gdintro']")
            If Not IsNothing(divProd3HtmlNode) Then
                Dim pdescHtmlNode As HtmlNode = divProd3HtmlNode.SelectSingleNode("./div[1]//div[@class='content']") ''特別推薦
                If Not IsNothing(pdescHtmlNode) Then strPdesc = pdescHtmlNode.InnerHtml
                Dim pspecHtmlNode As HtmlNode = divProd3HtmlNode.SelectSingleNode("./div[2]//div[@class='content']") ''商品規格
                If Not IsNothing(pspecHtmlNode) Then strPspec = pspecHtmlNode.InnerHtml
                Dim vdescHtmlNode As HtmlNode = divProd3HtmlNode.SelectSingleNode("./div[3]//div[@class='content']") ''商品保證
                If Not IsNothing(vdescHtmlNode) Then strValidDesc = vdescHtmlNode.InnerHtml
                Dim cdescHtmlNode As HtmlNode = divProd3HtmlNode.SelectSingleNode("./div[4]//div[@class='content']") ''商品運送
                If Not IsNothing(cdescHtmlNode) Then strCarriageDesc = cdescHtmlNode.InnerHtml
                Dim pInfoHtmlNode As HtmlNode = divProd3HtmlNode.SelectSingleNode("./div[5]//div[@class='content']") ''購買說明
                If Not IsNothing(pInfoHtmlNode) Then strPurchaseInfo = pInfoHtmlNode.InnerHtml
            End If


            ''Insert into TableProduct
        End If

        Return strContent
    End Function

End Module
