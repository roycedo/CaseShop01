Public Class clConfig
    Public Enum SearchEngine As Integer
        noSigned = 0
        Engadget = 1
        Area  =2
        ads = 3
        Yahoo = 4
        Google = 5
        iChannels = 6
        likeNews = 7
        PCHome = 8
    End Enum

    ''' <summary>
    ''' 隨機取得UserAgent
    ''' </summary>
    ''' <param name="intIndex">不隨機時請填入正確的UserAgent數字</param>
    ''' <param name="isRandom">是否隨機</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UserAgent(ByVal intIndex As Integer, Optional ByVal isRandom As Boolean = False) As String
        Dim strUserAgentList() As String = { _
            "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2049.0 Safari/537.36", _
            "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1985.67 Safari/537.36", _
            "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1985.67 Safari/537.36", _
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1944.0 Safari/537.36", _
            "Mozilla/5.0 (Windows NT 5.1; rv:31.0) Gecko/20100101 Firefox/31.0", _
            "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:29.0) Gecko/20120101 Firefox/29.0", _
            "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:25.0) Gecko/20100101 Firefox/29.0", _
            "Mozilla/5.0 (X11; OpenBSD amd64; rv:28.0) Gecko/20100101 Firefox/28.0", _
            "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0)", _
            "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/5.0; TheWorld)", _
            "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)", _
            "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)", _
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_6_8) AppleWebKit/537.13+ (KHTML, like Gecko) Version/5.1.7 Safari/534.57.2", _
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/534.55.3 (KHTML, like Gecko) Version/5.1.3 Safari/534.53.10", _
            "Opera/9.80 (Windows NT 6.0) Presto/2.12.388 Version/12.14", _
            "Opera/9.80 (Windows NT 6.1; WOW64; U; pt) Presto/2.10.229 Version/11.62", _
            "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) Presto/2.9.168 Version/11.52"
            }

        If isRandom Then
            Dim rng As Random = New Random
            intIndex = rng.Next(0, strUserAgentList.Length - 1)
        End If

        Return strUserAgentList(intIndex)
    End Function

    ''' <summary>
    ''' 隨機取得UserAgent
    ''' </summary>
    ''' <param name="intIndex">不隨機時請填入正確的UserAgent數字</param>
    ''' <param name="isRandom">是否隨機</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function UserAgent_Chester(ByVal intIndex As Integer, Optional ByVal isRandom As Boolean = False) As String
        Dim strUserAgentList() As String = { _
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.7; rv:7.0.1) Gecko/20100101 Firefox/7.0.1", _
            "Mozilla/5.0 (compatible; MSIE 9.0; Windows Phone OS 7.5; Trident/5.0; IEMobile/9.0)", _
            "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.1.9) Gecko/20100508 SeaMonkey/2.0.4", _
            "Opera/9.80 (S60; SymbOS; Opera Mobi/SYB-1107071606; U; en) Presto/2.8.149 Version/11.10", _
            "Mozilla/5.0 (Windows; U; MSIE 7.0; Windows NT 6.0; en-US)", _
            "Mozilla/5.0 (Linux; U; Android 4.0.3; ko-kr; LG-L160L Build/IML74K) AppleWebkit/534.30 (KHTML, like Gecko) Version/4.0 Mobile Safari/534.30", _
            "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_7; da-dk) AppleWebKit/533.21.1 (KHTML, like Gecko) Version/5.0.5 Safari/533.21.1", _
            "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:23.0) Gecko/20130406 Firefox/23.0", _
            "Opera/12.02 (Android 4.1; Linux; Opera Mobi/ADR-1111101157; U; en-US) Presto/2.9.201 Version/12.02", _
            "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:18.0) Gecko/20100101 Firefox/18.0", _
            "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533+ (KHTML, like Gecko) Element Browser 5.0", _
            "Mozilla/5.0 (BlackBerry; U; BlackBerry 9900; en) AppleWebKit/534.11+ (KHTML, like Gecko) Version/7.1.0.346 Mobile Safari/534.11+", _
            "Opera/9.80 (Windows NT 6.0) Presto/2.12.388 Version/12.14", _
            "Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25", _
            "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1468.0 Safari/537.36", _
            "Mozilla/5.0 (Linux; U; Android 2.3.5; en-us; HTC Vision Build/GRI40) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1", _
            "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0)", _
            "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/5.0; TheWorld)", _
            "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)", _
            "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)"
            }

        If isRandom Then
            Dim rng As Random = New Random
            intIndex = rng.Next(0, strUserAgentList.Length - 1)
        End If

        Return strUserAgentList(intIndex)
    End Function

End Class
