Imports System.Text.RegularExpressions

Public Class StringPlus

    Public Shared Function ExtractDomainFromURL(ByVal strURL As String) As String
        If strURL.Length > 0 Then
            If strURL.IndexOf("/") <> -1 Then
                Return strURL.Substring(0, strURL.IndexOf("/"))
            Else
                Return strURL
            End If
        Else
            Return ""
        End If
    End Function

    Public Shared Function ConvertBig5(ByVal strUtf As String) As String         Dim utf81 As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")         Dim big51 As System.Text.Encoding = System.Text.Encoding.GetEncoding("big5")          Dim strUtf81 As Byte() = utf81.GetBytes(strUtf.Trim())         Dim strBig51 As Byte() = System.Text.Encoding.Convert(utf81, big51, strUtf81)
        Dim big5Chars1 As Char() = New Char(big51.GetCharCount(strBig51, 0, strBig51.Length) - 1) {}         big51.GetChars(strBig51, 0, strBig51.Length, big5Chars1, 0)         Dim tempString1 As New String(big5Chars1)         Return tempString1     End Function

    Public Shared Function getCssAhref(ByVal strValue As String, ByVal strTitle As String) As String
        Dim strLink As String = String.Empty
        Dim intStart As Integer = strValue.IndexOf("href")
        Dim intEnd As Integer = strValue.IndexOf(strTitle)
        If intStart <> -1 And intEnd <> -1 Then
            strLink = strValue.Substring(intStart + 6, intEnd - 2 - (intStart + 6))
        End If
        Return strLink
    End Function

    Public Shared Function getCssAhrefLevel4(ByVal strValue As String, ByVal strTitle As String) As String
        Dim strLink As String = String.Empty
        Dim intStart As Integer = strValue.IndexOf("href")
        Dim intEnd As Integer = strValue.IndexOf(strTitle)
        If intStart <> -1 And intEnd <> -1 Then
            strLink = strValue.Substring(intStart + 6, intEnd - 2 - (intStart + 6))
        End If
        Return strLink
    End Function

End Class
