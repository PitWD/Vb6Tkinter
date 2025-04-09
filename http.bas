Attribute VB_Name = "http"
Option Explicit
' Connect to the network to get information, requires reference to "Microsoft WinHTTP Services, version 5.1"

' Parameters: sUrl - URL
' Parameters: postData - post content (if POST)
' Parameters: method - method, either "POST" or "GET"
' Parameters: cookies - cookies to be included in the post or get
' Returns: content returned by the post or get
Public Function HttpGetResponse(sUrl As String, Optional ByVal postData As String = "", Optional ByVal method As String = "GET", Optional ByVal cookies As String = "") As String
    Dim request As winhttp.WinHttpRequest
    If Len(Trim(cookies)) = 0 Then
        cookies = "a:x," ' If cookies are empty, set a random cookie to avoid errors
    End If
    
    On Error Resume Next
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With request
        .Open UCase(method), sUrl, True 'True: receive data synchronously
        .SetTimeouts 30000, 30000, 30000, 30000 ' Set timeout to 30 seconds
        .Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300 ' Very important (ignore errors)
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        .SetRequestHeader "Accept", "text/html, application/xhtml+xml, */*"
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .SetRequestHeader "Cookie", cookies
        .SetRequestHeader "Content-Length", Len(postData)
        .Send postData ' Start sending
        
        .WaitForResponse ' Wait for request
        'MsgBox WinHttp.Status ' Request status
        HttpGetResponse = .ResponseText ' Get the returned text (or other content)
    End With
    Set request = Nothing
End Function


