Attribute VB_Name = "vbaRequests"
Option Explicit

'                       Module vbaRequests
' The simple module for making requests to websites. Here is available support 
'         of GET, PUT, DELETE and other methods of requests.
' 
'             Author: tankalxat34 (Alexander Podstrechnyy)
'               https://github.com/tankalxat34/vbaRequests
'
' License:
' MIT License
'
' Copyright (c) 2022 Alexander Podstrechnyy
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


Public Function request(ByVal sURL As String, Optional ByVal typeRequest As String = "GET") As String
    ' Parameters:
    ' sURL - String - url to website
    ' typeRequest - Optional - String - type of request: GET, POST, OPTIONS and other.
    Dim oXMLHTTP
    On Error Resume Next
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    With oXMLHTTP
        .Open typeRequest, sURL, False
        .SetRequestHeader "Cache-Control", "max-age=0"
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.41 Safari/537.36 OPR/35.0.2066.10 (Edition beta)"
        .SetRequestHeader "Accept-Encoding", "deflate"
        .SetRequestHeader "Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4"
        .send
        request = .responseText
    End With
    Set oXMLHTTP = Nothing
End Function
