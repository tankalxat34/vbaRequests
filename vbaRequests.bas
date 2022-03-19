Attribute VB_Name = "vbaRequests"
Option Explicit

'                       Module vbaRequests
' The simple module for making requests to websites. Here is available support
'          of GET, POST, DELETE and other methods of requests.
'
'              tankalxat34 (Alexander Podstrechnyy)
'            https://github.com/tankalxat34/vbaRequests
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

Public Function createHeaders() As Object
    ' create the default dictionary with headers
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    headers.Add "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36 Edg/99.0.1150.39"
    headers.Add "Cache-Control", "max-age=0"
    headers.Add "Accept-Encoding", "deflate"
    headers.Add "Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4"
    
    Set createHeaders = headers
End Function


Public Function request(ByVal sURL As String, headersDictionary As Object, Optional ByVal username As String, Optional ByVal password As String, Optional ByVal typeRequest As String = "GET") As String
    ' Parameters:
    '|     Parameter     |            Type             |                                  Description                                  |
    '|-------------------|-----------------------------|-------------------------------------------------------------------------------|
    '| sURL              | String                      | The string URL of web-site                                                    |
    '| headersDictionary | Object Scripting.Dictionary | A dictionary containing headers for making a successful request to a website. |
    '|                   |                             | You can set the headers yourself, or use the "createHeaders"                  |
    '|                   |                             | function to automatically apply default headers to your request               |
    '| username          | String                      | String containing your username for login in website                          |
    '| password          | String                      | String containing your password or token for login in website                 |

    Dim oXMLHTTP
    Dim element As Variant
    
    On Error Resume Next
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    With oXMLHTTP
        .Open typeRequest, sURL, False
        
        ' set up all headers from headersDictionary
        For Each element In headersDictionary.Keys
            .SetRequestHeader element, headersDictionary.Item(element)
        Next
        
        ' check to available to set up username and password
        If username <> "" And password <> "" Then
            .SetRequestHeader "php-auth-user", username
            .SetRequestHeader "php-auth-pw", password
        End If
        
        ' send the request
        .send
        
        ' return request
        request = .responseText
    End With
    Set oXMLHTTP = Nothing
End Function
