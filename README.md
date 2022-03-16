# vbaRequests
![GitHub top language](https://img.shields.io/github/languages/top/tankalxat34/vbaRequests)
![skill](https://img.shields.io/badge/Microsoft%20Excel%20VBA-107C41?logo=microsoft&logoColor=white)

The simple module for making requests to websites

# Installation
[![test](https://img.shields.io/badge/-download-brightgreen?style=for-the-badge)](https://github.com/tankalxat34/vbaRequests/raw/main/vbaRequests.bas)

1. Download **[this file](https://github.com/tankalxat34/vbaRequests/raw/main/vbaRequests.bas)** or click to "DOWNLOAD" button.
2. Open your Microsoft Excel Book and in window `"Project - VBAProject"` click on free place and than *"Import file"*.
3. Choice the downloaded file.
4. Enjoy.

# Example
## Get your own IP
This code will show you your IP address
```vb
Sub helloworld()
    Dim userIP As String
    userIP = vbaRequests.request("https://ifconfig.me/ip")
    MsgBox userIP
End Sub
```

## IP verification
With this code you can check IP any user who are opening your Excel Book and if the user does not exists in IP-list - application will be closed.
```vb
Private Sub Workbook_Open()
    Dim userIP As String
    Dim listIP As String
    
    userIP = vbaRequests.request("https://ifconfig.me/ip")
    listIP = vbaRequests.request("https://raw.githubusercontent.com/tankalxat34/vba-ip-validation/main/ip_list.txt")
    
    If InStr(listIP, userIP) Then
        Exit Sub
    Else
        MsgBox "Your IP address is not confirmed in the system! The application will be closed!", vbCritical
        Application.Quit
    End If
End Sub
```
