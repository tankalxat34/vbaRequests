# vbaRequests
![GitHub top language](https://img.shields.io/github/languages/top/tankalxat34/vbaRequests)
![skill](https://img.shields.io/badge/Microsoft%20Excel%20VBA-107C41?logo=microsoft&logoColor=white)
![GitHub](https://img.shields.io/github/license/tankalxat34/vbaRequests?logo=github&logoColor=white)

The simple module for making requests to websites

Author: **[tankalxat34](https://github.com/tankalxat34)**

# Installation
[![test](https://img.shields.io/badge/-download-brightgreen?style=for-the-badge)](https://github.com/tankalxat34/vbaRequests/raw/main/vbaRequests.bas)

1. Download **[this file](https://github.com/tankalxat34/vbaRequests/raw/main/vbaRequests.bas)** or click to "DOWNLOAD" button.
2. Open your Microsoft Excel Book and in window `"Project - VBAProject"` click on free place and than *"Import file"*.
3. Choice the downloaded file.
4. Enjoy.

# Information about "request" public function

|     Parameter     |            Type             |                                  Description                                  |
|-------------------|-----------------------------|-------------------------------------------------------------------------------|
| sURL              | String                      | The string URL of web-site                                                    |
| headersDictionary | Object Scripting.Dictionary | A dictionary containing headers for making a successful request to a website. You can set the headers yourself, or use the "createHeaders" function to automatically apply default headers to your request                                             |
| username          | String                      | String containing your username for login in website                          |
| password          | String                      | String containing your password or token for login in website                 |
| typeRequest       | String                      | String of type for request: "GET", "POST", "PUT" and other types              |

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

## Response from GitHub
```vb
Sub githubResponse()
    Debug.Print vbaRequests.request("https://api.github.com/users/tankalxat34", _
            vbaRequests.createHeaders(), _
            "tankalxat34", "YOUR_TOKEN", _
            "GET")
End Sub
```
