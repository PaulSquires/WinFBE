'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWinHttpRequest.inc"
using Afx

' // Create an instance of the CWinHttp class
DIM pWHttp AS CWinHttpRequest

' // Open an HTTP connection to an HTTP resource
pWHttp.Open "GET", "http://microsoft.com"

' // Specify the user agent
pWHttp.SetOption(WinHttpRequestOption_UserAgentString, "A WinHttpRequest Example Program")

' // Send an HTTP request to the HTTP server
pWHttp.Send

IF pWHttp.GetLastResult = S_OK THEN
   ' // Get user agent string.
   DIM cvText AS CVAR = pWHttp.GetOption(WinHttpRequestOption_UserAgentString)
   PRINT cvText.ToStr
   ' // We can also use:
   ' PRINT pWHttp.GetOption(WinHttpRequestOption_UserAgentString).ToStr
   ' // Get URL
   PRINT pWHttp.GetOption(WinHttpRequestOption_URL).ToStr
   ' // Get URL Code Page.
   PRINT pWHttp.GetOption(WinHttpRequestOption_URLCodePage).ToStr
   ' // Convert percent symbols to escape sequences.
   PRINT pWHttp.GetOption(WinHttpRequestOption_EscapePercentInURL).ToStr
END IF

PRINT
PRINT "Press any key..."
SLEEP
