'#CONSOLE ON
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
   PRINT pWHttp.GetOption(WinHttpRequestOption_UserAgentString)
   ' // Get URL
   PRINT pWHttp.GetOption(WinHttpRequestOption_URL)
   ' // Get URL Code Page.
   PRINT pWHttp.GetOption(WinHttpRequestOption_URLCodePage)
   ' // Convert percent symbols to escape sequences.
   PRINT pWHttp.GetOption(WinHttpRequestOption_EscapePercentInURL)
END IF

PRINT
PRINT "Press any key..."
SLEEP
