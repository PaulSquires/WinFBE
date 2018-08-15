' ========================================================================================
' The following example shows how to open an HTTP connection, send an HTTP request, and
' returns the response body as a stream of bytes.
' ========================================================================================

'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWinHttpRequest.inc"
using Afx

' // Create an instance of the CWinHttp class
DIM pWHttp AS CWinHttpRequest

' // Open an HTTP connection to an HTTP resource
pWHttp.Open "GET", "http://www.microsoft.com/library/homepage/images/ms-banner.gif"

' // Send an HTTP request to the HTTP server
pWHttp.Send

' // Wait for response with a timeout of 5 seconds
DIM iSucceeded AS LONG = pWHttp.WaitForResponse(5)

' // Get the response body
DIM strResponseBody AS STRING = pWHttp.GetResponseStream

' // Save the buffer into a file
IF LEN(strResponseBody) THEN
   DIM fn AS LONG = FREEFILE
   OPEN "ms-banner.gif" FOR OUTPUT AS #fn
   PUT #fn, 1, strResponseBody
   CLOSE #fn
   PRINT "File saved"
ELSE
   PRINT "Failure"
END IF

PRINT
PRINT "Press any key..."
SLEEP
