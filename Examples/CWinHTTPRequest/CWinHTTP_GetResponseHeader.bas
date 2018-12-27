' ========================================================================================
' The following example shows how to open an HTTP connection, send an HTTP request, and
' returns the "Date" header from the response.
' ========================================================================================

#include "windows.bi"
#include "Afx/CWinHttpRequest.inc"
using Afx

' // Create an instance of the CWinHttp class
DIM pWHttp AS CWinHttpRequest

' // Open an HTTP connection to an HTTP resource
pWHttp.Open "GET", "http://microsoft.com"

' // Send an HTTP request to the HTTP server
pWHttp.Send

' // Wait for response with a timeout of 5 seconds
DIM iSucceeded AS LONG = pWHttp.WaitForResponse(5)

' // Get the response headers
DIM cbsResponseHeader AS CBSTR = pWHttp.GetResponseHeader("Date")
PRINT cbsResponseHeader

PRINT
PRINT "Press any key..."
SLEEP
