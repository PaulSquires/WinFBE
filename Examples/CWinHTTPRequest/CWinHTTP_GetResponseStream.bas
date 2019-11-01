' ========================================================================================
' The following example shows how to open an HTTP connection, sends an HTTP request to
' load an image, returns the response body as a stream of bytes and saves it to a file.
' ========================================================================================

'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWinHttpRequest.inc"
#include once "Afx/CStream.inc"
using Afx

' // Create an instance of the CWinHttp class
DIM pWHttp AS CWinHttpRequest

' // Open an HTTP connection to an HTTP resource
pWHttp.Open "GET", "https://i.ytimg.com/vi/nPUi1XNiRzQ/maxresdefault.jpg"

' // Send an HTTP request to the HTTP server
pWHttp.Send

' // Wait for response with a timeout of 5 seconds
DIM iSucceeded AS LONG = pWHttp.WaitForResponse(5)

SCOPE
   DIM st AS STRING = pWHttp.GetResponseBody
  ' // Open a file stream
   DIM pFileStream AS CFileStream
   IF pFileStream.Open("image.jpg", STGM_CREATE OR STGM_WRITE) = S_OK then
      pFileStream.WriteTextA(st)
   END IF
END SCOPE

PRINT
PRINT "Press any key..."
SLEEP
