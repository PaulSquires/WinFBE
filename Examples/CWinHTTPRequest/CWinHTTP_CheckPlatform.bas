' ========================================================================================
' The following example checks if the current platform is supported.
' ========================================================================================

'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWinHttpRequest.inc"
using Afx

IF AfxWinHttpCheckPlatform THEN
   PRINT "This platform is supported by WinHTTP."
ELSE
   PRINT "This platform is NOT supported by WinHTTP."
END IF

PRINT
PRINT "Press any key..."
SLEEP
