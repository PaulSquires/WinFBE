'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWmiDisp.inc"
using Afx

' // Connect to WMI using a moniker
' // Note: $ is used to avoid the pedantic warning of the compiler about escape characters
DIM pServices AS CWmiServices = $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
IF pServices.ServicesPtr = NULL THEN END

' // Execute a query
DIM hr AS HRESULT = pServices.ExecQuery("SELECT Caption, SerialNumber FROM Win32_BIOS")
IF hr <> S_OK THEN PRINT AfxWmiGetErrorCodeText(hr) : SLEEP : END

' // Get the number of objects retrieved
DIM nCount AS LONG = pServices.ObjectsCount
print "Number of objects: ", nCount
IF nCount = 0 THEN PRINT "No objects found" : SLEEP : END

' // Get a collection of named properties
IF pServices.GetNamedProperties <> S_OK THEN PRINT "Failed to get the named properties" : SLEEP : END 

' // Retrieve the value of the properties
'DIM cv AS CVAR = pServices.PropValue("Caption")
'PRINT cv.ToStr
PRINT pServices.PropValue("Caption")
PRINT pServices.PropValue("SerialNumber")

PRINT
PRINT "Press any key..."
SLEEP
