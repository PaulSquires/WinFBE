'#CONSOLE ON
#include "windows.bi"
#include "Afx/CWmiDisp.inc"
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
print "Count: ", nCount
IF nCount = 0 THEN PRINT "No objects found" : SLEEP : END

' // Enumerate the objects using the standard IEnumVARIANT enumerator (NextObject method)
' // and retrieve the properties using the CDispInvoke class.
DIM pDispServ AS CDispInvoke = pServices.NextObject
IF pDispServ.DispPtr THEN
   PRINT "Caption: "; pDispServ.Get("Caption")
   PRINT "Serial number: "; pDispServ.Get("SerialNumber")
END IF

PRINT
PRINT "Press any key..."
SLEEP
