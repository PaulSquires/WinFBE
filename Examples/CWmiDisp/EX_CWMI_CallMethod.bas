'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWmiDisp.inc"
using Afx

' // Connect to WMI using a moniker
DIM pServices AS CWmiServices = $"winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv"
IF pServices.ServicesPtr = NULL THEN END
 
' // Assign the WMI services object pointer to CDispInvoke
' // CWmiServices.ServicesObj returns an AddRefed pointer, whereas CWmiServices.ServicesPtr not.
DIM pDispServices AS CDispInvoke = CDispInvoke(pServices.ServicesObj)

' Parameters of the GetStringValue method:
' %HKEY_LOCAL_MACHINE ("2147483650") - The value must be specified as an string and in decimal, not hexadecimal.
' vDefKey = [IN]  "2147483650"
' vPath   = [IN]  "Software\Microsoft\Windows NT\CurrentVersion"
' vValue  = [OUT] "ProductName"

DIM cbsValue AS CBSTR
DIM cvRes AS CVAR = pDispServices.Invoke("GetStringValue", "2147483650", _
    $"Software\Microsoft\Windows NT\CurrentVersion", "ProductName", CVAR(cbsValue.vptr, VT_BSTR))
PRINT cbsValue

PRINT
PRINT "Press any key..."
SLEEP
