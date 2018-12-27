'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWmiDisp.inc"
using Afx

' // Connect to WMI using a moniker
DIM pServices AS CWmiServices = $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
IF pServices.ServicesPtr = NULL THEN END
 
' // Get an instance of the printer "OKI B410" --> change me
DIM hr AS HRESULT = pServices.Get("Win32_Printer.DeviceID='OKI B410'")
IF hr <> S_OK THEN PRINT AfxWmiGetErrorCodeText(hr) : SLEEP : END

' // Number of properties
PRINT "Number of properties: ", pServices.PropsCount
PRINT

' // Display some properties
PRINT "Port name: "; pServices.PropValue("PortName")
PRINT "Attributes: "; pServices.PropValue("Attributes")
PRINT "Paper sizes supported: "; pServices.PropValue("PaperSizesSupported")

PRINT
PRINT "Press any key..."
SLEEP
