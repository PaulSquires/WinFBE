'#CONSOLE ON
#include once "windows.bi"
#include once "Afx/CWmiDisp.inc"
using Afx

' // Connect to WMI using a moniker
'DIM pServices AS CWmiServices = $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
' Use the constructor for server connection, just for trying...
DIM pServices AS CWmiServices = CWmiServices(".", "root\cimv2")
IF pServices.ServicesPtr = NULL THEN END
 
 '// Get an instance of a file --> change me
DIM cbsPath AS CBSTR = ExePath & "\EX_CWMI_Get_02.bas"   ' --> change me
DIM hr AS HRESULT = pServices.Get("CIM_DataFile.Name='" & cbsPath & "'")
IF hr <> S_OK THEN PRINT AfxWmiGetErrorCodeText(hr) : SLEEP : END

' // Number of properties
PRINT "Number of properties: ", pServices.PropsCount
PRINT

' // Display some properties
PRINT "Relative path: "; pServices.PropValue("Path")
PRINT "FileName: "; pServices.PropValue("FileName")
PRINT "Extension: "; pServices.PropValue("Extension")
PRINT "Size: "; pServices.PropValue("Filesize")
'PRINT pServices.PropValue("LastModified")
PRINT "Date last modified: "; pServices.WmiDateToStr(pServices.PropValue("LastModified"), "dd-MM-yyyy")   ' // change the mask if needed

PRINT
PRINT "Press any key..."
SLEEP
