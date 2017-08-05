' ########################################################################################
' Microsoft Windows
' File: AfxWsh.bi
' Library name: IWshRuntimeLibrary
' Documentation string: Windows Script Host Object Model
' GUID: {F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}
' Version: 1.0, Locale ID = 0
' Path: C:\Windows\SysWOW64\wshom.ocx
' Attributes: 8 [&h00000008]  [HasDiskImage]
' ########################################################################################

#pragma once
#include once "Afx/AfxScrRun.bi"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' // ProgIDs (Program Identifiers)a
CONST AFX_PROGID_WScript_Network_1 = "WScript.Network.1"
CONST AFX_PROGID_WScript_Network_1 = "WScript.Network.1"
CONST AFX_PROGID_WScript_Shell_1 = "WScript.Shell.1"

' // Version independent ProgIDs
CONST AFX_PROGID_WScript_Network = "WScript.Network"
CONST AFX_PROGID_WScript_Network = "WScript.Network"
CONST AFX_PROGID_WScript_Shell = "WScript.Shell"

' // ClsIDs (Class identifiers)
CONST AFX_CLSID_IWshCollection_Class = "{F935DC28-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_CLSID_IWshEnvironment_Class = "{F935DC2A-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_CLSID_IWshNetwork_Class = "{F935DC26-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_CLSID_IWshShell_Class = "{F935DC22-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_CLSID_IWshShortcut_Class = "{F935DC24-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_CLSID_IWshURLShortcut_Class = "{F935DC2C-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_CLSID_WshCollection = "{387DAFF4-DA03-44D2-B0D1-80C927C905AC}"
CONST AFX_CLSID_WshEnvironment = "{F48229AF-E28C-42B5-BB92-E114E62BDD54}"
CONST AFX_CLSID_WshExec = "{08FED191-BE19-11D3-A28B-00104BD35090}"
CONST AFX_CLSID_WshNetwork = "{093FF999-1EA0-4079-9525-9614C3504B74}"
CONST AFX_CLSID_WshShell = "{72C24DD5-D70A-438B-8A42-98424B88AFB8}"
CONST AFX_CLSID_WshShortcut = "{A548B8E4-51D5-4661-8824-DAA1D893DFB2}"
CONST AFX_CLSID_WshURLShortcut = "{50E13488-6F1E-4450-96B0-873755403955}"

' // IIDs (Interface identifiers)
CONST AFX_IID_IWshCollection = "{F935DC27-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_IID_IWshEnvironment = "{F935DC29-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_IID_IWshExec = "{08FED190-BE19-11D3-A28B-00104BD35090}"
CONST AFX_IID_IWshNetwork = "{F935DC25-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_IID_IWshNetwork2 = "{24BE5A31-EDFE-11D2-B933-00104B365C9F}"
CONST AFX_IID_IWshShell = "{F935DC21-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_IID_IWshShell2 = "{24BE5A30-EDFE-11D2-B933-00104B365C9F}"
CONST AFX_IID_IWshShell3 = "{41904400-BE18-11D3-A28B-00104BD35090}"
CONST AFX_IID_IWshShortcut = "{F935DC23-1CF0-11D0-ADB9-00C04FD58A0B}"
CONST AFX_IID_IWshURLShortcut = "{F935DC2B-1CF0-11D0-ADB9-00C04FD58A0B}"

' // Enumerations

' // Documentation string: URLShortcut Object
ENUM WshWindowStyle
   WshHide = 0
   WshNormalFocus = 1
   WshMinimizedFocus = 2
   WshMaximizedFocus = 3
   WshNormalNoFocus = 4
   WshMinimizedNoFocus = 6
END ENUM

' // Documentation string: WSH Exec Object
ENUM WshExecStatus
   WshRunning = 0
   WshFinished = 1
   WshFailed = 2
END ENUM

#ifndef __Afx_IUnknown_INTERFACE_DEFINED__
#define __Afx_IUnknown_INTERFACE_DEFINED__
TYPE Afx_IUnknown AS Afx_IUnknown_
TYPE Afx_IUnknown_ EXTENDS OBJECT
	DECLARE ABSTRACT FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION AddRef () AS ULONG
	DECLARE ABSTRACT FUNCTION Release () AS ULONG
END TYPE
TYPE AFX_LPUNKNOWN AS Afx_IUnknown PTR
#endif

#ifndef __Afx_IDispatch_INTERFACE_DEFINED__
#define __Afx_IDispatch_INTERFACE_DEFINED__
TYPE Afx_IDispatch AS Afx_IDispatch_
TYPE Afx_IDispatch_  EXTENDS Afx_Iunknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) as HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif

' // Dual interfaces - Forward references
TYPE Afx_IWshCollection AS Afx_IWshCollection_
TYPE Afx_IWshEnvironment AS Afx_IWshEnvironment_
TYPE Afx_IWshExec AS Afx_IWshExec_
TYPE Afx_IWshNetwork AS Afx_IWshNetwork_
TYPE Afx_IWshNetwork2 AS Afx_IWshNetwork2_
TYPE Afx_IWshShell AS Afx_IWshShell_
TYPE Afx_IWshShell2 AS Afx_IWshShell2_
TYPE Afx_IWshShell3 AS Afx_IWshShell3_
TYPE Afx_IWshShortcut AS Afx_IWshShortcut_
TYPE Afx_IWshURLShortcut AS Afx_IWshURLShortcut_

' // Dual interfaces

' ########################################################################################
' Interface name: IWshCollection
' IID: {F935DC27-1CF0-11D0-ADB9-00C04FD58A0B}
' Documentation string: Generic Collection Object
' Attributes =  4416 [&h00001140] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 4
' ########################################################################################

#ifndef __Afx_IWshCollection_INTERFACE_DEFINED__
#define __Afx_IWshCollection_INTERFACE_DEFINED__

TYPE Afx_IWshCollection_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION Item (BYVAL Index AS VARIANT PTR, BYVAL out_Value AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Count (BYVAL out_Count AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_length (BYVAL out_Count AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION _NewEnum (BYVAL out_Enum AS Afx_IUnknown PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshEnvironment
' IID: {F935DC29-1CF0-11D0-ADB9-00C04FD58A0B}
' Documentation string: Environment Variables Collection Object
' Attributes =  4416 [&h00001140] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 6
' ########################################################################################

#ifndef __Afx_IWshEnvironment_INTERFACE_DEFINED__
#define __Afx_IWshEnvironment_INTERFACE_DEFINED__

TYPE Afx_IWshEnvironment_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL Name AS AFX_BSTR, BYVAL out_Value AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Item (BYVAL Name AS AFX_BSTR, BYVAL out_Value AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Count (BYVAL out_Count AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_length (BYVAL out_Count AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION _NewEnum (BYVAL out_Enum AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL Name AS AFX_BSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshExec
' IID: {08FED190-BE19-11D3-A28B-00104BD35090}
' Documentation string: WSH Exec Object
' Attributes =  4416 [&h00001140] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 7
' ########################################################################################

#ifndef __Afx_IWshExec_INTERFACE_DEFINED__
#define __Afx_IWshExec_INTERFACE_DEFINED__

TYPE Afx_IWshExec_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Status (BYVAL Status AS WshExecStatus PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StdIn (BYVAL ppts AS Afx_ITextStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StdOut (BYVAL ppts AS Afx_ITextStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StdErr (BYVAL ppts AS Afx_ITextStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ProcessID (BYVAL PID AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ExitCode (BYVAL ExitCode AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Terminate () AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshNetwork
' IID: {F935DC25-1CF0-11D0-ADB9-00C04FD58A0B}
' Documentation string: Network Object
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 13
' ########################################################################################

#ifndef __Afx_IWshNetwork_INTERFACE_DEFINED__
#define __Afx_IWshNetwork_INTERFACE_DEFINED__

TYPE Afx_IWshNetwork_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_UserDomain (BYVAL out_UserDomain AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_UserName (BYVAL out_UserName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_UserProfile (BYVAL out_UserProfile AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ComputerName (BYVAL out_ComputerName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Organization (BYVAL out_Organization AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Site (BYVAL out_Site AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MapNetworkDrive (BYVAL LocalName AS AFX_BSTR, BYVAL RemoteName AS AFX_BSTR, BYVAL UpdateProfile AS VARIANT PTR = NULL, BYVAL UserName AS VARIANT PTR = NULL, BYVAL Password AS VARIANT PTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveNetworkDrive (BYVAL Name AS AFX_BSTR, BYVAL Force AS VARIANT PTR = NULL, BYVAL UpdateProfile AS VARIANT PTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnumNetworkDrives (BYVAL out_Enum AS Afx_IWshCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddPrinterConnection (BYVAL LocalName AS AFX_BSTR, BYVAL RemoteName AS AFX_BSTR, BYVAL UpdateProfile AS VARIANT PTR = NULL, BYVAL UserName AS VARIANT PTR = NULL, BYVAL Password AS VARIANT PTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemovePrinterConnection (BYVAL Name AS AFX_BSTR, BYVAL Force AS VARIANT PTR = NULL, BYVAL UpdateProfile AS VARIANT PTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnumPrinterConnections (BYVAL out_Enum AS Afx_IWshCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetDefaultPrinter (BYVAL Name AS AFX_BSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshNetwork2
' IID: {24BE5A31-EDFE-11D2-B933-00104B365C9F}
' Documentation string: Network Object
' Attributes =  4416 [&h00001140] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IWshNetwork
' Number of methods = 1
' ########################################################################################

#ifndef __Afx_IWshNetwork2_INTERFACE_DEFINED__
#define __Afx_IWshNetwork2_INTERFACE_DEFINED__

TYPE Afx_IWshNetwork2_ EXTENDS Afx_IWshNetwork
   DECLARE ABSTRACT FUNCTION AddWindowsPrinterConnection (BYVAL PrinterName AS AFX_BSTR, BYVAL DriverName AS AFX_BSTR, BYVAL Port AS AFX_BSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshShell
' IID: {F935DC21-1CF0-11D0-ADB9-00C04FD58A0B}
' Documentation string: Shell Object Interface
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 9
' ########################################################################################

#ifndef __Afx_IWshShell_INTERFACE_DEFINED__
#define __Afx_IWshShell_INTERFACE_DEFINED__

TYPE Afx_IWshShell_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_SpecialFolders (BYVAL out_Folders AS Afx_IWshCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Environment (BYVAL Type AS VARIANT PTR = NULL, BYVAL out_Env AS Afx_IWshEnvironment PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Run (BYVAL Command AS AFX_BSTR, BYVAL WindowStyle AS VARIANT PTR = NULL, BYVAL WaitOnReturn AS VARIANT PTR = NULL, BYVAL out_ExitCode AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Popup (BYVAL Text AS AFX_BSTR, BYVAL SecondsToWait AS VARIANT PTR = NULL, BYVAL Title AS VARIANT PTR = NULL, BYVAL Type AS VARIANT PTR = NULL, BYVAL out_Button AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateShortcut (BYVAL PathLink AS AFX_BSTR, BYVAL out_Shortcut AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExpandEnvironmentStrings (BYVAL Src AS AFX_BSTR, BYVAL out_Dst AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RegRead (BYVAL Name AS AFX_BSTR, BYVAL out_Value AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RegWrite (BYVAL Name AS AFX_BSTR, BYVAL Value AS VARIANT PTR, BYVAL Type AS VARIANT PTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION RegDelete (BYVAL Name AS AFX_BSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshShell2
' IID: {24BE5A30-EDFE-11D2-B933-00104B365C9F}
' Documentation string: Shell Object Interface
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IWshShell
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_IWshShell2_INTERFACE_DEFINED__
#define __Afx_IWshShell2_INTERFACE_DEFINED__

TYPE Afx_IWshShell2_ EXTENDS Afx_IWshShell
   DECLARE ABSTRACT FUNCTION LogEvent (BYVAL Type AS VARIANT PTR, BYVAL Message AS AFX_BSTR, BYVAL Target AS AFX_BSTR, BYVAL out_Success AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AppActivate (BYVAL App AS VARIANT PTR, BYVAL Wait AS VARIANT PTR = NULL, BYVAL out_Success AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SendKeys (BYVAL Keys AS AFX_BSTR, BYVAL Wait AS VARIANT PTR = NULL) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshShell3
' IID: {41904400-BE18-11D3-A28B-00104BD35090}
' Documentation string: Shell Object Interface
' Attributes =  4416 [&h00001140] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IWshShell2
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_IWshShell3_INTERFACE_DEFINED__
#define __Afx_IWshShell3_INTERFACE_DEFINED__

TYPE Afx_IWshShell3_ EXTENDS Afx_IWshShell2
   DECLARE ABSTRACT FUNCTION Exec (BYVAL Command AS AFX_BSTR, BYVAL ppExec AS Afx_IWshExec PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CurrentDirectory (BYVAL out_Directory AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CurrentDirectory (BYVAL out_Directory AS AFX_BSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshShortcut
' IID: {F935DC23-1CF0-11D0-ADB9-00C04FD58A0B}
' Documentation string: Shortcut Object
' Attributes =  4416 [&h00001140] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 18
' ########################################################################################

#ifndef __Afx_IWshShortcut_INTERFACE_DEFINED__
#define __Afx_IWshShortcut_INTERFACE_DEFINED__

TYPE Afx_IWshShortcut_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_FullName (BYVAL out_FullName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Arguments (BYVAL out_Arguments AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Arguments (BYVAL out_Arguments AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Description (BYVAL out_Description AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Description (BYVAL out_Description AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Hotkey (BYVAL out_HotKey AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Hotkey (BYVAL out_HotKey AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IconLocation (BYVAL out_IconPath AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IconLocation (BYVAL out_IconPath AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_RelativePath (BYVAL rhs AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TargetPath (BYVAL out_Path AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TargetPath (BYVAL out_Path AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_WindowStyle (BYVAL out_ShowCmd AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_WindowStyle (BYVAL out_ShowCmd AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_WorkingDirectory (BYVAL out_WorkingDirectory AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_WorkingDirectory (BYVAL out_WorkingDirectory AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load (BYVAL PathLink AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Save () AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IWshURLShortcut
' IID: {F935DC2B-1CF0-11D0-ADB9-00C04FD58A0B}
' Documentation string: URLShortcut Object
' Attributes =  4416 [&h00001140] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 5
' ########################################################################################

#ifndef __Afx_IWshURLShortcut_INTERFACE_DEFINED__
#define __Afx_IWshURLShortcut_INTERFACE_DEFINED__

TYPE Afx_IWshURLShortcut_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_FullName (BYVAL out_FullName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TargetPath (BYVAL out_Path AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TargetPath (BYVAL out_Path AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load (BYVAL PathLink AS BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Save () AS HRESULT
END TYPE

#endif

' ########################################################################################

END NAMESPACE
