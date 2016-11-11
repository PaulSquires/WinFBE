' ########################################################################################
' Microsoft Windows
' File: AfxCOM.inc
' Contents: COM wrapper functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ========================================================================================
' Note: Don't forget to initialize the COM library with CoInitialize or CoInitializeEx
' or OleInitialize or you will get GPFs. Call CoUninitialize or OleUninitialize when ending
' the application.
' ========================================================================================

#pragma once
#include once "win/ole2.bi"
#include once "win/unknwnbase.bi"
#include once "win/ocidl.bi"
#include once "Afx/AfxWin.inc"
USING Afx

' // The definition for BSTR in the FreeBASIC headers was incorrectly changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' ########################################################################################
' Base types for declaring ABSTRACT interface methods in other types that inherit from
' these ones. Afx_IUnknown extends the built-in OBJECT type which provides run-time type
' information for all types derived from it using Extends. Extending the built-in Object
' type allows to add an extra hidden vtable pointer field at the top of the Type. The
' vtable is used to dispatch Virtual and Abstract methods and to access information for
' run-time type identification used by Operator Is.
' ########################################################################################

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

' ########################################################################################

' ========================================================================================
' Converts a string into a 16-byte (128-bit) Globally Unique Identifier (GUID)
' To be valid, the string must contain exactly 32 hexadecimal digits, delimited by hyphens
' and enclosed by curly braces. For example: {B09DE715-87C1-11D1-8BE3-0000F8754DA1}
' If pwszGuidText is omited, AfxGuid generates a new unique guid.
' Remarks: I have need to call the UuidCreate function dynamically because, at the time of
' writing, the library for the RPCRT4.DLL seems broken and the linker fails.
' ========================================================================================
PRIVATE FUNCTION AfxGuid (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
   DIM rguid AS GUID
   IF pwszGuidText = NULL THEN
      ' // Generate a new guid
      DIM AS ANY PTR pLib = DyLibLoad("RPCRT4.DLL")
      IF pLib  THEN
         DIM pProc AS FUNCTION (BYVAL Uuid AS UUID PTR) AS RPC_STATUS
         pProc = DyLibSymbol(pLib, "UuidCreate")
         IF pProc THEN pProc(@rguid)
         DyLibFree(pLib)
      END IF
   ELSE
      CLSIDFromString(pwszGuidText, @rGuid)
   END IF
   RETURN rguid
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a 38-byte human-readable guid string from a 16-byte GUID.
' ========================================================================================
PRIVATE FUNCTION AfxGuidText OVERLOAD (BYVAL classID AS CLSID PTR) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromCLSID(classID, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGuidText OVERLOAD (BYVAL classID AS CLSID) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromCLSID(@classID, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGuidText OVERLOAD (BYVAL riid AS REFIID) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromIID(riid, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with a specified ProgID or CLSID.
' Parameters:
' - wszProgID
'   The ProgID or the CLSID of the object to create.
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' wszLicKey (optional)
'   The license key as a unicode string.
' Return value:
'   The interface pointer or NULL.
' Usage examples:
'   DIM pDic AS IDictionary PTR
'   pDic = AfxNewCom("Scripting.Dictionary")
' -or-
'   pDic = AfxNewCom(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
' CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF wszProgID AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "") AS ANY PTR
   DIM hr AS HRESULT, classID AS CLSID, pUnk AS ANY PTR, pIClassFactory2 AS IClassFactory2 PTR
   IF VARPTR(wszProgID) = NULL THEN RETURN NULL
   IF INSTR(wszProgID, "{") THEN CLSIDFromString(wszProgID, @classID) ELSE CLSIDFromProgID(wszProgID, @classID)
   IF IsEqualGuid(@classID, @IID_NULL) THEN RETURN NULL
   IF LEN(wszLicKey) = 0 THEN
      CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   ELSE
      hr = CoGetClassObject(@ClassID, CLSCTX_SERVER, NULL, @IID_IClassFactory2, @pIClassFactory2)
      IF hr = S_OK AND pIClassFactory2 <> NULL THEN
         pIClassFactory2->lpVtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @IID_IUnknown, @wszLicKey, @pUnk)
      END IF
      IF pIClassFactory2 THEN pIClassFactory2->lpVtbl->Release(pIClassFactory2)
   END IF
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with a specified ProgID or CLSID.
' Parameters:
' - wszProgID = A CLSID in string format.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' Return value:
'   The interface pointer or NULL.
' Usage examples:
'   DIM pDic AS IDictionary PTR
'   pDic = AfxNewCom(CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'   CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' and IID_IDictionary as
'   CONST IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF wszClsID AS CONST WSTRING, BYREF wszIID AS CONST WSTRING) AS ANY PTR
   DIM classID AS CLSID, riid AS IID, pUnk AS ANY PTR
   IF VARPTR(wszClsID) = NULL OR VARPTR(wszIID) = NULL THEN RETURN NULL
   CLSIDFromProgID(wszClsID, @classID)
   CLSIDFromProgID(wszIID, @riid)
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @riid, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with a specified CLSID.
' Parameter:
' - classID = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' Return value:
'   The interface pointer or NULL.
' Usage examples:
'   DIM pDic AS IDictionary PTR
'   pDic = AfxNewCom(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
'   DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF classID AS CONST CLSID) AS ANY PTR
   DIM pUnk AS ANY PTR
   IF VARPTR(classID) = NULL THEN EXIT FUNCTION
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a single uninitialized object of the class associated with the specified CLSID and IID.
' Parameters:
' - classID = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' Return value:
'   The interface pointer or NULL.
' Usage examples:
'   DIM pDic AS IDictionary PTR
'   pDic = AfxNewCom(CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'   DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' and IID_IDictionary as
'   DIM IID_IDictionary AS IID = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID) AS ANY PTR
   DIM pUnk AS ANY PTR
   IF VARPTR(classID) = NULL OR VARPTR(riid) = NULL THEN EXIT FUNCTION
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @riid, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads the specified library from file and creates an instance of an object.
' Parameters:
' - wszLibName = Full path where the library is located.
' - rclsid = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' - wszLicKey = The license key.
' If it succeeds, returns a reference to the requested interface; otherwise, it returns null.
' Not every component is a suitable candidate for use under this overloaded AfxNewCom function.
'  - Only in-process servers (DLLs) are supported.
'  - Components that are system components or part of the operating system, such as XML,
'    Data Access, Internet Explorer, or DirectX, aren't supported
'  - Components that are part of an application, such Microsoft Office, aren't supported.
'  - Components intended for use as an add-in or a snap-in, such as an Office add-in or
'    a control in a Web browser, aren't supported.
'  - Components that manage a shared physical or virtual system resource aren't supported.
'  - Visual ActiveX controls aren't supported because they need to be initilized and
'    activated by the OLE container.
' Note: Do not use DyLibFree to unload the library once you have got a valid reference
' to an interface or your application will GPF. Before calling DyLibFree, all the
' interface references must be released. If you don't need to unload the library until
' the application ends, then you don't need to call DyLibFree because CoUninitialize
' closes the COM library on the current thread, unloads all DLLs loaded by the thread,
' frees any other resources that the thread maintains, and forces all RPC connections on
' the thread to close.
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "") AS ANY PTR

   DIM hr AS LONG, hLib AS HANDLE, pDisp AS ANY PTR
   DIM pIClassFactory AS IClassFactory PTR, pIClassFactory2 AS IClassFactory2 PTR

   ' // See if the library is already loaded in the address space
   hLib = GetModuleHandleW(wszLibName)
   ' // If it is not loaded, load it
   IF hLib = NULL THEN hLib = DyLibLoad(wszLibName)
   ' // If it fails, abort
   IF hLib = NULL THEN EXIT FUNCTION

   ' // Retrieve the address of the exported function DllGetClassObject
   DIM pfnDllGetClassObject AS FUNCTION (BYVAL rclsid AS CONST IID CONST PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppv AS LPVOID PTR) AS HRESULT
   pfnDllGetClassObject = DyLibSymbol(hLib, "DllGetClassObject")
   IF pfnDllGetClassObject = NULL THEN EXIT FUNCTION

   IF LEN(wszLicKey) = 0 THEN
      ' // Request a reference to the IClassFactory interface
      hr = pfnDllGetClassObject(@rclsid, @IID_IClassFactory, @pIClassFactory)
      IF hr <> S_OK THEN EXIT FUNCTION
      ' // Create an instance of the server or control
      hr = pIClassFactory->lpVtbl->CreateInstance(pIClassFactory, NULL, @riid, @pDisp)
      IF hr <> S_OK THEN
         pIClassFactory->lpVtbl->Release(pIClassFactory)
         EXIT FUNCTION
      END IF
   ELSE
      ' // Request a reference to the IClassFactory2 interface
      hr = pfnDllGetClassObject(@rclsid, @IID_IClassFactory, @pIClassFactory2)
      IF hr <> S_OK THEN EXIT FUNCTION
      ' // Create a licensed instance of the server or control
      hr = pIClassFactory2->lpVtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @riid, @wszLicKey, @pDisp)
      IF hr <> S_OK THEN
         pIClassFactory2->lpVtbl->Release(pIClassFactory2)
         EXIT FUNCTION
      END IF
   END IF

   IF pIClassFactory THEN pIClassFactory->lpVtbl->Release(pIClassFactory)
   IF pIClassFactory2 THEN pIClassFactory2->lpVtbl->Release(pIClassFactory2)
   RETURN pDisp

END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the wszClsid and wszIid parameters to GUIDs and calls the function above.
' ========================================================================================
PRIVATE FUNCTION AfxNewCom OVERLOAD (BYREF wszLibName AS CONST WSTRING, BYREF wszClsid AS CONST WSTRING, BYREF wszIid AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "") AS ANY PTR
   DIM rclsid AS CLSID, riid AS IID
   rclsid = AfxGuid(wszClsid)
   riid = AfxGuid(wszIid)
   RETURN AfxNewCom(wszLibName, rclsid, riid, wszLicKey)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a display name into a moniker that identifies the object named, and then binds
' to the object identified by the moniker.
' *** add an overloaded function for monikers *** ? Nor needed: call CoGetObject
' optional parameters riid, bind options.
' ========================================================================================
'PRIVATE FUNCTION AfxGetObject (BYVAL pwszDisplayName AS WSTRING PTR) AS IDispatch PTR
'   DIM hr AS LONG, pDisp AS IDispatch PTR
'   hr = CoGetObject(pwszDisplayName, NULL, @IID_IDispatch, @pDisp)
'   IF hr = S_OK THEN RETURN pDisp
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' The function returns the new reference count. This value is intended to be used only
' for test purposes.
' When the reference count on an object reaches zero, Release must cause the interface
' pointer to free itself. When the released pointer is the only existing reference to an
' object (whether the object supports single or multiple interfaces), the implementation
' must free the object.
' ========================================================================================
PRIVATE FUNCTION AfxSafeRelease (BYVAL pv AS ANY PTR) AS ULONG
   IF pv = NULL THEN RETURN 0
   RETURN cast(IUnknown PTR, pv)->lpvtbl->Release(pv)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. The method returns the new
' reference count. This value is intended to be used only for test purposes.
' This method should be called for every new copy of a pointer to an interface on an object.
' For example, if you are passing a copy of a pointer back from a method, you must call
' AddRef on that pointer. You must also call AddRef on a pointer before passing it as an
' in-out parameter to a method; the method will call IUnknown_Release before copying the
' out-value on top of it.
' Objects use a reference counting mechanism to ensure that the lifetime of the object
' includes the lifetime of references to it. You use AddRef to stabilize a copy of an
' interface pointer. It can also be called when the life of a cloned pointer must extend
' beyond the lifetime of the original pointer. The cloned pointer must be released by
' calling AfxSafeRelease.
' ========================================================================================
PRIVATE FUNCTION AfxSafeAddRef (BYVAL pv AS ANY PTR) AS ULONG
   IF pv = NULL THEN RETURN 0
   RETURN cast(IUnknown PTR, pv)->lpvtbl->AddRef(pv)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Establishes a connection between the connection point object and the client's sink.
' Parameters
' pUnk: [in] A pointer to the IUnknown of the object the client wants to connect with.
' pEvtObj: [in] A pointer to the client's IUnknown.
' riid: [in] The GUID of the connection point. Typically, this is the same as the outgoing
' interface managed by the connection point.
' pdwCookie: [out] A pointer to the cookie that uniquely identifies the connection.
' Return value: S_OK or an HRESULT value.
' ========================================================================================
PRIVATE FUNCTION AfxAdvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, BYVAL riid AS IID PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   IF pUnk = NULL OR pEvtObj = NULL OR riid = NULL OR pdwCookie = NULL THEN RETURN E_INVALIDARG
   ' // Query for the IConnectionPointContainer interface
   DIM pCPC AS IConnectionPointContainer PTR
   DIM hr AS HRESULT = IUnknown_QueryInterface(cast(IUnknown PTR, pUnk), @IID_IConnectionPointContainer, @pCPC)
   IF hr <> S_OK OR pCPC = NULL THEN RETURN hr
   ' // Query for the IConnectionPoint interface
   DIM pCP AS IConnectionPoint PTR
   hr = pCPC->lpvtbl->FindConnectionPoint(pCPC, riid, @pCP)
   IUnknown_Release(pCPC)
   IF hr <> S_OK OR pCP = NULL THEN RETURN hr
   ' // Establishes a connection between a connection point object and the client's sink.
   hr = pCP->lpvtbl->Advise(pCP, cast(IUnknown PTR, pEvtObj), pdwCookie)
   IUnknown_Release(pCP)
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxAdvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, BYREF riid AS CONST IID, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   RETURN AfxAdvise(pUnk, pEvtObj, cast(IID PTR, @riid), pdwCookie)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxAdvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYVAL pEvtObj AS ANY PTR, BYREF riid AS IID, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   RETURN AfxAdvise(pUnk, pEvtObj, cast(IID PTR, @riid), pdwCookie)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the events connection identified with the cookie returned by the Advise method.
' Parameters
' pUnk [in] A pointer to the IUnknown of the object the client wants to connect with.
' riid: [in] The GUID of the connection point. Typically, this is the same as the outgoing
' interface managed by the connection point.
' dwCookie: [in] The cookie that uniquely identifies the connection.
' Return value: S_OK or an HRESULT value.
' ========================================================================================
PRIVATE FUNCTION AfxUnadvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYVAL riid AS IID PTR, BYVAL dwCookie AS DWORD) AS HRESULT
   ' // Not a valid connection
   IF pUnk = NULL OR riid = NULL OR dwCookie = 0 THEN RETURN E_INVALIDARG
   ' // Query for the IConnectionPointContainer interface
   DIM pCPC AS IConnectionPointContainer PTR
   DIM hr AS HRESULT = IUnknown_QueryInterface(cast(IUnknown PTR, pUnk), @IID_IConnectionPointContainer, @pCPC)
   IF hr <> S_OK OR pCPC = NULL THEN RETURN hr
   ' // Query for the IConnectionPoint interface
   DIM pCP AS IConnectionPoint PTR
   hr = pCPC->lpvtbl->FindConnectionPoint(pCPC, riid, @pCP)
   IUnknown_Release(pCPC)
   IF hr <> S_OK OR pCP = NULL THEN RETURN hr
   ' // Terminates the advisory connection previously established with a call to AfxAdvise
   hr = pCP->lpvtbl->Unadvise(pCP, dwCookie)
   IUnknown_Release(pCP)
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxUnadvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYREF riid AS CONST IID, BYVAL dwCookie AS DWORD) AS HRESULT
   RETURN AfxUnadvise(pUnk, cast(IID PTR, @riid), dwCookie)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxUnadvise OVERLOAD (BYVAL pUnk AS ANY PTR, BYREF riid AS IID, BYVAL dwCookie AS DWORD) AS HRESULT
   RETURN AfxUnadvise(pUnk, cast(IID PTR, @riid), dwCookie)
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                *** VARIANT WRAPPERS ***
' ########################################################################################

' ========================================================================================
' For the API PROPVARIANT and VARIANT Functions see:
' https://msdn.microsoft.com/en-us/library/windows/desktop/bb762286%28v=vs.85%29.aspx
' ========================================================================================

' ========================================================================================
' Retrieves the element count of a variant structure.
' Note: Requires Windows XP SP2 or superior.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetElementCount (BYVAL pvarIn AS VARIANT PTR) AS ULONG
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetElementCount AS FUNCTION (BYVAL varIn AS VARIANT PTR) AS ULONG
   pVariantGetElementCount = DyLibSymbol(pLib, "VariantGetElementCount")
   IF pVariantGetElementCount = NULL THEN EXIT FUNCTION
   FUNCTION = pVariantGetElementCount(pvarIn)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the contents of a buffer stored in a VARIANT structure of type VT_ARRRAY | VT_UI1.
' Parameters:
' - pvarIn : [in] Reference to a source VARIANT structure.
' - pv     : [out] Pointer to a buffer of length cb bytes.
' - cb     : [in] The size of the pv buffer, in bytes. The buffer should be the same size as
'            the data to be extracted, or smaller.
' Return value:
' Returns one of the following values:
' - S_OK         : Data successfully extracted.
' - E_INVALIDARG : The VARIANT was not of type VT_ARRRAY | VT_UI1.
' - E_FAIL       : The VARIANT buffer value had fewer than cb bytes.
' Note: Requires Windows XP SP2 or superior.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToBuffer (BYVAL pvarIn AS VARIANT PTR, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToBuffer AS FUNCTION (BYVAL pvarIn AS VARIANT PTR, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS HRESULT
   pVariantToBuffer = DyLibSymbol(pLib, "VariantToBuffer")
   IF pVariantToBuffer = NULL THEN FUNCTION = E_FAIL
   FUNCTION = pVariantToBuffer(pvarIn, pv, cb)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the variant value of a variant structure to a string.
' Parameters:
' - pvarIn  : [in] Reference to a source variant structure.
' - pwszBuf : [out] Pointer to the extracted property value if one exists; otherwise, empty.
' - cchBuf  : [in] Specifies string length, in characters.
' Return value:
' If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToString (BYVAL pvarIn AS VARIANT PTR, BYVAL pwszBuf AS WSTRING PTR, BYVAL cchBuf AS UINT) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToString AS FUNCTION (BYVAL pVar AS VARIANT PTR, BYVAL pwszBuf AS WSTRING PTR, BYVAL cchBuf AS UINT) AS HRESULT
   pVariantToString = DyLibSymbol(pLib, "VariantToString")
   IF pVariantToString = NULL THEN FUNCTION = E_FAIL
   FUNCTION = pVariantToString(pvarIn, pwszBuf, cchBuf)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the variant value of a variant structure to a string.
' Parameters:
' - pvarIn   : [in] Reference to a source variant structure.
' - ppwszBuf : [out] Pointer to the extracted property value if one exists; otherwise, empty.
' Return value:
' If this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToStringAlloc (BYVAL pvarIn AS VARIANT PTR, BYVAL ppwszBuf AS WSTRING PTR PTR) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToStringAlloc AS FUNCTION (BYVAL pVar AS VARIANT PTR, BYVAL ppwszBuf AS WSTRING PTR PTR) AS HRESULT
   pVariantToStringAlloc = DyLibSymbol(pLib, "VariantToStringAlloc")
   IF pVariantToStringAlloc = NULL THEN FUNCTION = E_FAIL
   FUNCTION = pVariantToStringAlloc(pvarIn, ppwszBuf)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the contents of a VARIANT to a newly-allocated buffer.
' When pvarIn contains an array of bytes, it returns a pointer to an ansi BSTR with the raw
' contents, without unicode conversion.
' When pvarIn contains an array of other types, each element of the array is appended to
' the resulting string separated with a semicolon and a space.
' The returned pointer must be freed with SysFreeString.
' Parameters:
' - pvarIn = Pointer to the variant.
' - bClear = Clear the contents of the variant (TRUE or FALSE).
' Return value:
' - If the function succeeds, it returns the contents of the variant in string form;
'   if it fails, it returns an empty string and the contents of the variant aren't cleared.
' ========================================================================================
PRIVATE FUNCTION AfxVarToStr (BYVAL pvarIn AS VARIANT PTR, BYVAL bClear AS BOOLEAN = FALSE) AS CWSTR

   DIM cws AS CWSTR
   IF pvarIn = NULL THEN RETURN ""

   SELECT CASE pvarIn->vt

      CASE VT_R4   ' // float
         cws = WSTR(pvarIn->fltVal)

      CASE VT_ARRAY OR VT_R4   ' // array of floats
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS SINGLE PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            ' // The first one has not a leading space
            ' // and the last one has not a trailing ;
            IF i = 0 THEN cws += STR(*pvData)
            IF i = 0 AND cElements > 1 THEN cws += ";"
            IF i > 0 THEN cws += " " & WSTR(*pvData)
            IF i < cElements - 1 AND i > 0 THEN cws += ";"
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE VT_ARRAY OR VT_R8   ' // array of doubles
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS DOUBLE PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            ' // The first one has not a leading space
            ' // and the last one has not a trailing ;
            IF i = 0 THEN cws += STR(*pvData)
            IF i = 0 AND cElements > 1 THEN cws += ";"
            IF i > 0 THEN cws += " " & STR(*pvData)
            IF i < cElements - 1 AND i > 0 THEN cws += ";"
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE VT_ARRAY OR VT_I2, VT_ARRAY OR VT_BOOL   ' // array of shorts or booleans
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS SHORT PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            ' // The first one has not a leading space
            ' // and the last one has not a trailing ;
            IF i = 0 THEN cws += STR(*pvData)
            IF i = 0 AND cElements > 1 THEN cws += ";"
            IF i > 0 THEN cws += " " & STR(*pvData)
            IF i < cElements - 1 AND i > 0 THEN cws += ";"
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE VT_ARRAY OR VT_UI2   ' // array of unsigned shorts
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS USHORT PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            ' // The first one has not a leading space
            ' // and the last one has not a trailing ;
            IF i = 0 THEN cws += WSTR(*pvData)
            IF i = 0 AND cElements > 1 THEN cws += ";"
            IF i > 0 THEN cws += " " & WSTR(*pvData)
            IF i < cElements - 1 AND i > 0 THEN cws += ";"
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE VT_ARRAY OR VT_I4, VT_ARRAY OR VT_INT   ' // array of longs
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS LONG PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            ' // The first one has not a leading space
            ' // and the last one has not a trailing ;
            IF i = 0 THEN cws += STR(*pvData)
            IF i = 0 AND cElements > 1 THEN cws += ";"
            IF i > 0 THEN cws += " " & WSTR(*pvData)
            IF i < cElements - 1 AND i > 0 THEN cws += ";"
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE VT_ARRAY OR VT_UI4, VT_ARRAY OR VT_UINT   ' // array of unsigned longs
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS ULONG PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            ' // The first one has not a leading space
            ' // and the last one has not a trailing ;
            IF i = 0 THEN cws += WSTR(*pvData)
            IF i = 0 AND cElements > 1 THEN cws += ";"
            IF i > 0 THEN cws += " " & WSTR(*pvData)
            IF i < cElements - 1 AND i > 0 THEN cws += ";"
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE VT_ARRAY OR VT_I8   ' // array of longints
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS LONGINT PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            ' // The first one has not a leading space
            ' // and the last one has not a trailing ;
            IF i = 0 THEN cws += WSTR(*pvData)
            IF i = 0 AND cElements > 1 THEN cws += ";"
            IF i > 0 THEN cws += " " & WSTR(*pvData)
            IF i < cElements - 1 AND i > 0 THEN cws += ";"
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE VT_ARRAY OR VT_UI8   ' // array of unsigned longints
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS ULONGINT PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            ' // The first one has not a leading space
            ' // and the last one has not a trailing ;
            IF i = 0 THEN cws += WSTR(*pvData)
            IF i = 0 AND cElements > 1 THEN cws += ";"
            IF i > 0 THEN cws += " " & WSTR(*pvData)
            IF i < cElements - 1 AND i > 0 THEN cws += ";"
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE VT_ARRAY OR VT_VARIANT   ' // array of variants
         DIM cElements AS ULONG = AfxVariantGetElementCount(pvarIn)
         IF cElements < 1 THEN RETURN ""
         IF pvarIn->parray = NULL THEN RETURN ""
         ' // Access the data directly and convert it to string
         DIM pvData AS VARIANT PTR
         DIM hr AS HRESULT = SafeArrayAccessData(pvarIn->parray, @pvData)
         IF hr <> S_OK THEN RETURN ""
         FOR i AS LONG = 0 TO cElements - 1
            DIM ppwszBuf AS WSTRING PTR
            DIM hr AS HRESULT = AfxVariantToStringAlloc(pvData, @ppwszBuf)
            IF hr = S_OK AND ppwszBuf <> NULL THEN
               ' // The first one has not a leading space
               ' // and the last one has not a trailing ;
               IF i = 0 THEN cws += *ppwszBuf
               IF i = 0 AND cElements > 1 THEN cws += ";"
               IF i > 0 THEN cws += " " & *ppwszBuf
               IF i < cElements - 1 AND i > 0 THEN *ppwszBuf += ";"
               CoTaskMemFree ppwszBuf
            END IF
            pvData += 1
         NEXT
         SafeArrayUnaccessData pvarIn->parray

      CASE ELSE
         DIM ppwszBuf AS WSTRING PTR
         DIM hr AS HRESULT = AfxVariantToStringAlloc(pvarIn, @ppwszBuf)
         IF hr <> S_OK OR ppwszBuf = NULL THEN RETURN ""
         cws = *ppwszBuf
         CoTaskMemFree ppwszBuf

   END SELECT

   ' // Clear the passed variant
   IF bClear THEN VariantClear(pVarIn)
   ' // Return the string
   RETURN cws

END FUNCTION
' ========================================================================================
