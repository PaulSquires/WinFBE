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

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' // Renamed from tagPSTIME_FLAGS to AFX_tagPSTIME_FLAGS to avoid conflicts
' // if propvarutil.bi is ever added to the FB headers.
ENUM AFX_tagPSTIME_FLAGS
   PSTF_UTC = &h00000000
   PSTF_LOCAL = &h00000001
END ENUM
TYPE AFX_PSTIME_FLAGS AS LONG

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
' Takes a null terminated wide string as input, and returns a pointer to a new wide string
' allocated with CoTaskMemAlloc (SHStrDupW allocates memory using CoTaskMemAlloc).
' Free the returned string with CoTaskMemFree.
' Note: This is useful when we need to pass a pointer to a null terminated wide string to a
' function or method that will release it. If we pass a WSTRING it will GPF.
' If the length of the input string is 0, CoTaskMemAlloc allocates a zero-length item and
' returns a valid pointer to that item. If there is insufficient memory available,
' CoTaskMemAlloc returns NULL.
' ========================================================================================
PRIVATE FUNCTION AfxWstrAlloc (BYREF wszStr AS WSTRING) AS WSTRING PTR
   DIM ppwsz AS WSTRING PTR
   IF SHStrDupW(@wszStr, @ppwsz) = S_OK THEN RETURN ppwsz
END FUNCTION
' ========================================================================================

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

   DIM hr AS LONG, hLib AS HMODULE, pDisp AS ANY PTR
   DIM pIClassFactory AS IClassFactory PTR, pIClassFactory2 AS IClassFactory2 PTR

   ' // See if the library is already loaded in the address space
   hLib = GetModuleHandleW(wszLibName)
   ' // If it is not loaded, load it
   IF hLib = NULL THEN hLib = LoadLibraryW(wszLibName)
   ' // If it fails, abort
   IF hLib = NULL THEN EXIT FUNCTION

   ' // Retrieve the address of the exported function DllGetClassObject
   DIM pfnDllGetClassObject AS FUNCTION (BYVAL rclsid AS CONST IID CONST PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppv AS LPVOID PTR) AS HRESULT
   pfnDllGetClassObject = CAST(ANY PTR, GetProcAddress(hLib, "DllGetClassObject"))
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
' WARNING: Untested because I fon't have Office installed.
' If the requested object is in an EXE (out-of-process server), such Office applications,
' and it is running and registered in the Running Object Table (ROT), AfxGetCom will
' return a pointer to its interface. AfxAnyCom will first try to use an existing, running
' application if available, or it will create a new instance if not.
' Be aware that AfxGetCom can fail under if Office is running but not registered in the ROT.
' When an Office application starts, it does not immediately register its running objects.
' This optimizes the application's startup process. Instead of registering at startup, an
' Office application registers its running objects in the ROT once it loses focus. Therefore,
' if you attempt to use GetObject or GetActiveObject to attach to a running instance of an
' Office application before the application has lost focus, you might receive an error.
' See: https://support.microsoft.com/en-us/help/238610/getobject-or-getactiveobject-cannot-find-a-running-office-application
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGetCom OVERLOAD (BYREF wszProgID AS CONST WSTRING) AS ANY PTR
   DIM classID AS CLSID, pUnk AS ANY PTR
   CLSIDFromProgID(wszProgID, @classID)
   IF IsEqualGuid(@classID, @IID_NULL) THEN RETURN NULL
   IF GetActiveObject(@classID, NULL, @pUnk) = S_OK THEN RETURN pUnk
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGetCom OVERLOAD (BYREF classID AS CONST CLSID) AS ANY PTR
   DIM pUnk AS ANY PTR
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxAnyCom OVERLOAD (BYREF wszProgID AS CONST WSTRING) AS ANY PTR
   DIM classID AS CLSID, pUnk AS ANY PTR
   pUnk = AfxGetCom(wszProgID)
   IF pUnk THEN RETURN pUnk
   CLSIDFromProgID(wszProgID, @classID)
   IF IsEqualGuid(@classID, @IID_NULL) THEN RETURN NULL
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxAnyCom OVERLOAD (BYREF classID AS CONST CLSID) AS ANY PTR
   DIM pUnk AS ANY PTR
   IF GetActiveObject(@classID, NULL, @pUnk) = S_OK THEN RETURN pUnk
   CoCreateInstance(@classID, NULL, CLSCTX_INPROC_SERVER, @IID_IUnknown, @pUnk)
   RETURN pUnk
END FUNCTION
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
PRIVATE FUNCTION AfxSafeRelease (BYREF pv AS ANY PTR) AS ULONG
   IF pv = NULL THEN RETURN 0
   FUNCTION = cast(IUnknown PTR, pv)->lpvtbl->Release(pv)
   pv = NULL
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

' ========================================================================================
' Returns the description of the most recent OLE error in the current logical thread and
' clears the error state for the thread. It should be called as soon as possible after
' calling a method of an Automation interface and will only succeed if the object supports
' the IErrorInfo interface.
' ========================================================================================
PRIVATE FUNCTION AfxGetOleErrorInfo () AS CWSTR
   DIM pErrorInfo AS IErrorInfo PTR, cwsInfo AS CWSTR
   DIM hr AS HRESULT = GetErrorInfo(0, @pErrorInfo)
   IF hr <> S_OK THEN RETURN ""
   DIM bstrDescription AS AFX_BSTR
   hr = pErrorInfo->lpvtbl->GetDescription(pErrorInfo, @bstrDescription)
   IF hr = S_OK THEN
      cwsInfo = *bstrDescription
      SysFreeString bstrDescription
      AfxSafeRelease(pErrorInfo)
   END IF
   RETURN cwsInfo
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a standard IFontDisp object.
' Parameters:
' - wszFontName: The typeface name, e.g. "Segoe UI"
' - cySize: The point size, e.g. 9.
' - fWeight: Initial weight of the font. If the weight is below 550 (the average of
'      FW_NORMAL, 400, and FW_BOLD, 700), then the Bold property is also initialized to
'      FALSE. If the weight is above 550, the Bold property is set to TRUE.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: A pointer to the object or NULL on failure.
' Remarks: The returned font must be destroyed calling the release method of the IFontDisp
' interface when no longer needed to prevent memory leaks.
' Usage examples:
' DIM pFont AS IFontDisp PTR = AfxOleCreateFontDisp("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
' DIM pFont AS IFontDisp PTR = AfxOleCreateFontDisp("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
' DIM pFont AS IFontDisp PTR = AfxOleCreateFontDisp("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION AfxOleCreateFontDisp ( _
   BYREF wszFontName AS WSTRING, _                  ' __in  Font name
   BYVAL cySize AS LONGLONG, _                      ' __in  Point size
   BYVAL fWeight AS SHORT = FW_NORMAL, _            ' __in  Weight (FW_NORMAL, etc.)
   BYVAL fItalic AS WINBOOL = FALSE, _              ' __in  Italic state
   BYVAL fUnderline AS WINBOOL = FALSE, _           ' __in  Underline state
   BYVAL fStrikethrough AS WINBOOL = FALSE, _       ' __in  Strikethrough state
   BYVAL fCharset AS SHORT = DEFAULT_CHARSET _      ' __in  Character set
   ) AS IFontDisp PTR

   DIM tf AS FONTDESC
   tf.cbSizeOfStruct = SIZEOF(FONTDESC)
   tf.lpstrName =  @wszFontName
   tf.cySize.int64 = cySize * 10000
   tf.sWeight = fWeight
   tf.sCharset = fCharset
   tf.fItalic = fItalic
   tf.fUnderline = fUnderline
   tf.fStrikethrough = fStrikethrough
   DIM pFont AS IFontDisp PTR
   OleCreateFontIndirect(@tf, @IID_IDispatch, @pFont)
   RETURN pFont

END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxOleCreateFont ( _
   BYREF wszFontName AS WSTRING, _                  ' __in  Font name
   BYVAL cySize AS LONGLONG, _                      ' __in  Point size
   BYVAL fWeight AS SHORT = FW_NORMAL, _            ' __in  Weight (FW_NORMAL, etc.)
   BYVAL fItalic AS WINBOOL = FALSE, _              ' __in  Italic state
   BYVAL fUnderline AS WINBOOL = FALSE, _           ' __in  Underline state
   BYVAL fStrikethrough AS WINBOOL = FALSE, _       ' __in  Strikethrough state
   BYVAL fCharset AS SHORT = DEFAULT_CHARSET _      ' __in  Character set
   ) AS IFont PTR

   DIM tf AS FONTDESC
   tf.cbSizeOfStruct = SIZEOF(FONTDESC)
   tf.lpstrName =  @wszFontName
   tf.cySize.int64 = cySize * 10000
   tf.sWeight = fWeight
   tf.sCharset = fCharset
   tf.fItalic = fItalic
   tf.fUnderline = fUnderline
   tf.fStrikethrough = fStrikethrough
   DIM pFont AS IFont PTR
   OleCreateFontIndirect(@tf, @IID_IUnknown, @pFont)
   RETURN pFont

END FUNCTION
' ========================================================================================

' ########################################################################################
'                                *** CComPtr class ***
' ########################################################################################

' ========================================================================================
' - Usage:
' '#CONSOLE ON
' #INCLUDE ONCE "Afx/AfxCom.inc"
' #INCLUDE ONCE "Afx/AfxSapi.bi"
' using Afx
' ' // The COM library must be initialized to call AfxNewCom
' CoInitialize NULL
' ' // Create an instance of the Afx_ISpVoice interface
' DIM pSpVoice AS Afx_ISpVoice PTR = AfxNewCom("SAPI.SpVoice")
' DIM pCComPtrSpVoice AS CComPtr = pSpVoice
' ' // Call the Speak method
' pSpVoice->Speak("Hello World", 0, NULL)
' ' // Uninitialize the COM library
' CoUninitialize
' PRINT
' PRINT "Press any key..."
' SLEEP
' ========================================================================================
' ========================================================================================
' - Usage:
' '#CONSOLE ON
' #INCLUDE ONCE "Afx/AfxCom.inc"
' #INCLUDE ONCE "Afx/AfxSapi.bi"
' using Afx
' ' // The COM library must be initialized to call AfxNewCom
' CoInitialize NULL
' ' // Create an instance of the Afx_ISpVoice interface
' DIM pSpVoice AS CComPtr = AfxNewCom("SAPI.SpVoice")
' ' // Call the Speak method
' CAST(Afx_ISpVoice PTR, *pSpVoice)->Speak("Hello World", 0, NULL)
' ' // Uninitialize the COM library
' CoUninitialize
' PRINT
' PRINT "Press any key..."
' SLEEP
' ========================================================================================

' ========================================================================================
' CComPtr class
' ========================================================================================
TYPE CComPtr
   m_pUnk AS IUnknown PTR
   DECLARE CONSTRUCTOR (BYVAL pUnk AS ANY PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR Let (BYVAL pUnk AS ANY PTR)
   DECLARE SUB Attach (BYVAL pUnk AS ANY PTR)
   DECLARE FUNCTION Detach () AS ANY PTR
END TYPE
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CComPtr (BYVAL pUnk AS ANY PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   m_pUnk = pUnk
   IF fAddRef THEN AfxSafeAddRef(m_pUnk)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CComPtr
   AfxSafeRelease(m_pUnk)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CComPtr.Let (BYVAL pUnk AS ANY PTR)
   AfxSafeRelease(m_pUnk)
   m_pUnk= pUnk
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE SUB CComPtr.Attach (BYVAL pUnk AS ANY PTR)
   IF m_pUnk <> pUnk THEN
      AfxSafeRelease(m_pUnk)
      m_pUnk = pUnk
   END IF
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComPtr.Detach () AS ANY PTR
   DIM temp AS IUnknown PTR = m_pUnk
   m_pUnk = NULL
   RETURN temp
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR * (BYREF _ccomptr AS CComPtr) AS ANY PTR
   OPERATOR = _ccomptr.m_pUnk
END OPERATOR
' ========================================================================================

' ########################################################################################
'                                *** _CComPtr macro ***
' ########################################################################################

' ========================================================================================
' - Usage:
' '#CONSOLE ON
' #INCLUDE ONCE "Afx/AfxCom.inc"
' #INCLUDE ONCE "Afx/AfxSapi.bi"
' using Afx
' ' // The COM library must be initialized to call AfxNewCom
' CoInitialize NULL
' ' // Define the _CComPtrAfx_ISpVoice class
' _CComPtr(Afx_ISpVoice)
' ' // Create an instance of the _CComPtrAfx_ISpVoice class
' ' // and assign an instance of the Afx_ISpVoice interface to it
' DIM pSpVoice AS CComPtrAfx_ISpVoice = AfxNewCom("SAPI.SpVoice")
' ' // Call the Speak method
' DIM cwsText AS CWSTR = "Hello World"
' pSpVoice.vtbl->Speak(cwsText, 0, NULL)
' ' // Uninitialize the COM library
' CoUninitialize
' PRINT
' PRINT "Press any key..."
' SLEEP
' ========================================================================================

' ========================================================================================
' _CComPtr macro
' ========================================================================================
#macro _CComPtr(T)
#ifndef CComPtr##T
TYPE CComPtr##T
Private:
   DIM m_pUnk AS T PTR
   m_bUninitCOM AS BOOLEAN
Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pUnk AS T PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS T PTR
   DECLARE OPERATOR LET (BYVAL pUnk AS T PTR)
   DECLARE FUNCTION vtbl () AS T PTR
   DECLARE FUNCTION vptr () AS T PTR
END TYPE
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CComPtr##T
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' // The first time that is called, pUnk receives a NULL (?), the 2nd time, works!
CONSTRUCTOR CComPtr##T (BYVAL pUnk AS T PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Assign the passed pointer
   m_pUnk = pUnk
   ' // Increase the reference count if requested
   IF fAddRef THEN AfxSafeAddRef(m_pUnk)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
DESTRUCTOR CComPtr##T
   ' // Release the interface pointer
   AfxSafeRelease(m_pUnk)
   ' // Uninitialize the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CComPtr##T.CAST () AS T PTR
   ' // Return an addrefed interface pointer
   AfxSafeAddRef(m_pUnk)
   OPERATOR = m_pUnk
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR CComPtr##T.LET (BYVAL pUnk AS T PTR)
   ' // Release the interface pointer
   AfxSafeRelease(m_pUnk)
   ' // Assign the passed reference counted interface pointer
   m_pUnk = pUnk
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComPtr##T.vtbl () AS T PTR
   ' // Return the stored interface pointer
   RETURN m_pUnk
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CComPtr##T.vptr () AS T PTR
   ' // Release the interface pointer
   AfxSafeRelease(m_pUnk)
   ' // Return the address of the interface pointer
   RETURN @m_pUnk
END FUNCTION
' ========================================================================================
#endif
#endmacro

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
'            the data to be extracted.
' Return value:
' Returns one of the following values:
' - S_OK         : Data successfully extracted.
' - E_INVALIDARG : The VARIANT was not of type VT_ARRRAY OR VT_UI1.
' - E_FAIL       : The VARIANT buffer value had fewer than cb bytes.
' Note: Requires Windows XP SP2 or superior.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToBuffer (BYVAL pvarIn AS VARIANT PTR, BYVAL pv AS LPVOID, BYVAL cb AS ULONG) AS HRESULT
   FUNCTION = E_POINTER
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
   FUNCTION = E_POINTER
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
' Extracts the string property value of a variant structure. If no value exists, then the
' specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToStringWithDefault (BYVAL varIn AS VARIANT PTR, BYVAL pszDefault AS LPCWSTR) AS PCWSTR
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToStringWithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pszDefault AS LPCWSTR) AS PCWSTR
   pVariantToStringWithDefault = DyLibSymbol(pLib, "VariantToStringWithDefault")
   IF pVariantToStringWithDefault THEN FUNCTION = pVariantToStringWithDefault(varIn, pszDefault)
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
   FUNCTION = E_POINTER
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
' Extracts the contents of a VARIANT and returns them as a CWSTR.
' When pvarIn contains an array, each element of the array is appended to the resulting
' string separated with a semicolon and a space.
' For variants that contains an array of bytes, use AfxVariantToBuffer.
' Parameters:
' - pvarIn = Pointer to the variant.
' - bClear = Clear the contents of the variant (TRUE or FALSE).
' Return value:
' - If the function succeeds, it returns the contents of the variant in string form;
'   if it fails, it returns an empty string and the contents of the variant aren't cleared.
' ========================================================================================
PRIVATE FUNCTION AfxVarToStr OVERLOAD (BYVAL pvarIn AS VARIANT PTR, BYVAL bClear AS BOOLEAN = FALSE) AS CWSTR

   DIM cws AS CWSTR
   IF pvarIn = NULL THEN RETURN ""

   SELECT CASE pvarIn->vt

      ' // Unsupported by VariantToStringAlloc
      CASE VT_ERROR : cws = WSTR(pvarin->scode)
      CASE VT_R4 : cws = WSTR(pvarIn->fltVal)
      CASE VT_CY   ' // Currency
         DIM s AS STRING = STR(pvarIn->cyVal.int64 / 10000)
         DIM p AS LONG = INSTR(s, ".")
         DIM dec AS STRING
         IF p THEN
            dec = MID(s + "0000", p + 1, 4)
            s = LEFT(s, p) & dec
         END IF
         IF s = "0" THEN s = "0.0000"
         RETURN s

      ' // VT_BYREF variants
      CASE VT_BSTR OR VT_BYREF : IF pVarIn->pbstrVal THEN cws = *pVarin->pbstrVal
      CASE VT_BOOL OR VT_BYREF : IF pVarIn->pboolVal THEN cws = WSTR(*pVarIn->pboolVal)
      CASE VT_I1 OR VT_BYREF   : IF pVarIn->pcVal    THEN cws = WSTR(*pVarIn->pcVal)
      CASE VT_UI1 OR VT_BYREF  : IF pVarIn->pbVal    THEN cws = WSTR(*pVarIn->pbVal)
      CASE VT_I2 OR VT_BYREF   : IF pVarIn->piVal    THEN cws = WSTR(*pVarIn->piVal)
      CASE VT_UI2 OR VT_BYREF  : IF pVarIn->piVal    THEN cws = WSTR(*pVarIn->puiVal)
      CASE VT_INT OR VT_BYREF  : IF pVarIn->pintVal  THEN cws = WSTR(*pVarIn->pintVal)
      CASE VT_UINT OR VT_BYREF : IF pVarIn->puintVal THEN cws = WSTR(*pVarIn->puintVal)
      CASE VT_I4 OR VT_BYREF   : IF pVarIn->plVal    THEN cws = WSTR(*pVarIn->plVal)
      CASE VT_UI4 OR VT_BYREF  : IF pVarIn->pulVal   THEN cws = WSTR(*pVarIn->pulVal)
      CASE VT_I8 OR VT_BYREF   : IF pVarIn->pllVal   THEN cws = WSTR(*pVarIn->pllVal)
      CASE VT_UI8 OR VT_BYREF  : IF pVarIn->pullVal  THEN cws = WSTR(*pVarIn->pullVal)
      CASE VT_R4 OR VT_BYREF   : IF pvarIn->pfltVal  THEN cws = WSTR(*pvarIn->pfltVal)
      CASE VT_R8 OR VT_BYREF   : IF pvarIn->pdblVal  THEN cws = WSTR(*pvarIn->pdblVal)

      ' // Array of floats
      CASE VT_ARRAY OR VT_R4
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

      ' // Array of doubles
      CASE VT_ARRAY OR VT_R8
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

      ' // Array of signed shorts or bools
      CASE VT_ARRAY OR VT_I2, VT_ARRAY OR VT_BOOL
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

      ' // Array of unsigned shorts
      CASE VT_ARRAY OR VT_UI2
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

      ' // Array of signed longs
      CASE VT_ARRAY OR VT_I4, VT_ARRAY OR VT_INT
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

      ' // Array of unsigned longs
      CASE VT_ARRAY OR VT_UI4, VT_ARRAY OR VT_UINT
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

      ' // Array of longints
      CASE VT_ARRAY OR VT_I8
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

      ' // Array of unsigned longints
      CASE VT_ARRAY OR VT_UI8
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

      ' // Array of variants
      CASE VT_ARRAY OR VT_VARIANT
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

      ' // Other...
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

' ========================================================================================
' Allows to pass the variant by reference, i.e. AfxVarToStr(v), instead of using AfxVarToStr(@v)
' This also allows to use directly the retult of a function that returns a VARIANT or a CVAR,
' e.g. AfxVarToStr(AfxVarFromLong(83838), TRUE), AfxVarToStr(AfxCVarFromLong(83838)), while
' trying to use the first overloaded function, e.g. AfxVarToStr(@AfxVarFromLong(83838), TRUE),
' gives an apparently bogus compiler error: Error 7: Expected ")".
' ========================================================================================
PRIVATE FUNCTION AfxVarToStr OVERLOAD (BYREF varIn AS VARIANT, BYVAL bClear AS BOOLEAN = FALSE) AS CWSTR
   RETURN AfxVarToStr(@varIn, bClear)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure based on a string resource imbedded in an executable file.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromResource (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromResource AS FUNCTION (BYVAL hinst AS HINSTANCE, BYVAL id AS UINT, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromResource = DyLibSymbol(pLib, "InitVariantFromResource")
   IF pInitVariantFromResource THEN FUNCTION = pInitVariantFromResource(hinst, id, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with the contents of a buffer.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromBuffer (BYVAL pv AS VOID PTR, BYVAL cb AS UINT, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromBuffer AS FUNCTION (BYVAL pv AS VOID PTR, BYVAL cb AS UINT, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromBuffer = DyLibSymbol(pLib, "InitVariantFromBuffer")
   IF pInitVariantFromBuffer THEN FUNCTION = pInitVariantFromBuffer(pv, cb, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure based on a GUID. The structure is initialized as a VT_BSTR type.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromGUIDAsString (BYVAL guid AS IID PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromGUIDAsString AS FUNCTION (BYVAL guid AS IID PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromGUIDAsString = DyLibSymbol(pLib, "InitVariantFromGUIDAsString")
   IF pInitVariantFromGUIDAsString THEN FUNCTION = pInitVariantFromGUIDAsString(guid, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with the contents of a FILETIME structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromFileTime (BYVAL pft AS FILETIME PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromFileTime AS FUNCTION (BYVAL pft AS FILETIME PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromFileTime = DyLibSymbol(pLib, "InitVariantFromFileTime")
   IF pInitVariantFromFileTime THEN FUNCTION = pInitVariantFromFileTime(pft, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of FILETIME structures.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromFileTimeArray (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromFileTimeArray AS FUNCTION (BYVAL prgft AS FILETIME PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromFileTimeArray = DyLibSymbol(pLib, "InitVariantFromFileTimeArray")
   IF pInitVariantFromFileTimeArray THEN FUNCTION = pInitVariantFromFileTimeArray(prgft, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with a string stored in a STRRET structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromStrRet (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromStrRet AS FUNCTION (BYVAL pstrret AS STRRET PTR, BYVAL pidl AS PCUITEMID_CHILD, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromStrRet = DyLibSymbol(pLib, "InitVariantFromStrRet")
   IF pInitVariantFromStrRet THEN FUNCTION = pInitVariantFromStrRet(pstrret, pidl, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with a value stored in another VARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromVariantArrayElem (BYVAL varIn AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromVariantArrayElem AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromVariantArrayElem = DyLibSymbol(pLib, "InitVariantFromVariantArrayElem")
   IF pInitVariantFromVariantArrayElem THEN FUNCTION = pInitVariantFromVariantArrayElem(varIn, iElem, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure from an array of Boolean values.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromBooleanArray (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromBooleanArray AS FUNCTION (BYVAL prgf AS WINBOOL PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromBooleanArray = DyLibSymbol(pLib, "InitVariantFromBooleanArray")
   IF pInitVariantFromBooleanArray THEN FUNCTION = pInitVariantFromBooleanArray(prgf, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of 16-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromInt16Array (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromInt16Array AS FUNCTION (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromInt16Array = DyLibSymbol(pLib, "InitVariantFromInt16Array")
   IF pInitVariantFromInt16Array THEN FUNCTION = pInitVariantFromInt16Array(prgf, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of unsigned 16-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromUInt16Array (BYVAL prgf AS USHORT PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromUInt16Array AS FUNCTION (BYVAL prgf AS SHORT PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromUInt16Array = DyLibSymbol(pLib, "InitVariantFromUInt16Array")
   IF pInitVariantFromUInt16Array THEN FUNCTION = pInitVariantFromUInt16Array(prgf, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of 32-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromInt32Array (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromInt32Array AS FUNCTION (BYVAL prgn AS LONG PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromInt32Array = DyLibSymbol(pLib, "InitVariantFromInt32Array")
   IF pInitVariantFromInt32Array THEN FUNCTION = pInitVariantFromInt32Array(prgn, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of unsigned 32-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromUInt32Array (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromUInt32Array AS FUNCTION (BYVAL prgn AS ULONG PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromUInt32Array = DyLibSymbol(pLib, "InitVariantFromUInt32Array")
   IF pInitVariantFromUInt32Array THEN FUNCTION = pInitVariantFromUInt32Array(prgn, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of 64-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromInt64Array (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromInt64Array AS FUNCTION (BYVAL prgn AS LONGINT PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromInt64Array = DyLibSymbol(pLib, "InitVariantFromInt64Array")
   IF pInitVariantFromInt64Array THEN FUNCTION = pInitVariantFromInt64Array(prgn, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of unsigned 64-bit integer values.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromUInt64Array (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromUInt64Array AS FUNCTION (BYVAL prgn AS ULONGINT PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromUInt64Array = DyLibSymbol(pLib, "InitVariantFromUInt64Array")
   IF pInitVariantFromUInt64Array THEN FUNCTION = pInitVariantFromUInt64Array(prgn, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of values of type DOUBLE.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromDoubleArray (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromDoubleArray AS FUNCTION (BYVAL prgn AS DOUBLE PTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromDoubleArray = DyLibSymbol(pLib, "InitVariantFromDoubleArray")
   IF pInitVariantFromDoubleArray THEN FUNCTION = pInitVariantFromDoubleArray(prgn, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes a VARIANT structure with an array of strings.
' ========================================================================================
PRIVATE FUNCTION AfxVariantFromStringArray (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pInitVariantFromStringArray AS FUNCTION (BYVAL prgsz AS PCWSTR, BYVAL cElems AS ULONG, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pInitVariantFromStringArray = DyLibSymbol(pLib, "InitVariantFromStringArray")
   IF pInitVariantFromStringArray THEN FUNCTION = pInitVariantFromStringArray(prgsz, cElems, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the value of a Boolean property from a VARIANT structure. If no value can be
' extracted, then a default value is assigned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToBoolean (BYVAL varIn AS VARIANT PTR, BYVAL pfRet AS WINBOOL PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToBoolean AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pfRet AS WINBOOL PTR) AS HRESULT
   pVariantToBoolean = DyLibSymbol(pLib, "VariantToBoolean")
   IF pVariantToBoolean THEN FUNCTION = pVariantToBoolean(varIn, pfRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts a BOOL value from a VARIANT structure. If no value exists, then the specified
' default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToBooleanWithDefault (BYVAL varIn AS VARIANT PTR, BYVAL fDefault AS WINBOOL) AS WINBOOL
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToBooleanWithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL fDefault AS WINBOOL) AS WINBOOL
   pVariantToBooleanWithDefault = DyLibSymbol(pLib, "VariantToBooleanWithDefault")
   IF pVariantToBooleanWithDefault THEN FUNCTION = pVariantToBooleanWithDefault(varIn, fDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts the Int16 property value of a variant structure. If no value can be extracted,
' then a default value is assigned by this function.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt16 (BYVAL varIn AS VARIANT PTR, BYVAL piRet AS SHORT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt16 AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL piRet AS SHORT PTR) AS HRESULT
   pVariantToInt16 = DyLibSymbol(pLib, "VariantToInt16")
   IF pVariantToInt16 THEN FUNCTION = pVariantToInt16(varIn, piRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts an Int16 property value of a variant structure. If no value exists, then the
' specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt16WithDefault (BYVAL varIn AS VARIANT PTR, BYVAL iDefault AS SHORT) AS SHORT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   DIM pVariantToInt16WithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL iDefault AS SHORT) AS SHORT
   pVariantToInt16WithDefault = DyLibSymbol(pLib, "VariantToInt16WithDefault")
   IF pVariantToInt16WithDefault THEN FUNCTION = pVariantToInt16WithDefault(varIn, iDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an unsigned Int16 property value of a variant structure. If no value can be
' extracted, then a default value is assigned by this function.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt16 (BYVAL varIn AS VARIANT PTR, BYVAL puiRet AS USHORT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt16 AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL puiRet AS USHORT PTR) AS HRESULT
   pVariantToUInt16 = DyLibSymbol(pLib, "VariantToUInt16")
   IF pVariantToUInt16 THEN FUNCTION = pVariantToUInt16(varIn, puiRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts an unsigned Int16 property value of a variant structure. If no value exists,
' then the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt16WithDefault (BYVAL varIn AS VARIANT PTR, BYVAL uiDefault AS USHORT) AS USHORT
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt16WithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL uiDefault AS USHORT) AS USHORT
   pVariantToUInt16WithDefault = DyLibSymbol(pLib, "VariantToUInt16WithDefault")
   IF pVariantToUInt16WithDefault THEN FUNCTION = pVariantToUInt16WithDefault(varIn, uiDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an Int32 property value of a variant structure. If no value can be extracted,
' then a default value is assigned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt32 (BYVAL varIn AS VARIANT PTR, BYVAL plRet AS LONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt32 AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL plRet AS LONG PTR) AS HRESULT
   pVariantToInt32 = DyLibSymbol(pLib, "VariantToInt32")
   IF pVariantToInt32 THEN FUNCTION = pVariantToInt32(varIn, plRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts an Int32 property value of a variant structure. If no value exists, then the
' specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt32WithDefault (BYVAL varIn AS VARIANT PTR, BYVAL lDefault AS LONG) AS LONG
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt32WithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL lDefault AS LONG) AS LONG
   pVariantToInt32WithDefault = DyLibSymbol(pLib, "VariantToInt32WithDefault")
   IF pVariantToInt32WithDefault THEN FUNCTION = pVariantToInt32WithDefault(varIn, lDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts unsigned Int32 property value of a variant structure. If no value can be
' extracted, then a default value is assigned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt32 (BYVAL varIn AS VARIANT PTR, BYVAL pulRet AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt32 AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pulRet AS ULONG PTR) AS HRESULT
   pVariantToUInt32 = DyLibSymbol(pLib, "VariantToUInt32")
   IF pVariantToUInt32 THEN FUNCTION = pVariantToUInt32(varIn, pulRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts an unsigned Int32 property value of a variant structure. If no value currently
' exists, then the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt32WithDefault (BYVAL varIn AS VARIANT PTR, BYVAL ulDefault AS ULONG) AS ULONG
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt32WithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL ulDefault AS ULONG) AS ULONG
   pVariantToUInt32WithDefault = DyLibSymbol(pLib, "VariantToUInt32WithDefault")
   IF pVariantToUInt32WithDefault THEN FUNCTION = pVariantToUInt32WithDefault(varIn, ulDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an Int64 property value of a variant structure. If no value can be extracted,
' then a default value is assigned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt64 (BYVAL varIn AS VARIANT PTR, BYVAL pllRet AS LONGINT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt64 AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pllRet AS LONGINT PTR) AS HRESULT
   pVariantToInt64 = DyLibSymbol(pLib, "VariantToInt64")
   IF pVariantToInt64 THEN FUNCTION = pVariantToInt64(varIn, pllRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts an Int64 property value of a variant structure. If no value exists, then the
' specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt64WithDefault (BYVAL varIn AS VARIANT PTR, BYVAL llDefault AS LONGINT) AS LONGINT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt64WithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL llDefault AS LONGINT) AS LONGINT
   pVariantToInt64WithDefault = DyLibSymbol(pLib, "VariantToInt64WithDefault")
   IF pVariantToInt64WithDefault THEN FUNCTION = pVariantToInt64WithDefault(varIn, llDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts unsigned Int64 property value of a variant structure. If no value can be
' extracted, then a default value is assigned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt64 (BYVAL varIn AS VARIANT PTR, BYVAL pullRet AS ULONGINT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt64 AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pullRet AS ULONGINT PTR) AS HRESULT
   pVariantToUInt64 = DyLibSymbol(pLib, "VariantToUInt64")
   IF pVariantToUInt64 THEN FUNCTION = pVariantToUInt64(varIn, pullRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts an unsigned Int64 property value of a variant structure. If no value currently
' exists, then the specified default value is returned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt64WithDefault (BYVAL varIn AS VARIANT PTR, BYVAL ullDefault AS ULONGINT) AS ULONGINT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt64WithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL ullDefault AS ULONGINT) AS ULONGINT
   pVariantToUInt64WithDefault = DyLibSymbol(pLib, "VariantToUInt64WithDefault")
   IF pVariantToUInt64WithDefault THEN FUNCTION = pVariantToUInt64WithDefault(varIn, ullDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a DOUBLE value from a VARIANT structure. If no value can be extracted, then a
' default value is assigned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToDouble (BYVAL varIn AS VARIANT PTR, BYVAL pdblRet AS DOUBLE PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToDouble AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pdblRet AS DOUBLE PTR) AS HRESULT
   pVariantToDouble = DyLibSymbol(pLib, "VariantToDouble")
   IF pVariantToDouble THEN FUNCTION = pVariantToDouble(varIn, pdblRet)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts a DOUBLE value from a VARIANT structure. If no value can be extracted, then a
' default value is assigned.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToDoubleWithDefault (BYVAL varIn AS VARIANT PTR, BYVAL dblDefault AS DOUBLE) AS DOUBLE
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToDoubleWithDefault AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL dblDefault AS DOUBLE) AS DOUBLE
   pVariantToDoubleWithDefault = DyLibSymbol(pLib, "VariantToDoubleWithDefault")
   IF pVariantToDoubleWithDefault THEN FUNCTION = pVariantToDoubleWithDefault(varIn, dblDefault)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a GUID property value of a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToGUID (BYVAL varIn AS VARIANT PTR, BYVAL pguid AS GUID PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToGUID AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pguid AS GUID PTR) AS HRESULT
   pVariantToGUID = DyLibSymbol(pLib, "VariantToGUID")
   IF pVariantToGUID THEN FUNCTION = pVariantToGUID(varIn, pguid)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a GUID property value of a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToDosDateTime (BYVAL varIn AS VARIANT PTR, BYVAL pwDate AS WORD PTR, BYVAL pwTime AS WORD PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToDosDateTime AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pwDate AS WORD PTR, BYVAL pwTime AS WORD PTR) AS HRESULT
   pVariantToDosDateTime = DyLibSymbol(pLib, "VariantToDosDateTime")
   IF pVariantToDosDateTime THEN FUNCTION = pVariantToDosDateTime(varIn, pwDate, pwTime)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' If the source variant is a VT_BSTR, extracts string and places it into a STRRET structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToStrRet (BYVAL varIn AS VARIANT PTR, BYVAL pstrret AS STRRET PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToStrRet AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL pstrret AS STRRET PTR) AS HRESULT
   pVariantToStrRet = DyLibSymbol(pLib, "VariantToStrRet")
   IF pVariantToStrRet THEN FUNCTION = pVariantToStrRet(varIn, pstrret)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a FILETIME structure from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToFileTime (BYVAL varIn AS VARIANT PTR, BYVAL stfOut AS AFX_PSTIME_FLAGS, BYVAL pftOut AS FILETIME PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToFileTime AS FUNCTION (BYVAL varIn AS VARIANT PTR, BYVAL stfOut AS AFX_PSTIME_FLAGS, BYVAL pftOut AS FILETIME PTR) AS HRESULT
   pVariantToFileTime = DyLibSymbol(pLib, "VariantToFileTime")
   IF pVariantToFileTime THEN FUNCTION = pVariantToFileTime(varIn, stfOut, pftOut)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an array of Boolean values from a VARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToBooleanArray (BYVAL pvar AS VARIANT PTR, BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToBooleanArray AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgf AS WINBOOL PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToBooleanArray = DyLibSymbol(pLib, "VariantToBooleanArray")
   IF pVariantToBooleanArray THEN FUNCTION = pVariantToBooleanArray(pvar, prgf, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Allocates an array of BOOL values then extracts data from a VARIANT structure into that array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToBooleanArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgf AS WINBOOL PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToBooleanArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgf AS WINBOOL PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToBooleanArrayAlloc = DyLibSymbol(pLib, "VariantToBooleanArrayAlloc")
   IF pVariantToBooleanArrayAlloc THEN FUNCTION = pVariantToBooleanArrayAlloc(pvar, pprgf, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt16Array (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt16Array AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS SHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToInt16Array = DyLibSymbol(pLib, "VariantToInt16Array")
   IF pVariantToInt16Array THEN FUNCTION = pVariantToInt16Array(pvar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt16ArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS SHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt16ArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS SHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToInt16ArrayAlloc = DyLibSymbol(pLib, "VariantToInt16ArrayAlloc")
   IF pVariantToInt16ArrayAlloc THEN FUNCTION = pVariantToInt16ArrayAlloc(pvar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt16Array (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt16Array AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS USHORT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToUInt16Array = DyLibSymbol(pLib, "VariantToUInt16Array")
   IF pVariantToUInt16Array THEN FUNCTION = pVariantToUInt16Array(pvar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt16ArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS USHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt16ArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS USHORT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToUInt16ArrayAlloc = DyLibSymbol(pLib, "VariantToUInt16ArrayAlloc")
   IF pVariantToUInt16ArrayAlloc THEN FUNCTION = pVariantToUInt16ArrayAlloc(pvar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt32Array (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt32Array AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS LONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToInt32Array = DyLibSymbol(pLib, "VariantToInt32Array")
   IF pVariantToInt32Array THEN FUNCTION = pVariantToInt32Array(pvar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt32ArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS LONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt32ArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS LONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToInt32ArrayAlloc = DyLibSymbol(pLib, "VariantToInt32ArrayAlloc")
   IF pVariantToInt32ArrayAlloc THEN FUNCTION = pVariantToInt32ArrayAlloc(pvar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int16 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt32Array (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt32Array AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS ULONG PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToUInt32Array = DyLibSymbol(pLib, "VariantToUInt32Array")
   IF pVariantToUInt32Array THEN FUNCTION = pVariantToUInt32Array(pvar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated unsigned Int32 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt32ArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS ULONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt32ArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS ULONG PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToUInt32ArrayAlloc = DyLibSymbol(pLib, "VariantToUInt32ArrayAlloc")
   IF pVariantToUInt32ArrayAlloc THEN FUNCTION = pVariantToUInt32ArrayAlloc(pvar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int64 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt64Array (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt64Array AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS LONGINT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToInt64Array = DyLibSymbol(pLib, "VariantToInt64Array")
   IF pVariantToInt64Array THEN FUNCTION = pVariantToInt64Array(pvar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated Int64 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToInt64ArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS LONGINT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToInt64ArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS LONGINT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToInt64ArrayAlloc = DyLibSymbol(pLib, "VariantToInt64ArrayAlloc")
   IF pVariantToInt64ArrayAlloc THEN FUNCTION = pVariantToInt64ArrayAlloc(pvar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into an unsigned Int64 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt64Array (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt64Array AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS ULONGINT PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToUInt64Array = DyLibSymbol(pLib, "VariantToUInt64Array")
   IF pVariantToUInt64Array THEN FUNCTION = pVariantToUInt64Array(pvar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated Int64 array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToUInt64ArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS ULONGINT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToUInt64ArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS ULONGINT PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToUInt64ArrayAlloc = DyLibSymbol(pLib, "VariantToUInt64ArrayAlloc")
   IF pVariantToUInt64ArrayAlloc THEN FUNCTION = pVariantToUInt64ArrayAlloc(pvar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts an array of DOUBLE values from a VARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToDoubleArray (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToDoubleArray AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgn AS DOUBLE PTR, BYVAL crgn AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToDoubleArray = DyLibSymbol(pLib, "VariantToDoubleArray")
   IF pVariantToDoubleArray THEN FUNCTION = pVariantToDoubleArray(pvar, prgn, crgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Allocates an array of DOUBLE values then extracts data from a VARIANT structure into that array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToDoubleArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS DOUBLE PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToDoubleArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgn AS DOUBLE PTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToDoubleArrayAlloc = DyLibSymbol(pLib, "VariantToDoubleArrayAlloc")
   IF pVariantToDoubleArrayAlloc THEN FUNCTION = pVariantToDoubleArrayAlloc(pvar, pprgn, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts data from a vector structure into a string array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToStringArray (BYVAL pvar AS VARIANT PTR, BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToStringArray AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL prgsz AS PWSTR, BYVAL crgsz AS ULONG, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToStringArray = DyLibSymbol(pLib, "VariantToStringArray")
   IF pVariantToStringArray THEN FUNCTION = pVariantToStringArray(pvar, prgsz, crgsz, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Extracts data from a vector structure into a newly-allocated string array.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToStringArrayAlloc (BYVAL pvar AS VARIANT PTR, BYVAL pprgsz AS PWSTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToStringArrayAlloc AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL pprgsz AS PWSTR PTR, BYVAL pcElem AS ULONG PTR) AS HRESULT
   pVariantToStringArrayAlloc = DyLibSymbol(pLib, "VariantToStringArrayAlloc")
   IF pVariantToStringArrayAlloc THEN FUNCTION = pVariantToStringArrayAlloc(pvar, pprgsz, pcElem)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single Boolean element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetBooleanElem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pfVal AS WINBOOL PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetBooleanElem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pfVal AS WINBOOL PTR) AS HRESULT
   pVariantGetBooleanElem = DyLibSymbol(pLib, "VariantGetBooleanElem")
   IF pVariantGetBooleanElem THEN FUNCTION = pVariantGetBooleanElem(pvar, iElem, pfVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single Int16 element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetInt16Elem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pfVal AS SHORT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetInt16Elem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pfVal AS SHORT PTR) AS HRESULT
   pVariantGetInt16Elem = DyLibSymbol(pLib, "VariantGetInt16Elem")
   IF pVariantGetInt16Elem THEN FUNCTION = pVariantGetInt16Elem(pvar, iElem, pfVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single unsigned Int16 element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetUInt16Elem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS USHORT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetUInt16Elem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS USHORT PTR) AS HRESULT
   pVariantGetUInt16Elem = DyLibSymbol(pLib, "VariantGetUInt16Elem")
   IF pVariantGetUInt16Elem THEN FUNCTION = pVariantGetUInt16Elem(pvar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single Int32 element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetInt32Elem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS LONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetInt32Elem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS LONG PTR) AS HRESULT
   pVariantGetInt32Elem = DyLibSymbol(pLib, "VariantGetInt32Elem")
   IF pVariantGetInt32Elem THEN FUNCTION = pVariantGetInt32Elem(pvar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single unsigned Int32 element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetUInt32Elem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONG PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetUInt32Elem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONG PTR) AS HRESULT
   pVariantGetUInt32Elem = DyLibSymbol(pLib, "VariantGetUInt32Elem")
   IF pVariantGetUInt32Elem THEN FUNCTION = pVariantGetUInt32Elem(pvar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single Int64 element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetInt64Elem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS LONGINT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetInt64Elem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS LONGINT PTR) AS HRESULT
   pVariantGetInt64Elem = DyLibSymbol(pLib, "VariantGetInt64Elem")
   IF pVariantGetInt64Elem THEN FUNCTION = pVariantGetInt64Elem(pvar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single unsigned Int64 element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetUInt64Elem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONGINT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetUInt64Elem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS ULONGINT PTR) AS HRESULT
   pVariantGetUInt64Elem = DyLibSymbol(pLib, "VariantGetUInt64Elem")
   IF pVariantGetUInt64Elem THEN FUNCTION = pVariantGetUInt64Elem(pvar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts one double element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetDoubleElem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS DOUBLE PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetDoubleElem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL pnVal AS DOUBLE PTR) AS HRESULT
   pVariantGetDoubleElem = DyLibSymbol(pLib, "VariantGetDoubleElem")
   IF pVariantGetDoubleElem THEN FUNCTION = pVariantGetDoubleElem(pvar, iElem, pnVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Extracts a single wide string element from a variant structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantGetStringElem (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL ppszVal AS PWSTR PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantGetStringElem AS FUNCTION (BYVAL pvar AS VARIANT PTR, BYVAL iElem AS ULONG, BYVAL ppszVal AS PWSTR PTR) AS HRESULT
   pVariantGetStringElem = DyLibSymbol(pLib, "VariantGetStringElem")
   IF pVariantGetStringElem THEN FUNCTION = pVariantGetStringElem(pvar, iElem, ppszVal)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Frees the memory and references used by an array of VARIANT structures stored in an array.
' ========================================================================================
PRIVATE SUB AfxClearVariantArray (BYVAL pvars AS VARIANT PTR, BYVAL cvars AS UINT)
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT SUB
   DIM pClearVariantArray AS SUB (BYVAL pvars AS VARIANT PTR, BYVAL cvars AS UINT)
   pClearVariantArray = DyLibSymbol(pLib, "ClearVariantArray")
   IF pClearVariantArray THEN pClearVariantArray(pvars, cvars)
   DyLibFree(pLib)
END SUB
' ========================================================================================

' ========================================================================================
' Compares two variant structures, based on default comparison rules.
' ========================================================================================
PRIVATE FUNCTION AfxVariantCompare (BYVAL var1 AS VARIANT PTR, BYVAL var2 AS VARIANT PTR) AS INT_
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantCompare AS FUNCTION (BYVAL var1 AS VARIANT PTR, BYVAL var2 AS VARIANT PTR) AS INT_
   pVariantCompare = DyLibSymbol(pLib, "VariantCompare")
   IF pVariantCompare THEN FUNCTION = pVariantCompare(var1, var2)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the contents of a PROPVARIANT structure to a VARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxPropVariantToVariant (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pPropVariantToVariant AS FUNCTION (BYVAL pPropVar AS PROPVARIANT PTR, BYVAL pVar AS VARIANT PTR) AS HRESULT
   pPropVariantToVariant = DyLibSymbol(pLib, "PropVariantToVariant")
   IF pPropVariantToVariant THEN FUNCTION = pPropVariantToVariant(pPropVar, pVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the contents of a VARIANT structure to a PROPVARIANT structure.
' ========================================================================================
PRIVATE FUNCTION AfxVariantToPropVariant (BYVAL pVar AS VARIANT PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("propsys.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pVariantToPropVariant AS FUNCTION (BYVAL pVar AS VARIANT PTR, BYVAL pPropVar AS PROPVARIANT PTR) AS HRESULT
   pVariantToPropVariant = DyLibSymbol(pLib, "VariantToPropVariant")
   IF pVariantToPropVariant THEN FUNCTION = pVariantToPropVariant(pVar, pPropVar)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Variant types
' ========================================================================================
PRIVATE FUNCTION AfxGetVarType (BYVAL pvar AS VARIANT PTR) AS VARTYPE
   RETURN pvar->vt
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsVarTypeFloat (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN (vt = VT_R4 OR vt = VT_R8)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsVariantArray (BYVAL pvar AS VARIANT PTR) AS BOOLEAN
   RETURN (pvar->vt AND VT_ARRAY)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsVariantString (BYVAL pvar AS VARIANT PTR) AS BOOLEAN
   RETURN AfxVariantToStringWithDefault(pvar, NULL) <> NULL
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsVarTypeSignedInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   DIM fRet AS BOOLEAN
   SELECT CASE vt
      CASE VT_I1, VT_I2, VT_I4, VT_I8
         fRet = TRUE
   END SELECT
   RETURN fRet
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsVarTypeUnsignedInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   DIM fRet AS BOOLEAN
   SELECT CASE vt
      CASE VT_UI1, VT_UI2, VT_UI4, VT_UI8
         fRet = TRUE
   END SELECT
   RETURN fRet
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsVarTypeInteger (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN AfxIsVarTypeSignedInteger(vt) OR AfxIsVarTypeUnsignedInteger(vt)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxIsVarTypeNumber (BYVAL vt AS VARTYPE) AS BOOLEAN
   RETURN (AfxIsVarTypeInteger(vt)) OR (AfxIsVarTypeFloat(vt))
END FUNCTION
' ========================================================================================

END NAMESPACE
