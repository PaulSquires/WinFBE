' ########################################################################################
' Microsoft Windows
' File: AfxRegExp.bi
' Library name: VBScript_RegExp_55
' Documentation string: Microsoft VBScript Regular Expressions 5.5
' GUID: {3F4DACA7-160D-11D2-A8E9-00104B365C9F}
' Version: 5.5, Locale ID = 0
' Path: C:\Windows\SysWOW64\vbscript.dll\3
' Attributes: 8 [&h00000008]  [HasDiskImage]
' ########################################################################################

#pragma once
#include once "win/control.bi."

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' // ProgIDs (Program Identifiers)
CONST AFX_PROGID_VBScript_RegExp = "VBScript.RegExp"

' // ClsIDs (Class identifiers)
CONST AFX_CLSID_Match = "{3F4DACA5-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_CLSID_MatchCollection = "{3F4DACA6-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_CLSID_RegExp = "{3F4DACA4-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_CLSID_SubMatches = "{3F4DACC0-160D-11D2-A8E9-00104B365C9F}"

' // IIDs (Interface identifiers)
CONST AFX_IID_IMatch = "{3F4DACA1-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_IID_IMatch2 = "{3F4DACB1-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_IID_IMatchCollection = "{3F4DACA2-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_IID_IMatchCollection2 = "{3F4DACB2-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_IID_IRegExp = "{3F4DACA0-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_IID_IRegExp2 = "{3F4DACB0-160D-11D2-A8E9-00104B365C9F}"
CONST AFX_IID_ISubMatches = "{3F4DACB3-160D-11D2-A8E9-00104B365C9F}"

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
TYPE Afx_IMatch AS Afx_IMatch_
TYPE Afx_IMatch2 AS Afx_IMatch2_
TYPE Afx_IMatchCollection AS Afx_IMatchCollection_
TYPE Afx_IMatchCollection2 AS Afx_IMatchCollection2_
TYPE Afx_IRegExp AS Afx_IRegExp_
TYPE Afx_IRegExp2 AS Afx_IRegExp2_
TYPE Afx_ISubMatches AS Afx_ISubMatches_

' // Dual interfaces

' ########################################################################################
' Interface name: IMatch
' IID: {3F4DACA1-160D-11D2-A8E9-00104B365C9F}
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_IMatch_INTERFACE_DEFINED__
#define __Afx_IMatch_INTERFACE_DEFINED__

TYPE Afx_IMatch_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL pValue AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FirstIndex (BYVAL pFirstIndex AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Length (BYVAL pLength AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IMatch2
' IID: {3F4DACB1-160D-11D2-A8E9-00104B365C9F}
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 4
' ########################################################################################

#ifndef __Afx_IMatch2_INTERFACE_DEFINED__
#define __Afx_IMatch2_INTERFACE_DEFINED__

TYPE Afx_IMatch2_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL pValue AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FirstIndex (BYVAL pFirstIndex AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Length (BYVAL pLength AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SubMatches (BYVAL ppSubMatches AS Afx_IDispatch PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IMatchCollection
' IID: {3F4DACA2-160D-11D2-A8E9-00104B365C9F}
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_IMatchCollection_INTERFACE_DEFINED__
#define __Afx_IMatchCollection_INTERFACE_DEFINED__

TYPE Afx_IMatchCollection_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL index AS LONG, BYVAL ppMatch AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL pCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL ppEnum AS Afx_IUnknown PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IMatchCollection2
' IID: {3F4DACB2-160D-11D2-A8E9-00104B365C9F}
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_IMatchCollection2_INTERFACE_DEFINED__
#define __Afx_IMatchCollection2_INTERFACE_DEFINED__

TYPE Afx_IMatchCollection2_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL index AS LONG, BYVAL ppMatch AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL pCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL ppEnum AS Afx_IUnknown PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IRegExp
' IID: {3F4DACA0-160D-11D2-A8E9-00104B365C9F}
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 9
' ########################################################################################

#ifndef __Afx_IRegExp_INTERFACE_DEFINED__
#define __Afx_IRegExp_INTERFACE_DEFINED__

TYPE Afx_IRegExp_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Pattern (BYVAL pPattern AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Pattern (BYVAL pPattern AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IgnoreCase (BYVAL pIgnoreCase AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IgnoreCase (BYVAL pIgnoreCase AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Global (BYVAL pGlobal AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Global (BYVAL pGlobal AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION Execute (BYVAL sourceString AS AFX_BSTR, BYVAL ppMatches AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Test (BYVAL sourceString AS AFX_BSTR, BYVAL pMatch AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Replace (BYVAL sourceString AS AFX_BSTR, BYVAL replaceString AS AFX_BSTR, BYVAL pDestString AS AFX_BSTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: IRegExp2
' IID: {3F4DACB0-160D-11D2-A8E9-00104B365C9F}
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 11
' ########################################################################################

#ifndef __Afx_IRegExp2_INTERFACE_DEFINED__
#define __Afx_IRegExp2_INTERFACE_DEFINED__

TYPE Afx_IRegExp2_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Pattern (BYVAL pPattern AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Pattern (BYVAL pPattern AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IgnoreCase (BYVAL pIgnoreCase AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IgnoreCase (BYVAL pIgnoreCase AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Global (BYVAL pGlobal AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Global (BYVAL pGlobal AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Multiline (BYVAL pMultiline AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Multiline (BYVAL pMultiline AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION Execute (BYVAL sourceString AS AFX_BSTR, BYVAL ppMatches AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Test (BYVAL sourceString AS AFX_BSTR, BYVAL pMatch AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Replace (BYVAL sourceString AS AFX_BSTR, BYVAL replaceVar AS VARIANT, BYVAL pDestString AS AFX_BSTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISubMatches
' IID: {3F4DACB3-160D-11D2-A8E9-00104B365C9F}
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_ISubMatches_INTERFACE_DEFINED__
#define __Afx_ISubMatches_INTERFACE_DEFINED__

TYPE Afx_ISubMatches_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL index AS LONG, BYVAL pSubMatch AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL pCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL ppEnum AS Afx_IUnknown PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

END NAMESPACE
