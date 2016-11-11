' ########################################################################################
' Microsoft Windows
' File: COleCon.inc
' Contents: OLE Container
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#INCLUDE ONCE "win/ole2.bi"
#include once "win/unknwnbase.bi"
#include once "win/unknwn.bi"
#include once "win/olectl.bi"
#include once "win/docobj.bi"
#include once "win/shlguid.bi"
#include once "win/shobjidl.bi"
'#INCLUDE ONCE "vbinterf.inc"
#include once "win/ocidl.bi"

NAMESPACE Afx

CONST OC_CLASSNAME = "COleCon"   ' // Container's class name

' ========================================================================================
' Macro for debug
' To allow debugging, define _OC_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _OC_DEBUG_
   #define _OC_DEBUG_ 0
#endif
#ifndef _OC_DP_
   #define _OC_DP_ 1
   #MACRO OC_DP(st)
      #if (_OC_DEBUG_ = 1)
         OutputDebugStringW(st)
      #endif
   #ENDMACRO
#endif
' ========================================================================================

' // The definition for BSTR in the FreeBASIC headers was incorrectly changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

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
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif

#ifndef __Afx_IOleObject_INTERFACE_DEFINED__
#define __Afx_IOleObject_INTERFACE_DEFINED__
TYPE Afx_IOleObject AS Afx_IOleObject_
TYPE Afx_IOleObject_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION SetClientSite (BYVAL pClientSite AS IOleClientSite PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetClientSite (BYVAL ppClientSite AS IOleClientSite PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetHostNames (BYVAL szContainerApp AS LPCOLESTR, BYVAL szContainerObj AS LPCOLESTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Close (BYVAL dwSaveOption AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetMoniker (BYVAL dwWhichMoniker AS DWORD, BYVAL pmk AS IMoniker PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetMoniker (BYVAL dwAssign AS DWORD, BYVAL dwWhichMoniker AS DWORD, BYVAL ppmk AS IMoniker PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION InitFromData (BYVAL pDataObject AS IDataObject ptr, BYVAL fCreation AS WINBOOL, BYVAL dwReserved AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetClipboardData (BYVAL dwReserved AS DWORD, BYVAL ppDataObject AS IDataObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DoVerb (BYVAL iVerb AS LONG, BYVAL lpmsg AS LPMSG, BYVAL pActiveSite AS IOleClientSite PTR, BYVAL lindex AS LONG, BYVAL hwndParent AS HWND, BYVAL lprcPosRect AS LPCRECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnumVerbs (BYVAL ppEnumOleVerb AS IEnumOLEVERB PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Update () AS HRESULT
   DECLARE ABSTRACT FUNCTION IsUpToDate () AS HRESULT
   DECLARE ABSTRACT FUNCTION GetUserClassID (BYVAL pClsid AS CLSID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetUserType (BYVAL dwFormOfType AS DWORD, BYVAL pszUserType AS LPOLESTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetExtent (BYVAL dwDrawAspect AS DWORD, BYVAL psizel AS SIZEL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetExtent (BYVAL dwDrawAspect AS DWORD, BYVAL psizel AS SIZEL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Advise (BYVAL pAdvSink AS IAdviseSink PTR, BYVAL pdwConnection AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Unadvise (BYVAL dwConnection AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnumAdvise (BYVAL ppenumAdvise AS IEnumSTATDATA PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetMiscStatus (BYVAL dwAspect AS DWORD, BYVAL pdwStatus AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetColorScheme (BYVAL pLogpal AS LOGPALETTE PTR) AS HRESULT
END TYPE
TYPE AFX_LPOLEOBJECT AS Afx_IOleObject PTR
#endif

#ifndef __Afx_IOleInPlaceObject_INTERFACE_DEFINED__
#define __Afx_IOleInPlaceObject_INTERFACE_DEFINED__
TYPE Afx_IOleInPlaceObject AS Afx_IOleInPlaceObject_
TYPE Afx_IOleInPlaceObject_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION InPlaceDeactivate () AS HRESULT
   DECLARE ABSTRACT FUNCTION UIDeactivate () AS HRESULT
   DECLARE ABSTRACT FUNCTION SetObjectRects (BYVAL lprcPosRect AS LPCRECT, BYVAL lprcClipRect AS LPCRECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReactivateAndUndo () AS HRESULT
END TYPE
TYPE AFX_LPOLEINPLACEOBJECT AS Afx_IOleInPlaceObject PTR
#endif

#ifndef __Afx_IOleClientSite_INTERFACE_DEFINED__
#define __Afx_IOleClientSite_INTERFACE_DEFINED__
TYPE Afx_IOleClientSite AS Afx_IOleClientSite_
TYPE Afx_IOleClientSite_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION SaveObject () AS HRESULT
   DECLARE ABSTRACT FUNCTION GetMoniker (BYVAL dwAssign AS DWORD, BYVAL dwWhichMoniker AS DWORD, BYVAL ppmk AS IMoniker PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetContainer (BYVAL ppContainer AS IOleContainer PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowObject () AS HRESULT
   DECLARE ABSTRACT FUNCTION OnShowWindow (BYVAL fShow AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION RequestNewObjectLayout () AS HRESULT
END TYPE
TYPE AFX_LPOLECLIENTSITE AS Afx_IOleClientSite PTR
#endif

#ifndef __Afx_IOleInPlaceActiveObject_INTERFACE_DEFINED__
#define __Afx_IOleInPlaceActiveObject_INTERFACE_DEFINED__
TYPE Afx_IOleInPlaceActiveObject AS Afx_IOleInPlaceActiveObject_
TYPE Afx_IOleInPlaceActiveObject_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetActiveObject (BYVAL pActiveObject AS IOleInPlaceActiveObject PTR, BYVAL pszObjName AS LPCOLESTR) AS HRESULT
END TYPE
TYPE AFX_LPOLEINPLACEACTIVEOBJECT AS Afx_IOleInPlaceActiveObject PTR
#endif

#ifndef __Afx_IServiceProvider_INTERFACE_DEFINED__
#define __Afx_IServiceProvider_INTERFACE_DEFINED__
TYPE Afx_IServiceProvider AS Afx_IServiceProvider_
TYPE Afx_IServiceProvider_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION QueryService (BYVAL guidService AS const GUID const PTR, BYVAL riid AS const IID const PTR, BYVAL ppvObject AS ANY PTR PTR) AS HRESULT
END TYPE
TYPE AFX_LPSERVICEPROVIDER AS Afx_IServiceProvider PTR
#endif

#ifndef __Afx_IParseDisplayName_INTERFACE_DEFINED__
#define __Afx_IParseDisplayName_INTERFACE_DEFINED__
TYPE Afx_IParseDisplayName AS Afx_IParseDisplayName_
TYPE Afx_IParseDisplayName_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT
END TYPE
TYPE AFX_LPPARSEDISPLAYNAME AS Afx_IParseDisplayName PTR
#endif

#ifndef __Afx_IOleContainer_INTERFACE_DEFINED__
#define __Afx_IOleContainer_INTERFACE_DEFINED__
TYPE Afx_IOleContainer AS Afx_IOleContainer_
TYPE Afx_IOleContainer_ EXTENDS Afx_IParseDisplayName
   DECLARE ABSTRACT FUNCTION EnumObjects (BYVAL grfFlags AS DWORD, BYVAL ppenum AS IEnumUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION LockContainer (BYVAL fLock AS WINBOOL) AS HRESULT
END TYPE
TYPE AFX_LPOLECONTAINER AS Afx_IOleContainer PTR
#endif

#ifndef __Afx_IOleWindow_INTERFACE_DEFINED__
#define __Afx_IOleWindow_INTERFACE_DEFINED__
TYPE Afx_IOleWindow AS Afx_IOleWindow_
TYPE Afx_IOleWindow_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
END TYPE
TYPE AFX_LPOLEWINDOW AS Afx_IOleWindow PTR
#endif

#ifndef __Afx_IOleInPlaceSite_INTERFACE_DEFINED__
#define __Afx_IOleInPlaceSite_INTERFACE_DEFINED__
TYPE Afx_IOleInPlaceSite AS Afx_IOleInPlaceSite_
TYPE Afx_IOleInPlaceSite_ EXTENDS Afx_IOleWindow
	DECLARE ABSTRACT FUNCTION CanInPlaceActivate () AS HRESULT
	DECLARE ABSTRACT FUNCTION OnInPlaceActivate () AS HRESULT
	DECLARE ABSTRACT FUNCTION OnUIActivate () AS HRESULT
	DECLARE ABSTRACT FUNCTION GetWindowContext (BYVAL ppFrame AS IOleInPlaceFrame PTR PTR, BYVAL ppDoc as IOleInPlaceUIWindow PTR PTR, BYVAL lprcPosRect AS LPRECT, BYVAL lprcClipRect AS LPRECT, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
	DECLARE ABSTRACT FUNCTION Scroll (BYVAL scrollExtant AS SIZE) AS HRESULT
	DECLARE ABSTRACT FUNCTION OnUIDeactivate (BYVAL fUndoable AS WINBOOL) AS HRESULT
	DECLARE ABSTRACT FUNCTION OnInPlaceDeactivate () AS HRESULT
	DECLARE ABSTRACT FUNCTION DiscardUndoState () AS HRESULT
	DECLARE ABSTRACT FUNCTION DeactivateAndUndo () AS HRESULT
	DECLARE ABSTRACT FUNCTION OnPosRectChange (BYVAL lprcPosRect AS LPCRECT) AS HRESULT
END TYPE
TYPE AFX_LPOLEINPLACESITE AS Afx_IOleInPlaceSite PTR
#endif

#ifndef __Afx_IOleInPlaceSiteEx_INTERFACE_DEFINED__
#define __Afx_IOleInPlaceSiteEx_INTERFACE_DEFINED__
TYPE Afx_IOleInPlaceSiteEx AS Afx_IOleInPlaceSiteEx_
TYPE Afx_IOleInPlaceSiteEx_ EXTENDS Afx_IOleInPlaceSite
	DECLARE ABSTRACT FUNCTION OnInPlaceActivateEx (BYVAL pfNoRedraw AS WINBOOL PTR, BYVAL dwFlags AS DWORD) AS HRESULT
	DECLARE ABSTRACT FUNCTION OnInPlaceDeactivateEx (BYVAL fNoRedraw AS WINBOOL) AS HRESULT
	DECLARE ABSTRACT FUNCTION RequestUIActivate () AS HRESULT
END TYPE
TYPE AFX_LPOLEINPLACESITEEX AS Afx_IOleInPlaceSiteEx PTR
#endif

#ifndef __Afx_IOleInPlaceUIWindow_INTERFACE_DEFINED__
#define __Afx_IOleInPlaceUIWindow_INTERFACE_DEFINED__
TYPE Afx_IOleInPlaceUIWindow AS Afx_IOleInPlaceUIWindow_
TYPE Afx_IOleInPlaceUIWindow_ EXTENDS Afx_IOleWindow_
   DECLARE ABSTRACT FUNCTION GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL AS LPCOLESTR) AS HRESULT
END TYPE
TYPE AFX_LPOLEINPLACEUIWINDOW AS Afx_IOleInPlaceUIWindow PTR
#endif

#ifndef __Afx_IOleInPlaceFrame_INTERFACE_DEFINED__
#define __Afx_IOleInPlaceFrame_INTERFACE_DEFINED__
TYPE Afx_IOleInPlaceFrame AS Afx_IOleInPlaceFrame_
TYPE Afx_IOleInPlaceFrame_ EXTENDS Afx_IOleInPlaceUIWindow
   DECLARE ABSTRACT FUNCTION InsertMenus (BYVAL hmenuShared AS HMENU, BYVAL lpMenuWidths AS LPOLEMENUGROUPWIDTHS) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetMenu (BYVAL hmenuShared AS HMENU, BYVAL holemenu AS HOLEMENU, BYVAL hwndActiveObject AS HWND) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveMenus (BYVAL hmenuShared AS HMENU) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetStatusText (BYVAL pszStatusText AS LPCOLESTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnableModeless (BYVAL fEnable AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION TranslateAccelerator (BYVAL lpmsg AS LPMSG, BYVAL wID AS WORD) AS HRESULT
END TYPE
TYPE AFX_LPOLEINPLACEUIWINDOW AS Afx_IOleInPlaceUIWindow PTR
#endif

#ifndef __Afx_IOleControlSite_INTERFACE_DEFINED__
#define __Afx_IOleControlSite_INTERFACE_DEFINED__
TYPE Afx_IOleControlSite AS Afx_IOleControlSite_
TYPE Afx_IOleControlSite_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION OnControlInfoChanged () AS HRESULT
   DECLARE ABSTRACT FUNCTION LockInPlaceActive (BYVAL bLock AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetExtendedControl (BYVAL ppDisp AS IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION TransformCoords (BYVAL pptlHimetric AS POINTL PTR, BYVAL pptfContainer AS POINTF PTR, BYVAL dwFlags AS DWORD ) AS HRESULT
   DECLARE ABSTRACT FUNCTION TranslateAccelerator (BYVAL lpMsg As LPMSG, BYVAL dwModifiers As DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION OnFocus (BYVAL bGotFocus AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowPropertyFrame () AS HRESULT
END TYPE
TYPE AFX_LPOLECONTROLSITE AS Afx_IOleControlSite PTR
#endif

#ifndef __Afx_ISimpleFrameSite_INTERFACE_DEFINED__
#define __Afx_ISimpleFrameSite_INTERFACE_DEFINED__
TYPE Afx_ISimpleFrameSite AS Afx_ISimpleFrameSite_
TYPE Afx_ISimpleFrameSite_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION PreMessageFilter (BYVAL hWnd AS HWND, BYVAL msg AS UINT, BYVAL wp AS WPARAM, BYVAL lp AS LPARAM, BYVAL plResult AS LRESULT PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION PostMessageFilter (BYVAL hWnd AS HWND, BYVAL msg AS UINT, BYVAL wp AS WPARAM, BYVAL lp AS LPARAM, BYVAL plResult AS LRESULT PTR, BYVAL dwCookie AS DWORD) AS HRESULT
END TYPE
TYPE AFX_LPSIMPLEFRAMESITE AS Afx_ISimpleFrameSite PTR
#endif

END NAMESPACE

