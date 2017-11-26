' ########################################################################################
' Microsoft Windows
' File: CAxHost_IOleControlSite.inc
' Contents: OLE Container class
' Implementation of the IOleControlSite interface.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' this CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
'                          *** CAxHost_IOleControlSite class ***
' ########################################################################################

' ========================================================================================
' IOleControlSite interface implemented
' Provides the methods that enable a site object to manage each embedded control within a
' container. A site object provides IOleControlSite as well as other site interfaces such
' as IOleClientSite and IOleInPlaceSite. When a control requires the services expressed
' through this interface, it will query one of the other client site interfaces for IOleControlSite.
' ========================================================================================
TYPE CAxHost_IOleControlSite EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION OnControlInfoChanged () AS HRESULT
   DECLARE VIRTUAL FUNCTION LockInPlaceActive (BYVAL bLock AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetExtendedControl (BYVAL ppDisp AS IDispatch PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION TransformCoords (BYVAL pptlHimetric AS POINTL PTR, BYVAL pptfContainer AS POINTF PTR, BYVAL dwFlags AS DWORD ) AS HRESULT
   DECLARE VIRTUAL FUNCTION TranslateAccelerator (BYVAL lpMsg As LPMSG, BYVAL dwModifiers As DWORD) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnFocus (BYVAL bGotFocus AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION ShowPropertyFrame () AS HRESULT

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pAxHost AS CAxHost PTR)
   DECLARE DESTRUCTOR

Private:
   cRef AS ULONG = 0
   m_pAxHost AS CAxHost PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CAxHost_IOleControlSite
   CAXH_DP("CAxHost_IOleControlSite Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IOleControlSite (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IOleControlSite constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IOleControlSite
   CAXH_DP("### CAxHost_IOleControlSite destructor")
   m_pAxHost->m_pIOleControlSiteImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IOleControlSite AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IOleControlSite Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Informs the container that the control's CONTROLINFO structure has changed and that the
' container should call the control's IOleControl.GetControlInfo for an update.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.OnControlInfoChanged () AS HRESULT
   FUNCTION = S_OK ' Returned in all cases
END FUNCTION
' ========================================================================================

' ========================================================================================
' Optional. Indicates whether or not this control should remain in-place active, regardless
' of possible deactivation events.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.LockInPlaceActive (BYVAL fLock AS LONG) AS HRESULT
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Requests an IDispatch pointer to the extended control that the container uses to wrap
' the real control.
' Not required. This method is only necessary for containers that support extended controls.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.GetExtendedControl (BYVAL ppDisp AS IDispatch PTR PTR) AS HRESULT
   IF ppDisp = NULL THEN RETURN E_POINTER
   *ppDisp = NULL
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts between a POINTL structure expressed in HIMETRIC units (as is standard in OLE)
' and a POINTF structure expressed in units the container specifies.
' HIMETRIC_PER_INCH = 2540   ' 1 inch = 2.54 mm
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.TransformCoords (BYVAL pPtlHimetric AS POINTL PTR, BYVAL pPtfContainer AS POINTF PTR, BYVAL dwFlags AS DWORD) AS HRESULT
   DIM hDC AS HDC = GetDC(NULL)
   IF hDC = NULL THEN RETURN E_NOTIMPL
   IF (dwFlags AND XFORMCOORDS_HIMETRICTOCONTAINER) THEN
      pPtfContainer->x = MulDiv(pPtlHimetric->x, GetDeviceCaps(hDC, LOGPIXELSX), 2540)
      pPtfContainer->y = MulDiv(pPtlHimetric->y, GetDeviceCaps(hDC, LOGPIXELSY), 2540)
   ELSEIF (dwFlags AND XFORMCOORDS_CONTAINERTOHIMETRIC) THEN
      pPtlHimetric->x = MulDiv(pPtfContainer->x, 2540, GetDeviceCaps(hDC, LOGPIXELSX))
      pPtlHimetric->y = MulDiv(pPtfContainer->y, 2540, GetDeviceCaps(hDC, LOGPIXELSY))
   END IF
   ReleaseDC(NULL, hDC)
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Instructs the container to process a specified keystroke.
' Not required. This method can return S_FALSE and perform no action.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.TranslateAccelerator (BYVAL pMsg AS LPMSG, BYVAL grfModifiers AS DWORD) AS HRESULT
   FUNCTION = S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether the embedded control in this control site has gained or lost the focus.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.OnFocus (BYVAL fGotFocus AS LONG) AS HRESULT
   RETURN S_OK ' Returned in all cases
END FUNCTION
' ========================================================================================

' ========================================================================================
' Instructs the container to show a property page frame for the control object if the
' container so desires.
' Not required. This method is necessary only for containers that want to include their
' own property pages to handle extended control properties in addition to those provided
' by a control.
' ========================================================================================
FUNCTION CAxHost_IOleControlSite.ShowPropertyFrame () AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
