' ########################################################################################
' Microsoft Windows
' File: CAxHost_IOleInPlaceSite.inc
' Contents: Implementation of the IOleInPlaceSite interface
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
'                          *** CAxHost_IOleInPlaceSite class ***
' ########################################################################################

' ========================================================================================
' IOleInPlaceSite interface implemented
' Manages the interaction between the container and the object's in-place client site.
' Recall that the client site is the display site for embedded objects, and provides position
' and conceptual information about the object.
' This interface provides methods that manage in-place objects. With IOleInPlaceSite, you
' can determine if an object can be activated and manage its activation and deactivation.
' You can notify the container when one of its objects is being activated and inform the
' container that a composite menu will replace the container's regular menu. It provides
' methods that make it possible for the in-place object to retrieve the window object
' hierarchy, and the position in the parent window where the object should place its
' in-place activation window. Finally, it determines how the container scrolls the object,
' manages the object undo state, and notifies the object when its borders have changed.
' ========================================================================================
TYPE CAxHost_IOleInPlaceSite EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
	DECLARE VIRTUAL FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
	DECLARE VIRTUAL FUNCTION ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
	DECLARE VIRTUAL FUNCTION CanInPlaceActivate () AS HRESULT
	DECLARE VIRTUAL FUNCTION OnInPlaceActivate () AS HRESULT
	DECLARE VIRTUAL FUNCTION OnUIActivate () AS HRESULT
	DECLARE VIRTUAL FUNCTION GetWindowContext (BYVAL ppFrame AS IOleInPlaceFrame PTR PTR, BYVAL ppDoc as IOleInPlaceUIWindow PTR PTR, BYVAL lprcPosRect AS LPRECT, BYVAL lprcClipRect AS LPRECT, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
	DECLARE VIRTUAL FUNCTION Scroll (BYVAL scrollExtant AS SIZE) AS HRESULT
	DECLARE VIRTUAL FUNCTION OnUIDeactivate (BYVAL fUndoable AS WINBOOL) AS HRESULT
	DECLARE VIRTUAL FUNCTION OnInPlaceDeactivate () AS HRESULT
	DECLARE VIRTUAL FUNCTION DiscardUndoState () AS HRESULT
	DECLARE VIRTUAL FUNCTION DeactivateAndUndo () AS HRESULT
	DECLARE VIRTUAL FUNCTION OnPosRectChange (BYVAL lprcPosRect AS LPCRECT) AS HRESULT

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
CONSTRUCTOR CAxHost_IOleInPlaceSite
   CAXH_DP("CAxHost_IOleInPlaceSite Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IOleInPlaceSite (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IOleInPlaceSite constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IOleInPlaceSite
   CAXH_DP("### CAxHost_IOleInPlaceSite destructor")
   m_pAxHost->m_pIOleInPlaceSiteImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IOleInPlaceSite AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.Release () AS ULONG
   this.cRef -= 1
   CAXH_DP("CAxHost_IOleInPlaceSite Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.GetWindow(BYVAL phwnd AS HWND PTR) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite GetWindow")
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pAxHost->m_hwndContainer
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Optional. Determines whether context-sensitive help mode should be entered during an
' in-place activation session.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite ContextSensitiveHelp")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether or not the container can activate the object in place.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.CanInPlaceActivate () AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite CanInPlaceActivate")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that one of its objects is being activated in place.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.OnInPlaceActivate () AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite OnInPlaceActivate")
   m_pAxHost->m_bInPlaceActive = TRUE
   IF m_pAxHost->m_pIOleObject THEN
      OleLockRunning(CAST(IUnknown PTR, CAST(ULONG_PTR, m_pAxHost->m_pIOleObject)), TRUE, FALSE)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that the object is about to be activated in place and that the object
' is going to replace the container's main menu with an in-place composite menu.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.OnUIActivate () AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite OnUIActivate")
   m_pAxHost->m_bUIActive = TRUE
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the in-place object to retrieve the window interfaces that form the window object
' hierarchy, and the position in the parent window where the object's in-place activation
' window should be placed.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.GetWindowContext (BYVAL ppFrame AS IOleInPlaceFrame PTR PTR, BYVAL ppDoc as IOleInPlaceUIWindow PTR PTR, BYVAL lprcPosRect AS LPRECT, BYVAL lprcClipRect AS LPRECT, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite GetWindowContext")

   IF ppFrame = NULL OR ppDoc = NULL THEN RETURN E_INVALIDARG

   ' Required: We must provide our IOleInPlaceFrame, even if we aren't
   ' going to use any of its methods, or the object will be embedded
   ' but not in-place activated.
   IF m_pAxHost->m_pIOleInPlaceFrameImpl <> NULL THEN
      *ppFrame = cast(IOleInPlaceFrame PTR, cast(ANY PTR, m_pAxHost->m_pIOleInPlaceFrameImpl))
   ELSE
      m_pAxHost->QueryInterface(@IID_IOleInPlaceFrame, @m_pAxHost->m_pIOleInPlaceFrameImpl)
      *ppFrame = cast(IOleInPlaceFrame PTR, cast(ANY PTR, m_pAxHost->m_pIOleInPlaceFrameImpl))
   END IF
   *ppDoc = NULL

   IF lprcPosRect THEN
      GetClientRect m_pAxHost->m_hWndContainer, lprcPosRect
      GetClientRect m_pAxHost->m_hWndContainer, lprcClipRect
   END IF

   IF lpFrameInfo THEN
      lpFrameInfo->cb = SIZEOF(OLEINPLACEFRAMEINFO)
      lpFrameInfo->fMDIApp = FALSE
      lpFrameInfo->hWndFrame = m_pAxHost->m_hWndContainer
      lpFrameInfo->hAccel = 0
      lpFrameInfo->cAccelEntries = 0
   END IF

   RETURN S_OK

END FUNCTION
' ========================================================================================

' ========================================================================================
' Tells the container to scroll the view of the object by a specified number of pixels.
' Not required. This method can return S_FALSE and perform no action.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.Scroll (BYVAL scrollExtant AS SIZE) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite Scroll")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container on deactivation that it should reinstall its user interface and take
' focus, and whether or not the object has an undoable state.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.OnUIDeactivate (BYVAL fUndoable AS WINBOOL) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite OnUIDeactivate")
   m_pAxHost->m_bUIActive = FALSE
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that the object is no longer active in place.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.OnInPlaceDeactivate () AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite OnInPlaceDeactivate")
   m_pAxHost->m_bInPlaceActive = FALSE
   IF m_pAxHost->m_pIOleObject THEN
      OleLockRunning(cast(IUnknown PTR, CAST(ULONG_PTR, m_pAxHost->m_pIOleObject)), FALSE, FALSE)
   END IF
   AfxSafeRelease(m_pAxHost->m_pIOleInPlaceObject)
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tells the container that the object no longer has any undo state and that the container
' should not call IOleInPlaceObject::ReActivateAndUndo.
' Not required. This method can return S_OK and perform no action.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.DiscardUndoState () AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite DiscardUndoState")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes the container to end the in-place session, deactivate the object, and revert to its
' own saved undo state.
' Not required. This method must implement deactivation, but undo is optional.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.DeactivateAndUndo () AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite DeactivateAndUndo")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates the object's extents have changed.
' ========================================================================================
FUNCTION CAxHost_IOleInPlaceSite.OnPosRectChange (BYVAL lprcPosRect AS LPCRECT) AS HRESULT
   CAXH_DP("CAxHost_IOleInPlaceSite OnPosRectChange")
   IF lprcPosRect = NULL THEN RETURN E_INVALIDARG
   IF m_pAxHost->m_pIOleInPlaceObject THEN
      m_pAxHost->m_pIOleInPlaceObject->SetObjectRects(lprcPosRect, lprcPosRect)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================
