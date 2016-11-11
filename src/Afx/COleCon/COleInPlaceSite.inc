' ########################################################################################
' Microsoft Windows
' File: COleInPlaceSite.inc
' Contents: Implementation of the IOleInPlaceSite interface
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

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
TYPE Afx_IOleInPlaceSiteImpl EXTENDS Afx_IOleInPlaceSite

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
   DECLARE CONSTRUCTOR (BYVAL pData AS Afx_OleConDispatch PTR)
   DECLARE DESTRUCTOR

   cRef AS DWORD                       ' // Reference count
   m_pData AS Afx_OleConDispatch PTR   ' // Pointer to the Afx_OleConDispatch object

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR Afx_IOleInPlaceSiteImpl
   OC_DP("Afx_IOleInPlaceSiteImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IOleInPlaceSiteImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IOleInPlaceSiteImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IOleInPlaceSiteImpl
   OC_DP("Afx_IOleInPlaceSiteImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.GetWindow(BYVAL phwnd AS HWND PTR) AS HRESULT
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pData->m_hwnd
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Optional. Determines whether context-sensitive help mode should be entered during an
' in-place activation session.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether or not the container can activate the object in place.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.CanInPlaceActivate () AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that one of its objects is being activated in place.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.OnInPlaceActivate () AS HRESULT
   m_pData->m_bInPlaceActive = TRUE
   IF m_pData->m_pIOleObject THEN
      DIM pIOleObject AS ANY PTR = m_pData->m_pIOleObject
      OleLockRunning(pIOleObject, TRUE, FALSE)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that the object is about to be activated in place and that the object
' is going to replace the container's main menu with an in-place composite menu.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.OnUIActivate () AS HRESULT
   m_pData->m_bUIActive = TRUE
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the in-place object to retrieve the window interfaces that form the window object
' hierarchy, and the position in the parent window where the object's in-place activation
' window should be placed.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.GetWindowContext (BYVAL ppFrame AS IOleInPlaceFrame PTR PTR, BYVAL ppDoc as IOleInPlaceUIWindow PTR PTR, BYVAL lprcPosRect AS LPRECT, BYVAL lprcClipRect AS LPRECT, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT

   IF ppFrame = NULL OR ppDoc = NULL THEN RETURN E_INVALIDARG

   ' Required: We must provide our IOleInPlaceFrame, even if we aren't
   ' going to use any of its methods, or the object will be embedded
   ' but not in-place activated.
   IF m_pData->m_pIOleInPlaceFrameImpl <> NULL THEN
      DIM pIOleInPlaceFrame AS ANY PTR = m_pData->m_pIOleInPlaceFrameImpl
      *ppFrame = cast(IOleInPlaceFrame PTR, pIOleInPlaceFrame)
   END IF

   IF m_pData->m_pIOleInPlaceUIWindowImpl <> NULL THEN
      DIM pIOleInPlaceUIWindow AS ANY PTR = m_pData->m_pIOleInPlaceUIWindowImpl
      *ppDoc = cast(IOleInPlaceUIWindow PTR, pIOleInPlaceUIWindow)
   END IF

   IF lprcPosRect THEN
      GetClientRect m_pData->m_hWnd, lprcPosRect
      GetClientRect m_pData->m_hWnd, lprcClipRect
   END IF

   IF lpFrameInfo THEN
      lpFrameInfo->cb = SIZEOF(OLEINPLACEFRAMEINFO)
      lpFrameInfo->fMDIApp = FALSE
      lpFrameInfo->hWndFrame = m_pData->m_hWnd
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
FUNCTION Afx_IOleInPlaceSiteImpl.Scroll (BYVAL scrollExtant AS SIZE) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container on deactivation that it should reinstall its user interface and take
' focus, and whether or not the object has an undoable state.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.OnUIDeactivate (BYVAL fUndoable AS WINBOOL) AS HRESULT
   m_pData->m_bUIActive = FALSE
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that the object is no longer active in place.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.OnInPlaceDeactivate () AS HRESULT
   m_pData->m_bInPlaceActive = FALSE
   IF m_pData->m_pIOleObject THEN
      DIM pIOleInPlaceObject AS ANY PTR = m_pData->m_pIOleInPlaceObject
      OleLockRunning(cast(IUnknown PTR, pIOleInPlaceObject), FALSE, FALSE)
   END IF
   IF m_pData->m_pIOleInPlaceObject THEN
      DIM pIOleInPlaceObject AS ANY PTR = m_pData->m_pIOleInPlaceObject
      IUnknown_Release(cast(IUnknown PTR, pIOleInPlaceObject))
      m_pData->m_pIOleInPlaceObject = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tells the container that the object no longer has any undo state and that the container
' should not call IOleInPlaceObject::ReActivateAndUndo.
' Not required. This method can return S_OK and perform no action.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.DiscardUndoState () AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes the container to end the in-place session, deactivate the object, and revert to its
' own saved undo state.
' Not required. This method must implement deactivation, but undo is optional.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.DeactivateAndUndo () AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates the object's extents have changed.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteImpl.OnPosRectChange (BYVAL lprcPosRect AS LPCRECT) AS HRESULT
   IF lprcPosRect = NULL THEN RETURN E_INVALIDARG
   IF m_pData->m_pIOleInPlaceObject THEN
      m_pData->m_pIOleInPlaceObject->SetObjectRects(lprcPosRect, lprcPosRect)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' IOleInPlaceSiteEx interface implemented
' Ole Container implementation of the IOleInPlaceSiteEx interface.
' The IOleInPlaceSiteEx interface provides an additional set of activation and
' deactivation notification methods that enable an object to avoid unnecessary flashing on
' the screen when the object is activated and deactivated.
' When an object is activated, it does not know if its visual display is already correct.
' When the object is deactivated, the container does not know if the visual display is
' correct. To avoid a redraw and the associated screen flicker in both cases, the
' container can provide this extension to IOleInPlaceSite.
' ========================================================================================
TYPE Afx_IOleInPlaceSiteExImpl EXTENDS Afx_IOleInPlaceSiteEx

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
	DECLARE VIRTUAL FUNCTION OnInPlaceActivateEx (BYVAL pfNoRedraw AS WINBOOL PTR, BYVAL dwFlags AS DWORD) AS HRESULT
	DECLARE VIRTUAL FUNCTION OnInPlaceDeactivateEx (BYVAL fNoRedraw AS WINBOOL) AS HRESULT
	DECLARE VIRTUAL FUNCTION RequestUIActivate () AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pData AS Afx_OleConDispatch PTR)
   DECLARE DESTRUCTOR

   cRef AS DWORD                       ' // Reference count
   m_pData AS Afx_OleConDispatch PTR   ' // Pointer to the Afx_OleConDispatch object

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR Afx_IOleInPlaceSiteExImpl
   OC_DP("Afx_IOleInPlaceSiteExImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IOleInPlaceSiteExImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IOleInPlaceSiteExImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IOleInPlaceSiteExImpl
   OC_DP("Afx_IOleInPlaceSiteExImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.GetWindow(BYVAL phwnd AS HWND PTR) AS HRESULT
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pData->m_hwnd
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Optional. Determines whether context-sensitive help mode should be entered during an
' in-place activation session.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether or not the container can activate the object in place.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.CanInPlaceActivate () AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that one of its objects is being activated in place.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.OnInPlaceActivate () AS HRESULT
   m_pData->m_bInPlaceActive = TRUE
   IF m_pData->m_pIOleObject THEN
      DIM pIOleObject AS ANY PTR = m_pData->m_pIOleObject
      OleLockRunning(pIOleObject, TRUE, FALSE)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that the object is about to be activated in place and that the object
' is going to replace the container's main menu with an in-place composite menu.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.OnUIActivate () AS HRESULT
   m_pData->m_bUIActive = TRUE
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the in-place object to retrieve the window interfaces that form the window object
' hierarchy, and the position in the parent window where the object's in-place activation
' window should be placed.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.GetWindowContext (BYVAL ppFrame AS IOleInPlaceFrame PTR PTR, BYVAL ppDoc as IOleInPlaceUIWindow PTR PTR, BYVAL lprcPosRect AS LPRECT, BYVAL lprcClipRect AS LPRECT, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT

   IF ppFrame = NULL OR ppDoc = NULL THEN RETURN E_INVALIDARG

   ' Required: We must provide our IOleInPlaceFrame, even if we aren't
   ' going to use any of its methods, or the object will be embedded
   ' but not in-place activated.
   IF m_pData->m_pIOleInPlaceFrameImpl <> NULL THEN
      DIM pIOleInPlaceFrame AS ANY PTR = m_pData->m_pIOleInPlaceFrameImpl
      *ppFrame = cast(IOleInPlaceFrame PTR, pIOleInPlaceFrame)
   END IF

   IF m_pData->m_pIOleInPlaceUIWindowImpl <> NULL THEN
      DIM pIOleInPlaceUIWindow AS ANY PTR = m_pData->m_pIOleInPlaceUIWindowImpl
      *ppDoc = cast(IOleInPlaceUIWindow PTR, pIOleInPlaceUIWindow)
   END IF

   IF lprcPosRect THEN
      GetClientRect m_pData->m_hWnd, lprcPosRect
      GetClientRect m_pData->m_hWnd, lprcClipRect
   END IF

   IF lpFrameInfo THEN
      lpFrameInfo->cb = SIZEOF(OLEINPLACEFRAMEINFO)
      lpFrameInfo->fMDIApp = FALSE
      lpFrameInfo->hWndFrame = m_pData->m_hWnd
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
FUNCTION Afx_IOleInPlaceSiteExImpl.Scroll (BYVAL scrollExtant AS SIZE) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container on deactivation that it should reinstall its user interface and take
' focus, and whether or not the object has an undoable state.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.OnUIDeactivate (BYVAL fUndoable AS WINBOOL) AS HRESULT
   m_pData->m_bUIActive = FALSE
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that the object is no longer active in place.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.OnInPlaceDeactivate () AS HRESULT
   m_pData->m_bInPlaceActive = FALSE
   IF m_pData->m_pIOleObject THEN
      DIM pIOleInPlaceObject AS ANY PTR = m_pData->m_pIOleInPlaceObject
      OleLockRunning(cast(IUnknown PTR, pIOleInPlaceObject), FALSE, FALSE)
   END IF
   IF m_pData->m_pIOleInPlaceObject THEN
      DIM pIOleInPlaceObject AS ANY PTR = m_pData->m_pIOleInPlaceObject
      IUnknown_Release(cast(IUnknown PTR, pIOleInPlaceObject))
      m_pData->m_pIOleInPlaceObject = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Tells the container that the object no longer has any undo state and that the container
' should not call IOleInPlaceObject::ReActivateAndUndo.
' Not required. This method can return S_OK and perform no action.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.DiscardUndoState () AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes the container to end the in-place session, deactivate the object, and revert to its
' own saved undo state.
' Not required. This method must implement deactivation, but undo is optional.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.DeactivateAndUndo () AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates the object's extents have changed.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.OnPosRectChange (BYVAL lprcPosRect AS LPCRECT) AS HRESULT
   IF lprcPosRect = NULL THEN RETURN E_INVALIDARG
   IF m_pData->m_pIOleInPlaceObject THEN
      m_pData->m_pIOleInPlaceObject->SetObjectRects(lprcPosRect, lprcPosRect)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Called by the embedded object to determine if it needs to redraw itself upon activation.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.OnInPlaceActivateEx (BYVAL pfNoRedraw AS WINBOOL PTR, BYVAL dwFlags AS DWORD) AS HRESULT
   m_pData->m_bInPlaceActive = TRUE
   IF m_pData->m_pIOleObject THEN
      DIM pIOleObject AS ANY PTR = m_pData->m_pIOleObject
      OleLockRunning(pIOleObject, TRUE, FALSE)
   END If
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container of whether the object needs to be redrawn upon deactivation.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.OnInPlaceDeactivateEx ( BYVAL pfNoRedraw AS LONG) AS HRESULT
   IF m_pData->m_pIOleObject THEN
      DIM pIOleObject AS ANY PTR = m_pData->m_pIOleObject
      OleLockRunning(pIOleObject, FALSE, FALSE)
   END IF
   IF m_pData->m_pIOleInPlaceObject THEN
      DIM pIOleInPlaceObject AS ANY PTR = m_pData->m_pIOleInPlaceObject
      IUnknown_Release(cast(IUnknown PTR, pIOleInPlaceObject))
      m_pData->m_pIOleInPlaceObject = NULL
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies the container that the object is about to enter the UI-active state.
' ========================================================================================
FUNCTION Afx_IOleInPlaceSiteExImpl.RequestUIActivate () AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================
