' ########################################################################################
' Microsoft Windows
' File: COleInPlaceFrame.inc
' Contents: Implementation of the IOleInPlaceFrame interface
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ========================================================================================
' IOleWindow interface implemented
' The IOleWindow interface provides methods that allow an application to obtain the handle
' to the various windows that participate in in-place activation, and also to enter and
' exit context-sensitive help mode.
' ========================================================================================
TYPE Afx_IOleWindowImpl EXTENDS Afx_IOleWindow

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pData AS Afx_OleConDispatch PTR)
   DECLARE DESTRUCTOR

   cRef AS DWORD                       ' // Reference count
   m_pData AS Afx_OleConDispatch PTR   ' // Pointer to the Afx_OleConDispatch class

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR Afx_IOleWindowImpl
   OC_DP("Afx_IOleWindowImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IOleWindowImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IOleWindowImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IOleWindowImpl
   OC_DP("Afx_IOleWindowImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IOleWindowImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleWindowImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleWindowImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION Afx_IOleWindowImpl.GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pData->m_hwnd
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether context-sensitive help mode should be entered during an in-place
' activation session.
' ========================================================================================
FUNCTION Afx_IOleWindowImpl.ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' IOleInPlaceUIWindow interface implemented
' Implemented by container applications and used by object applications to negotiate border
' space on the document or frame window. The container provides a RECT structure in which
' the object can place toolbars and other similar controls, determines if tools can in fact
' be installed around the object's window frame, allocates space for the border, and
' establishes a communication channel between the object and each frame and document window.
' ========================================================================================
TYPE Afx_IOleInPlaceUIWindowImpl EXTENDS Afx_IOleInPlaceUIWindow

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   DECLARE VIRTUAL FUNCTION RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL pszObjName AS LPCOLESTR) AS HRESULT

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
CONSTRUCTOR Afx_IOleInPlaceUIWindowImpl
   OC_DP("Afx_IOleInPlaceUIWindowImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IOleInPlaceUIWindowImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IOleInPlaceUIWindowImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IOleInPlaceUIWindowImpl
   OC_DP("Afx_IOleInPlaceUIWindowImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pData->m_hwnd
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether context-sensitive help mode should be entered during an in-place
' activation session.
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies a RECT structure for toolbars and controls.
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   IF lprectBorder = NULL THEN RETURN E_INVALIDARG ELSE RETURN INPLACE_E_NOTOOLSPACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if tools can be installed around object's window frame.
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   IF pborderwidths = NULL THEN RETURN E_INVALIDARG ELSE RETURN INPLACE_E_NOTOOLSPACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allocates space for the border.
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides for direct communication between the object and each document and frame window.
' SetActiveObject will be called with both NULL and non-NULL pointers.
' ========================================================================================
FUNCTION Afx_IOleInPlaceUIWindowImpl.SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL pszObjName AS LPCOLESTR) AS HRESULT
   IF m_pData->m_pIOleInPlaceActiveObject <> NULL THEN
      DIM pIOleInPlaceActiveObject AS ANY PTR = m_pData->m_pIOleInPlaceActiveObject
      IUnknown_Release(cast(IUnknown PTR, pIOleInPlaceActiveObject))
      m_pData->m_pIOleInPlaceActiveObject = NULL
   END IF
   IF pActiveObject THEN
      IUnknown_AddRef(pActiveObject)
      m_pData->m_pIOleInPlaceActiveObject = cast(ANY PTR, pActiveObject)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' IOleInPlaceFrame interface implemented
' Controls the container's top-level frame window. This control involves allowing the
' container to insert its menu group into the composite menu, install the composite menu
' into the appropriate window frame, and remove the container's menu elements from the
' composite menu. It sets and displays status text relevant to the in-place object. It
' also enables or disables the frame's modeless dialog boxes, and translates accelerator
' keystrokes intended for the container's frame.
' ========================================================================================
TYPE Afx_IOleInPlaceFrameImpl EXTENDS Afx_IOleInPlaceFrame

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   DECLARE VIRTUAL FUNCTION RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL pszObjName AS LPCOLESTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION InsertMenus (BYVAL hmenuShared AS HMENU, BYVAL lpMenuWidths AS LPOLEMENUGROUPWIDTHS) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetMenu (BYVAL hmenuShared AS HMENU, BYVAL holemenu AS HOLEMENU, BYVAL hwndActiveObject AS HWND) AS HRESULT
   DECLARE VIRTUAL FUNCTION RemoveMenus (BYVAL hmenuShared AS HMENU) AS HRESULT
   DECLARE VIRTUAL FUNCTION SetStatusText (BYVAL pszStatusText AS LPCOLESTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION EnableModeless (BYVAL fEnable AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION TranslateAccelerator (BYVAL lpmsg AS LPMSG, BYVAL wID AS WORD) AS HRESULT

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
CONSTRUCTOR Afx_IOleInPlaceFrameImpl
   OC_DP("Afx_IOleInPlaceFrameImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IOleInPlaceFrameImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IOleInPlaceFrameImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IOleInPlaceFrameImpl
   OC_DP("Afx_IOleInPlaceFrameImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle to one of the windows participating in in-place activation
' (frame, document, parent, or in-place object window).
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.GetWindow (BYVAL phwnd AS HWND PTR) AS HRESULT
   IF phwnd = NULL THEN RETURN E_INVALIDARG
   *phwnd = m_pData->m_hwnd
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether context-sensitive help mode should be entered during an in-place
' activation session.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.ContextSensitiveHelp(BYVAL fEnterMode AS BOOL) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies a RECT structure for toolbars and controls.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.GetBorder (BYVAL lprectBorder AS LPRECT) AS HRESULT
   IF lprectBorder = NULL THEN RETURN E_INVALIDARG ELSE RETURN INPLACE_E_NOTOOLSPACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if tools can be installed around object's window frame.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.RequestBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   IF pborderwidths = NULL THEN RETURN E_INVALIDARG ELSE RETURN INPLACE_E_NOTOOLSPACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allocates space for the border.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.SetBorderSpace (BYVAL pborderwidths AS LPCBORDERWIDTHS) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides for direct communication between the object and each document and frame window.
' SetActiveObject will be called with both NULL and non-NULL pointers.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.SetActiveObject (BYVAL pActiveObject AS LPOLEINPLACEACTIVEOBJECT, BYVAL pszObjName AS LPCOLESTR) AS HRESULT
   IF m_pData->m_pIOleInPlaceActiveObject <> NULL THEN
      DIM pIOleInPlaceActiveObject AS ANY PTR = m_pData->m_pIOleInPlaceActiveObject
      IUnknown_Release(cast(IUnknown PTR, pIOleInPlaceActiveObject))
      m_pData->m_pIOleInPlaceActiveObject = NULL
   END IF
   IF pActiveObject THEN
      IUnknown_AddRef(pActiveObject)
      m_pData->m_pIOleInPlaceActiveObject = cast(ANY PTR, pActiveObject)
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Allows container to insert menus.
' Not required. This method is necessary only for containers with menu UI.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.InsertMenus (BYVAL hmenuShared AS HMENU, BYVAL lpMenuWidths AS LPOLEMENUGROUPWIDTHS) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a composite menu to window frame.
' Not required. This method is necessary only for containers with menu UI.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.SetMenu (BYVAL hmenuShared AS HMENU, BYVAL holemenu AS HOLEMENU, BYVAL hwndActiveObject AS HWND) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes a container's menu elements.
' Not required. This method is necessary only for containers with menu UI.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.RemoveMenus (BYVAL hmenuShared AS HMENU) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets and displays status text about.
' Not required. This method is necessary only for containers that have a status line.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.SetStatusText (BYVAL pszStatusText AS LPCOLESTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Optional. Enables or disables modeless dialog boxes.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.EnableModeless (BYVAL fEnable AS BOOL) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Optional. Translates keystrokes.
' ========================================================================================
FUNCTION Afx_IOleInPlaceFrameImpl.TranslateAccelerator (BYVAL lpmsg AS LPMSG, BYVAL wID AS WORD) AS HRESULT
   RETURN S_FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' IOleWindow interface implemented
' Provides simple frame controls that act as simple containers for other nested controls.
' Some controls merely contain other controls. In such cases, the simple control container,
' called a simple frame, need not implement all container requirements. It can delegate
' most of the interface calls from its contained controls to the outer container that
' manages the simple frame. To support what are called simple frame controls, a container
' implements this interface as well as other site interfaces such as IOleControlSite.
' An example of a simple frame control is a group box that only needs to capture a few
' keystrokes for its contained controls to implement the correct tab or arrow key behavior,
' but does not want to handle every other message. Through the two methods of this interface,
' the simple frame control passes messages to its control site both before and after its
' own processing. If that site is itself a simple frame, it can continue to pass messages
' up the chain.
' Implement this interface on a control object to support nested controls without requiring
' the control to itself be a full container.
' ========================================================================================
TYPE Afx_ISimpleFrameSiteImpl EXTENDS Afx_ISimpleFrameSite

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION PreMessageFilter (BYVAL hWnd AS HWND, BYVAL msg AS UINT, BYVAL wp AS WPARAM, BYVAL lp AS LPARAM, BYVAL plResult AS LRESULT PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION PostMessageFilter (BYVAL hWnd AS HWND, BYVAL msg AS UINT, BYVAL wp AS WPARAM, BYVAL lp AS LPARAM, BYVAL plResult AS LRESULT PTR, BYVAL dwCookie AS DWORD) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pData AS Afx_OleConDispatch PTR)
   DECLARE DESTRUCTOR

   cRef AS DWORD                       ' // Reference count
   m_pData AS Afx_OleConDispatch PTR   ' // Pointer to the Afx_OleConDispatch class

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR Afx_ISimpleFrameSiteImpl
   OC_DP("Afx_ISimpleFrameSiteImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_ISimpleFrameSiteImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_ISimpleFrameSiteImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_ISimpleFrameSiteImpl
   OC_DP("Afx_ISimpleFrameSiteImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_ISimpleFrameSiteImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_ISimpleFrameSiteImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_ISimpleFrameSiteImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides a site with the opportunity to process a message that is received by a
' control's own window before the control itself does any processing.
' ========================================================================================
FUNCTION Afx_ISimpleFrameSiteImpl.PreMessageFilter (BYVAL hWnd AS HWND, BYVAL msg AS UINT, BYVAL wp AS WPARAM, BYVAL lp AS LPARAM, BYVAL plResult AS LRESULT PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   IF plResult = NULL OR pdwCookie = NULL THEN RETURN E_POINTER
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends the simple frame site a message that is received by a control's own window after
' both ISimpleFrameSite.PreMessageFilter and the control have had a chance to process the
' message.
' ========================================================================================
FUNCTION Afx_ISimpleFrameSiteImpl.PostMessageFilter (ByVal hWnd as HWND, byval msg as UINT, byval wp as WPARAM, byval lp as LPARAM, byval plResult as LRESULT ptr, byval dwCookie as DWORD) as HRESULT
   IF plResult = NULL THEN RETURN E_POINTER
   RETURN S_FALSE
END FUNCTION
' ========================================================================================
