' ########################################################################################
' Microsoft Windows
' File: COleConSite.inc
' Contents: Implementation of the IOleClientSite interface
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ========================================================================================
' IOleClientSite interface implemented
' Provides the primary means by which an embedded object obtains information about the
' location and extent of its display site, its moniker, its user interface, and other
' resources provided by its container. An object server calls IOleClientSite to request
' services from the container. A container must provide one instance of IOleClientSite for
' every compound-document object it contains.
' Note: This interface is not supported for use across machine boundaries.
' ========================================================================================
TYPE Afx_IOleClientSiteImpl EXTENDS Afx_IOleClientSite

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION SaveObject () AS HRESULT
   DECLARE VIRTUAL FUNCTION GetMoniker (BYVAL AS DWORD, BYVAL AS DWORD, BYVAL AS LPMONIKER PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetContainer (BYVAL AS LPOLECONTAINER PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ShowObject () AS HRESULT
   DECLARE VIRTUAL FUNCTION OnShowWindow (BYVAL AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION RequestNewObjectLayout () AS HRESULT

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
CONSTRUCTOR Afx_IOleClientSiteImpl
   OC_DP("Afx_IOleClientSiteImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IOleClientSiteImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IOleClientSiteImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IOleClientSiteImpl
   OC_DP("Afx_IOleClientSiteImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves embedded object.
' Not required. This method is necessary only if persistence is supported.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.SaveObject () AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Requests object's moniker.
' Not required. This method is necessary only if the container supports linking to
' controls within its own form or document.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.GetMoniker (BYVAL dwAssign AS DWORD, BYVAL dwWhichMoniker AS DWORD, BYVAL ppmk AS LPMONIKER PTR) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Requests pointer to object's container.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.GetContainer (BYVAL ppContainer AS LPOLECONTAINER PTR) AS HRESULT
   IF ppContainer = NULL THEN RETURN E_INVALIDARG
   RETURN m_pdata->QueryInterface(@IID_IOleContainer, ppContainer)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Asks container to display object.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.ShowObject () AS HRESULT
   ShowWindow m_pData->m_hwnd, SW_SHOW
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies container when object becomes visible or invisible.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.OnShowWindow (BYVAL fShow AS BOOL) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Asks container to resize display site.
' ========================================================================================
FUNCTION Afx_IOleClientSiteImpl.RequestNewObjectLayout () AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
