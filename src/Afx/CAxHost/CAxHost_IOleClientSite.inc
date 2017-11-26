' ########################################################################################
' Microsoft Windows
' File: CAxHost_IOleClientSite.inc
' Contents: OLE Container class
' Implementation of the IOleClientSite interface.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' this CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
'                          *** CAxHost_IOleClientSite class ***
' ########################################################################################

' ========================================================================================
' IOleClientSite interface implemented
' Provides the primary means by which an embedded object obtains information about the
' location and extent of its display site, its moniker, its user interface, and other
' resources provided by its container. An object server calls IOleClientSite to request
' services from the container. A container must provide one instance of IOleClientSite for
' every compound-document object it contains.
' Note: This interface is not supported for use across machine boundaries.
' ========================================================================================
TYPE CAxHost_IOleClientSite EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION SaveObject () AS HRESULT
   DECLARE VIRTUAL FUNCTION GetMoniker (BYVAL AS DWORD, BYVAL AS DWORD, BYVAL AS LPMONIKER PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetContainer (BYVAL AS LPOLECONTAINER PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ShowObject () AS HRESULT
   DECLARE VIRTUAL FUNCTION OnShowWindow (BYVAL AS BOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION RequestNewObjectLayout () AS HRESULT

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
CONSTRUCTOR CAxHost_IOleClientSite
   CAXH_DP("CAxHost_IOleClientSite Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IOleClientSite (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IOleClientSite constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IOleClientSite
   CAXH_DP("### CAxHost_IOleClientSite destructor")
   m_pAxHost->m_pIOleClientSiteImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IOleClientSite AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IOleClientSite Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves embedded object.
' Not required. This method is necessary only if persistence is supported.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.SaveObject () AS HRESULT
   CAXH_DP("CAxHost_IOleClientSite SaveObject")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Requests object's moniker.
' Not required. This method is necessary only if the container supports linking to
' controls within its own form or document.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.GetMoniker (BYVAL dwAssign AS DWORD, BYVAL dwWhichMoniker AS DWORD, BYVAL ppmk AS LPMONIKER PTR) AS HRESULT
   CAXH_DP("CAxHost_IOleClientSite GetMoniker")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Requests pointer to object's container.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.GetContainer (BYVAL ppContainer AS LPOLECONTAINER PTR) AS HRESULT
   CAXH_DP("CAxHost_IOleClientSite GetContainer")
   IF ppContainer = NULL THEN RETURN E_INVALIDARG
   *ppContainer = NULL
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Asks container to display object.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.ShowObject () AS HRESULT
   CAXH_DP("CAxHost_IOleClientSite ShowObject")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies container when object becomes visible or invisible.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.OnShowWindow (BYVAL fShow AS BOOL) AS HRESULT
   CAXH_DP("CAxHost_IOleClientSite OnShowWindow")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Asks container to resize display site.
' ========================================================================================
FUNCTION CAxHost_IOleClientSite.RequestNewObjectLayout () AS HRESULT
   CAXH_DP("CAxHost_IOleClientSite RequestNewObjectLayout")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
