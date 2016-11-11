' ########################################################################################
' Microsoft Windows
' File: CServiceProvider.inc
' Contents: Implementation of the IServiceProvider interface
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ========================================================================================
' This function must be called by ordinal because it isn't exported by name in XP and below.
' ========================================================================================
PRIVATE FUNCTION AfxOleConQueryService ( _
   BYVAL punk AS IUnknown PTR _                           ' __in IUnknown* punk
 , BYVAL guidService AS REFGUID _                         ' __in REFGUID guidService
 , BYVAL riid AS REFGUID _                                ' __in REFIID riid
 , BYVAL ppvOut AS LPVOID PTR _                           ' __out void ** ppvOut
 ) AS HRESULT                                             ' HRESULT

   ' // Load the shell library
   DIM AS ANY PTR pLib = DyLibLoad("Shlwapi.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   ' // Get the address of the function (ordinal 176)
   DIM pAfxOleConQueryService AS FUNCTION (BYVAL punk AS IUnknown PTR, BYVAL guidService AS REFGUID, _
       BYVAL riid AS REFGUID, BYVAL ppvOut AS LPVOID PTR) AS HRESULT
   pAfxOleConQueryService = DyLibSymbol(pLib, 176)
   IF pAfxOleConQueryService = NULL THEN
      FUNCTION = E_FAIL
   ELSE
      FUNCTION = AfxOleConQueryService(pUnk, guidService, riid, ppvOut)
   END IF
   FreeLibrary pLib

END FUNCTION
' ========================================================================================

' ========================================================================================
' IServiceProvider interface implemented
' Provides a generic access mechanism to locate a GUID-identified service.
' ========================================================================================
TYPE Afx_IServiceProviderImpl EXTENDS Afx_IServiceProvider

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION QueryService (BYVAL guidService AS const GUID const PTR, BYVAL riid AS const IID const PTR, BYVAL ppvObject AS ANY PTR PTR) AS HRESULT

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
CONSTRUCTOR Afx_IServiceProviderImpl
   OC_DP("Afx_IServiceProviderImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IServiceProviderImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IServiceProviderImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IServiceProviderImpl
   OC_DP("Afx_IServiceProviderImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IServiceProviderImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IServiceProviderImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IServiceProviderImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Acts as the factory method for any services exposed through an implementation of
' IServiceProvider.
' ========================================================================================
FUNCTION Afx_IServiceProviderImpl.QueryService (BYVAL guidService AS const GUID const PTR, BYVAL riid AS const IID const PTR, BYVAL ppvObject AS ANY PTR PTR) AS HRESULT
   OC_DP("Afx_IServiceProviderImpl QueryService")
   ' // Handle SID_SInPlaceBrowser (needed to enable the filesystem object to navigate
   ' // in-place within the WebBrowser control when running on Windows 7)
   IF IsEqualIID(@guidService, @SID_SInPlaceBrowser) THEN
      RETURN AfxOleConQueryService(cast(IUnknown PTR, m_pData->m_pOcx), @SID_SShellBrowser, riid, @ppvObject)
   END If
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================
