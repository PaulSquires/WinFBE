' ########################################################################################
' Microsoft Windows
' File: COleContainer.inc
' Contents: Implementation of the IOleContainer interface
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ========================================================================================
' IParseDisplayName interface implemented
' Parses a displayable name string to convert it into a moniker for custom moniker implementations.
' Display name parsing is necessary when the end user inputs a string to identify a component,
' as in the following situations:
' - A compound document application that supports linked components typically supports the
'   Edit.Links... dialog box. Through this dialog box, the end user can enter a display name
'   to specify a new link source for a specified linked component. The compound document needs
'   to have this input string converted into a moniker.
' - A script language such as the macro language of a spreadsheet can allow textual references
'   to a component. The language's interpreter needs to have such a reference converted into
'    a moniker in order to execute the macro.
' This interface is not supported for use across machine boundaries.
' ========================================================================================
TYPE Afx_IParseDisplayNameImpl EXTENDS Afx_IParseDisplayName

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT

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
CONSTRUCTOR Afx_IParseDisplayNameImpl
   OC_DP("Afx_IParseDisplayNameImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IParseDisplayNameImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IParseDisplayNameImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IParseDisplayNameImpl
   OC_DP("Afx_IParseDisplayNameImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IParseDisplayNameImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IParseDisplayNameImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IParseDisplayNameImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses the display name returning a moniker corresponding to it.
' Not required. This method is only necessary for containers that link to controls or
' other embedded objects.
' ========================================================================================
FUNCTION Afx_IParseDisplayNameImpl.ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ########################################################################################
' Ole Container implementation of the IParseDisplayName interface.
' This interface, as well as support for nested simple frames is optional.
' The IParseDisplayName interface parses a displayable name string to convert it into a
' moniker for custom moniker implementations. Display name parsing is necessary when the
' end user inputs a string to identify a component, as in the following situations:
' * A compound document application that supports linked components typically supports
'   the Edit:Links... dialog box. Through this dialog box, the end user can enter a
'   display name to specify a new link source for a specified linked component. The
'   compound document needs to have this input string converted into a moniker.
' * A script language such as the macro language of a spreadsheet can allow textual
'   references to a component. The language's interpreter needs to have such a reference
'   converted into a moniker in order to execute the macro.
' ########################################################################################

' ========================================================================================
' IOleContainer interface implemented
' ========================================================================================
TYPE Afx_IOleContainerImpl EXTENDS Afx_IOleContainer

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION EnumObjects (BYVAL grfFlags AS DWORD, BYVAL ppenum AS IEnumUnknown PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION LockContainer (BYVAL fLock AS WINBOOL) AS HRESULT

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
CONSTRUCTOR Afx_IOleContainerImpl
   OC_DP("Afx_IOleContainerImpl CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR Afx_IOleContainerImpl (BYVAL pData AS Afx_OleConDispatch PTR)
   OC_DP("Afx_IOleContainerImpl CONSTRUCTOR - pData = " & WSTR(pData))
   m_pData = pData
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IOleContainerImpl
   OC_DP("Afx_IOleClientSiteImpl DESTRUCTOR")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IOleContainerImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   RETURN m_pdata->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleContainerImpl.AddRef () AS ULONG
   RETURN m_pdata->AddRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_IOleContainerImpl.Release () AS ULONG
   RETURN m_pdata->Release
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses the display name returning a moniker corresponding to it.
' Not required. This method is only necessary for containers that link to controls or
' other embedded objects.
' ========================================================================================
FUNCTION Afx_IOleContainerImpl.ParseDisplayName (BYVAL pbc AS IBindCtx PTR, BYVAL pszDisplayName AS LPOLESTR, BYVAL pchEaten AS ULONG PTR, BYVAL ppmkOut AS IMoniker PTR PTR) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enumerates objects in a container.
' Not required. This method is only necessary to enumerate ActiveX controls.
' ========================================================================================
FUNCTION Afx_IOleContainerImpl.EnumObjects (BYVAL grfFlags AS DWORD, BYVAL ppenum AS IEnumUnknown PTR PTR) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Keeps container running until explicitly released.
' Not required. This method is only necessary for containers that link to controls or
' other embedded objects.
' ========================================================================================
FUNCTION Afx_IOleContainerImpl.LockContainer (BYVAL fLock AS WINBOOL) AS HRESULT
   m_pData->m_bLocked = fLock
   FUNCTION = S_OK
END FUNCTION
' ========================================================================================
