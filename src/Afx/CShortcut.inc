' ########################################################################################
' Microsoft Windows
' File: CShorcut.inc
' Contents: Windows shortcuts
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxWsh.bi"
#include once "Afx/AfxCOM.inc"
#include once "Afx/CWSTR.inc"

NAMESPACE Afx

' ========================================================================================
' CShortcut class
' ========================================================================================
TYPE CShortcut

   Private:
      m_bUninitCOM AS BOOLEAN
      m_Result AS HRESULT              ' // HRESULT code
      m_pLnk AS Afx_IWshShortcut PTR   ' // IWshShortcut interface

   Public:
      DECLARE CONSTRUCTOR (BYREF cbsPathName AS CBSTR)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE PROPERTY FullName () AS CBSTR
      DECLARE PROPERTY Arguments () AS CBSTR
      DECLARE PROPERTY Arguments (BYREF cbsArguments AS CBSTR)
      DECLARE PROPERTY Description () AS CBSTR
      DECLARE PROPERTY Description (BYREF cbsDescription AS CBSTR)
      DECLARE PROPERTY Hotkey () AS CBSTR
      DECLARE PROPERTY Hotkey (BYREF cbsHotkey AS CBSTR)
      DECLARE PROPERTY IconLocation () AS CBSTR
      DECLARE PROPERTY IconLocation (BYREF cbsIconLocation AS CBSTR)
      DECLARE PROPERTY RelativePath (BYREF cbsRelativePath AS CBSTR)
      DECLARE PROPERTY TargetPath () AS CBSTR
      DECLARE PROPERTY TargetPath (BYREF cbsTargetPath AS CBSTR)
      DECLARE PROPERTY WindowStyle () AS INT_
      DECLARE PROPERTY WindowStyle (BYVAL nstyle AS INT_)
      DECLARE PROPERTY WorkingDirectory () AS CBSTR
      DECLARE PROPERTY WorkingDirectory (BYREF cbsWorkingDirectory AS CBSTR)
      DECLARE FUNCTION Save () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CURLShortcut class
' ========================================================================================
TYPE CURLShortcut

   Private:
      m_bUninitCOM AS BOOLEAN
      m_Result AS HRESULT                 ' // HRESULT code
      m_pLnk AS Afx_IWshURLShortcut PTR   ' // IWshURLShortcut interface

   Public:
      DECLARE CONSTRUCTOR (BYREF cbsPathName AS CBSTR)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE PROPERTY FullName () AS CBSTR
      DECLARE PROPERTY TargetPath () AS CBSTR
      DECLARE PROPERTY TargetPath (BYREF cbsTargetPath AS CBSTR)
      DECLARE FUNCTION Save () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CShortcut constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CShortcut (BYREF cbsPathName AS CBSTR)

   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE

   ' // Create an instance of the IWshShortcut interface
   DIM pWsh AS Afx_IWshShell PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell_ AS GUID = (&hF935DC21, &h1CF0, &h11D0, {&hAD, &hB9, &h00, &hC0, &h4F, &hD5, &h8A, &h0B})
   m_Result = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell_, @pWsh)
   IF pWsh = NULL THEN m_Result = E_NOINTERFACE
   IF pWsh THEN
      m_Result = pWsh->CreateShortcut(cbsPathName, cast(Afx_IDIspatch PTR PTR, @m_pLnk))
      pWsh->Release
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CShortcut destructor
' ========================================================================================
PRIVATE DESTRUCTOR CShortcut
   ' // Release the interface
   IF m_pLnk THEN m_pLnk->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CShortcut.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the fully qualified path to the shortcut's target.
' ========================================================================================
PRIVATE PROPERTY CShortcut.FullName () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrFullName AS AFX_BSTR
   m_Result = m_pLnk->get_FullName(@bstrFullName)
   PROPERTY = bstrFullName
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets/sets the arguments for a shortcut, or identifies a shortcut's arguments.
' ========================================================================================
PRIVATE PROPERTY CShortcut.Arguments () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrArguments AS AFX_BSTR
   m_Result = m_pLnk->get_Arguments(@bstrArguments)
   PROPERTY = bstrArguments
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.Arguments (BYREF cbsArguments AS CBSTR)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_Arguments(cbsArguments)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets/sets the shortcut description.
' ========================================================================================
PRIVATE PROPERTY CShortcut.Description () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrDescription AS AFX_BSTR
   m_Result = m_pLnk->get_Description(@bstrDescription)
   PROPERTY = bstrDescription
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.Description (BYREF cbsDescription AS CBSTR)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_Description(cbsDescription)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a key-combination to a shortcut, or identifies the key-combination assigned to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.Hotkey () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrHotkey AS AFX_BSTR
   m_Result = m_pLnk->get_Hotkey(@bstrHotkey)
   PROPERTY = bstrHotkey
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.Hotkey (BYREF cbsHotkey AS CBSTR)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_Hotkey(cbsHotkey)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.IconLocation () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrIconLocation AS AFX_BSTR
   m_Result = m_pLnk->get_IconLocation(@bstrIconLocation)
   PROPERTY = bstrIconLocation
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.IconLocation (BYREF cbsIconLocation AS CBSTR)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_IconLocation(cbsIconLocation)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a relative path to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.RelativePath (BYREF cbsRelativePath AS CBSTR)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_RelativePath(cbsRelativePath)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.TargetPath () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrTargetPath AS AFX_BSTR
   m_Result = m_pLnk->get_TargetPath(@bstrTargetPath)
   PROPERTY = bstrTargetPath
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.TargetPath (BYREF cbsTargetPath AS CBSTR)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_TargetPath(cbsTargetPath)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a window style to a shortcut, or identifies the type of window style used by a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.WindowStyle () AS INT_
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM nStyle AS INT_
   m_Result = m_pLnk->get_Windowstyle(@nStyle)
   PROPERTY = nStyle
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.WindowStyle (BYVAL nstyle AS INT_)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_Windowstyle(nStyle)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a working directory to a shortcut, or identifies the working directory used by a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.WorkingDirectory () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrWorkingDirectory AS AFX_BSTR
   m_Result = m_pLnk->get_WorkingDirectory(@bstrWorkingDirectory)
   PROPERTY = bstrWorkingDirectory
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.WorkingDirectory (BYREF cbsWorkingDirectory AS CBSTR)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_WorkingDirectory(cbsWorkingDirectory)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Saves a shortcut object to disk.
' ========================================================================================
PRIVATE FUNCTION CShortcut.Save () AS HRESULT
   IF m_pLnk = NULL THEN m_Result = E_POINTER : RETURN m_Result
   m_Result = m_pLnk->Save
END FUNCTION
' ========================================================================================

' ========================================================================================
' CURLShortcut constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CURLShortcut (BYREF cbsPathName AS CBSTR)

   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE

   ' // Create an instance of the IWshURLShortcut interface
   DIM pWsh AS Afx_IWshShell PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell_ AS GUID = (&hF935DC21, &h1CF0, &h11D0, {&hAD, &hB9, &h00, &hC0, &h4F, &hD5, &h8A, &h0B})
   m_Result = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell_, @pWsh)
   IF pWsh = NULL THEN m_Result = E_NOINTERFACE
   IF pWsh THEN
      m_Result = pWsh->CreateShortcut(cbsPathName, cast(Afx_IDIspatch PTR PTR, @m_pLnk))
      pWsh->Release
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CURLShortcut destructor
' ========================================================================================
PRIVATE DESTRUCTOR CURLShortcut
   ' // Release the interface
   IF m_pLnk THEN m_pLnk->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CURLShortcut.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the fully qualified path to the shortcut's target.
' ========================================================================================
PRIVATE PROPERTY CURLShortcut.FullName () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrFullName AS AFX_BSTR
   m_Result = m_pLnk->get_FullName(@bstrFullName)
   PROPERTY = bstrFullName
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CURLShortcut.TargetPath () AS CBSTR
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrTargetPath AS AFX_BSTR
   m_Result = m_pLnk->get_TargetPath(@bstrTargetPath)
   PROPERTY = bstrTargetPath
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CURLShortcut.TargetPath (BYREF cbsTargetPath AS CBSTR)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_TargetPath(cbsTargetPath)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Saves a shortcut object to disk.
' ========================================================================================
PRIVATE FUNCTION CURLShortcut.Save () AS HRESULT
   IF m_pLnk = NULL THEN m_Result = E_POINTER : RETURN m_Result
   m_Result = m_pLnk->Save
END FUNCTION
' ========================================================================================

END NAMESPACE
