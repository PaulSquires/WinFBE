' ########################################################################################
' Microsoft Windows
' File: RichOleDemo2.bas
' Contents: Demonstrates how to load a rich edit text file from a resource file and how to
' implement the IRichEditOleCallback interface, which allows to display embedded images.
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxRichEdit.inc"
#INCLUDE ONCE "Afx/AfxCOM.inc"
USING Afx

CONST IDC_RICHEDIT = 1001

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Implementation of the IRichEditOleCallback interface
' ========================================================================================

#ifndef __Afx_IRichEditOleCallback_INTERFACE_DEFINED__
#define __Afx_IRichEditOleCallback_INTERFACE_DEFINED__
TYPE Afx_IRichEditOleCallback AS Afx_IRichEditOleCallback_
TYPE Afx_IRichEditOleCallback_ EXTENDS Afx_IUnknown
	DECLARE ABSTRACT FUNCTION GetNewStorage (BYVAL lplpstg AS LPSTORAGE PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION GetInPlaceContext (BYVAL lplpFrame AS LPOLEINPLACEFRAME PTR, BYVAL lplpDoc AS LPOLEINPLACEUIWINDOW PTR, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
	DECLARE ABSTRACT FUNCTION ShowContainerUI (BYVAL fShow AS WINBOOL) AS HRESULT
	DECLARE ABSTRACT FUNCTION QueryInsertObject (BYVAL lpclsid AS LPCLSID, BYVAL lpstg AS LPSTORAGE, BYVAL cp AS LONG) AS HRESULT
	DECLARE ABSTRACT FUNCTION DeleteObject (BYVAL lpoleobj AS LPOLEOBJECT) AS HRESULT
	DECLARE ABSTRACT FUNCTION QueryAcceptData (BYVAL lpdataobj AS LPDATAOBJECT, BYVAL lpcfFormat AS CLIPFORMAT PTR, BYVAL reco AS DWORD, BYVAL fReally AS WINBOOL, BYVAL hMetaPict AS HGLOBAL) AS HRESULT
	DECLARE ABSTRACT FUNCTION ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
	DECLARE ABSTRACT FUNCTION GetClipboardData (BYVAL lpchrg AS CHARRANGE PTR, BYVAL reco AS DWORD, BYVAL lplpdataobj AS LPDATAOBJECT PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION GetDragDropEffect (BYVAL fDrag AS WINBOOL, BYVAL grfKeyState AS DWORD, BYVAL pdwEffect AS LPDWORD) AS HRESULT
	DECLARE ABSTRACT FUNCTION GetContextMenu (BYVAL seltype AS WORD, BYVAL lpoleobj AS LPOLEOBJECT, BYVAL lpchrg AS CHARRANGE PTR, BYVAL lphmenu AS HMENU PTR) AS HRESULT
END TYPE
TYPE AFX_LPRICHEDITOLECALLBACK AS Afx_IRichEditOleCallback PTR
#endif

' ========================================================================================
TYPE Afx_IRichEditOleCallbackImpl EXTENDS Afx_IRichEditOleCallback

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL AS REFIID, BYVAL AS PVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULong
   DECLARE VIRTUAL FUNCTION GetNewStorage (BYVAL lplpstg AS LPSTORAGE PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetInPlaceContext (BYVAL lplpFrame AS LPOLEINPLACEFRAME PTR, BYVAL lplpDoc AS LPOLEINPLACEUIWINDOW PTR, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
   DECLARE VIRTUAL FUNCTION ShowContainerUI (BYVAL fShow AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION QueryInsertObject (BYVAL lpclsid AS LPCLSID, BYVAL lpstg AS LPSTORAGE, BYVAL cp AS LONG) AS HRESULT
   DECLARE VIRTUAL FUNCTION DeleteObject (BYVAL lpoleobj AS LPOLEOBJECT) AS HRESULT
   DECLARE VIRTUAL FUNCTION QueryAcceptData (BYVAL lpdataobj AS LPDATAOBJECT, BYVAL lpcfFormat AS CLIPFORMAT PTR, BYVAL reco AS DWORD, BYVAL fReally AS WINBOOL, BYVAL hMetaPict AS HGLOBAL) AS HRESULT
   DECLARE VIRTUAL FUNCTION ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetClipboardData (BYVAL lpchrg AS CHARRANGE PTR, BYVAL reco AS DWORD, BYVAL lplpdataobj AS LPDATAOBJECT PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetDragDropEffect (BYVAL fDrag AS WINBOOL, BYVAL grfKeyState AS DWORD, BYVAL pdwEffect AS LPDWORD) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetContextMenu (BYVAL seltype AS WORD, BYVAL lpoleobj AS LPOLEOBJECT, BYVAL lpchrg AS CHARRANGE PTR, BYVAL lphmenu AS HMENU PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   cRef AS DWORD                       ' // Reference count

END TYPE
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
CONSTRUCTOR Afx_IRichEditOleCallbackImpl
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR Afx_IRichEditOleCallbackImpl
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   DIM IID_IRichEditOleCallback_ AS GUID = (&h00020D03, &h0000, &h0000, {&hC0, &h00, &h00, &h00, &h00, &h00, &h00, &h46})
   IF IsEqualIID(riid, @IID_IRichEditOleCallback_) OR IsEqualIID(riid, @IID_IUnknown) OR IsEqualIID(riid, @IID_IDispatch) THEN
      *ppvObj = @this
      ' // Not really needed, since this is not a COM object
      cast(Afx_IUnknown PTR, *ppvObj)->AddRef
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.AddRef () AS ULONG
   RETURN 2
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.Release () AS ULONG
   RETURN 1
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.GetNewStorage (BYVAL lplpstg AS LPSTORAGE PTR) AS HRESULT
   DIM hr AS HRESULT, pILockBytes AS ILockBytes PTR
   hr = CreateILockBytesOnHGlobal(NULL, TRUE, @pILockBytes)
   IF FAILED(hr) THEN RETURN hr
   hr = StgCreateDocfileOnILockBytes(pILockBytes, _
        STGM_SHARE_EXCLUSIVE OR STGM_READWRITE OR STGM_CREATE, 0, lplpstg)
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.GetInPlaceContext (BYVAL lplpFrame AS LPOLEINPLACEFRAME PTR, BYVAL lplpDoc AS LPOLEINPLACEUIWINDOW PTR, BYVAL lpFrameInfo AS LPOLEINPLACEFRAMEINFO) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.ShowContainerUI (BYVAL fShow AS WINBOOL) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.QueryInsertObject (BYVAL lpclsid AS LPCLSID, BYVAL lpstg AS LPSTORAGE, BYVAL cp AS LONG) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.DeleteObject (BYVAL lpoleobj AS LPOLEOBJECT) AS HRESULT
   RETURN S_OK
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.QueryAcceptData (BYVAL lpdataobj AS LPDATAOBJECT, BYVAL lpcfFormat AS CLIPFORMAT PTR, BYVAL reco AS DWORD, BYVAL fReally AS WINBOOL, BYVAL hMetaPict AS HGLOBAL) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.ContextSensitiveHelp (BYVAL fEnterMode AS WINBOOL) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.GetClipboardData (BYVAL lpchrg AS CHARRANGE PTR, BYVAL reco AS DWORD, BYVAL lplpdataobj AS LPDATAOBJECT PTR) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.GetDragDropEffect (BYVAL fDrag AS WINBOOL, BYVAL grfKeyState AS DWORD, BYVAL pdwEffect AS LPDWORD) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================
' ========================================================================================
FUNCTION Afx_IRichEditOleCallbackImpl.GetContextMenu (BYVAL seltype AS WORD, BYVAL lpoleobj AS LPOLEOBJECT, BYVAL lpchrg AS CHARRANGE PTR, BYVAL lphmenu AS HMENU PTR) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use  manifest
'   AfxSetProcessDPIAware

   ' // Initialize the COM library
   CoInitialize NULL

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "Rich Ole Demo", @WndProc)
   ' // Set the client size
   pWindow.SetClientSize(750, 450)
   ' // Center the window
   pWindow.Center

   ' // Add a rich edit control without coordinates (it will be resized in WM_SIZE)
   DIM hRichEdit AS HWND = pWindow.AddControl("RichEdit", , IDC_RICHEDIT, "RichEdit box", , , , , _
   WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR WS_VSCROLL OR ES_MULTILINE OR ES_WANTRETURN OR ES_NOHIDESEL OR ES_SAVESEL)

   ' // Raise the 32,767 default characters limit to no limit
   SendMessageW(hRichEdit, EM_EXLIMITTEXT, 0, -1)
   ' // Set the left margin
   SendMessageW(hRichEdit, EM_SETMARGINS, EC_LEFTMARGIN, 20)
   ' // Specify which notifications the control sends to its parent window
   SendMessageW(hRichEdit, EM_SETEVENTMASK, 0, ENM_CHANGE)

   ' // Set the IRichEditOleCallback object.
   ' // The control calls the AddRef function for the object before returning.
   DIM pRichEditOleCallback AS Afx_IRichEditOleCallbackImpl PTR = NEW Afx_IRichEditOleCallbackImpl
   SendMessageW(hRichEdit, EM_SETOLECALLBACK, 0, cast(LPARAM, pRichEditOleCallback))

   ' // Load the file
   RichEdit_LoadRtfFromResourceW hRichEdit, hInstance, "RTFFILE"

   ' // Set the focus in the control
   SetFocus hRichEdit

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Delete the class
   IF pRichEditOleCallback THEN Delete pRichEditOleCallback

   ' // Uninitialize the COM library
   CoUninitialize

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            ' // If ESC key pressed, close the application sending an WM_CLOSE message
            CASE IDCANCEL
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_SIZE
         ' // If the window isn't minimized, resize it
         IF wParam <> SIZE_MINIMIZED THEN
            ' // Resize the Rich Edit control
            DIM hRichEdit AS HWND = GetDlgItem(hwnd, IDC_RICHEDIT)
            MoveWindow hRichEdit, 10, 10, LOWORD(lParam) - 20, HIWORD(lParam) - 20, CTRUE
         END IF

    	CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
