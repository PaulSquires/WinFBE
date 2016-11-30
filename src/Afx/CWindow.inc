' ########################################################################################
' File: CWindow.inc
' Version: 1.0
' Release candidate 26
' Contents: A wrapper class to create a SDK main window and add controls to it.
' Operating system: Microsoft Windows
' Compiler: FreeBasic 32 & 64-bit, Unicode.
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

#include once "windows.bi"
#include once "Afx/AfxCtl.inc"
#include once "Afx/AfxWin.inc"

NAMESPACE Afx

' ========================================================================================
' CScrollWindow class
' ========================================================================================
TYPE CScrollWindow

   Private:
      m_hwnd AS HWND             ' // Handle of the window
      m_ClientRect AS RECT       ' // Coordinates of the window's client area
      m_nHorzInc AS LONG         ' // Horizontal increment
      m_nVertInc AS LONG         ' // Vertical increment
      m_HScrollMax AS LONG       ' // Maximum horizontal scroll value
      m_VScrollMax AS LONG       ' // Maximum vertical scroll value
      m_HScrollPos AS LONG       ' // Horizontal scroll position
      m_VScrollPos AS LONG       ' // Vertical scroll position
      m_HorzUnits AS LONG        ' // Amount, in device units, of horizontal scrolling
      m_VertUnits AS LONG        ' // Amount, in device units, of vertical scrolling

   Public:
      DECLARE CONSTRUCTOR (BYVAL hwnd AS HWND)
      DECLARE DESTRUCTOR
      DECLARE SUB SetupScrollbars
      DECLARE SUB ResetScrollbars
      DECLARE SUB SetClientRect (BYVAL pRect AS RECT PTR)
      DECLARE SUB SetScrollingUnits (BYVAL nHorzUnits AS LONG, BYVAL nVertUnits AS LONG)
      DECLARE SUB OnSize (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB OnVScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB OnHScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB OnMouseWheel (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

END TYPE
' ========================================================================================

' ========================================================================================
' CWindow class
' ========================================================================================
TYPE CWindow

   Private:
      m_hwnd AS HWND                                  ' // Window handle
      m_hInstance AS HINSTANCE                        ' // Instance handle
      m_hFont AS HFONT                                ' // Default font handle
      m_hAccel AS HACCEL                              ' // Accelerator table handle
      m_DPI AS LONG = 96                              ' // Design-time DPI
      m_rx AS SINGLE = 1                              ' // Horizontal scaling ratio
      m_ry AS SINGLE = 1                              ' // Vertical scaling ratio
      m_wAtom AS ATOM                                 ' // Class atom
      m_wszClassName AS WSTRING * 260                 ' // Class name
      m_DefaultFontSize AS LONG                       ' // Default font size
      m_wszDefaultFontName AS WSTRING * LF_FACESIZE   ' // Default font name
      DIM m_rgUserData(0 TO 99) AS LONG_PTR
      DIM m_rgAccelEntries(ANY) AS ACCEL
      #ifdef USEMDI
      DIM m_hwndClient AS HWND                        ' // MDI window handle
      DIM m_wszMDIClassName AS WSTRING * 260          ' // MDI Class name
      #endif
      m_hRichEditLib AS HMODULE                       ' // Rich Edit moudle handle
      m_pScrollWindow AS CScrollWindow PTR            ' // Pointer to the scroll window class

   Public:
      DECLARE CONSTRUCTOR (BYREF wszClassName AS CONST WSTRING = "")
      DECLARE DESTRUCTOR
      DECLARE PROPERTY DPI () AS SINGLE
      DECLARE PROPERTY DPI (BYVAL dpi AS SINGLE)
      DECLARE PROPERTY UserData (BYVAL idx AS LONG) AS LONG_PTR
      DECLARE PROPERTY UserData (BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)
      DECLARE FUNCTION DoEvents (BYVAL nCmdShow AS LONG = 0) AS LONG
      DECLARE FUNCTION Create (BYVAL hParent AS HWND = NULL, BYREF wszTitle AS WSTRING = "", BYVAL lpfnWndProc AS WNDPROC = NULL, _
         BYVAL x AS LONG = CW_USEDEFAULT, BYVAL y AS LONG = CW_USEDEFAULT, BYVAL nWidth AS LONG = CW_USEDEFAULT, BYVAL nHeight AS LONG = CW_USEDEFAULT, _
         BYVAL dwStyle AS DWORD = WS_OVERLAPPEDWINDOW OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS, _
         BYVAL dwExStyle AS DWORD = WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE) AS HWND
      DECLARE SUB SetClientSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB Center (BYVAL hwnd AS HWND = NULL, BYVAL hwndParent AS HWND = NULL)
      DECLARE PROPERTY hWindow () AS HWND
      DECLARE PROPERTY hWindow (BYVAL hwnd AS HWND)
      DECLARE PROPERTY InstanceHandle () AS HINSTANCE
      DECLARE PROPERTY InstanceHandle (BYVAL hInst AS HINSTANCE)
      DECLARE PROPERTY ClassStyle () AS ULONG_PTR
      DECLARE PROPERTY ClassStyle (BYVAL dwStyle AS ULONG_PTR)
      DECLARE PROPERTY WindowStyle () AS ULONG_PTR
      DECLARE PROPERTY WindowStyle (BYVAL dwStyle AS ULONG_PTR)
      DECLARE PROPERTY WindowExStyle () AS ULONG_PTR
      DECLARE PROPERTY WindowExStyle (BYVAL dwExStyle AS ULONG_PTR)
      DECLARE PROPERTY DefaultFontName () AS STRING
      DECLARE PROPERTY DefaultFontName (BYREF wszFontName AS WSTRING)
      DECLARE PROPERTY DefaultFontSize () AS LONG
      DECLARE PROPERTY DefaultFontSize (BYVAL nPointSize AS LONG)
      DECLARE PROPERTY Font () AS HFONT
      DECLARE PROPERTY Font (BYVAL hFont AS HFONT)
      DECLARE PROPERTY Brush () AS HBRUSH
      DECLARE PROPERTY Brush (BYVAL hbrBackground AS HBRUSH)
      DECLARE PROPERTY BigIcon (BYVAL hIcon AS HICON)
      DECLARE PROPERTY SmallIcon (BYVAL hIcon AS HICON)
      DECLARE PROPERTY rxRatio () AS SINGLE
      DECLARE PROPERTY rxRatio (BYVAL rx AS SINGLE)
      DECLARE PROPERTY ryRatio (BYVAL rx AS SINGLE)
      DECLARE PROPERTY ryRatio () AS SINGLE
      DECLARE FUNCTION ScaleX (BYVAL cx AS SINGLE) AS SINGLE
      DECLARE FUNCTION ScaleY (BYVAL cy AS SINGLE) AS SINGLE
      DECLARE FUNCTION UnScaleX (BYVAL cx AS SINGLE) AS SINGLE
      DECLARE FUNCTION UnScaleY (BYVAL cy AS SINGLE) AS SINGLE
      DECLARE PROPERTY Width () AS LONG
      DECLARE PROPERTY Height () AS LONG
      DECLARE PROPERTY ClientHeight () AS LONG
      DECLARE PROPERTY ClientWidth () AS LONG
      DECLARE PROPERTY ControlWidth (BYVAL hwnd AS HWND) AS LONG
      DECLARE PROPERTY ControlHeight (BYVAL hwnd AS HWND) AS LONG
      DECLARE PROPERTY ControlClientWidth (BYVAL hwnd AS HWND) AS LONG
      DECLARE PROPERTY ControlClientHeight (BYVAL hwnd AS HWND) AS LONG
      DECLARE SUB GetWindowRect OVERLOAD (BYVAL lpRect AS LPRECT)
      DECLARE FUNCTION GetWindowRect OVERLOAD () AS RECT
      DECLARE SUB GetClientRect OVERLOAD (BYVAL lpRect AS LPRECT)
      DECLARE FUNCTION GetClientRect OVERLOAD () AS RECT
      DECLARE SUB GetControlWindowRect OVERLOAD (BYVAL hwnd AS HWND, BYVAL lpRect AS LPRECT)
      DECLARE FUNCTION GetControlWindowRect OVERLOAD (BYVAL hwnd AS HWND) AS RECT
      DECLARE SUB GetControlClientRect OVERLOAD (BYVAL hwnd AS HWND, BYVAL lpRect AS LPRECT)
      DECLARE FUNCTION GetControlClientRect OVERLOAD (BYVAL hwnd AS HWND) AS RECT
      DECLARE SUB GetWorkArea OVERLOAD (BYVAL lpRect AS LPRECT)
      DECLARE FUNCTION GetWorkArea OVERLOAD () AS RECT
      DECLARE FUNCTION ControlHandle (BYVAL cID AS LONG) AS HWND
      DECLARE PROPERTY ScrollWindowPtr () AS CScrollWindow PTR
      DECLARE PROPERTY ScrollWindowPtr (BYVAL pScrollWindow AS CScrollWindow PTR)
      DECLARE SUB Resize ()
      DECLARE FUNCTION SetWindowPos (BYVAL hwnd AS HWND, BYVAL hwndInsertAfter AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL uFlags AS UINT) AS BOOLEAN
      DECLARE FUNCTION MoveWindow (BYVAL hwnd AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL bRepaint AS BOOLEAN) AS BOOLEAN
      DECLARE FUNCTION CreateFont (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, BYVAL lWeight AS LONG = 0, _
              BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS HFONT
      DECLARE FUNCTION SetFont (BYREF wszFaceName AS WSTRING, BYVAL lPointSize AS LONG, BYVAL lWeight AS LONG = 0, _
              BYVAL bItalic AS UBYTE = FALSE, BYVAL bUnderline AS UBYTE = FALSE, BYVAL bStrikeOut AS UBYTE = FALSE, BYVAL bCharSet AS UBYTE = DEFAULT_CHARSET) AS BOOLEAN
      DECLARE FUNCTION AddControl (BYREF wszClassName AS WSTRING, BYVAL hParent AS HWND = NULL, BYVAL cID AS INTEGER = 0, _
              BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
              BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS LONG = -1, BYVAL dwExStyle AS LONG = -1, BYVAL lpParam AS LONG_PTR = 0, _
              BYVAL pWndProc AS WNDPROC = NULL, BYVAL uIdSubclass AS UINT_PTR = &HFFFFFFFF, BYVAL dwRefData AS DWORD_PTR = NULL) AS HWND
      DECLARE PROPERTY AccelHandle () AS HACCEL
      DECLARE PROPERTY AccelHandle (BYVAL hAccel AS HACCEL)
      DECLARE SUB AddAccelerator OVERLOAD (BYVAL fvirt AS UBYTE, BYVAL wKey AS WORD, BYVAL cmd AS WORD)
      DECLARE SUB AddAccelerator OVERLOAD (BYVAL fvirt AS UBYTE, BYREF wszKey AS WSTRING, BYVAL cmd AS WORD)
      DECLARE FUNCTION CreateAcceleratorTable () AS HACCEL
      DECLARE SUB DestroyAcceleratorTable ()
      #ifdef USEMDI
      DECLARE FUNCTION CreateMDIWindow (BYVAL cID AS LONG, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL dwStyle AS DWORD, BYVAL dwExStyle AS DWORD, BYVAL hSubMenu AS HANDLE, BYVAL lpfnWndProc AS WNDPROC) AS HWND
      DECLARE PROPERTY MDIClassName (BYREF wszClassName AS WSTRING)
      DECLARE PROPERTY hwndClient () AS HWND
      #endif
END TYPE
' ========================================================================================

' =====================================================================================
' CTabPage class
' Creates a generic window used as a tab page of a tab control.
' =====================================================================================
TYPE CTabPage EXTENDS CWindow

   Private:
      m_hTabPage AS HWND                     ' // Tab page handle

   Public:
      DECLARE DESTRUCTOR
      DECLARE PROPERTY hTabPage () AS HWND
      DECLARE FUNCTION InsertPage (BYVAL hTab AS HWND, BYVAL nPage AS LONG, BYREF wszTitle AS WSTRING = "", _
         BYVAL nImage AS LONG = -1, BYVAL lpfnWndProc AS WNDPROC = NULL, _
         BYVAL dwStyle AS DWORD = WS_CLIPSIBLINGS OR WS_CLIPCHILDREN, BYVAL dwExStyle AS DWORD = 0) AS HWND

END TYPE
' ========================================================================================

' ########################################################################################
'                              *** CSCROLLWINDOW CLASS ***
' ########################################################################################

' ========================================================================================
' CScrollWindow class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CScrollWindow (BYVAL hwnd AS HWND)
   m_hwnd = hwnd
   m_HorzUnits = 10
   m_VertUnits = 10
   GetClientRect(m_hwnd, @m_ClientRect)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CScrollWindow class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CScrollWindow
'   OutputDebugStringW "CScrollWindow.Destructor"
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Sets the amount, in device units, of horizontal and vertical scrolling
' ========================================================================================
PRIVATE SUB CScrollWindow.SetClientRect (BYVAL pRect AS RECT PTR)
   IF pRect = NULL THEN EXIT SUB
   m_CLientRect = *pRect
END SUB
' ========================================================================================

' ========================================================================================
' Sets the amount, in device units, of horizontal and vertical scrolling
' ========================================================================================
PRIVATE SUB CScrollWindow.SetScrollingUnits (BYVAL nHorzUnits AS LONG, BYVAL nVertUnits AS LONG)
   m_HorzUnits = nHorzUnits
   m_VertUnits = nVertUnits
END SUB
' ========================================================================================

' ========================================================================================
' Resets the window scrolling information
' ========================================================================================
PRIVATE SUB CScrollWindow.ResetScrollbars
   ScrollWindow(m_hwnd, m_HScrollPos * m_HorzUnits, 0, NULL, NULL)
   ScrollWindow(m_hwnd, 0, m_VScrollPos * m_HorzUnits, NULL, NULL)
   m_HScrollPos = 0
   m_VScrollPos = 0
   SetScrollPos(m_hwnd, SB_HORZ, m_HScrollPos, CTRUE)
   SetScrollPos(m_hwnd, SB_VERT, m_VScrollPos, CTRUE)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the scroll information
' ========================================================================================
PRIVATE SUB CScrollWindow.SetupScrollbars

   DIM bMaximized AS LONG, tempRect AS RECT, wp AS WINDOWPLACEMENT

   GetClientRect(m_hwnd, @tempRect)
   GetWindowPlacement(m_hwnd, @wp)
   bMaximized = (wp.showCmd = SW_MAXIMIZE)

   ' // Calculate how many scrolling increments for the client area
   m_nHorzInc = ((m_ClientRect.Right - m_ClientRect.Left) - (tempRect.Right  - tempRect.Left) ) \ m_HorzUnits
   m_nVertInc = ((m_ClientRect.Bottom - m_ClientRect.Top) - (tempRect.bottom - tempRect.Top ) ) \ m_HorzUnits

   '// Set the vertical and horizontal scrolling info
   m_HScrollMax = MAX(0, m_nHorzInc)
   m_HScrollPos = MIN(m_HScrollPos, m_HScrollMax)
   SetScrollRange(m_hwnd, SB_HORZ, 0, m_HScrollMax, FALSE)
   SetScrollPos(m_hwnd, SB_HORZ, m_HScrollPos, CTRUE)
   m_VScrollMax = MAX(0, m_nVertInc)
   m_VScrollPos = MIN(m_VScrollPos, m_VScrollMax)
   SetScrollRange(m_hwnd, SB_VERT, 0, m_VScrollMax, FALSE)
   SetScrollPos(m_hwnd, SB_VERT, m_VScrollPos, TRUE)

END SUB
' ========================================================================================

' ========================================================================================
' Handle vertical scrollbar messages
' WM_VSCROLL
' nScrollCode = (int) LOWORD(wParam); // scroll bar value
' nPos = (short int) HIWORD(wParam);  // scroll box position
' hwndScrollBar = (HWND) lParam;      // handle of scroll bar
' ========================================================================================
PRIVATE SUB CScrollWindow.OnVScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

   DIM nInc AS LONG, iMove AS LONG

   SELECT CASE LOWORD(wParam)
      CASE SB_TOP        : nInc = -m_VScrollPos
      CASE SB_BOTTOM     : nInc = m_VScrollMax - m_VScrollPos
      CASE SB_LINEUP     : nInc = -1
      CASE SB_LINEDOWN   : nInc = 1
      CASE SB_PAGEUP     : nInc = MIN(-1, -m_nVertInc)
      CASE SB_PAGEDOWN   : nInc = MAX(1, m_nVertInc)
      CASE SB_THUMBTRACK : nInc = HIWORD(wParam) - m_VScrollPos
      CASE ELSE          : nInc = 0
   END SELECT

   nInc = MAX(-m_VScrollPos, MIN(nInc, m_VScrollMax - m_VScrollPos))
   IF nInc THEN
      m_VScrollPos += nInc
      iMove = -m_HorzUnits * nInc
      ScrollWindow(m_hwnd, 0, iMove, NULL, NULL)
      SetScrollPos(m_hwnd, SB_VERT, m_VScrollPos, CTRUE)
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' Handle horizontal scrollbar messages
' WM_HSCROLL
' nScrollCode = (int) LOWORD(wParam);  // scroll bar value
' nPos = (short int) HIWORD(wParam);   // scroll box position
' hwndScrollBar = (HWND) lParam;       // handle of scroll bar
' ========================================================================================
PRIVATE SUB CScrollWindow.OnHScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

   DIM nInc AS LONG, iMove AS LONG

   SELECT CASE LOWORD(wParam)
      CASE SB_TOP        : nInc = -m_HScrollPos
      CASE SB_BOTTOM     : nInc = m_HScrollMax - m_HScrollPos
      CASE SB_LINEUP     : nInc = -1
      CASE SB_LINEDOWN   : nInc = 1
      CASE SB_PAGEUP     : nInc = -m_HorzUnits
      CASE SB_PAGEDOWN   : nInc = m_HorzUnits
      CASE SB_THUMBTRACK : nInc = HIWORD(wParam) - m_HScrollPos
      CASE ELSE          : nInc = 0
   END SELECT

   nInc = MAX(-m_HScrollPos, MIN(nInc, m_HScrollMax - m_HScrollPos))
   IF nInc THEN
      m_HScrollPos += nInc
      iMove = -m_HorzUnits * nInc
      ScrollWindow(m_hwnd, iMove, 0, NULL, NULL)
      SetScrollPos(m_hwnd, SB_HORZ, m_HScrollPos, CTRUE)
  END IF

END SUB
' ========================================================================================

' ========================================================================================
' Handle WM_SIZE messges
' WM_SIZE
' fwSizeType = wParam;      // resizing flag
' nWidth = LOWORD(lParam);  // width of client area
' nHeight = HIWORD(lParam); // height of client area
' ========================================================================================
PRIVATE SUB CScrollWindow.OnSize (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
   this.ResetScrollbars
   this.SetupScrollbars
END SUB
' ========================================================================================

' ########################################################################################
'                                 *** CWINDOW CLASS ***
' ########################################################################################

' // CWindow class forward declarations
DECLARE FUNCTION CWindow_RegisterClass (BYREF wszClassName AS CONST WSTRING, BYVAL hInstance AS HINSTANCE, BYVAL lpfnWndProc AS WNDPROC) AS ATOM
DECLARE FUNCTION CWindow_WindowProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION CWindow_CloseEnumProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG
' // MDI (Multiple document interface)
#ifdef USEMDI
DECLARE FUNCTION CWindow_RegisterMDIClass (BYREF wsClassName AS WSTRING, BYVAL hInstance AS HINSTANCE, BYVAL lpfnWndProc AS WNDPROC) AS WORD
DECLARE FUNCTION CWindowMDIProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
#endif


' ========================================================================================
' Registers the window class.
' Parameter:
' - lpfnWndProc = Pointer to the window procedure.
' Return Value:
'   An atom that uniquely identifies the class.
' Remarks:
'   The values of several of the members of this structure can be retrieved with
'   GetClassLongPtrW and set with SetClassLongPtrW.
' Note: The class uses CS_HREDRAW OR CS_VREDRAW as default window styles. Without them,
' the background is not repainted and the controls leave garbage in it when resized. With
' them, windows with many controls cause heavy flicker. To avoid flicker, you can change
' the windows style using e.g. pWindow.ClassStyle = CS_DBLCLKS and take care yourself of
' repainting.
' ========================================================================================
PRIVATE FUNCTION CWindow_RegisterClass (BYREF wszClassName AS CONST WSTRING, BYVAL hInstance AS HINSTANCE, BYVAL lpfnWndProc AS WNDPROC) AS ATOM

   DIM    wAtom AS ATOM                     ' // Atom
   DIM    wcexw AS WNDCLASSEXW              ' // WNDCLASSEXW structure
   STATIC nCount AS LONG = 0                ' // Counter
   STATIC m_wszClassName AS WSTRING * 256   ' // class name

   IF LEN(wszClassName) THEN
      m_wszClassName = wszClassName
   ELSE
      m_wszClassName = "FBWindowClass"
      m_wszClassName = m_wszClassName & ":" & WSTR(nCount)
   END IF

   ' // Default handler
   IF lpfnWndProc = NULL THEN lpfnWndProc = @CWindow_WindowProc

   ' // Fill the WNDCLASSEXW structure
   WITH wcexw
      .cbSize        = SIZEOF(wcexw)
      .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
      .lpfnWndProc   = lpfnWndProc
      .cbClsExtra    = 0
      .cbWndExtra    = SIZEOF(HANDLE)
      .hInstance     = hInstance
      .hCursor       = ..LoadCursorW(NULL, CAST(LPCWSTR, IDC_ARROW))
      .hbrBackground = CAST(HBRUSH, COLOR_3DFACE + 1)
      .lpszMenuName  = NULL
      .lpszClassName = @m_wszClassName
      .hIcon         = 0
      .hIconSm       = 0
   END WITH

   ' // Register the class
   wAtom = RegisterClassExW(@wcexw)
   ' // Increment the class counter
   IF wAtom THEN nCount = nCount + 1
   ' // Return the atom
   FUNCTION = wAtom

END FUNCTION
' ========================================================================================

' ========================================================================================
' Processes window messages
' // Uses IsDialogMessage in the message pump
' // Note: To process arrow keys, characters, enter, insert, backspace or delete keys, set USEDLGMSG = 0.
' // Or you can leave it as is and process the WM_GETDLGCODE message:
' // CASE WM_GETDLGCODE
' //    FUNCTION = DLGC_WANTALLKEYS
' // If you are only interested in arrow keys and characters...
' // CASE WM_GETDLGCODE
' //    FUNCTION = DLGC_WANTARROWS OR DLGC_WANTCHARS
#ifndef USEDLGMSG
#define USEDLGMSG 1
#endif
' ========================================================================================
PRIVATE FUNCTION CWindow.DoEvents (BYVAL nCmdShow AS LONG = 0) AS LONG

   DIM uMsg AS MSG
   IF m_hwnd = NULL THEN EXIT FUNCTION
   ' // Show the window and update its client area
   IF nCmdShow = 0 THEN .ShowWindow m_hwnd, SW_SHOW ELSE .ShowWindow m_hwnd, nCmdShow
   .UpdateWindow m_hwnd
   ' // Message loop
   WHILE .GetMessageW(@uMsg, NULL, 0, 0)
#ifdef USEMDI
      ' // Processes accelerator keystrokes for window menu command
      ' // of the multiple document interface (MDI) child windows
      IF m_hwndClient = NULL OR .TranslateMDISysAccel(m_hwndClient, @uMsg) = 0 THEN
#endif
         ' // Processes accelerator keys for menu commands
         IF m_hAccel = NULL OR .TranslateAcceleratorW(m_hwnd, m_hAccel, @uMsg) = 0 THEN
#if (USEDLGMSG >= 1)
            ' // Determines whether a message is intended for the specified
            ' // dialog box and, if it is, processes the message.
            IF .IsDialogMessageW(m_hwnd, @uMsg) = 0 THEN
               ' // Translates virtual-key messages into character messages.
               .TranslateMessage @uMsg
               ' // Dispatches a message to a window procedure.
               .DispatchMessageW @uMsg
            END IF
#else
            ' // Translates virtual-key messages into character messages.
            .TranslateMessage @uMsg
            ' // Dispatches a message to a window procedure.
            .DispatchMessageW @uMsg
#endif
         END IF
#ifdef USEMDI
      END IF
#endif
   WEND
   FUNCTION = uMsg.wParam

END FUNCTION
' ========================================================================================

' ========================================================================================
' Default CWindow callback function.
' ========================================================================================
PRIVATE FUNCTION CWindow_WindowProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC pWindow AS CWindow PTR
#ifdef USEMDI
   STATIC hwndClient AS .HWND    ' // Handle of the MDI client window
   DIM    hwndActive AS .HWND    ' // Active window
   DIM    hMdi AS HANDLE         ' // MDI child window handle
   ' // MDI client window handle
   IF hwndClient = NULL AND pWindow <> NULL THEN hwndClient = pWindow->hwndClient
#endif

   SELECT CASE AS CONST uMsg

      CASE WM_CREATE
         ' // Retrieve a reference to the CWindow class from the MDI child window handle
         DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
         pWindow = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
         EXIT FUNCTION

      CASE WM_SYSCOMMAND
         ' // Capture this message and send a WM_CLOSE message
         ' // Note: Needed with some OCXs, that otherwise remain in memory
         IF (wParam AND &HFFF0) = SC_CLOSE THEN
            SendMessageW hwnd, WM_CLOSE, 0, 0
            EXIT FUNCTION
         END IF

      CASE WM_COMMAND

         SELECT CASE LOWORD(wParam)
            CASE IDCANCEL
               IF HIWORD(wParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

         #ifdef USEMDI
         IF hwndClient THEN
            ' // Pass unprocessed messages to the active MDI child and then to DefFrameProc()
            hwndActive = CAST(HANDLE, .SendMessageW(hwndClient, WM_MDIGETACTIVE, 0, 0))
            IF .IsWindow(hwndActive) THEN .SendMessageW hwndActive, WM_COMMAND, wParam, lParam
         END IF
         #endif

      CASE WM_NOTIFY
'         DIM ptnmhdr AS NMHDR PTR
'         ptnmhdr = lParam
'         SELECT CASE @ptnmhdr.idFrom
'            ' ...
'            ' ...
'         END SELECT

         #ifdef USEMDI
         IF hwndClient THEN
            ' // Pass unprocessed messages to the active MDI child and then to DefFrameProc()
            hwndActive = CAST(HANDLE, .SendMessageW(hwndClient, WM_MDIGETACTIVE, 0, 0))
            IF .IsWindow(hwndActive) <> 0 THEN .SendMessageW hwndActive, WM_NOTIFY, wParam, lParam
         END IF
         #endif

'      CASE WM_SIZE
'         #ifdef USEMDI
'         ' // If the window isn't minimized, resize it (sample code)
'         DIM rc AS RECT
'         IF wParam <> SIZE_MINIMIZED THEN
'            IF hwndClient THEN
'               .GetClientRect hwnd, @rc
'               .MoveWindow hwndClient, rc.Left, rc.Top, rc.Right + 2, rc.Bottom + 2, TRUE
'            END IF
'         END IF
'         #endif
         ' // Note: This message is not passed to DefFrameProc when space
         ' // is being reserved in the client area of the MDI frame
         ' // or controls on the MDI frame are resizeable.
'         EXIT FUNCTION

      #ifdef USEMDI
      CASE WM_CLOSE
         IF hwndClient THEN
            ' // Attempt to close all MDI child windows
            .EnumChildWindows hwndClient, @CWindow_CloseEnumProc, 0
            ' // If child windows are still open abort closing the application
            IF .GetWindow(hwndClient, GW_CHILD) THEN EXIT FUNCTION
         END IF
      #endif

      #ifdef USEMDI
      CASE WM_QUERYENDSESSION
         IF hwndClient THEN
            ' // Attempt to close all MDI child windows
            .EnumChildWindows hwndClient, @CWindow_CloseEnumProc, 0
            ' // If child windows are still open abort closing the application
            IF .GetWindow(hwndClient, GW_CHILD) THEN EXIT FUNCTION
         END IF
      #endif

      CASE WM_DESTROY
         ' // Close the main window
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

#ifdef USEMDI
   IF hwndClient THEN
   ' // The DefFrameProc function provides default processing for any window
   ' // messages that the window procedure of a multiple-document interface (MDI)
   ' // frame window does not process. All window messages that are not explicitly
   ' // processed by the window procedure must be passed to the DefFrameProc
   ' // function, not the DefWindowProc function.
      FUNCTION = .DefFrameProcW(hwnd, hwndClient, uMsg, wParam, lParam)
   ELSE
   ' // The DefWindowProc function calls the default window procedure to provide
   ' // default processing for any window messages that an application does not process.
   ' // This function ensures that every message is processed. DefWindowProc
   ' // is called with the same parameters received by the window procedure.
      FUNCTION = .DefWindowProcW(hwnd, uMsg, wParam, lParam)
   END IF
#else
   FUNCTION = .DefWindowProcW(hwnd, uMsg, wParam, lParam)
#endif

END FUNCTION
' ========================================================================================

' ========================================================================================
' CWindow class constructor
' Usage:
'    DIM pWindow AS CWindow
' -or-
'    DIM pWindow AS CWindow = "MyClassName" (change it as needed)
' Remarks:
' Using the fist syntax, CWindow will use "FBWindowClass:" and a number as the window class.
' Using the second syntax, CWindows will use the passed string as the class name.
' ========================================================================================
PRIVATE CONSTRUCTOR CWindow (BYREF wszClassName AS CONST WSTRING = "")

   ' // Class name
   m_wszClassName = wszClassName
   ' // Instance handle
   m_hInstance = .GetModuleHandleW(NULL)
   ' // Scale windows according to the DPI setting.
   IF AfxUseDpiScaling = 0 THEN this.DPI = 96 ELSE this.DPI = -1
   ' // Default font name
   IF LEN(m_wszDefaultFontName) = 0 THEN
      IF AfxWindowsVersion >= 600 AND CLNG(AfxIsProcessDpiAware) THEN m_wszDefaultFontName = "Segoe UI" ELSE m_wszDefaultFontName = "Tahoma"
   END IF
   ' // Default font size
   IF m_DefaultFontSize = 0 THEN
      IF AfxWindowsVersion >= 600 AND CLNG(AfxIsProcessDpiAware) THEN m_DefaultFontSize = 9 ELSE m_DefaultFontSize = 8
   END IF
   ' // Initialize the common controls library
   DIM icc AS INITCOMMONCONTROLSEX
   icc.dwSize = SIZEOF(icc)
   icc.dwICC  = ICC_NATIVEFNTCTL_CLASS OR ICC_COOL_CLASSES OR ICC_BAR_CLASSES OR _
                ICC_TAB_CLASSES OR ICC_USEREX_CLASSES OR ICC_WIN95_CLASSES OR _
                ICC_STANDARD_CLASSES OR ICC_ANIMATE_CLASS OR ICC_DATE_CLASSES OR _
                ICC_HOTKEY_CLASS OR ICC_INTERNET_CLASSES OR ICC_LISTVIEW_CLASSES OR _
                ICC_PAGESCROLLER_CLASS OR ICC_PROGRESS_CLASS OR ICC_TREEVIEW_CLASSES OR _
                ICC_UPDOWN_CLASS
   .InitCommonControlsEx(@icc)

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CWindow class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWindow

   IF m_hFont THEN .DeleteObject m_hFont
   IF m_hAccel THEN .DestroyAcceleratorTable(m_hAccel)
   IF m_wszClassName <> "" THEN .UnregisterClassW(m_wszClassName, m_hInstance)
   IF m_hRichEditLib THEN FreeLibrary m_hRichEditLib
   IF m_pScrollWindow THEN Delete m_pScrollWindow

END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the DPI (dots per pixel) used by the application.
' ========================================================================================
PRIVATE PROPERTY CWindow.DPI () AS SINGLE
   PROPERTY = m_DPI
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the DPI (dots per pixel) to be used by the application.
' Pass -1 to use the value returned by the GetDeviceCaps API function.
' Note: Set this value if you want to make your application High-DPI aware.
' The main window, controls and fonts will be scaled if the user changes the DPI setting.
' Don't change the DPI value once the main window has been created.
' ========================================================================================
PRIVATE PROPERTY CWindow.DPI (BYVAL nDPI AS SINGLE)

   DIM hDC AS HDC
   m_DPI = nDPI
   ' // Get the screen device context
   hDC = .GetDC(NULL)
   IF m_DPI < 0 THEN
      m_DPI = .GetDeviceCaps(hDC, LOGPIXELSX)
      nDPI = m_DPI
   END IF
   ' // Resolution ratio = current resolution / 96
   IF nDPI < 0 THEN
      m_rx = (.GetDeviceCaps(hDC, LOGPIXELSX) / 96)
      m_ry = (.GetDeviceCaps(hDC, LOGPIXELSY) / 96)
   ELSE
      m_rx = nDPI / 96
      m_ry = nDPI / 96
   END IF
   .ReleaseDC NULL, hDC

END PROPERTY
' ========================================================================================

' =====================================================================================
' Retrieves a value from the user data area of a CWindow.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE PROPERTY CWindow.UserData (BYVAL idx AS LONG) AS LONG_PTR
   IF idx < 0 OR idx > 99 THEN EXIT PROPERTY
   PROPERTY = m_rgUserData(idx)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets a value in the user data area of a CWindow.
' idx = The index number of the user data value to retrieve, in the range 0 to 99 inclusive.
' =====================================================================================
PRIVATE PROPERTY CWindow.UserData (BYVAL idx AS LONG, BYVAL newValue AS LONG_PTR)
   IF idx < 0 OR idx > 99 THEN EXIT PROPERTY
   m_rgUserData(idx) = newValue
END PROPERTY
' =====================================================================================

' ========================================================================================
' Window creation
' Parameters:
' - hParent     = Parent window handle
' - wszTitle    = Window caption
' - lpfnWndProc = Address of the callback function
' - x           = Horizontal position
' - y           = Vertical position
' - nWidth      = Window width
' - nHeight     = Window height
' - dwStyle     = Window style
' - dwExStyle   = Extended style
' Remarks: As the last parameter we are passing a pointer to the class to allow its use
' in the WM_CREATE message, e.g.
'    CASE WM_CREATE
'       DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
'       DIM pWindow AS CWindow PTR = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
'       IF pWindow THEN pWindow->AddControl("Button", hwnd, IDCANCEL, "&Close", 350, 250, 75, 23)
' -or-
'    CASE WM_CREATE
'       DIM pWindow AS CWindow PTR = AfxCWindowPtr(CAST(CREATESTRUCT PTR, lParam))
'       IF pWindow THEN pWindow->AddControl("Button", hwnd, IDCANCEL, "&Close", 350, 250, 75, 23)
' -or-
'    CASE WM_CREATE
'       DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
'       IF pWindow THEN pWindow->AddControl("Button", hwnd, IDCANCEL, "&Close", 350, 250, 75, 23)
' ========================================================================================
PRIVATE FUNCTION CWindow.Create (BYVAL hParent AS HWND = NULL, BYREF wszTitle AS WSTRING = "", BYVAL lpfnWndProc AS WNDPROC = NULL, _
   BYVAL x AS LONG = CW_USEDEFAULT, BYVAL y AS LONG = CW_USEDEFAULT, BYVAL nWidth AS LONG = CW_USEDEFAULT, BYVAL nHeight AS LONG = CW_USEDEFAULT, _
   BYVAL dwStyle AS UINT = WS_OVERLAPPEDWINDOW OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS, _
   BYVAL dwExStyle AS UINT = WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE) AS HWND

   IF m_hwnd <> NULL THEN EXIT FUNCTION
   m_wAtom = CWindow_RegisterClass(m_wszClassName, m_hInstance, lpfnWndProc)
   IF m_wAtom = 0 THEN EXIT FUNCTION
   ' // Create a default font
   IF m_hFont = NULL THEN m_hFont = this.CreateFont(m_wszDefaultFontName, m_DefaultFontSize, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET)
   ' // Create the window
   m_hwnd = .CreateWindowExW(dwExStyle, CAST(LPCWSTR, CAST(ULONG_PTR, CAST(WORD, m_wAtom))), wszTitle, dwStyle, _
            IIF(x = CW_USEDEFAULT, CW_USEDEFAULT, x * m_rx), _
            IIF(y = CW_USEDEFAULT, CW_USEDEFAULT, y * m_ry), _
            IIF(nWidth = CW_USEDEFAULT, CW_USEDEFAULT, nWidth * m_rx), _
            IIF(nHeight = CW_USEDEFAULT, CW_USEDEFAULT, nHeight * m_ry), _
            hParent, NULL, m_hInstance, CAST(HANDLE, @this))
   IF m_hwnd = NULL THEN EXIT FUNCTION
   ' // Get the class name
   .GetClassNameW(m_hwnd, @m_wszClassName, SIZEOF(m_wszClassName) / 2)
   FUNCTION = m_hwnd
   ' // Set the font
   IF m_hFont THEN .SendMessageW m_hwnd, WM_SETFONT, CAST(WPARAM, m_hFont), FALSE
   ' // Store the class pointer
   .SetWindowLongPtrW(m_hwnd, 0, CAST(LONG_PTR, @this))

END FUNCTION
' ========================================================================================

' ========================================================================================
' Adjusts the bounding rectangle of the window based on the desired size of the client area.
' ========================================================================================
PRIVATE SUB CWindow.SetClientSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
   AfxSetWindowClientSize(m_hwnd, nWidth, nHeight, m_rx, m_ry)
END SUB
' ========================================================================================

' =====================================================================================
' Centers a window on the screen or over another window.
' It also ensures that the placement is done within the work area.
' Parameters:
' - hwnd = Handle of the window.
' - hwndParent = [optional] Handle of the parent window.
' =====================================================================================
PRIVATE SUB CWindow.Center (BYVAL hwnd AS HWND = NULL, BYVAL hwndParent AS HWND = NULL)
   IF hwnd = NULL THEN hwnd = m_hwnd
   AfxCenterWindow(hwnd, hwndParent)
END SUB
' =====================================================================================

' =====================================================================================
' Returns the window handle
' =====================================================================================
PRIVATE PROPERTY CWindow.hWindow () AS HWND
   PROPERTY = m_hwnd
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the window handle
' =====================================================================================
PRIVATE PROPERTY CWindow.hWindow (BYVAL hwnd AS HWND)
   IF m_hwnd = NULL THEN m_hwnd = hwnd
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the instance handle
' =====================================================================================
PRIVATE PROPERTY CWindow.InstanceHandle () AS HINSTANCE
   PROPERTY = m_hInstance
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the instance handle
' =====================================================================================
PRIVATE PROPERTY CWindow.InstanceHandle (BYVAL hInst AS HINSTANCE)
   IF m_hInstance = NULL THEN m_hInstance = hInst
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the window class style.
' =====================================================================================
PRIVATE PROPERTY CWindow.ClassStyle () AS ULONG_PTR
   IF m_hwnd = NULL THEN EXIT PROPERTY
   PROPERTY = .GetClassLongPtrW(m_hwnd, GCL_STYLE)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the window class style.
' =====================================================================================
PRIVATE PROPERTY CWindow.ClassStyle (BYVAL dwStyle AS ULONG_PTR)
   IF m_hwnd = NULL THEN EXIT PROPERTY
   .SetClassLongPtrW(m_hwnd, GCL_STYLE, dwStyle)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the window style.
' =====================================================================================
PRIVATE PROPERTY CWindow.WindowStyle () AS ULONG_PTR
   IF m_hwnd = NULL THEN EXIT PROPERTY
   PROPERTY = .GetWindowLongPtrW(m_hwnd, GWL_STYLE)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the window style.
' Example to create a popup window: pWindow.WindowStyle = WS_POPUPWINDOW OR WS_CAPTION
' =====================================================================================
PRIVATE PROPERTY CWindow.WindowStyle (BYVAL dwStyle AS ULONG_PTR)
   IF m_hwnd = NULL THEN EXIT PROPERTY
   .SetWindowLongPtrW(m_hwnd, GWL_STYLE, dwStyle)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the window extended style.
' =====================================================================================
PRIVATE PROPERTY CWindow.WindowExStyle () AS ULONG_PTR
   IF m_hwnd = NULL THEN EXIT PROPERTY
   PROPERTY = .GetWindowLongPtrW(m_hwnd, GWL_EXSTYLE)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the window extended style.
' =====================================================================================
PRIVATE PROPERTY CWindow.WindowExStyle (BYVAL dwExStyle AS ULONG_PTR)
   IF m_hwnd = NULL THEN EXIT PROPERTY
   .SetWindowLongPtrW(m_hwnd, GWL_EXSTYLE, dwExStyle)
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets/sets the default font name
' =====================================================================================
PRIVATE PROPERTY CWindow.DefaultFontName () AS STRING
   PROPERTY = m_wszDefaultFontName
END PROPERTY
' =====================================================================================
PRIVATE PROPERTY CWindow.DefaultFontName (BYREF wszFontName AS WSTRING)
   IF VARPTR(wszFontName) <> NULL AND LEN(wszFontName) > 0 THEN m_wszDefaultFontName = wszFontName
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets/sets the size in points of the default font.
' =====================================================================================
PRIVATE PROPERTY CWindow.DefaultFontSize () AS LONG
   PROPERTY = m_DefaultFontSize
END PROPERTY
' =====================================================================================
PRIVATE PROPERTY CWindow.DefaultFontSize (BYVAL nPointSize AS LONG)
   m_DefaultFontSize = nPointSize
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the handle of the font used as default.
' =====================================================================================
PRIVATE PROPERTY CWindow.Font () AS HFONT
   PROPERTY = m_hFont
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the handle of the font used as default.
' =====================================================================================
PRIVATE PROPERTY CWindow.Font (BYVAL hFont AS HFONT)
   IF m_hFont THEN .DeleteObject m_hFont
   m_hFont = hFont
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the background brush.
' =====================================================================================
PRIVATE PROPERTY CWindow.Brush () AS HBRUSH
   IF m_hwnd = NULL THEN EXIT PROPERTY
   PROPERTY = CAST(HBRUSH, .GetClassLongPtrW(m_hwnd, GCLP_HBRBACKGROUND))
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the background brush.
' Handle to the class background brush. This member can be a handle to the physical
' brush to be used for painting the background, or it can be a color value. A color
' value must be one of the standard system colors (the value 1 must be added
' to the chosen color), e.g. COLOR_WINDOW + 1.
' You can also use CreateSolidBrush to create a logical brush with a solid color, e.g.
' CreateSolidBrush(BGR(0, 0, 255)
' =====================================================================================
PRIVATE PROPERTY CWindow.Brush (BYVAL hbrBackground AS HBRUSH)
   IF m_hwnd = NULL THEN EXIT PROPERTY
   .SetClassLongPtrW m_hwnd, GCLP_HBRBACKGROUND, CAST(LONG_PTR, hbrBackground)
   .InvalidateRect m_hwnd, NULL, CTRUE
   .UpdateWindow m_hwnd
END PROPERTY
' =====================================================================================

' =====================================================================================
' Associates a new large icon with the main window. The system displays the large icon
' in the ALT+TAB dialog box.
' Parameters:
' - hIcon = Handle to the new large icon.
'   If this parameter is NULL, the icon is removed.
' =====================================================================================
PRIVATE PROPERTY CWindow.BigIcon (BYVAL hIcon AS HICON)
   .SendMessageW(m_hwnd, WM_SETICON, CAST(WPARAM, ICON_BIG), CAST(LPARAM, hIcon))
END PROPERTY
' =====================================================================================

' =====================================================================================
' Associates a new small icon with the main window. The system displays the small icon
' in the in the window caption.
' Parameters:
' - hIcon = Handle to the new small icon.
'   If this parameter is NULL, the icon is removed.
' =====================================================================================
PRIVATE PROPERTY CWindow.SmallIcon (BYVAL hIcon AS HICON)
   .SendMessageW(m_hwnd, WM_SETICON, CAST(WPARAM, ICON_SMALL), CAST(LPARAM, hIcon))
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the horizontal scaling ratio
' =====================================================================================
PRIVATE PROPERTY CWindow.rxRatio () AS SINGLE
   PROPERTY = m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the horizontal scaling ratio
' =====================================================================================
PRIVATE PROPERTY CWindow.rxRatio (BYVAL rx AS SINGLE)
   m_rx = rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Gets the vertical scaling ratio
' =====================================================================================
PRIVATE PROPERTY CWindow.ryRatio () AS SINGLE
   PROPERTY = m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Sets the vertical scaligng ratio
' =====================================================================================
PRIVATE PROPERTY CWindow.ryRatio (BYVAL ry AS SINGLE)
   m_ry = ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Scales an horizontal coordinate according the DPI setting.
' =====================================================================================
PRIVATE FUNCTION CWindow.ScaleX (BYVAL cx AS SINGLE) AS SINGLE
   FUNCTION = cx * m_rx
END FUNCTION
' =====================================================================================

' =====================================================================================
' Scales a vertical coordinate according the DPI setting.
' =====================================================================================
PRIVATE FUNCTION CWindow.ScaleY (BYVAL cy AS SINGLE) AS SINGLE
   FUNCTION = cy * m_ry
END FUNCTION
' =====================================================================================

' =====================================================================================
' Unscales an horizontal coordinate according the DPI setting.
' =====================================================================================
PRIVATE FUNCTION CWindow.UnScaleX (BYVAL cx AS SINGLE) AS SINGLE
   FUNCTION = cx / m_rx
END FUNCTION
' =====================================================================================

' =====================================================================================
' Unscales a vertical coordinate according the DPI setting.
' =====================================================================================
PRIVATE FUNCTION CWindow.UnScaleY (BYVAL cy AS SINGLE) AS SINGLE
   FUNCTION = cy / m_ry
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns the width of the window
' =====================================================================================
PRIVATE PROPERTY CWindow.Width () AS LONG
   DIM rc AS RECT
   .GetWindowRect m_hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Right - rc.Left) / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the height of the window
' =====================================================================================
PRIVATE PROPERTY CWindow.Height () AS LONG
   DIM rc AS RECT
   .GetWindowRect m_hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Bottom - rc.Top) / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the client width of the window
' =====================================================================================
PRIVATE PROPERTY CWindow.ClientWidth () AS LONG
   DIM rc AS RECT
   .GetClientRect m_hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Right - rc.Left) / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the client height of the window
' =====================================================================================
PRIVATE PROPERTY CWindow.ClientHeight () AS LONG
   DIM rc AS RECT
   .GetClientRect m_hwnd, @rc
   ' // Divide by m_ry to make the result High DPI aware
   PROPERTY = (rc.Bottom - rc.Top) / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the width of the specified window
' =====================================================================================
PRIVATE PROPERTY CWindow.ControlWidth (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   .GetWindowRect hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Right - rc.Left) / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the height of the specified window
' =====================================================================================
PRIVATE PROPERTY CWindow.ControlHeight (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   .GetWindowRect hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Bottom - rc.Top) / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the client width of the specified window
' =====================================================================================
PRIVATE PROPERTY CWindow.ControlClientWidth (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   .GetClientRect hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Right - rc.Left) / m_rx
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the client height of the specified window
' =====================================================================================
PRIVATE PROPERTY CWindow.ControlClientHeight (BYVAL hwnd AS HWND) AS LONG
   DIM rc AS RECT
   .GetClientRect hwnd, @rc
   ' // Divide by m_rx to make the result High DPI aware
   PROPERTY = (rc.Bottom - rc.Top) / m_ry
END PROPERTY
' =====================================================================================

' =====================================================================================
' Retrieves the dimensions of the bounding rectangle of the main window.
' =====================================================================================
PRIVATE SUB CWindow.GetWindowRect OVERLOAD (BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   .GetWindowRect(m_hwnd, lpRect)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetWindowRect OVERLOAD () AS RECT
   DIM rc AS RECT
   .GetWindowRect(m_hwnd, @rc)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the coordinates of the main window's client area.
' =====================================================================================
PRIVATE SUB CWindow.GetClientRect OVERLOAD (BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   .GetClientRect(m_hwnd, lpRect)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetClientRect OVERLOAD () AS RECT
   DIM rc AS RECT
   .GetClientRect(m_hwnd, @rc)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the dimensions of the bounding rectangle of the specified window.
' =====================================================================================
PRIVATE SUB CWindow.GetControlWindowRect (BYVAL hwnd AS HWND, BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   .GetWindowRect(hwnd, lpRect)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetControlWindowRect (BYVAL hwnd AS HWND) AS RECT
   DIM rc AS RECT
   .GetWindowRect(hwnd, @rc)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the dimensions of a window's client area.
' =====================================================================================
PRIVATE SUB CWindow.GetControlClientRect OVERLOAD (BYVAL hwnd AS HWND, BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   .GetClientRect(hwnd, lpRect)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetControlClientRect OVERLOAD (BYVAL hwnd AS HWND) AS RECT
   DIM rc AS RECT
   .GetClientRect(hwnd, @rc)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves the size of the work area on the primary display monitor. The work area is the
' portion of the screen not obscured by the system taskbar or by application desktop toolbars.
' =====================================================================================
PRIVATE SUB CWindow.GetWorkArea OVERLOAD (BYVAL lpRect AS LPRECT)
   IF lpRect = NULL THEN EXIT SUB
   SystemParametersInfoW(SPI_GETWORKAREA, 0, lpRect, 0)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   lpRect->Left   /= m_rx
   lpRect->Right  /= m_rx
   lpRect->Top    /= m_ry
   lpRect->Bottom /= m_ry
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CWindow.GetWorkArea OVERLOAD () AS RECT
   DIM rc AS RECT
   SystemParametersInfoW(SPI_GETWORKAREA, 0, @rc, 0)
   ' // Divide by m_rx and m_ry to make the result High DPI aware
   rc.Left   /= m_rx
   rc.Right  /= m_rx
   rc.Top    /= m_ry
   rc.Bottom /= m_ry
   FUNCTION = rc
END FUNCTION
' =====================================================================================

' =====================================================================================
' Retrieves a handle to the child control specified by its identifier.
' =====================================================================================
PRIVATE FUNCTION CWindow.ControlHandle (BYVAL cID AS LONG) AS HWND
   FUNCTION = .GetDlgItem(m_hwnd, cID)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Returns a pointer to the scroll window class
' =====================================================================================
PRIVATE PROPERTY CWindow.ScrollWindowPtr () AS CScrollWindow PTR
   PROPERTY = m_pScrollWindow
END PROPERTY
' =====================================================================================
' =====================================================================================
' Sets a pointer to the scroll window class
' =====================================================================================
PRIVATE PROPERTY CWindow.ScrollWindowPtr (BYVAL pScrollWindow AS CScrollWindow PTR)
   IF m_pScrollWindow THEN Delete m_pScrollWindow
   m_pScrollWindow = pScrollWindow
END PROPERTY
' =====================================================================================

' =====================================================================================
' Resizes the window
' =====================================================================================
PRIVATE SUB CWindow.Resize
   .SendMessageW m_hwnd, WM_SIZE, CAST(WPARAM, SIZE_RESTORED), MAKELONG(this.ClientWidth, this.ClientHeight)
END SUB
' =====================================================================================

' =====================================================================================
' Changes the size, position, and Z order of a child, pop-up, or top-level window.
' =====================================================================================
PRIVATE FUNCTION CWindow.SetWindowPos (BYVAL hwnd AS HWND, BYVAL hwndInsertAfter AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL cx AS LONG, BYVAL cy AS LONG, BYVAL uFlags AS UINT) AS BOOLEAN
   FUNCTION = .SetWindowPos(hwnd, hwndInsertAfter, x * m_rx, y * m_ry, cx * m_rx, cy * m_ry, uFlags)
END FUNCTION
' =====================================================================================

' =====================================================================================
' Changes the position and dimensions of the specified window.
' =====================================================================================
PRIVATE FUNCTION CWindow.MoveWindow (BYVAL hwnd AS HWND, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL bRepaint AS BOOLEAN) AS BOOLEAN
   FUNCTION = .MoveWindow(hwnd, x * m_rx, y * m_ry, nWidth * m_rx, nHeight * m_ry, bRepaint)
END FUNCTION
' =====================================================================================

' ========================================================================================
' Creates a High DPI aware logical font.
' - wszFaceName = The typeface name.
' - lPointSize = The point size.
' - lWeight = The weight of the font in the range 0 through 1000. For example, 400 is normal
'      and 700 is bold. If this value is zero, a default weight is used.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: The handle of the font or NULL on failure.
' Remarks: The returned font must be destroyed with DeleteObject or the macro DeleteFont
' when no longer needed to prevent memory leaks.
' Usage examples:
'   hFont = CWindow.CreateFont("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
'   hFont = CWindow.CreateFont("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
'   hFont = CWindow.CreateFont("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION CWindow.CreateFont ( _
   BYREF wszFaceName AS WSTRING, _                  ' __in Typeface name of font
   BYVAL lPointSize  AS LONG, _                     ' __in Point size
   BYVAL lWeight     AS LONG = 0, _                 ' __in Font weight(bold etc.)
   BYVAL bItalic     AS UBYTE = FALSE, _            ' __in CTRUE = italic
   BYVAL bUnderline  AS UBYTE = FALSE, _            ' __in CTRUE = underline
   BYVAL bStrikeOut  AS UBYTE = FALSE, _            ' __in CTRUE = strikeout
   BYVAL bCharSet    AS UBYTE = DEFAULT_CHARSET _   ' __in character set
   ) AS HFONT                                       ' Handle of font or NULL on failure.

   DIM tlfw AS LOGFONTW
   DIM wstrFaceName AS WSTRING * LF_FACESIZE
   IF VARPTR(wszFacename) <> NULL AND LEN(wszFacename) > 0 THEN wstrFaceName = wszFaceName ELSE wstrFaceName = m_wszDefaultFontName

   ' // The result must be 1, unless the programmer has set a different DPI with the set DPI property
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   lPointSize = (lPointSize * m_DPI) \ .GetDeviceCaps(hDC, LOGPIXELSY)

   tlfw.lfHeight         = -MulDiv(lPointSize, .GetDeviceCaps(hDC, LOGPIXELSY), 72)  ' logical font height
   tlfw.lfWidth          =  0                                                        ' average character width
   tlfw.lfEscapement     =  0                                                        ' escapement
   tlfw.lfOrientation    =  0                                                        ' orientation angles
   tlfw.lfWeight         =  lWeight                                                  ' font weight
   tlfw.lfItalic         =  bItalic                                                  ' italic(CTRUE/FALSE)
   tlfw.lfUnderline      =  bUnderline                                               ' underline(CTRUE/FALSE)
   tlfw.lfStrikeOut      =  bStrikeOut                                               ' strikeout(CTRUE/FALSE)
   tlfw.lfCharSet        =  bCharset                                                 ' character set
   tlfw.lfOutPrecision   =  OUT_TT_PRECIS                                            ' output precision
   tlfw.lfClipPrecision  =  CLIP_DEFAULT_PRECIS                                      ' clipping precision
   tlfw.lfQuality        =  DEFAULT_QUALITY                                          ' output quality
   tlfw.lfPitchAndFamily =  FF_DONTCARE                                              ' pitch and family
   tlfw.lfFaceName       =  wstrFaceName                                             ' typeface name

   .ReleaseDC HWND_DESKTOP, hDC
   FUNCTION = .CreateFontIndirectW(@tlfw)

END FUNCTION
' ========================================================================================

' =====================================================================================
' Creates a High DPI aware logical font and sets it as the default font.
' Usage examples:
'   CWindow.SetFont("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
'   CWindow.SetFont("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
'   CWindow.SetFont("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' Return Value = TRUE or FALSE.
' =====================================================================================
PRIVATE FUNCTION CWindow.SetFont ( _
   BYREF wszFaceName AS WSTRING, _                  ' __in Typeface name of font
   BYVAL lPointSize  AS LONG, _                     ' __in Point size
   BYVAL lWeight     AS LONG = 0, _                 ' __in Font weight(bold etc.)
   BYVAL bItalic     AS UBYTE = FALSE, _            ' __in CTRUE = italic
   BYVAL bUnderline  AS UBYTE = FALSE, _            ' __in CTRUE = underline
   BYVAL bStrikeOut  AS UBYTE = FALSE, _            ' __in CTRUE = strikeout
   BYVAL bCharSet    AS UBYTE = DEFAULT_CHARSET _   ' __in character set
   ) AS BOOLEAN                                     ' Result

   DIM hFont AS HFONT
   DIM wstrFaceName AS WSTRING * LF_FACESIZE
   IF VARPTR(wszFacename) <> NULL AND LEN (wszFacename) > 0 THEN wstrFaceName = wszFaceName ELSE wstrFaceName = m_wszDefaultFontName
   hFont = this.CreateFont(wstrFaceName, lPointSize, lWeight, bItalic, bUnderline, bStrikeOut, bCharSet)
   IF hFont THEN
      IF m_hFont THEN .DeleteObject m_hFont
      m_hFont = hFont
      m_wszDefaultFontName = wstrFaceName
      m_DefaultFontSize = lPointSize
      FUNCTION = TRUE
   ELSE
      FUNCTION = FALSE
   END IF

END FUNCTION
' =====================================================================================

' =====================================================================================
' Adds a control to the window
' =====================================================================================
PRIVATE FUNCTION CWindow.AddControl ( _
   BYREF wszClassName AS WSTRING, _                       ' // Class name
   BYVAL hParent AS HWND = NULL, _                        ' // Parent window handle
   BYVAL cID AS INTEGER = 0, _                            ' // Control identifier
   BYREF wszTitle AS WSTRING = "", _                      ' // Control caption
   BYVAL x AS LONG = 0, _                                 ' // Horizontal position
   BYVAL y AS LONG = 0, _                                 ' // Vertical position
   BYVAL nWidth AS LONG = 0, _                            ' // Control width
   BYVAL nHeight AS LONG = 0, _                           ' // Control height
   BYVAL dwStyle AS LONG = -1, _                          ' // Control style
   BYVAL dwExStyle AS LONG = -1, _                        ' // Extended style
   BYVAL lpParam AS LONG_PTR = 0, _                       ' // Pointer to custom data
   BYVAL pWndProc AS WNDPROC = NULL, _                    ' // Address of the window callback procedure
   BYVAL uIdSubclass AS UINT_PTR = &HFFFFFFFF, _          ' // The subclass ID
   BYVAL dwRefData as DWORD_PTR = NULL _                  ' // Pointer to reference data
   ) AS HWND                                              ' // Control handle

   DIM hCtl AS HWND
   IF LEN(wszClassName) = 0 THEN EXIT FUNCTION
   IF hParent = NULL THEN hParent = m_hwnd
   DIM bSetFont AS LONG = CTRUE
   ' // Window styles
   DIM wsClassName AS WSTRING * 260
   wsClassName = wszClassName

   SELECT CASE UCASE(wsClassName)
      CASE "BUTTON"
         ' Adds a button to the window
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER
         IF dwStyle = BS_FLAT THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER OR BS_FLAT
         IF dwStyle = BS_DEFPUSHBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_DEFPUSHBUTTON
         IF dwStyle = BS_OWNERDRAW THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_OWNERDRAW
         #if _WIN32_WINNT = &h0602
         IF dwStyle = BS_SPLITBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_SPLITBUTTON
         IF dwStyle = BS_DEFSPLITBUTTON THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_CENTER OR BS_VCENTER OR BS_DEFSPLITBUTTON
         #endif
         dwExStyle = 0
      CASE "CUSTOMBUTTON", "OWNERDRAWBUTTON"
         ' Adds an ownerdraw button to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_OWNERDRAW
         dwExStyle = 0
      CASE "RADIOBUTTON", "OPTION"
         ' Adds a radio button to the window.
         ' Note: In PowerBASIC this control is called "Option".
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER
         IF dwStyle = WS_GROUP THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER OR WS_GROUP
         dwExStyle = 0
      CASE "CHECKBOX"
         ' Adds a checkbox to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTOCHECKBOX OR BS_LEFT OR BS_VCENTER
         dwExStyle = 0
      CASE "CHECK3STATE"
         ' Adds a 3 state checkbox to the window.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_AUTO3STATE OR BS_LEFT OR BS_VCENTER
         dwExStyle = 0
      CASE "LABEL"
         ' Adds a label to the window.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SS_LEFT OR WS_GROUP OR SS_NOTIFY
         dwExStyle = 0
      CASE "BITMAPLABEL"
         ' Adds an image label to the window.
         ' You must delete the bitmap before the application ends.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_BITMAP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "ICONLABEL"
         ' Adds an image label to the window.
         ' You must delete the icon before the application ends.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_ICON
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "BITMAPBUTTON"
         ' Adds an image button to the window.
         ' You must delete the bitmap before the application ends.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_BITMAP
         dwExStyle = 0
      CASE "ICONBUTTON"
         ' Adds an image button to the window.
         ' You must delete the icon before the application ends.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_ICON
         dwExStyle = 0
      CASE "CUSTOMLABEL"
         ' Adds an ownerdraw label to the window.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR SS_OWNERDRAW
         dwExStyle = 0
         bSetFont = FALSE
      CASE "FRAME", "FRAMEWINDOW"
         ' Adds a frame to the window.
         ' Note: This is not the same that PowerBASIC DDT's Frame control, that in fact is a Group Box.
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPSIBLINGS OR WS_GROUP OR SS_BLACKFRAME
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "GROUPBOX"
         ' Adds a group box to the window.
         ' Note: This is the same that DDT's frame control.
         wsClassName = "Button"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPSIBLINGS OR WS_GROUP OR BS_GROUPBOX
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
      CASE "LINE"
         ' Adds an horizontal line to the window
         wsClassName = "Static"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SS_ETCHEDFRAME
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_TRANSPARENT
         bSetFont = FALSE
      CASE "EDIT", "TEXTBOX"
         ' Adds an edit control to the window.
         wsClassName = "Edit"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR ES_AUTOHSCROLL
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "EDITMULTILINE", "MULTILINETEXTBOX"
         ' Adds an edit control to the window.
         wsClassName = "Edit"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR WS_VSCROLL OR ES_LEFT OR ES_AUTOHSCROLL OR ES_MULTILINE OR ES_NOHIDESEL OR ES_WANTRETURN
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "COMBOBOX"
         ' Adds a combo box to the window.
         IF dwStyle = -1 THEN dwStyle = WS_CHILD OR WS_VISIBLE OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR CBS_DROPDOWN OR CBS_HASSTRINGS OR CBS_SORT
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "COMBOBOXEX", "COMBOBOXEX32"
         ' Adds a combo box ex to the window.
         wsClassName = "ComboBoxEx32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR CBS_DROPDOWNLIST
         dwExStyle = 0
      CASE "LISTBOX"
         ' Adds a list box to the window.
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_HSCROLL OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR LBS_STANDARD OR LBS_HASSTRINGS OR LBS_SORT OR LBS_NOTIFY
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "PROGRESSBAR", "MSCTLS_PROGRESS32"
         ' Adds a progress bar to the window.
         wsClassName = "msctls_progress32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE
         dwExStyle = 0
         bSetFont = FALSE
      CASE "HEADER", "SYSHEADER32"
         ' Adds an header control to the window.
         wsClassName = "SysHeader32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR CCS_TOP OR HDS_HORZ OR HDS_BUTTONS
         dwExStyle = 0
      CASE "TREEVIEW", "SYSTREEVIEW32"
         ' Adds a tree view control to the window.
         wsClassName = "SysTreeView32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR TVS_HASBUTTONS OR TVS_HASLINES OR TVS_LINESATROOT OR TVS_SHOWSELALWAYS
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "LISTVIEW", "SYSLISTVIEW32"
         ' Adds a list view control to the window.
         wsClassName = "SysListView32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_TABSTOP OR LVS_REPORT OR LVS_SHOWSELALWAYS OR LVS_SHAREIMAGELISTS OR LVS_AUTOARRANGE OR LVS_EDITLABELS OR LVS_ALIGNTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "TOOLBAR", "TOOLBARWINDOW32"
         ' Adds a toolbar control to the window.
         wsClassName = "ToolbarWindow32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_TOP OR WS_BORDER OR TBSTYLE_FLAT OR TBSTYLE_TOOLTIPS
         dwExStyle = 0
      CASE "REBAR", "REBARWINDOW32"
         ' Adds a rebar control to the window.
         wsClassName = "ReBarWindow32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_BORDER OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_NODIVIDER OR RBS_VARHEIGHT OR RBS_BANDBORDERS
         dwExStyle = 0
      CASE "DATETIMEPICKER", "SYSDATETIMEPICK32"
         ' Adds a date time picker control to the window.
         wsClassName = "SysDateTimePick32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR DTS_SHORTDATEFORMAT
         dwExStyle = 0
      CASE "MONTHCALENDAR", "MONTHCAL", "SYSMONTHCAL32"
         ' Adds a month calendar control to the window.
         wsClassName = "SysMonthCal32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "IPADDRESS", "SYSIPADDRESS32"
         ' Adds an IPAddress control to the window.
         wsClassName = "SysIPAddress32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "HOTKEY", "MSCTLS_HOTKEY32"
         ' Adds an hotkey control to the window.
         wsClassName = "msctls_hotkey32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
      CASE "ANIMATE", "ANIMATION", "SYSANIMATE32"
         ' Adds an animation control to the window.
         wsClassName = "SysAnimate32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR ACS_TRANSPARENT
         dwExStyle = 0
      CASE "SYSLINK"
         ' Adds a SysLink control to the window.
         ' Note: The SysLink control is defined in the ComCtl32.dll version 6 and requires a manifest
         ' or directive that specifies that version 6 of the DLL should be used if it is available.
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
         dwExStyle = 0
         bSetFont = FALSE
      CASE "PAGER", "SYSPAGER"
         ' Adds a Pager control to the window.
         wsClassName = "SysPager"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR PGS_HORZ
         dwExStyle = 0
         bSetFont = FALSE
      CASE "TAB", "TABCONTROL", "SYSTABCONTROL32"
         ' Adds a Tab control to the window.
         wsClassName = "SysTabControl32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_GROUP OR WS_TABSTOP OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR TCS_TABS OR TCS_SINGLELINE OR TCS_RAGGEDRIGHT
         IF dwExStyle = -1 THEN dwExStyle = 0
         dwExStyle =  dwExStyle OR WS_EX_CONTROLPARENT
      CASE "STATUSBAR", "MSCTLS_STATUSBAR32"
         ' Adds a StatusBar control to the window.
         wsClassName = "msctls_statusbar32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_BOTTOM OR SBARS_SIZEGRIP
         dwExStyle = 0
      CASE "SIZEBAR", "SIZEBOX", "SIZEGRIP"
         ' Adds a size box to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR SBS_SIZEGRIP OR SBS_SIZEBOXBOTTOMRIGHTALIGN
         dwExStyle = 0
         bSetFont = FALSE
         nWidth = GetSystemMetrics(SM_CXVSCROLL)
         nHeight = GetSystemMetrics(SM_CYHSCROLL)
         DIM rcClient AS RECT = this.GetClientRect
         x = rcClient.Right - nWidth
         y = rcClient.Bottom - nHeight
      CASE "HSCROLLBAR"
         ' Adds an horizontal scroll bar to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR SBS_HORZ
         dwExStyle = 0
         bSetFont = FALSE
      CASE "VSCROLLBAR"
         ' Adds a vertical scroll bar to the window.
         wsClassName = "Scrollbar"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR SBS_VERT
         dwExStyle = 0
         bSetFont = FALSE
      CASE "TRACKBAR", "MSCTLS_TRACKBAR32"
         wsClassName = "msctls_trackbar32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR TBS_AUTOTICKS OR TBS_HORZ OR TBS_BOTTOM OR TBS_TOOLTIPS
         dwExStyle = 0
         bSetFont = FALSE
      CASE "UPDOWN", "MSCTLS_UPDOWN32"
         wsClassName = "msctls_updown32"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR UDS_WRAP OR UDS_ARROWKEYS OR UDS_ALIGNRIGHT OR UDS_SETBUDDYINT
         dwExStyle = 0
         bSetFont = FALSE
      CASE "RICHEDIT", "RichEdit50W"
         IF dwStyle = -1 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR ES_LEFT OR WS_HSCROLL OR WS_VSCROLL OR ES_AUTOHSCROLL OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_WANTRETURN OR ES_NOHIDESEL OR ES_SAVESEL
         IF dwExStyle = -1 THEN dwExStyle = WS_EX_CLIENTEDGE
         wsClassName = "RichEdit50W"
         m_hRichEditLib = CAST(HMODULE, LoadLibraryW("MSFTEDIT.DLL"))
   END SELECT
   ' // Don't allow negative values for the styles
   IF dwStyle = -1 THEN dwStyle = 0
   IF dwExStyle = -1 THEN dwExStyle = 0
   ' // Make sure that the control has the WS_CHILD style
   dwStyle = dwStyle OR WS_CHILD
   ' // Create the control
   hCtl = .CreateWindowExW(dwExStyle, wsClassName, wszTitle, dwStyle, x * m_rx, y * m_ry, nWidth * m_rx, nHeight * m_ry, _
          hParent, CAST(HMENU, CAST(LONG_PTR, cID)), m_hInstance, CAST(LPVOID, lpParam))
   IF hCtl = NULL THEN EXIT FUNCTION
   ' // Set the font
   IF m_hFont THEN
      IF bSetFont THEN .SendMessageW hCtl, WM_SETFONT, CAST(WPARAM, m_hFont), CTRUE
   END IF
   SELECT CASE UCASE(wszClassName)
      CASE "LISTBOX"
         ' // Adjust the height of the control so that the integral height
         ' // is based on the new font rather than the default SYSTEM_FONT
         SetWindowPos hCtl, NULL, x, y, nWidth, nHeight, SWP_NOZORDER
      CASE "DATETIMEPICKER", "SYSDATETIMEPICK32"
         ' // Sets the font to be used by the date and time picker control's child month calendar control.
         IF m_hFont THEN SendMessageW hCtl, DTM_SETMCFONT, CAST(WPARAM, m_hFont), CTRUE
      CASE "PROGRESSBAR", "MSCTLS_PROGRESS32"
         ' // Set the default range
         .SendMessageW hCtl, PBM_SETRANGE32, 0, 100
         ' // Set the default initial value
         .SendMessageW hCtl, PBM_SETPOS, 0, 0
      CASE "TRACKBAR", "MSCTLS_TRACKBAR32"
         ' // Set the default range values
         .SendMessageW hCtl, TBM_SETRANGEMIN, CTRUE, 0
         .SendMessageW hCtl, TBM_SETRANGEMAX, CTRUE, 100
         ' // Set the default page size
         .SendMessageW hCtl, TBM_SETPAGESIZE, 0, 10
      CASE "UPDOWN", "MSCTLS_UPDOWN32"
         ' // Set the default base
         .SendMessageW hCtl, UDM_SETBASE, 10, 0
         ' // Set the default range values
         .SendMessageW hCtl, UDM_SETRANGE32, 100, 0
         ' // Set the default initial value
         .SendMessageW hCtl, UDM_SETPOS32, 0, 0
         ' // Correct for Windows using a default size for the updown control
         this.SetWindowPos hCtl, NULL, x, y, nWidth, nHeight, SWP_NOZORDER
      CASE "HSCROLLBAR", "VSCROLLBAR"
         ' // Initialize the scroll bar with default values
         DIM tsi AS SCROLLINFO
         tsi.cbSize = SIZEOF(tsi)
         tsi.fMask  = SIF_PAGE OR SIF_POS OR SIF_RANGE
         tsi.nMin   = 0
         tsi.nMax   = 100
         tsi.nPage  = 0
         tsi.nPos   = 0
         .SetScrollInfo hCtl, SB_CTL, @tsi, CTRUE
      CASE "TOOLBAR", "TOOLBARWINDOW32"
         ' // Set the button size
         DIM AS LONG nButtonWidth, nButtonHeight
         nButtonWidth = LOWORD(.SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0)) * m_rx
         nButtonHeight = HIWORD(.SendMessageW(hCtl, TB_GETBUTTONSIZE, 0, 0)) * m_ry
         .SendMessageW hCtl, TB_SETBUTTONSIZE, 0, MAKELONG(nButtonWidth, nButtonHeight)
         ' // Send this message for backward compatibility
         .SendMessageW hCtl, TB_BUTTONSTRUCTSIZE, SIZEOF(TBBUTTON), 0
      CASE "BITMAPLABEL"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = .LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = .LoadImageW(m_hInstance, wszTitle, IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN .SendMessageW(hCtl, STM_SETIMAGE, IMAGE_BITMAP, CAST(LPARAM, hImage))
      CASE "ICONLABEL"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = .LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = .LoadImageW(m_hInstance, wszTitle, IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN .SendMessageW(hCtl, STM_SETIMAGE, IMAGE_ICON, CAST(LPARAM, hImage))
      CASE "BITMAPBUTTON"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = .LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = .LoadImageW(m_hInstance, wszTitle, IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN .SendMessageW(hCtl, BM_SETIMAGE, IMAGE_BITMAP, CAST(LPARAM, hImage))
      CASE "ICONBUTTON"
         ' // Loads the image
         DIM hImage AS HANDLE, wID AS WORD, dwID AS DWORD
         IF LEFT(wszTitle, 1) = "#" THEN
            wID = VAL(MID(wszTitle, 2))
            dwID = MAKELONG(wID, 0)
            hImage = .LoadImageW(m_hInstance, CAST(LPCWSTR, CAST(ULONG_PTR, dwID)), IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         ELSE
            hImage = .LoadImageW(m_hInstance, wszTitle, IMAGE_ICON, 0, 0, LR_DEFAULTCOLOR)
         END IF
         IF hImage THEN .SendMessageW(hCtl, BM_SETIMAGE, IMAGE_ICON, CAST(LPARAM, hImage))
   END SELECT
   ' // Subclass the control if pWndProc is not null
   IF pWndProc <> NULL THEN
      IF uIdSubclass = &HFFFFFFFF THEN
         .SetPropW(hCtl, "OLDWNDPROC", CAST(HANDLE, .SetWindowLongPtrW(hCtl, GWLP_WNDPROC, CAST(LONG_PTR, pWndProc))))
      ELSE
         .SetWindowSubclass hCtl, CAST(SUBCLASSPROC, pWndProc), uIdSubclass, dwRefData
      END IF
   END IF
   FUNCTION = hCtl
END FUNCTION
' =====================================================================================

' =====================================================================================
' Gets/Sets the accelerator table handle
' =====================================================================================
PRIVATE PROPERTY CWindow.AccelHandle () AS HACCEL
   PROPERTY = m_hAccel
END PROPERTY
' =====================================================================================
' =====================================================================================
PRIVATE PROPERTY CWindow.AccelHandle (BYVAL hAccel AS HACCEL)
   IF m_hAccel THEN .DestroyAcceleratorTable(m_hAccel)
   IF UBOUND(m_rgAccelEntries) - LBOUND(m_rgAccelEntries) > -1 THEN ERASE m_rgAccelEntries
   m_hAccel = hAccel
END PROPERTY
' =====================================================================================
' =====================================================================================
' Adds an accelerator key to the table.
' =====================================================================================
PRIVATE SUB CWindow.AddAccelerator OVERLOAD (BYVAL fvirt AS UBYTE, BYVAL wKey AS WORD, BYVAL cmd AS WORD)
   REDIM PRESERVE m_rgAccelEntries(UBOUND(m_rgAccelEntries) + 1) AS ACCEL
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).fvirt = fvirt
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).key = wKey
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).cmd = cmd
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB CWindow.AddAccelerator OVERLOAD (BYVAL fvirt AS UBYTE, BYREF wszKey AS WSTRING, BYVAL cmd AS WORD)
   REDIM PRESERVE m_rgAccelEntries(UBOUND(m_rgAccelEntries) + 1) AS ACCEL
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).fvirt = fvirt
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).key = ASC(wszKey)
   m_rgAccelEntries(UBOUND(m_rgAccelEntries)).cmd = cmd
END SUB
' =====================================================================================
' =====================================================================================
' Creates the accelerator table.
' =====================================================================================
PRIVATE FUNCTION CWindow.CreateAcceleratorTable () AS HACCEL
   IF UBOUND(m_rgAccelEntries) - LBOUND(m_rgAccelEntries) = -1 THEN EXIT FUNCTION
   IF m_hAccel THEN .DestroyAcceleratorTable(m_hAccel)
   m_hAccel = .CreateAcceleratorTableW(CAST(LPACCEL, @m_rgAccelEntries(0)), UBOUND(m_rgAccelEntries) - LBOUND(m_rgAccelEntries) + 1)
   FUNCTION = m_hAccel
END FUNCTION
' =====================================================================================
' =====================================================================================
' Destroys the accelerator table.
' =====================================================================================
PRIVATE SUB CWindow.DestroyAcceleratorTable
   IF m_hAccel THEN .DestroyAcceleratorTable(m_hAccel)
   IF UBOUND(m_rgAccelEntries) - LBOUND(m_rgAccelEntries) > -1 THEN ERASE m_rgAccelEntries
   m_hAccel = NULL
END SUB
' =====================================================================================

#ifdef USEMDI
' =====================================================================================
' Optional. Sets de MDI class name
' =====================================================================================
PRIVATE PROPERTY CWindow.MDIClassName (BYREF wszClassName AS WSTRING)
   m_wszMDIClassName = wszClassName
END PROPERTY
' =====================================================================================

' =====================================================================================
' Returns the MDI client window handle
' =====================================================================================
PRIVATE PROPERTY CWindow.hwndClient () AS HWND
   PROPERTY = m_hwndClient
END PROPERTY
' =====================================================================================

' ========================================================================================
' Registers the MDI window class.
' Parameter:
' - lpfnWndProc = Pointer to the window procedure.
' Return Value:
'   An atom that uniquely identifies the class.
' Remarks:
'   The values of several of the members of this structure can be retrieved with
'   GetClassLong and set with SetClassLong.
' ========================================================================================
PRIVATE FUNCTION CWindow_RegisterMDIClass (BYREF wsClassName AS WSTRING, BYVAL hInstance AS HINSTANCE, BYVAL lpfnWndProc AS WNDPROC) AS WORD

   STATIC wAtom AS ATOM                     ' // Atom
   DIM    wcexw AS WNDCLASSEXW              ' // WNDCLASSEXW structure

   ' // Already registered?
   IF wAtom THEN
      FUNCTION = wAtom
      EXIT FUNCTION
   END IF

   ' // MDI window class name
   DIM  wszClassName AS WSTRING * 256
   IF LEN(wsClassName) THEN wszClassName = wsClassName ELSE wszClassName = "FBFrameClass"

   ' // Default handler
   IF lpfnWndProc = NULL THEN lpfnWndProc = @CWindowMDIProc

   ' // Fill the WNDCLASSEX structure
   WITH wcexw
      .cbSize        = SIZEOF(wcexw)
      .style         = CS_DBLCLKS
      .lpfnWndProc   = lpfnWndProc
      .cbClsExtra    = 0
      .cbWndExtra    = 0
      .hInstance     = hInstance
      .hCursor       = ..LoadCursorW(NULL, CAST(LPCWSTR, IDC_ARROW))
      .hbrBackground = CAST(HBRUSH, COLOR_APPWORKSPACE + 1)
      .lpszMenuName  = NULL
      .lpszClassName = @wszClassName
      .hIcon         = 0
      .hIconSm       = 0
   END WITH

   ' // Register the class
   wAtom = .RegisterClassExW(@wcexw)

   ' // Return the atom
   FUNCTION = wAtom

END FUNCTION
' ========================================================================================

' =====================================================================================
' Creates a MDI window
' =====================================================================================
PRIVATE FUNCTION CWindow.CreateMDIWindow ( _
   BYVAL cID AS LONG, _                  ' // Control identifier
   BYVAL x AS LONG, _                    ' // Horizontal position
   BYVAL y AS LONG, _                    ' // Vertical position
   BYVAL nWidth AS LONG, _               ' // Window width
   BYVAL nHeight AS LONG, _              ' // Window height
   BYVAL dwStyle AS DWORD, _             ' // Window style(s)
   BYVAL dwExStyle AS DWORD, _           ' // Extended style
   BYVAL hSubMenu AS HANDLE, _           ' // MDI window submenu
   BYVAL lpfnWndProc AS WNDPROC _        ' // Address of the callback function
   ) AS HWND                             ' // Window handle

   ' // Exit if the main window has not already been created
   IF m_hwnd = NULL THEN EXIT FUNCTION
   ' // Exit if the window has already been created
   IF m_hwndClient THEN EXIT FUNCTION
   ' // Default size values
   IF nWidth = 0 THEN
      nWidth = CW_USEDEFAULT
      IF x = 0 THEN
         x = CW_USEDEFAULT
      END IF
    END IF
   ' // Register the class
   DIM wMDIAtom AS WORD
   wMDIAtom = CWindow_RegisterMDIClass(m_wszMDIClassName, m_hInstance, lpfnWndProc)
   IF wMDIAtom = 0 THEN EXIT FUNCTION
   ' // Create the window
   DIM tccs AS CLIENTCREATESTRUCT
   ' // Note: If you use 1 as the identifier for the first child, the
   ' // identifiers of the windows will conflict with the dialog box
   ' // identifiers, e.g. the first window with ID_OK, the 2nd with
   ' // IDCANCEL, etc.
   tccs.idFirstChild = 13
   tccs.hWindowMenu  = hSubMenu
   IF dwStyle = 0 THEN dwStyle = WS_VISIBLE OR WS_HSCROLL OR WS_VSCROLL OR WS_CLIPCHILDREN OR WS_CLIPSIBLINGS
   IF dwExStyle = 0 THEN dwExStyle = WS_EX_CLIENTEDGE
   ' // Make sure that the control has the WS_CHILD style
   dwStyle = dwStyle OR WS_CHILD
   m_hwndClient = .CreateWindowExW(dwExStyle, "MDICLIENT", "", dwStyle, _
                  x * m_rx, y * m_ry, nWidth * m_rx, nHeight * m_ry, m_hwnd, CAST(HMENU, CAST(LONG_PTR, cID)), m_hInstance, CAST(LPVOID, @tccs))
   IF m_hwndClient = NULL THEN EXIT FUNCTION
   ' // Resize the window
   .SendMessageW m_hwnd, WM_SIZE, SIZE_RESTORED, MAKELONG (this.CLientWidth, this.ClientHeight)
   FUNCTION = m_hwndClient
END FUNCTION
' =====================================================================================

' ========================================================================================
' Default CWindow MDI callback function.
' ========================================================================================
PRIVATE FUNCTION CWindowMDIProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE AS CONST uMsg

'      CASE WM_CREATE
         ' // Retrieve a reference to the CWindow class
'         DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
'         DIM pWindow AS CWindow PTR = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
'         EXIT FUNCTION

'      CASE WM_MDIACTIVATE
'         IF lParam = hwnd THEN
'         END IF
'         EXIT FUNCTION

      CASE WM_SETFOCUS
         ' // Set the keyboard focus to the first control that is
         ' // visible, not disabled, and has the WS_TABSTOP style
         .SetFocus .GetNextDlgTabItem(hwnd, NULL, FALSE)

'      CASE WM_SIZE
'         IF wParam <> SIZE_MINIMIZED THEN
            ' Resize the window and its controls
'         END IF

'      CASE WM_DESTROY
'         ' // Do cleanunp if needed, such removing properties attached
'         ' // to the MDI child window.
'         EXIT FUNCTION

   END SELECT

   ' // The DefMDIChildProc function provides default processing for any window
   ' // message that the window procedure of a multiple-document interface (MDI)
   ' // child window does not process. A window message not processed by the window
   ' // procedure must be passed to the DefMDIChildProcW function, not to the
   ' // DefWindowProcW function.
   FUNCTION = .DefMDIChildProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' =====================================================================================
' Destroys all MDI child windows.
' =====================================================================================
PRIVATE FUNCTION CWindow_CloseEnumProc ( _
   BYVAL hwnd   AS HWND, _                  ' // Handle of enumerated child window
   BYVAL lParam AS LPARAM _                 ' // Unused
   ) AS LONG

   ' // Skip icon title windows
   IF .GetWindow(hwnd, GW_OWNER) <> 0 THEN
      FUNCTION = CTRUE
      EXIT FUNCTION
   END IF

   .SendMessageW GetParent(hwnd), WM_MDIRESTORE, CAST(WPARAM, hwnd), 0
   IF .SendMessageW(hwnd, WM_QUERYENDSESSION, 0, 0) = 0 THEN
      FUNCTION = CTRUE
      EXIT FUNCTION
   END IF

   .SendMessageW .GetParent(hwnd), WM_MDIDESTROY, CAST(WPARAM, hwnd), 0

   FUNCTION = CTRUE

END FUNCTION
#endif
' =====================================================================================

' ########################################################################################
'                                 *** CTABPAGE CLASS ***
' ########################################################################################

' ========================================================================================
' CTabPage class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTabPage

   DIM i AS LONG, nCount AS LONG, hTab AS HWND, tci AS TCITEMW
   ' // Get the handle of the tab control
   hTab = .GetParent(m_hTabPage)
   IF hTab THEN
      ' // Get the number of items
      nCount = .SendMessageW(hTab, TCM_GETITEMCOUNT, 0, 0)
      ' // Ask to return the value of the lParam member
      tci.mask = TCIF_PARAM
      ' // Get information of the items
      FOR i = 0 TO nCount - 1
         IF .SendMessageW(hTab, TCM_GETITEMW, i, CAST(lParam, @tci)) THEN
            IF tci.lParam = @this THEN
               ' // Delete the tab item
               .SendMessageW(hTab, TCM_DELETEITEM, i, 0)
               ' // Destroy the window of the tab page
               .DestroyWindow m_hTabPage
               ' // Set the focus in the first tab
               IF nCount > 1 THEN .SendMessageW hTab, TCM_SETCURFOCUS, 0, 0
               EXIT FOR
            END IF
         END IF
      NEXT
   ELSE
      ' // Destroy the window of the tab page
      .DestroyWindow m_hTabPage
   END IF

END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Adds a tab page and creates a generic window that will be associated with the page.
' Parameters:
' - hTab        = [in] A handle to the tab control.
' - nPage       = [in] The zero based position of the page to be inserted.
' - wszTitle    = [in, opt] The text to be displayed of the tab area.
' - nImage      = [in, opt] Index in the tab control's image list, or -1 if there is no image for the tab.
' - lpfnWndProc = [in, opt] Address of the window callback procedure.
' Return Value:
' - The handle of the new tab page.
' ========================================================================================
PRIVATE FUNCTION CTabPage.InsertPage (BYVAL hTab AS HWND, BYVAL nPage AS LONG, BYREF wszTitle AS WSTRING = "", _
   BYVAL nImage AS LONG = -1, BYVAL lpfnWndProc AS WNDPROC = NULL, _
   BYVAL dwStyle AS DWORD = WS_CLIPSIBLINGS OR WS_CLIPCHILDREN, BYVAL dwExStyle AS DWORD = 0) AS HWND

   DIM AS LONG x, y, nWidth, nHeight, cItems
   DIM rc AS RECT, tci AS TCITEMW

   IF IsWindow(hTab) = NULL THEN EXIT FUNCTION
   IF lpfnWndProc = NULL THEN lpfnWndProc = CAST(HANDLE, .GetClassLongPtrW(this.hWindow, GWLP_WNDPROC))

   dwStyle = dwStyle OR WS_CHILD
   dwExStyle = dwExStyle OR WS_EX_CONTROLPARENT
   this.ClassStyle = CS_DBLCLKS

   cItems = .SendMessageW(hTab, TCM_GETITEMCOUNT, 0, 0)
   IF nPage < 0 OR nPage > cItems THEN nPage = cItems
   tci.mask    = TCIF_TEXT OR TCIF_IMAGE OR TCIF_PARAM
   tci.pszText = @wszTitle
   tci.iImage  = nImage
   tci.lParam  = CAST(lParam, @this)
   .SendMessageW hTab, TCM_INSERTITEMW, nPage, CAST(lParam, @tci)
   .GetWindowRect(hTab, @rc)
   .SendMessageW hTab, TCM_ADJUSTRECT, FALSE, CAST(lParam, @rc)
   .MapWindowPoints NULL, hTab, CAST(LPPOINT, @rc), 2
   ' // Adjust for High DPI because create will resize the values
   rc.Left   /= this.rxRatio
   rc.Right  /= this.rxRatio
   rc.Top    /= this.ryRatio
   rc.Bottom /= this.ryRatio
   ' // Calculate coordinates and size
   x = rc.Left
   y = rc.Top
   nWidth  = max(1, rc.Right - rc.Left)
   nHeight = max(1, rc.Bottom - rc.Top)
   m_hTabPage = this.Create(hTab, wszTitle, lpfnWndProc, x, y, nWidth, nHeight, dwStyle, dwExStyle)
   FUNCTION = m_hTabPage

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the window handle of the tab page
' ========================================================================================
PRIVATE PROPERTY CTabPage.hTabPage () AS HWND
   PROPERTY = m_hTabPage
END PROPERTY
' ========================================================================================

' ########################################################################################
'                                *** HELPER FUNCTIONS ***
' ########################################################################################

' =====================================================================================
' Resizes all the tab pages associated with a tab control
' =====================================================================================
PRIVATE FUNCTION AfxResizeTabPages (BYVAL hTab AS HWND) AS BOOLEAN
   IF hTab = NULL THEN EXIT FUNCTION
   DIM nCount AS LONG, i AS LONG, tci AS TCITEMW, pTabPage AS CTabPage PTR
   ' // Get the number of items
   nCount = .SendMessageW(hTab, TCM_GETITEMCOUNT, 0, 0)
   IF nCount = 0 THEN EXIT FUNCTION
   ' // Ask to return the value of the lParam member
   tci.mask = TCIF_PARAM
   ' // Get information of the items
   FOR i = 0 TO nCount - 1
      IF .SendMessageW(hTab, TCM_GETITEMW, i, CAST(LPARAM, @tci)) THEN
         IF tci.lParam THEN
            pTabPage = CAST(CTabPage PTR, tci.lParam)
            ' // Retrieve the size of the tab control window
            DIM rcParent AS RECT
            .GetWindowRect(hTab, @rcParent)
            ' // Calculates the tab control's display area given its window rectangle
            .SendMessageW(hTab, TCM_ADJUSTRECT, FALSE, CAST(LPARAM, @rcParent))
            ' // Convert to window coordinates
            .MapWindowPoints(NULL, hTab, CAST(LPPOINT, @rcParent), 2)
            ' // Move the tab page
            .MoveWindow(pTabPage->hTabPage, rcParent.Left, rcParent.Top, _
               rcParent.Right - rcParent.Left, rcParent.Bottom - rcParent.Top, CTRUE)
         END IF
      END IF
   NEXT
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detroys the specified tab page
' Parameters:
' - hTab = Handle to the tab control.
' - idx  = Zero based index of the tab.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxDestroyTabPage (BYVAL hTab AS HWND, BYVAL idx AS LONG) AS BOOLEAN
   IF hTab = NULL THEN EXIT FUNCTION
   DIM tci AS TCITEMW, pTabPage AS CTabPage PTR
   ' // Ask to return the value of the lParam member
   tci.mask = TCIF_PARAM
   IF SendMessageW(hTab, TCM_GETITEMW, idx, CAST(LPARAM, @tci)) = 0 THEN EXIT FUNCTION
   IF tci.lParam THEN
      pTabPage = CAST(CTabPage PTR, tci.lParam)
      IF pTabPage THEN
         Delete pTabPage
         FUNCTION = TRUE
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detroys all the tab pages
' Parameter:
' - hTab = Handle to the tab control.
' Return value: TRUE or FALSE.
' ========================================================================================
PRIVATE FUNCTION AfxDestroyAllTabPages (BYVAL hTab AS HWND) AS BOOLEAN
   IF hTab = NULL THEN EXIT FUNCTION
   DIM tci AS TCITEMW, pTabPage AS CTabPage PTR
   ' // Ask to return the value of the lParam member
   tci.mask = TCIF_PARAM
   DO
      IF SendMessageW(hTab, TCM_GETITEMW, 0, CAST(LPARAM, @tci)) = 0 THEN EXIT DO
      IF tci.lParam THEN
         pTabPage = CAST(CTabPage PTR, tci.lParam)
         IF pTabPage THEN
            Delete pTabPage
            FUNCTION = TRUE
         END IF
      END IF
   LOOP
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWindow class given the handle of its associated window handle.
' To retrieve it from the handle of any of its child windows or controls, use AfxCWindowOwnerPtr.
' ========================================================================================
PRIVATE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL hwnd AS HWND) AS CWindow PTR
   FUNCTION = CAST(CWindow PTR, .GetWindowLongPtrW(hwnd, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWindow class given the a pointer to the CREATESTRUCT structure.
' ========================================================================================
PRIVATE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL lParam AS LPARAM) AS CWindow PTR
   DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
   FUNCTION = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWindow class given the a pointer to the CREATESTRUCT structure.
' ========================================================================================
PRIVATE FUNCTION AfxCWindowPtr OVERLOAD (BYVAL pCreateStruct AS CREATESTRUCT PTR) AS CWindow PTR
   FUNCTION = CAST(CWindow PTR, pCreateStruct->lpCreateParams)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CWindow class given the handle of the window created with it
' or the handle of any of it's children.
' ========================================================================================
PRIVATE FUNCTION AfxCWindowOwnerPtr OVERLOAD (BYVAL hwnd AS HWND) AS CWindow PTR
   IF hwnd = NULL THEN EXIT FUNCTION
   DIM hRootOwner AS .HWND = .GetAncestor(hwnd, GA_ROOTOWNER)
   IF hRootOwner = NULL THEN EXIT FUNCTION
   FUNCTION = CAST(CWindow PTR, .GetWindowLongPtrW(hRootOwner, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CTabPage class given the handle of the tab control to which the
' tab page is associated and the zero-based tab index. If nTabIdx is ommited, the function
' will return the pointer of the selected tab, if any.
' ========================================================================================
PRIVATE FUNCTION AfxCTabPagePtr OVERLOAD (BYVAL hTab AS HWND, BYVAL idx AS LONG) AS CTabPage PTR
   IF hTab = NULL THEN EXIT FUNCTION
   IF idx = -1 THEN idx = SendMessageW(hTab, TCM_GETCURSEL, 0, 0)
   IF idx = -1 THEN EXIT FUNCTION   ' No tab selected
   ' // Ask to return the value of the lParam member
   DIM tci AS TCITEMW
   tci.mask = TCIF_PARAM
   IF .SendMessageW(hTab, TCM_GETITEMW, idx, CAST(LPARAM, @tci)) THEN
      IF tci.lParam THEN
         FUNCTION = CAST(CTabPage PTR, tci.lParam)
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CScrollWindow class given the handle of the tab control to which the
' tab page is associated and the zero-based tab index. If nTabIdx is ommited, the function
' will return the pointer associated to the selected tab, if any.
' ========================================================================================
PRIVATE FUNCTION AfxScrollTabPagePtr (BYVAL hTab AS HWND, BYVAL idx AS LONG) AS CScrollWindow PTR
   FUNCTION = AfxCTabPagePtr(hTab, idx)->ScrollWindowPtr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CScrollWindow class given the handle of the window attached to it.
' ========================================================================================
PRIVATE FUNCTION AfxScrollWindowPtr (BYVAL hwnd AS HWND) AS CScrollWindow PTR
   FUNCTION = cast(CWindow PTR, GetWindowLongPtr(hwnd, 0))->ScrollWindowPtr
END FUNCTION
' ========================================================================================

END NAMESPACE

