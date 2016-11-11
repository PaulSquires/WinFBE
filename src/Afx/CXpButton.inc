' ########################################################################################
' Microsoft Windows
' File: CXpButton.inc
' Contents: Themed button control
' Draws a themed button in Windows XP or superior and a classic button in the other Windows versions.
' Do not use it with Windows 95 because the TrackMouseEvent is not available in that OS.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/tmschema.bi"
#include once "win/uxtheme.bi"
#include once "Afx/CWindow.inc"

NAMESPACE Afx

' ========================================================================================
' Image class
' ========================================================================================
CONST XPBI_NORMAL = 1
CONST XPBI_HOT = 2
CONST XPBI_DISABLED = 3
' ========================================================================================

' ========================================================================================
' Image position
' ========================================================================================
CONST XPBI_NONE = &H0                                   ' // No image
CONST XPBI_LEFT = &H1                                   ' // Left (default)
CONST XPBI_RIGHT = &H2                                  ' // Right
CONST XPBI_CENTER = &H4                                 ' // Center
CONST XPBI_VCENTER = &H8                                ' // Vertically centered
CONST XPBI_TOP = &H10                                   ' // Top
CONST XPBI_BOTTOM = &H20                                ' // Bottom
CONST XPBI_ABOVE = XPBI_TOP OR XPBI_CENTER              ' // Above the text
CONST XPBI_BELOW = XPBI_BOTTOM OR XPBI_CENTER           ' // Below the text
CONST XPBI_CENTERCENTER = XPBI_CENTER OR XPBI_VCENTER   ' // Center-center (no text)
' ========================================================================================

' ========================================================================================
' CXpButton class
' ========================================================================================
TYPE CXpButton

   Private:
      m_hCtl              AS HWND          ' // Button handle
      m_fState            AS LONG          ' // Button state
      m_dwTextFlags       AS DWORD         ' // Format text flags
      m_hCursor           AS HCURSOR       ' // Cursor handle
      m_hFont             AS HFONT         ' // Default font handle
      m_hImage            AS HANDLE        ' // Image handle
      m_ImageType         AS LONG          ' // Image type
      m_ImageWidth        AS LONG          ' // Image width
      m_ImageHeight       AS LONG          ' // Image height
      m_ImagePos          AS LONG          ' // Image position
      m_ImageMargin       AS LONG          ' // Image margin (in pixels)
      m_TextMargin        AS LONG          ' // Text margin (in pixels)
      m_hHotImage         AS HANDLE        ' // Hot image handle
      m_hDisabledImage    AS HANDLE        ' // Disabled image handle
      m_bIsToggle         AS LONG          ' // Button is a toggle button
      m_bToggled          AS LONG          ' // Toggle flag
      m_bIsThemed         AS LONG          ' // Application is themed
      m_bClkDown          AS LONG          ' // Button received the button down message

   Public:
      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hWindow () AS HWND
      DECLARE SUB Redraw
      DECLARE SUB SetFont (BYVAL hFont AS HFONT, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetFont () AS HFONT
      DECLARE SUB SetTextFormat (BYVAL dwTextFlags AS DWORD, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetTextFormat () AS DWORD
      DECLARE SUB SetCursor (BYVAL hCursor AS HCURSOR)
      DECLARE FUNCTION GetCursor () AS HCURSOR
      DECLARE SUB SetToggle (BYVAL fState AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetToggle () AS LONG
      DECLARE SUB SetToggleState (BYVAL fToggle AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetToggleState () AS LONG
      DECLARE SUB SetImageSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetImageWidth (BYVAL nWidth AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetImageWidth () AS LONG
      DECLARE SUB SetImageHeight (BYVAL nHeight AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetImageHeight () AS LONG
      DECLARE SUB SetImagePos (BYVAL nPos AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetImagePos () AS LONG
      DECLARE SUB SetImageMargin (BYVAL nMargin AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetImageMargin () AS LONG
      DECLARE SUB SetTextMargin (BYVAL nMargin AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetTextMargin () AS LONG
      DECLARE FUNCTION GetButtonState () AS LONG
      DECLARE FUNCTION GetImageType () AS LONG
      DECLARE FUNCTION IsThemed () AS LONG
      DECLARE SUB EnableTheming
      DECLARE SUB DisableTheming
      DECLARE FUNCTION GetImage (BYVAL ImageState AS LONG) AS HANDLE
      DECLARE SUB UxDrawPushButton (BYVAL hDC AS HDC)
      DECLARE SUB SetImage (BYVAL hImage AS HANDLE, BYVAL ImageType AS LONG, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetIcon (BYVAL hIcon AS HICON, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetBitmap (BYVAL hBitmap AS HBITMAP, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetIconFromFile (BYVAL pwszPath AS WSTRING PTR, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetBitmapFromFile (BYVAL pwszPath AS WSTRING PTR, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE STATIC FUNCTION CXPButtonProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CXpButton class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CXpButton (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)

   ' // Register the class
   DIM wAtom AS ATOM
   DIM wcexw AS WNDCLASSEXW
   DIM wszClassName AS WSTRING * 260 = "AFX_XPBUTTON"
   IF .GetClassInfoExW(.GetModuleHandleW(NULL), @wszClassName, @wcexw) = 0 THEN
      ' // Fill the WNDCLASSEXW structure
      WITH wcexw
         .cbSize        = SIZEOF(wcexw)
         .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
         .lpfnWndProc   = @CXPButtonProc
         .cbClsExtra    = 0
         .cbWndExtra    = SIZEOF(HANDLE)
         .hInstance     = ..GetModuleHandleW(NULL)
         .hCursor       = ..LoadCursorW(NULL, CAST(LPCWSTR, IDC_ARROW))
         .hbrBackground = NULL
         .lpszMenuName  = NULL
         .lpszClassName = @wszClassName
         .hIcon         = NULL
         .hIconSm       = NULL
      END WITH
      wAtom = .RegisterClassExW(@wcexw)
   END IF

   ' // Checks if the application is themed
   IF .IsAppThemed AND .IsThemeActive THEN m_bIsThemed = CTRUE
   ' // Default values
   m_dwTextFlags = DT_CENTER OR DT_VCENTER OR DT_SINGLELINE
   m_ImagePos = XPBI_LEFT OR XPBI_VCENTER
   m_ImageMargin = 4
   m_TextMargin = 4
   m_hCursor = .LoadCursor(NULL, IDC_ARROW)

   ' // Create the control
   IF dwStyle = 0 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER
   IF pWindow THEN m_hCtl = pWindow->AddControl(wszClassName, pWindow->hWindow, cID, wszTitle, x, y, nWidth, nHeight, dwStyle, dwExStyle, lpParam)
   IF m_hCtl THEN
      .SetWindowLongPtrW m_hCtl, 0, CAST(LONG_PTR, @this)
      ' // Set the same font used by the parent
      DIM lfw AS LOGFONTW
      IF pWindow->Font THEN
         IF .GetObjectW(pWindow->Font, SIZEOF(lfw), @lfw) THEN m_hFont = CreateFontIndirectW(@lfw)
      END IF
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CXpButton class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CXpButton
   ' // Deletes the images
   IF m_hImage THEN .DeleteObject m_hImage
   IF m_hHotImage THEN .DeleteObject m_hHotImage
   IF m_hDisabledImage THEN .DeleteObject m_hDisabledImage
   ' // Deletes the font
   IF m_hFont THEN .DeleteObject m_hFont
END DESTRUCTOR
' ========================================================================================


' ========================================================================================
' Window procedure
' ========================================================================================
PRIVATE FUNCTION CXpButton.CXPButtonProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM pButton AS CXpButton PTR
   DIM hDc AS HDC, pPaint AS PAINTSTRUCT
   DIM rc AS RECT, pt AS POINT

   SELECT CASE uMsg

      CASE WM_CREATE
         EXIT FUNCTION

      CASE WM_GETDLGCODE
         DIM pMsg AS MSG PTR, dwMsg AS DWORD
         IF lParam = 0 THEN
            ' // Enables mnemonics
            FUNCTION = DLGC_BUTTON OR DLGC_UNDEFPUSHBUTTON
            EXIT FUNCTION
         ELSE
            ' // We want to handle the return key
            pMsg = CAST(MSG PTR, lParam)
            IF pMsg > 0 THEN
               dwMsg = pMsg->Message
               IF dwMsg = WM_KEYDOWN OR dwMsg = WM_CHAR THEN
                  IF pMsg->wParam = VK_RETURN THEN
                     FUNCTION = DLGC_WANTALLKEYS
                     EXIT FUNCTION
                  END IF
               END IF
            END IF
         END IF
         EXIT FUNCTION

      CASE WM_ERASEBKGND
         ' // Don't erase the background to avoid flicker
         FUNCTION = 1
         EXIT FUNCTION

      CASE WM_PAINT
         ' // Draws the button
         hDc = .BeginPaint(hwnd, @pPaint)
         pButton = CAST(CXpButton PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pButton THEN pButton->UxDrawPushButton(hDc)
         .EndPaint hwnd, @pPaint
         EXIT FUNCTION

      CASE WM_PRINTCLIENT
         ' // Draws the button
         pButton = CAST(CXpButton PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pButton THEN pButton->UxDrawPushButton(CAST(.HDC, wParam))
         EXIT FUNCTION

      CASE WM_COMMAND, WM_NOTIFY
         ' // Forwards the message to the parent window
         .SendMessageW GetParent(hwnd), uMsg, wParam, lParam

      CASE WM_LBUTTONDBLCLK
        ' // Forwards this message for rapid button clicking...
        .SendMessageW hwnd, WM_LBUTTONDOWN, wParam, lParam
        EXIT FUNCTION

      CASE WM_LBUTTONDOWN
         ' // Redraws the button in pushed state
         pButton = CAST(CXpButton PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pButton = NULL THEN EXIT FUNCTION
         .GetWindowRect pButton->hWindow, @rc
         .GetCursorPos @pt
         IF .PtInRect(@rc, pt) THEN
            pButton->m_bClkDown = CTRUE
            IF pButton->m_bToggled = FALSE THEN
               pButton->m_bToggled = CTRUE
            ELSE
               pButton->m_bToggled = FALSE
            END IF
            pButton->m_fState = BST_PUSHED
            .SetCapture pButton->hWindow
            IF .GetFocus <> pButton->hWindow THEN .SetFocus pButton->hWindow
            pButton->Redraw
         END IF
         EXIT FUNCTION

      CASE WM_LBUTTONUP
         ' // Redraws the button in normal state
         pButton = CAST(CXpButton PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pButton = NULL THEN EXIT FUNCTION
         .ReleaseCapture
         .GetWindowRect pButton->hWindow, @rc
         .GetCursorPos @pt
         pButton->m_fState = 0
         pButton->Redraw
         ' // Forwards the message to the parent window
         IF .IsWindowEnabled(pButton->hWindow) THEN
            IF .PtInRect(@rc, pt) THEN
               IF pButton->m_bClkDown THEN .SendMessageW .GetParent(pButton->hWindow), WM_COMMAND, MAKELONG(GetDlgCtrlId(pButton->hWindow), BN_CLICKED), CAST(.LPARAM, pButton->hWindow)
            END IF
         END IF
         pButton->m_bClkDown = FALSE
         EXIT FUNCTION

      CASE WM_KEYDOWN
         ' // Redraws the button in pushed state
         pButton = CAST(CXpButton PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pButton = NULL THEN EXIT FUNCTION
         SELECT CASE wParam
            CASE VK_RETURN, VK_SPACE
               ' // Toggles the toggle flag
               IF pButton->m_bToggled = FALSE THEN
                  pButton->m_bToggled = CTRUE
               ELSE
                  pButton->m_bToggled = FALSE
               END IF
               ' // Redraws the button in pushed state
               pButton->m_fState = BST_PUSHED
               pButton->Redraw
         END SELECT
         EXIT FUNCTION

      CASE WM_KEYUP
         ' // Redraws the button in normal state
         pButton = CAST(CXpButton PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pButton = NULL THEN EXIT FUNCTION
         IF pButton->m_fState = BST_PUSHED THEN
            pButton->m_fState = 0
            pButton->Redraw
            ' // Forwards the message to the parent window
            IF .IsWindowEnabled(pButton->hWindow) THEN
               .SendMessageW .GetParent(pButton->hWindow), WM_COMMAND, MAKELONG(GetDlgCtrlId(pButton->hWindow), BN_CLICKED), CAST(.LPARAM, pButton->hWindow)
            END IF
         END IF
         EXIT FUNCTION

      CASE WM_ENABLE
         ' // Redraws the button
         .InvalidateRect hwnd, NULL, 0
         .UpdateWindow hwnd
         EXIT FUNCTION

      CASE WM_SETFOCUS
         ' // Redraws the button
         .InvalidateRect hwnd, NULL, 0
         .UpdateWindow hwnd
         EXIT FUNCTION

      CASE WM_KILLFOCUS
         ' // Redraws the button
         .InvalidateRect hwnd, NULL, 0
         .UpdateWindow hwnd
         EXIT FUNCTION

      CASE WM_SETCURSOR
         ' // Establishes the cursor shape
         pButton = CAST(CXpButton PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pButton = NULL THEN EXIT FUNCTION
         IF pButton->GetCursor THEN
            .SetCursor pButton->GetCursor
            FUNCTION = CTRUE
            EXIT FUNCTION
         END IF

      CASE WM_MOUSEMOVE
         ' // Tracks the mouse movement and stores the hot state
         DIM trackMouse AS TRACKMOUSEEVENT
         IF .GetPropW(hwnd, "HOT") = 0 THEN
            trackMouse.cbSize = SIZEOF(trackMouse)
            trackMouse.dwFlags = TME_LEAVE
            trackMouse.hwndTrack = hwnd
            trackMouse.dwHoverTime = 1
            .TrackMouseEvent(@trackMouse)
            .SetPropW hwnd, "HOT", CAST(HANDLE, CTRUE)
            .InvalidateRect hwnd, NULL, 0
            .UpdateWindow hwnd
         END IF
         EXIT FUNCTION

      CASE WM_MOUSELEAVE
         ' // Removes the hot state and redraws the button
         .RemovePropW hwnd, "HOT"
         .InvalidateRect hwnd, NULL, 0
         .UpdateWindow hwnd
         EXIT FUNCTION

   END SELECT

   ' // Default processing for other messages.
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the button
' ========================================================================================
PRIVATE FUNCTION CXpButton.hWindow () AS HWND
   FUNCTION = m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Redraws the button
' ========================================================================================
PRIVATE SUB CXpButton.Redraw
   .InvalidateRect m_hCtl, NULL, 0
   .UpdateWindow m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Sets the font for the button
' Parameters:
' - hFont = Font handle
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetFont (BYVAL hFont AS HFONT, BYVAL fRedraw AS LONG = FALSE)
   IF hFont = NULL THEN EXIT SUB
   IF m_hFont THEN .DeleteObject m_hFont
   m_hFont = hFont
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Returns the font handle
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetFont () AS HFONT
   FUNCTION = m_hFont
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text format
' Parameters:
' - dwTextFlags = Format values (default DT_CENTER OR DT_VCENTER OR DT_SINGLELINE)
'   - DT_BOTTOM = Renders the text string at the bottom of the display rectangle.
'     This value is used only with the DT_SINGLELINE value.
'   - DT_CALCRECT = Determines the width and height of the display rectangle.
'   - DT_CENTER = Centers text horizontally in the display rectangle.
'   - DT_EDITCONTRO = Duplicates the text-displaying characteristics of a multiline edit
'     control. Specifically, the average character width is calculated in the same manner
'     as for an edit control, and the function does not display a partially visible last
'     line.
'   - DT_END_ELLIPSIS = Truncates a text string that is wider than the display rectangle
'     and adds an ellipsis to indicate the truncation. The string is not modified unless
'     the DT_MODIFYSTRING flag is specified.
'   - DT_EXPANDTABS = Expands tab characters. The default number of characters per tab
'     is eight. The DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS, and DT_END_ELLIPSIS values cannot
'     be used with the DT_EXPANDTABS value.
'   - DT_EXTERNALLEADING = Includes the external leading of a font in the line height.
'     Normally, external leading is not included in the height of a line of text.
'   - DT_HIDEPREFIX = Ignores the prefix character & in the text. The letter that follows
'     is not underlined, but other prefix characters are still processed. For example:
'        input string: "A&bc&&d"
'        normal: "Abc&d"
'        DT_HIDEPREFIX: "Abc&d"
'   - DT_LEFT = Aligns text to the left.
'   - DT_MODIFYSTRING = Modifies a string to match the displayed text. This value has
'     no effect unless DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is specified.
'   - DT_NOCLIP = Draws the text string without clipping the display rectangle.
'   - DT_NOFULLWIDTHCHARBREAK = Prevents a line break at a double-byte character set
'     (DBCS), so that the line-breaking rule is equivalent to single-byte character set
'     (SBCS). This can be used, for example, to make icon labels written in Korean text
'     more readable. This value has no effect unless DT_WORDBREAK is specified.
'   - DT_NOPREFIX = Turns off processing of prefix characters. Normally, DrawThemeText
'     interprets the prefix character & as a directive to underscore the character that
'     follows, and the prefix characters && as a directive to print a single &. By
'     specifying DT_NOPREFIX, this processing is turned off. For example:
'        input string: "A&bc&&d"
'        normal: "Abc&d"
'        DT_NOPREFIX: "A&bc&&d"
'   - DT_PATH_ELLIPSIS = Replaces characters in the middle of text with an ellipsis so
'     that the result fits in the display rectangle. If the string contains backslash (\)
'     characters, DT_PATH_ELLIPSIS preserves as much as possible of the text after the
'     last backslash. The string is not modified unless the DT_MODIFYSTRING flag is
'     specified.
'   - DT_PREFIXONLY = Draws only an underline at the position of the character following
'     the prefix character &. Normally DrawThemeText interprets the & as a directive to
'     underline the character that follows and the prefix characters && as a directive to
'     print a single &. By specifying DT_PREFIXONLY, no characters are drawn, only an
'     underline. White spaces are placed in the positions where characters would normally
'     appear. For example:
'        input string: "A&bc&&d"
'        normal: "Abc&d"
'        DT_PREFIXONLY: " _   "
'   - DT_RIGHT = Aligns text to the right.
'   - DT_RTLREADING = Lays out text in right-to-left order for bidirectional text, for
'     example, text in a Hebrew or Arabic font. The default direction for text is
'     left-to-right.
'   - DT_SINGLELINE = Displays text on a single line. Carriage returns and line feeds do
'     not break the line.
'   - DT_TABSTOPS = Sets tab stops.
'   - DT_TOP = Renders the text at the top of the display rectangle.
'   - DT_VCENTER = Centers text vertically. This value is used only with the DT_SINGLELINE
'     value.
'   - DT_WORDBREAK = Breaks lines between words if a word would extend past the edge of
'     the display rectangle. A carriage return/line feed (CR/LF) sequence also breaks the
'     line.
'   - DT_WORD_ELLIPSIS = Truncates any word that does not fit in the display rectangle
'     and adds an ellipsis.
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetTextFormat (BYVAL dwTextFlags AS DWORD, BYVAL fRedraw AS LONG = FALSE)
   m_dwTextFlags = dwTextFlags
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Returns the text format
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetTextFormat () AS DWORD
   FUNCTION = m_dwTextFlags
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the cursor for the button
' Parameter:
' - hCursor = Cursor handle
' ========================================================================================
PRIVATE SUB CXpButton.SetCursor (BYVAL hCursor AS HCURSOR)
   IF m_hCursor THEN DestroyCursor m_hCursor
   m_hCursor = hCursor
END SUB
' ========================================================================================

' ========================================================================================
' Gets the cursor for the button
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetCursor () AS HCURSOR
   FUNCTION = m_hCursor
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets button to toggle state (TRUE) or to pushbutton state (FALSE)
' Parameters:
' - fToggle = Toggle state (TRUE or FALSE)
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetToggle (BYVAL fToggle AS LONG, BYVAL fRedraw AS LONG = FALSE)
   IF fToggle <> 0 THEN fToggle = CTRUE
   m_bIsToggle = fToggle
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Gets the button toggle state
' Returns CTRUE if the button is toggled or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetToggle () AS LONG
   FUNCTION = m_bIsToggle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set toggle state to pushed (TRUE) or unpushed (FALSE)
' Parameters:
' - fState = Toggled state (TRUE or FALSE)
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetToggleState (BYVAL fState AS LONG, BYVAL fRedraw AS LONG = FALSE)
   IF fState <> 0 THEN fState = -1
   m_bToggled = fState
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Returns TRUE if button is toggled (down) or FALSE otherwise
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetToggleState () AS LONG
   FUNCTION = m_bToggled
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the size of the image
' Parameters:
' - nWidth = Width of the image, in pixels.
' - nHeight = Height of the image, in pixels.
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetImageSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_ImageWidth = nWidth
   m_ImageHeight = nHeight
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Sets the width of the image
' Parameters:
' - nWidth = Width of the image, in pixels.
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetImageWidth (BYVAL nWidth AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_ImageWidth = nWidth
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Gets the width of the image.
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetImageWidth () AS LONG
   FUNCTION = m_ImageWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the height of the image
' Parameters:
' - nHeight = Height of the image, in pixels.
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetImageHeight (BYVAL nHeight AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_ImageHeight = nHeight
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Gets the height of the image.
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetImageHeight () AS LONG
   FUNCTION = m_ImageHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the position of the image
' Parameters:
' - nPos = Position.
'     %XPBI_NONE = &H0                                    ' // No image
'     %XPBI_LEFT = &H1                                    ' // Left (default)
'     %XPBI_RIGHT = &H2                                   ' // Right
'     %XPBI_CENTER = &H4                                  ' // Center
'     %XPBI_VCENTER = &H8                                 ' // Vertically centered
'     %XPBI_TOP = &H10                                    ' // Top
'     %XPBI_BOTTOM = &H20                                 ' // Bottom
'     %XPBI_ABOVE = %XPBI_TOP OR %XPBI_CENTER             ' // Above the text
'     %XPBI_BELOW = %XPBI_BOTTOM OR %XPBI_CENTER          ' // Below the text
'     %XPBI_CENTERCENTER = %XPBI_CENTER OR %XPBI_VCENTER  ' // Center-center (no text)
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetImagePos (BYVAL nPos AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_ImagePos = nPos
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Gets the position of the image.
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetImagePos () AS LONG
   FUNCTION = m_ImagePos
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the image margin
' Parameters:
' - nMargin = Margin (in pixels)
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetImageMargin (BYVAL nMargin AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_ImageMargin = nMargin
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Gets the margin of the image.
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetImageMargin () AS LONG
   FUNCTION = m_ImageMargin
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text margin
' Parameters:
' - nMargin = Margin (in pixels)
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetTextMargin (BYVAL nMargin AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_TextMargin = nMargin
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Gets the text margin.
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetTextMargin () AS LONG
   FUNCTION = m_TextMargin
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the button state
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetButtonState () AS LONG
   FUNCTION = m_fState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the image type
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetImageType () AS LONG
   FUNCTION = m_ImageType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns CTRUE if themes are enabled or FALSE otherwise
' ========================================================================================
PRIVATE FUNCTION CXpButton.IsThemed () AS LONG
   FUNCTION = m_bIsThemed
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables theming
' ========================================================================================
PRIVATE SUB CXpButton.EnableTheming
   m_bIsThemed = CTRUE
   this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Disables theming
' ========================================================================================
PRIVATE SUB CXpButton.DisableTheming
   m_bIsThemed = FALSE
   this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Returns the handle of the image
' Parameters:
' - ImageState =
'      XPBI_NORMAL = 1
'      XPBI_HOT = 2
'      XPBI_DISABLED = 3
' ========================================================================================
PRIVATE FUNCTION CXpButton.GetImage (BYVAL ImageState AS LONG) AS HANDLE
   SELECT CASE ImageState
      CASE XPBI_NORMAL
         FUNCTION = m_hImage
      CASE XPBI_HOT
         FUNCTION = m_hHotImage
      CASE XPBI_DISABLED
         FUNCTION = m_hDisabledImage
   END SELECT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws the button
' ========================================================================================
PRIVATE SUB CXpButton.UxDrawPushButton (BYVAL hDC AS HDC)

   DIM hTheme AS HTHEME                   ' // Theme handle
   DIM rcClient AS RECT                   ' // Original coordinates of the button client area
   DIM rc AS RECT                         ' // Coordinates of the button client area
   DIM rcContent AS RECT                  ' // Coordinates of the button content
   DIM iStateId AS LONG                   ' // Button state
   DIM bIsPressed AS LONG                 ' // Boolean (pressed state)
   DIM bIsFocused AS LONG                 ' // Boolean (focused state)
   DIM bIsDisabled AS LONG                ' // Boolean (disabled state)
   DIM iy AS LONG                         ' // x-coordinate of the image
   DIM ix AS LONG                         ' // y-coordinate of the image
   DIM hImage AS HANDLE                   ' // Image handle
   DIM fImageDrawn AS LONG                ' // Boolean
   DIM lStyle AS LONG                     ' // Button styles
   DIM uState AS DWORD                    ' // Initial state of the frame control
   DIM wszCaption AS WSTRING * 260        ' // Button text

   ' // Selects the font
   IF hDc = 0 THEN EXIT SUB
   .SelectObject hDC, m_hFont

   ' // Button style
   IF m_bIsThemed = FALSE THEN lStyle = .GetWindowLongPtrW(m_hCtl, GWL_STYLE)

   ' // Gets the text of the button
   .GetWindowTextW(m_hCtl, wszCaption, SIZEOF(wszCaption))

   ' // Opens the theme data for the specified window and class
   IF m_bIsThemed THEN
      hTheme = .OpenThemeData(m_hCtl, "Button")
      IF hTheme = 0 THEN EXIT SUB
   END IF

   ' // Gets the button state
   IF m_hCtl = .GetFocus THEN bIsFocused = CTRUE
   IF (m_fState AND BST_PUSHED) = BST_PUSHED THEN bIsPressed = CTRUE
   IF .IsWindowEnabled(m_hCtl) = FALSE THEN bIsDisabled = CTRUE

   ' // Sets the button state for the theme
   iStateId = PBS_NORMAL
   IF bIsPressed THEN iStateId = PBS_PRESSED
   IF bIsDisabled THEN iStateId = PBS_DISABLED
   IF iStateId = PBS_NORMAL THEN
      IF bIsFocused THEN iStateId = PBS_DEFAULTED
      IF .GetPropW(m_hCtl, "HOT") THEN iStateId = PBS_HOT
   END IF

   ' // Toggled button
   IF m_bIsToggle AND m_bToggled THEN iStateId = PBS_PRESSED

   ' // Gets the coordinates of the button's client area
   .GetClientRect m_hCtl, @rcClient
   rc = rcClient

   ' // Draws the button
   IF m_bIsThemed THEN
      ' // Increase 1 pixel to include the edge
      .InflateRect @rc, 1, 1
      ' // Draws the theme-specified border and fills for the "iPartId" and "iStateId".
      .DrawThemeBackground(hTheme, hDc, BP_PUSHBUTTON, iStateId, @rc, NULL)
      ' // Gets the size of the content for the theme-defined background
      .GetThemeBackgroundContentRect(hTheme, hDc, BP_PUSHBUTTON, iStateId, @rc, @rcContent)
   ELSE
      ' // Uses GDI to draw the button
      rcContent = rc
      IF bIsFocused THEN
         IF m_bIsToggle = FALSE OR bIsPressed = FALSE THEN
            .FrameRect hDc, @rcContent, GetSysColorBrush(COLOR_WINDOWTEXT)
         END IF
         .InflateRect @rcContent, -1, -1
      END IF
      IF m_bIsToggle THEN
         IF iStateId = PBS_PRESSED THEN
            .DrawEdge hDc, @rcContent, EDGE_SUNKEN, BF_RECT OR BF_MIDDLE OR BF_SOFT
         ELSE
            IF (lStyle AND BS_FLAT) = BS_FLAT THEN
               .DrawEdge hDc, @rcContent, EDGE_RAISED, BF_RECT OR BF_MIDDLE OR BF_SOFT OR BF_FLAT
            ELSE
               .DrawEdge hDc, @rcContent, EDGE_RAISED, BF_RECT OR BF_MIDDLE OR BF_SOFT
            END IF
         END IF
      ELSE
         .FillRect hDc, @rcContent, GetSysColorBrush(COLOR_BTNFACE)
         IF bIsPressed THEN
            .FrameRect hDc, @rcContent, GetSysColorBrush(COLOR_BTNSHADOW)
         ELSE
            uState = DFCS_BUTTONPUSH
            IF iStateId = PBS_HOT THEN uState = uState OR DFCS_HOT
            IF (lStyle AND BS_FLAT) = BS_FLAT THEN uState = uState OR DFCS_FLAT
            .DrawFrameControl hDc, @rcContent, DFC_BUTTON, uState
         END IF
      END IF
   END IF

   ' // Calculates the y-coordinate of the image
   IF (m_ImagePos AND XPBI_VCENTER) = XPBI_VCENTER THEN
      iy = ((rcContent.Bottom - rcContent.Top - m_ImageHeight) \ 2) + rcContent.Top
   ELSEIF (m_ImagePos AND XPBI_BOTTOM) = XPBI_BOTTOM THEN
      iy = rcContent.Bottom - m_ImageHeight
   ELSE
      iy = rcContent.Top
   END IF

   ' // Calculates the x-coordinate of the image
   IF (m_ImagePos AND XPBI_RIGHT) = XPBI_RIGHT THEN
      ix = rcContent.Right - m_ImageMargin - m_ImageWidth
   ELSEIF (m_ImagePos AND XPBI_LEFT) = XPBI_LEFT THEN
      ix = rcContent.Left + m_ImageMargin
   ELSEIF (m_ImagePos AND XPBI_CENTER) = XPBI_CENTER THEN
      ix = ((rcClient.Right - rcClient.Left) \ 2) - (m_ImageWidth \ 2)
   ELSE
      ix = rcContent.Left + m_ImageMargin
   END IF

   ' // If state = pressed, show the icon slightly displaced
   IF iStateId = PBS_PRESSED THEN
      ix = ix + 1
      iy = iy + 1
   END IF

   ' // Draws the image (don't draw it if the icon for normal state doesn't exist)
   IF m_hImage <> NULL AND m_ImagePos <> XPBI_NONE THEN
      IF iStateId = PBS_HOT THEN hImage = m_hHotImage
      IF iStateId = PBS_DISABLED THEN hImage = m_hDisabledImage
      IF hImage = 0 THEN hImage = m_hImage
      IF m_ImageType = IMAGE_BITMAP THEN
         .DrawStateW hDc, NULL, NULL, CAST(LPARAM, hImage), 0, ix, iy, 0, 0, DST_BITMAP
      ELSE
         .DrawIconEx hDc, ix, iy, hImage, m_ImageWidth, m_ImageHeight, 0, NULL, DI_NORMAL
      END IF
      fImageDrawn = CTRUE
   END IF

   ' // Calculates the coordinates of the text
   IF fImageDrawn THEN
      IF (m_ImagePos AND XPBI_LEFT) = XPBI_LEFT THEN
         IF (m_dwTextFlags AND DT_CENTER) <> DT_CENTER AND _
            (m_dwTextFlags AND DT_RIGHT) <> DT_RIGHT THEN
            IF (m_ImagePos AND XPBI_ABOVE) = XPBI_ABOVE THEN
               rcContent.Left = rcContent.Left + m_TextMargin
            ELSE
               rcContent.Left = rcContent.Left + m_TextMargin + m_ImageWidth + m_ImageMargin
            END IF
         END IF
      ELSEIF (m_ImagePos AND XPBI_RIGHT) = XPBI_RIGHT THEN
         IF (m_dwTextFlags AND DT_RIGHT) = DT_RIGHT THEN
            IF (m_ImagePos AND XPBI_ABOVE) = XPBI_ABOVE THEN
               rcContent.Right = rcContent.Right - m_TextMargin
            ELSE
               rcContent.Right = rcContent.Right - m_TextMargin - m_ImageWidth - m_ImageMargin
            END IF
         END IF
      END IF
   END IF

   IF (m_ImagePos AND XPBI_CENTERCENTER) <> XPBI_CENTERCENTER THEN
      IF (m_dwTextFlags AND DT_CENTER) = DT_CENTER THEN
         ' Do nothing
      ELSEIF (m_dwTextFlags AND DT_RIGHT) = DT_RIGHT THEN
         rcContent.Right = rcContent.Right - m_TextMargin
      ELSEIF (m_dwTextFlags AND DT_LEFT) = DT_LEFT THEN
         rcContent.Left = rcContent.Left + m_TextMargin
      END IF
      IF m_bIsThemed THEN
         ' // Draws the text using the theme-specified color and font for the "iPartId" and "iStateId"
         .DrawThemeText(hTheme, hDc, BP_PUSHBUTTON, iStateId, @wszCaption, -1, m_dwTextFlags, 0, @rcContent)
      ELSE
         ' // Uses GDI to draw the text
         .SetBkMode hDc, TRANSPARENT
         IF bIsDisabled THEN
            .OffsetRect @rcContent, 1, 1
            .SetTextColor hDc, GetSysColor(COLOR_3DHILIGHT)
            .DrawTextW hDc, @wszCaption, -1, @rcContent, m_dwTextFlags
            .OffsetRect @rcContent, -1, -1
            .SetTextColor hDc, GetSysColor(COLOR_3DSHADOW)
            .DrawTextW hDc, @wszCaption, -1, @rcContent, m_dwTextFlags
         ELSE
            .SetTextColor hDc, GetSysColor(COLOR_BTNTEXT)
            .SetBkColor hDc, GetSysColor(COLOR_BTNFACE)
            .DrawTextW hDc, @wszCaption, -1, @rcContent, m_dwTextFlags
         END IF
      END IF
   END IF

   IF m_bIsThemed THEN
      ' // Draws the focus rect
      IF bIsFocused THEN
         .InflateRect @rc, -3, -3
         .DrawFocusRect hDc, @rc
      END IF
      ' // Closes the theme data handle
      .CloseThemeData(hTheme)
   ELSE
      ' // Draws the focus rect
      IF bIsFocused THEN
         .InflateRect @rc, -4, -4
         .DrawFocusRect hDc, @rc
      END IF
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' Sets the image for the button
' Parameters:
' - hImage = Image's handle
' - ImageType = IMAGE_ICON or IMAGE_BITMAP
' - ImageState =
'      XPBI_NORMAL = 1
'      XPBI_HOT = 2
'      XPBI_DISABLED = 3
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetImage (BYVAL hImage AS HANDLE, BYVAL ImageType AS LONG, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
   IF hImage = NULL THEN EXIT SUB
   IF ImageType <> IMAGE_ICON AND ImageType <> IMAGE_BITMAP THEN EXIT SUB
   IF ImageState <> XPBI_NORMAL AND ImageState <> XPBI_HOT AND ImageState <> XPBI_DISABLED THEN EXIT SUB
   ' // Destroys the old image
   SELECT CASE ImageState
      CASE XPBI_NORMAL
         IF m_hImage THEN DeleteObject m_hImage
         m_hImage = hImage
      CASE XPBI_HOT
         IF m_hHotImage THEN DeleteObject m_hHotImage
         m_hHotImage = hImage
      CASE XPBI_DISABLED
         IF m_hDisabledImage THEN DeleteObject m_hDisabledImage
         m_hDisabledImage = hImage
   END SELECT
   ' // Image type
   m_ImageType = ImageType
   ' // Default width and height of the image
   m_ImageWidth = 16
   m_ImageHeight = 16
   DIM tii AS ICONINFO, tbm AS BITMAP
   SELECT CASE ImageType
      CASE IMAGE_ICON
         IF .GetIconInfo(hImage, @tii) THEN
            IF .GetObject(tii.hbmMask, SIZEOF(tbm), @tbm) THEN
               m_ImageWidth = tbm.bmWidth
               m_ImageHeight = tbm.bmHeight
            END IF
         END IF
      CASE IMAGE_BITMAP
         IF .GetObject(hImage, SIZEOF(tbm), @tbm) THEN
            m_ImageWidth = tbm.bmWidth
            m_ImageHeight = tbm.bmHeight
         END IF
   END SELECT
   ' // Redraws the button
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Sets the icon for the button
' Parameters:
' - hIcon = Icon's handle
' - ImageState =
'      XPBI_NORMAL = 1
'      XPBI_HOT = 2
'      XPBI_DISABLED = 3
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetIcon (BYVAL hIcon AS HICON, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
   this.SetImage(hIcon, IMAGE_ICON, ImageState, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the bitmap for the button
' Parameters:
' - hBitmap = Bitmap's handle
' - ImageState =
'      XPBI_NORMAL = 1
'      XPBI_HOT = 2
'      XPBI_DISABLED = 3
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetBitmap (BYVAL hBitmap AS HBITMAP, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
   this.SetImage(hBitmap, IMAGE_BITMAP, ImageState, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Loads an icon from file and sets it as the image of the button.
' Parameters:
' - pwszPath = Full path of the icon's file.
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetIconFromFile (BYVAL pwszPath AS WSTRING PTR, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
   DIM hIcon AS HICON = .LoadImageW(NULL, pwszPath, IMAGE_ICON, 0, 0, LR_LOADFROMFILE)
   this.SetImage(hIcon, IMAGE_ICON, ImageState, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Loads a bitmap from file and sets it as the image of the button.
' Parameters:
' - pwszPath = Full path of the bitmap file.
' - fRedraw = TRUE or FALSE (redraws the button to reflect the changes)
' ========================================================================================
PRIVATE SUB CXpButton.SetBitmapFromFile (BYVAL pwszPath AS WSTRING PTR, BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
   DIM hBitmap AS HBITMAP = LoadImageW(NULL, pwszPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
   this.SetImage(hBitmap, IMAGE_BITMAP, ImageState, fRedraw)
END SUB
' ========================================================================================

END NAMESPACE
