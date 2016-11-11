' ########################################################################################
' Microsoft Windows
' File: CPgBar3D.inc
' Contents: Progress bar control
' Based on the PGBAR3D control originally written by Börje Hagsten.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/CWindow.inc"

NAMESPACE Afx

' ========================================================================================
' wParam colors for PGB_SETBARCOL message
' ========================================================================================
CONST PGB_SILVER = 0
CONST PGB_RED    = 1
CONST PGB_GREEN  = 2
CONST PGB_BLUE   = 3
CONST PGB_CYAN   = 4   ' // Blue-green
CONST PGB_VIOLET = 5   ' // Red-blue
CONST PGB_GOLD   = 6   ' // Yellow
CONST PGB_BRONZE = 7   ' // Brown
' ========================================================================================

' ========================================================================================
' wParam colors for PGB_SETTXTON message
' ========================================================================================
CONST PGB_TEXTNONE    = 0
CONST PGB_TEXTAUTO    = 1
CONST PGB_TEXTCUSTOM  = 2
' ========================================================================================

' ========================================================================================
' wParam colors for PGB_SETBARDIR message
' ========================================================================================
CONST PGB_BD_LEFTTORIGHT = 0
CONST PGB_BD_RIGHTTOLEFT = 1
CONST PGB_BD_BOTTOMTOTOP = 2
CONST PGB_BD_TOPTOBOTTOM = 3
' ========================================================================================

' ========================================================================================
' wParam colors for PGB_SETGRADIENTDIR message
' ========================================================================================
CONST PGB_GD_HORIZONTAL = 0
CONST PGB_GD_VERTICAL   = 1
' ========================================================================================

' ========================================================================================
'Custom control messages
' ========================================================================================
CONST PGB_SETMAX         = WM_USER + 100   ' // wParam sets max number of steps
CONST PGB_STEPUP         = WM_USER + 103   ' // Increases step while < max; wParam and lParam must be 0
CONST PGB_STEPDN         = WM_USER + 104   ' // Decreases step while > 0; wParam and lParam must be 0
CONST PGB_SETVALUE       = WM_USER + 105   ' // wParam sets progess bar value, lParam controls redraw
CONST PGB_BUILDBARS      = WM_USER + 109   ' // Build/rebuild the progress bars, lParam controls redraw
CONST PGB_REFRESH        = WM_USER + 110   ' // Redraw the control; wParam and lParam must be 0

CONST PGB_GETMAX         = WM_USER + 120   ' // Returns max number of steps
CONST PGB_GETVALUE       = WM_USER + 121   ' // Returns step value
CONST PGB_GETTXTON       = WM_USER + 122   ' // Returns txtOnOff value
CONST PGB_GETTXTPOS      = WM_USER + 123   ' // Returns text position in control
CONST PGB_GETTXTCOLBAR   = WM_USER + 124   ' // Returns bar text color
CONST PGB_GETTXTCOLBKG   = WM_USER + 125   ' // Returns background text color
CONST PGB_GETCOLBKG      = WM_USER + 126   ' // Returns background color
CONST PGB_GETBARCOL      = WM_USER + 127   ' // Returns bar color scheme
CONST PGB_GETBARDIR      = WM_USER + 128   ' // Returns bar direction, 0:left/right 1:upside down 2:bottom/top 3:top/bottom
CONST PGB_GETGRADIENTDIR = WM_USER + 129   ' // Returns gradient direction - 0:horizontal, 1:vertical
CONST PGB_GETTXTANGLE    = WM_USER + 130   ' // Returns rotated font

CONST PGB_SETTXTON       = WM_USER + 150   ' // lParam sets: 0 = no text, 1 = auto text (%), 2 = custom text
CONST PGB_SETTXTBAR      = WM_USER + 151   ' // wParam points to text text for bar, lParam controls redraw
CONST PGB_SETTXTBKG      = WM_USER + 152   ' // wParam points to text text for background, lParam controls redraw
CONST PGB_SETTXTPOS      = WM_USER + 153   ' // wParam sets text position in control
CONST PGB_SETTXTCOLBAR   = WM_USER + 154   ' // wParam sets bar text color

CONST PGB_SETTXTCOLBKG   = WM_USER + 155   ' // wParam sets background text color
CONST PGB_SETCOLBKG      = WM_USER + 156   ' // wParam sets background color, lParam controls rebuild of control
CONST PGB_SETBARCOL      = WM_USER + 157   ' // wParam sets bar color scheme, lParam controls rebuild of control
CONST PGB_SETBARDIR      = WM_USER + 159   ' // wParam sets bar direction, 0:left/right 1:upside down 2:bottom/top 3:top/bottom, lParam controls rebuild of control
CONST PGB_SETGRADIENTDIR = WM_USER + 160   ' // wParam sets gradient direction - 0:horizontal, 1:vertical, lParam controls rebuild of control
CONST PGB_SETTXTANGLE    = WM_USER + 161   ' // wParam sets rotated font, lParam controls rebuild of control
CONST PGB_SETFONT        = WM_USER + 162   ' // wParam = Handle of the font, lParam controls rebuild of control
CONST PGB_GETFONT        = WM_USER + 162   ' // Returns the hanfle of the font
' ========================================================================================

' ========================================================================================
' CPgBar3D class
' ========================================================================================
TYPE CPgBar3D

   Private:
      m_hCtl        AS HWND            ' // Button handle
      m_pStep       AS LONG            ' // For tracking what step we are on
      m_pMax        AS LONG            ' // for storing max number of steps, usually 100 (%)
      m_direction   AS LONG            ' // Bar direction - left to right, or right to left?
      m_gradientDir AS LONG            ' // Gradient direction - left to right, or right to left?
      m_txtAngle    AS LONG            ' // Store given text angle
      m_bkgColor    AS LONG            ' // Background color
      m_barCol      AS LONG            ' // Bar color scheme
      m_txtColBar   AS LONG            ' // Custom text color in bar
      m_txtColBkg   AS LONG            ' // Custom text color on background
      m_txtOnOff    AS LONG            ' // 0 = no text, 1 = auto text (%), 2 = custom text
      m_txtPos      AS LONG            ' // Text position in control, see DrawText API..
      m_txtBkg      AS WSTRING * 255   ' // Text to be painted on background, increase/decrease size to suit your needs
      m_txtBar      AS WSTRING * 255   ' // text to be painted on bar, increase/decrease size to suit your needs
      m_hbBack      AS HBRUSH          ' // Handle for background brush
      m_barDC       AS HDC             ' // Memory DC for progress bar
      m_barBit      AS HBITMAP         ' // Handle to progress bar bitmap
      m_barDC2      AS HDC             ' // Memory DC for progress bar buffer
      m_barBit2     AS HBITMAP         ' // Handle to progress bar buffer bitmap
      m_memDc       AS HDC             ' // Memory DC for main buffer
      m_hBit        AS HBITMAP         ' // Handle to main buffer bitmap
      m_hFont       AS HFONT           ' // Handle to font used by the control
      m_hRotateFont AS HFONT           ' // Handle to rotated font style
      DIM m_PalClr(192) AS LONG        ' // Array of colors for the gradients

   Public:

      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hWindow () AS HWND
      DECLARE SUB Redraw
      DECLARE SUB Clear
      DECLARE SUB CreateGradientBars
      DECLARE FUNCTION GetMaxSteps () AS LONG
      DECLARE SUB SetMaxSteps (BYVAL nMaxSteps AS LONG)
      DECLARE FUNCTION GetStep () AS LONG
      DECLARE SUB SetStep (BYVAL nStepValue AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetTextOn () AS LONG
      DECLARE SUB SetTextOn (BYVAL nMode AS LONG)
      DECLARE FUNCTION GetTextPos () AS LONG
      DECLARE SUB SetTextPos (BYVAL nPos AS LONG)
      DECLARE FUNCTION GetBarColor () AS LONG
      DECLARE SUB SetBarColor (BYVAL nColor AS LONG)
      DECLARE FUNCTION GetBarTextColor () AS LONG
      DECLARE SUB SetBarTextColor (BYVAL nColor AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetBarDirection () AS LONG
      DECLARE SUB SetBarDirection (BYVAL nDirection AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetBackColor () AS LONG
      DECLARE SUB SetBackColor (BYVAL nColor AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetTextBackColor () AS LONG
      DECLARE SUB SetTextBackColor (BYVAL nColor AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetGradientDirection () AS LONG
      DECLARE SUB SetGradientDirection (BYVAL nAngle AS LONG)
      DECLARE FUNCTION GetTextAngle () AS LONG
      DECLARE SUB SetTextAngle (BYVAL nAngle AS LONG, BYVAL fRebuild AS LONG = FALSE)
      DECLARE SUB StepUp
      DECLARE SUB StepDown
      DECLARE SUB BuildBar (BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetBarText (BYREF wszText AS WSTRING, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetBackText (BYREF wszText AS WSTRING, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetFont (BYVAL hFont AS HFONT, BYVAL fRedraw AS LONG = FALSE)
      DECLARE FUNCTION GetFont () AS HFONT
      DECLARE STATIC FUNCTION CPgBar3DProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CPgBar3D class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CPgBar3D (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)

   ' // Register the class
   DIM wAtom AS ATOM
   DIM wcexw AS WNDCLASSEXW
   DIM wszClassName AS WSTRING * 260 = "AFX_PGBAR3D"
   IF .GetClassInfoExW(.GetModuleHandleW(NULL), @wszClassName, @wcexw) = 0 THEN
      ' // Fill the WNDCLASSEXW structure
      WITH wcexw
         .cbSize        = SIZEOF(wcexw)
         .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
         .lpfnWndProc   = @CPgBar3DProc
         .cbClsExtra    = 0
         .cbWndExtra    = SIZEOF(HANDLE)   ' // make room to store a pointer to the class
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

   ' // Initial values
   m_txtOnOff  = 0
   m_txtPos    = DT_SINGLELINE OR DT_CENTER OR DT_VCENTER OR DT_NOCLIP OR DT_NOPREFIX
   m_txtColBar = BGR(0, 0, 0)
   m_txtColBkg = BGR(255, 255, 0)
   m_bkgColor  = BGR(128, 128, 128)
   m_barCol    = 0
   m_hbBack    = CreateSolidBrush(m_bkgColor)

   ' // Create the control
   IF dwStyle = 0 THEN dwStyle = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP
   IF pWindow THEN m_hCtl = pWindow->AddControl(wszClassName, pWindow->hWindow, cID, wszTitle, x, y, nWidth, nHeight, dwStyle, dwExStyle, lpParam)
   IF m_hCtl THEN
      .SetWindowLongPtrW m_hCtl, 0, CAST(LONG_PTR, @this)
      IF LEN(wszTitle) THEN .SetWindowTextW m_hCtl, wszTitle
      ' // Set the same font used by the parent
      DIM lfw AS LOGFONTW
      IF pWindow->Font THEN
         IF .GetObjectW(pWindow->Font, SIZEOF(lfw), @lfw) THEN m_hFont = CreateFontIndirectW(@lfw)
      END IF
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CPgBar3D class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CPgBar3D
   ' // Free resources
   IF m_hFont THEN DeleteObject m_hFont
   IF m_hRotateFont THEN DeleteObject m_hRotateFont
   IF m_hbBack      THEN DeleteObject m_hbBack
   IF m_hbit        THEN DeleteObject SelectObject(m_memDC, m_hbit)
   IF m_memDC       THEN DeleteDC m_memDC
   IF m_barBit      THEN DeleteObject SelectObject(m_barDC, m_barBit)
   IF m_barDC       THEN DeleteDC m_barDC
   IF m_barBit2     THEN DeleteObject SelectObject(m_barDC2, m_barBit2)
   IF m_barDC2      THEN DeleteDC m_barDC2
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.CPgBar3DProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM pPgBar AS CPgBar3D PTR
   DIM hDC AS HDC, ps AS PAINTSTRUCT
   DIM xPos AS LONG, yPos AS LONG, rc AS RECT, rcTxt AS RECT
   DIM hFontOld AS HFONT, lpSize AS SIZEL, tp AS WSTRING PTR

   SELECT CASE uMsg

      CASE WM_CREATE
         EXIT FUNCTION

      CASE WM_ERASEBKGND
         ' // Don't erase the background to avoid flicker
         FUNCTION = 1
         EXIT FUNCTION

      CASE WM_PAINT
         ' // Get a pointer to the class
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         ' // Get the size of the control
         .GetClientRect hwnd, @rc
         ' // Clear the background
         .FillRect pPgBar->m_memDC, @rc, pPgBar->m_hbBack
         ' // Calculate the coordinates
         xPos = pPgBar->m_pStep * rc.Right / pPgBar->m_pMax
         yPos = pPgBar->m_pStep * rc.Bottom / pPgBar->m_pMax
         ' // If it has text...
         IF pPgBar->m_txtOnOff THEN
            IF pPgBar->m_txtOnOff = 1 THEN
               ' // Auto text to paint on bar
               pPgBar->m_txtBar = WSTR(pPgBar->m_pStep) & "%"
               ' // Auto text to paint on background
               pPgBar->m_txtBkg = pPgBar->m_txtBar
            END IF
            ' // Copy rect for drawtext
            rcTxt = rc
            ' // If rotated font...
            IF pPgBar->m_hRotateFont THEN
               ' // Store original fonts for later use (is the same in both DCs)
               hFontOld = .SelectObject(pPgBar->m_memDC, pPgBar->m_hRotateFont)
               hFontOld = .SelectObject(pPgBar->m_barDC2, pPgBar->m_hRotateFont)
               IF pPgBar->m_direction = 1 THEN
                  ' // Direction upside down
                  pPgBar->m_txtPos = DT_SINGLELINE OR DT_CENTER OR DT_VCENTER OR DT_NOCLIP OR DT_NOPREFIX
                  .GetTextExtentPoint32W(pPgBar->m_memDC, pPgBar->m_txtBar, LEN(pPgBar->m_txtBar), @lpSize)
                  rcTxt.Bottom = rcTxt.Bottom + lpSize.cy * 2
                  rcTxt.Right  = rcTxt.Right + lpSize.cx * 2
               ELSEIF pPgBar->m_direction = 2 THEN
                  ' // Direction bottom to top
                  pPgBar->m_txtPos = DT_SINGLELINE OR DT_LEFT OR DT_VCENTER OR DT_NOCLIP OR DT_NOPREFIX
                  .GetTextExtentPoint32W(pPgBar->m_memDC, pPgBar->m_txtBar, LEN(pPgBar->m_txtBar), @lpSize)
                  rcTxt.Left   = (rcTxt.Right - lpSize.cy) / 2
                  rcTxt.Bottom = rcTxt.Bottom + lpSize.cx * 1.25
               ELSEIF pPgBar->m_direction = 3 THEN
                  ' // Direction top to bottom
                 pPgBar->m_txtPos = DT_SINGLELINE OR DT_LEFT OR DT_VCENTER OR DT_NOCLIP OR DT_NOPREFIX
                 .GetTextExtentPoint32W(pPgBar->m_memDC, pPgBar->m_txtBar, LEN(pPgBar->m_txtBar), @lpSize)
                 rcTxt.Left = (rcTxt.Right + lpSize.cy) / 2
                 rcTxt.Top = rcTxt.Top - lpSize.cx / 1.35
               END IF
            ELSE
               ' // Store original fonts for later use (is the same in both DCs)
               hFontOld = .SelectObject(pPgBar->m_memDC, pPgBar->m_hFont)
               hFontOld = .SelectObject(pPgBar->m_barDC2, pPgBar->m_hFont)
            END IF
            ' // Paint original bar to buffer
            .BitBlt pPgBar->m_barDC2, 0, 0, rc.Right, rc.Bottom, pPgBar->m_barDC, 0, 0, SRCCOPY
            ' // Set color on bar
            .SetTextColor pPgBar->m_barDC2, pPgBar->m_txtColBar
            ' // Draw text on bar
            .DrawTextW pPgBar->m_barDC2, pPgBar->m_txtBar, -1, @rcTxt, pPgBar->m_txtPos
            ' // Set color on background
            .SetTextColor pPgBar->m_memDC, pPgBar->m_txtColBkg
            ' // Draw text on background
            .DrawTextW pPgBar->m_memDC, pPgBar->m_txtBkg, -1, @rcTxt, pPgBar->m_txtPos
            ' // Paint proper part of gradient bar
            IF pPgBar->m_direction = 0 THEN
               ' // Direction left to right, with text
               .BitBlt pPgBar->m_memDC, 0, 0, xPos, rc.Bottom, pPgBar->m_barDC2, 0, 0, SRCCOPY
            ELSEIF pPgBar->m_direction = 1 THEN  'RIGHT TO LEFT - WITH TEXT
               ' // Direction right to left, with text
               .BitBlt pPgBar->m_memDC, rc.Right - xPos, 0, xPos, rc.Bottom, pPgBar->m_barDC2, rc.Right - xPos, 0, SRCCOPY
            ELSEIF pPgBar->m_direction = 2 THEN
               ' // Direction bottom to top, with text
               .BitBlt pPgBar->m_memDC, 0, rc.Bottom - yPos, rc.Right, rc.Bottom, pPgBar->m_barDC2, 0, rc.Bottom - yPos, SRCCOPY
            ELSEIF pPgBar->m_direction = 3 THEN
               ' // Direction top to bottom, with text
               .BitBlt pPgBar->m_memDC, 0, 0, rc.Right, yPos, pPgBar->m_barDC2, 0, 0, SRCCOPY
            END IF
         ELSE
            IF pPgBar->m_direction = 0 THEN
               ' // Direction left ro right, without text
               .BitBlt pPgBar->m_memDC, 0, 0, xPos, rc.Bottom, pPgBar->m_barDC, 0, 0, SRCCOPY
            ELSEIF pPgBar->m_direction = 1 THEN
               ' // Direction right to left, without text
               .BitBlt pPgBar->m_memDC, rc.Right - xPos, 0, xPos, rc.Bottom, pPgBar->m_barDC, rc.Right - xPos, 0, SRCCOPY
            ELSEIF pPgBar->m_direction = 2 THEN  'BOTTOM TO TOP - NO TEXT
               ' // Direction bottom to top, without text
               .BitBlt pPgBar->m_memDC, 0, rc.Bottom - yPos, rc.Right, rc.Bottom, pPgBar->m_barDC, 0, rc.Bottom - yPos, SRCCOPY
            ELSEIF pPgBar->m_direction = 3 THEN  'TOP TO BOTTOM - NO TEXT
               ' // Top to bottom, without text
               .BitBlt pPgBar->m_memDC, 0, 0, rc.Right, yPos, pPgBar->m_barDC, 0, 0, SRCCOPY
            END IF
         END IF
         ' // Begin secreen painting
         .BeginPaint hwnd, @ps
         .BitBlt ps.hDC, 0, 0, rc.Right, rc.Bottom, pPgBar->m_memDC, 0, 0, SRCCOPY
         IF hFontOld THEN
            ' // Select the original fonts back
            .SelectObject(pPgBar->m_memDC, hFontOld)
            .SelectObject(pPgBar->m_barDC2, hFontOld)
         END IF
         .EndPaint hwnd, @ps
         EXIT FUNCTION

      CASE PGB_STEPUP
         ' // Increases step while < max; wParam and lParam shall be 0
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pPgBar->m_pStep < pPgBar->m_pMax THEN
            pPgBar->m_pStep = pPgBar->m_pStep + 1
            .SendMessageW hwnd, PGB_REFRESH, 0, 0
         END IF
         EXIT FUNCTION

      ' // Step down while > 0
      CASE PGB_STEPDN
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pPgBar->m_pStep > 0 THEN
            pPgBar->m_pStep = pPgBar->m_pStep - 1
            .SendMessageW hwnd, PGB_REFRESH, 0, 0
         END IF
         EXIT FUNCTION

      ' // Build gradient bars
      CASE PGB_BUILDBARS
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->CreateGradientBars
         IF lParam THEN .SendMessageW hwnd, PGB_REFRESH, 0, 0
         EXIT FUNCTION

      ' // Redraws the control
      CASE PGB_REFRESH
        .InvalidateRect hwnd, NULL, 0
        .UpdateWindow hWnd
         EXIT FUNCTION

      ' // wParam sets progess bar value, lParam controls redraw
      CASE PGB_SETVALUE
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_pStep = MIN(pPgBar->m_pMax, wParam)
         IF lParam THEN .SendMessageW hwnd, PGB_REFRESH, 0, 0
         EXIT FUNCTION

      ' // Return the current step value
      CASE PGB_GETVALUE
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_pStep
         EXIT FUNCTION

     ' // Set max number of steps
      CASE PGB_SETMAX
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_pMax = wParam
         EXIT FUNCTION

      ' // Get max number of steps
      CASE PGB_GETMAX
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_pMax
         EXIT FUNCTION

      ' // Set bar color
      CASE PGB_SETBARCOL
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_barCol = wParam * 24 + 1
         ' // Rebuild the bars
         IF lParam THEN .SendMessageW hwnd, PGB_BUILDBARS, 0, 1
         EXIT FUNCTION

      ' // Return bar color scheme
      CASE PGB_GETBARCOL
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_barCol / 24
         EXIT FUNCTION
      ' // Set background color via wParam

      CASE PGB_SETCOLBKG
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_bkgColor = wParam
         ' // Delete old brush
         IF pPgBar->m_hbBack THEN DeleteObject pPgBar->m_hbBack
         ' // Create background color brush
         pPgBar->m_hbBack = CreateSolidBrush(pPgBar->m_bkgColor)
         IF lParam THEN .SendMessageW hwnd, PGB_BUILDBARS, 1, 0
         EXIT FUNCTION

      ' // Return background color
      CASE PGB_GETCOLBKG
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_bkgColor
         EXIT FUNCTION

      ' // Bar direction; left to right = 0, default
      CASE PGB_SETBARDIR
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_direction = wParam
         IF lParam THEN .SendMessageW hwnd, PGB_BUILDBARS, 0, 1
         EXIT FUNCTION

      ' // Return bar direction
      CASE PGB_GETBARDIR
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_direction
         EXIT FUNCTION

      ' // Set gradient direction; horizontal = 0, default
      CASE PGB_SETGRADIENTDIR
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_gradientDir = wParam
         IF lParam THEN .SendMessageW hwnd, PGB_BUILDBARS, 0, 1
         EXIT FUNCTION

      ' // Return gradient direction
      CASE PGB_GETGRADIENTDIR
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_gradientDir
         EXIT FUNCTION

      ' // Set text on/off
      CASE PGB_SETTXTON
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_txtOnOff = lParam
         EXIT FUNCTION

      ' // Return text on/off setting
      CASE PGB_GETTXTON
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_txtOnOff
         EXIT FUNCTION

      ' // Set text position in the control
      CASE PGB_SETTXTPOS
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_txtPos = wParam
         EXIT FUNCTION

      ' // Return txtPos setting
      CASE PGB_GETTXTPOS
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_txtPos
         EXIT FUNCTION

      ' // Set text color
      CASE PGB_SETTXTCOLBAR
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_txtColBar = wParam
         EXIT FUNCTION

      ' // Return bar text color
      CASE PGB_GETTXTCOLBAR
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_txtColBar
         EXIT FUNCTION

      ' // Set text background color
      CASE PGB_SETTXTCOLBKG
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_txtColBkg = wParam
         EXIT FUNCTION

      ' // Return background text color
      CASE PGB_GETTXTCOLBKG
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = pPgBar->m_txtColBkg
         EXIT FUNCTION

      ' // Set the text of the bar
      CASE PGB_SETTXTBAR
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         tp = CAST(WSTRING PTR, wParam)
         pPgBar->m_txtBar = *tp
         IF lParam THEN .SendMessageW hwnd, PGB_REFRESH, 0, 0
         EXIT FUNCTION

      ' // Set text of the background
      CASE PGB_SETTXTBKG
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         tp = CAST(WSTRING PTR, wParam)
         pPgBar->m_txtBkg = *tp
         IF lParam THEN .SendMessageW hwnd, PGB_REFRESH, 0, 0
         EXIT FUNCTION

     ' // Set the font handle
     CASE PGB_SETFONT
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pPgBar->m_hFont THEN DeleteObject pPgBar->m_hFont
         pPgBar->m_hFont = CAST(HFONT, wParam)
         IF lParam THEN .SendMessageW hwnd, PGB_REFRESH, 0, 0
         EXIT FUNCTION

     ' // Get the font handle
     CASE PGB_SETFONT
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         FUNCTION = CAST(LRESULT, pPgBar->m_hFont)
         EXIT FUNCTION

     ' // Set angle of the font
     CASE PGB_SETTXTANGLE
         pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
         pPgBar->m_txtAngle = wParam
         DIM lfw AS LOGFONTW
         .GetObjectW(pPgBar->m_hFont, SIZEOF(lfw), @lfw)
         lfw.lfescapement  = wParam * 10   ' // Angle is given in tenths of degrees
         lfw.lforientation = wParam * 10   ' // Both should be same
         lfw.lfWeight = FW_BOLD            ' // Whatever, this one looks something like system font..
         pPgBar->m_hRotateFont = .CreateFontIndirectW(@lfw)   ' // Create the font and store its handle
         IF lParam THEN .SendMessageW hwnd, PGB_BUILDBARS, 0, 1   ' // Rebuild and refresh if lParam says so
         EXIT FUNCTION

      ' // Return eventual text angle
      CASE PGB_GETTXTANGLE
        pPgBar = CAST(CPgBar3D PTR, .GetWindowLongPtrW(hwnd, 0))
        FUNCTION = pPgBar->m_txtAngle
        EXIT FUNCTION

   END SELECT

   ' // Default processing for other messages.
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the control
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.hWindow () AS HWND
   FUNCTION = m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Redraws the control
' ========================================================================================
PRIVATE SUB CPgBar3D.Redraw
   .SendMessageW(this.m_hCtl, PGB_REFRESH, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Clears the control
' ========================================================================================
PRIVATE SUB CPgBar3D.Clear
   DIM txtOnOff AS LONG = .SendMessageW(this.m_hCtl, PGB_GETTXTON, 0, 0)
   .SendMessageW(this.m_hCtl, PGB_SETTXTON, 0, 0)
   .SendMessageW(this.m_hCtl, PGB_SETVALUE, 0, 1)
   .SendMessageW(this.m_hCtl, PGB_SETTXTON, 0, txtOnOff)
END SUB
' ========================================================================================

' ========================================================================================
' Creates the gradient bars
' ========================================================================================
PRIVATE SUB CPgBar3D.CreateGradientBars

   DIM hDC AS HDC, ic AS LONG, jj AS LONG, hPen AS HPEN, kk AS SINGLE, L AS SINGLE, rc AS RECT

   ' // Get the handle of the device context
   hDC = .GetDC(m_hCtl)
   IF hDC = NULL THEN EXIT SUB

   ' // Free resources
   IF m_hbit    THEN DeleteObject SelectObject(m_memDC, m_hbit)
   IF m_memDC   THEN DeleteDC m_memDC
   IF m_barBit  THEN DeleteObject SelectObject(m_barDC, m_barBit)
   IF m_barDC   THEN DeleteDC m_barDC
   IF m_barBit2 THEN DeleteObject SelectObject(m_barDC2, m_barBit2)
   IF m_barDC2  THEN DeleteDC m_barDC2

   ' // Get the width and height of the control
   .GetClientRect m_hCtl, @rc

   ' // Create new resources
   m_memDC   = CreateCompatibleDC(hDC)
   m_hbit    = CreateCompatibleBitmap(hDC, rc.Right, rc.Bottom)
   m_hbit    = SelectObject(m_memDC, m_hbit)
   m_barDC   = CreateCompatibleDC(hDC)
   m_barBit  = CreateCompatibleBitmap(hDC, rc.Right, rc.Bottom)
   m_barBit  = SelectObject(m_barDC, m_barBit)
   m_barDC2  = CreateCompatibleDC(hDC)
   m_barBit2 = CreateCompatibleBitmap(hDC, rc.Right, rc.Bottom)
   m_barBit2 = SelectObject(m_barDC2, m_barBit2)

   ' // Set the text background modes
   SetBkMode m_memDC, TRANSPARENT
   SetBkMode m_barDC2, TRANSPARENT

   ' // Release the device context
   .ReleaseDC m_hCtl, hDC

   ' // Palette of colors
   jj = 1
   FOR ic = 117 TO 255 STEP 6         '0, gray table 1-24
     m_PalClr(jj) = BGR(ic, ic, ic) : jj += 1
   NEXT
   FOR ic = 117 TO 255 STEP 6         '1, red table 25-48
     m_PalClr(jj) = BGR(ic, ic - 117, ic - 117) : jj += 1
   NEXT
   FOR ic = 117 TO 255 STEP 6         '2, green table 49-72
     m_PalClr(jj) = BGR(ic - 117, ic, ic - 117) : jj += 1
   NEXT
   FOR ic = 117 TO 255 STEP 6         '3, blue table 73-96
     m_PalClr(jj) = BGR(ic - 117, ic - 117, ic) : jj += 1
   NEXT
   FOR ic = 117 TO 255 STEP 6         '4, blue-green table 97-120
     m_PalClr(jj) = BGR(ic - 117, ic, ic) : jj += 1
   NEXT
   FOR ic = 117 TO 255 STEP 6         '5, violet table 121-144
     m_PalClr(jj) = BGR(ic, ic - 117, ic) : jj += 1
   NEXT
   FOR ic = 117 TO 255 STEP 6         '6, gold table 145-168
     m_PalClr(jj) = BGR(MIN&(ic + 64, 255), ic, ic - 117) : jj += 1
   NEXT
   FOR ic = 117 TO 255 STEP 6         '7, brown table 169-192
     m_PalClr(jj) = BGR(MIN&(ic + 16, 255), ic - 48, ic - 117) : jj += 1
   NEXT
   jj -=1

   IF m_gradientDir = 0 THEN
      ' // Horizontal bar
      jj = rc.Bottom - 1
   ELSE
      ' // Vertical bar
      jj = rc.Right - 1
   END IF
   kk = m_barCol

   ' // Calculate steps for color
   L = 1 / ((jj / 2) / 24)

   ' // Draw the gradient bar
   FOR ic = 0 TO jj
      hPen = CreatePen(PS_SOLID, 1, m_PalClr(INT(kk)))   ' // Create pen
      hPen = SelectObject(m_barDC, hPen)                 ' // Select pen into DC, store original pen
      IF m_gradientDir = 0 THEN                          ' // Horizontal bar
         MoveToEx m_barDC, 0, ic, NULL                   ' // Move into position...
         LineTo m_barDC, rc.Right, ic                    ' //   ...and draw a line from left to right
      ELSE                                               ' // Vertical bar
         MoveToEx m_barDC, ic, 0, NULL                   ' // Move into position...
         LineTo m_barDC, ic, rc.Bottom                   ' //   ...and draw a line from top to bottom
      END IF
      DeleteObject SelectObject(m_barDC, hPen)           ' // Delete pen to avoid memory leaks
      IF ic < jj / 2 - 1 THEN
         kk = MIN(m_barCol + 23, kk + L)
      ELSE
         kk = MAX(m_barCol, kk - L)
      END IF
   NEXT

END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the maximum number of steps
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetMaxSteps () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETMAX, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetMaxSteps (BYVAL nMaxSteps AS LONG)
   .SendMessageW(this.m_hCtl, PGB_SETMAX, nMaxSteps, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the step value
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetStep () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETVALUE, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetStep (BYVAL nStepValue AS LONG, BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETVALUE, nStepValue, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the text mode
' 0 = no text, 1 = auto text (%), 2 = custom text
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetTextOn () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETTXTON, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetTextOn (BYVAL nMode AS LONG)
   .SendMessageW(this.m_hCtl, PGB_SETTXTON, 0, nMode)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the text position
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetTextPos () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETTXTPOS, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetTextPos (BYVAL nPos AS LONG)
   .SendMessageW(this.m_hCtl, PGB_SETTXTPOS, nPos, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the bar color
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetBarColor () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETBARCOL, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetBarColor (BYVAL nColor AS LONG)
   .SendMessageW(this.m_hCtl, PGB_SETBARCOL, nColor, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the bar text color
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetBarTextColor () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETTXTCOLBAR, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetBarTextColor (BYVAL nColor AS LONG, BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETTXTCOLBAR, nColor, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the bar direction
' 0:left/right 1:upside down 2:bottom/top 3:top/bottom
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetBarDirection () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETBARDIR, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetBarDirection (BYVAL nDirection AS LONG, BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETBARDIR, nDirection, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the background color
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetBackColor () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETCOLBKG, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetBackColor (BYVAL nColor AS LONG, BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETCOLBKG, nColor, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the text color
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetTextBackColor () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETTXTCOLBKG, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetTextBackColor (BYVAL nColor AS LONG, BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETTXTCOLBKG, nColor, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the gradient direction
' 0:horizontal, 1:vertical
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetGradientDirection () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_SETGRADIENTDIR, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetGradientDirection (BYVAL nDirection AS LONG)
   .SendMessageW(this.m_hCtl, PGB_SETGRADIENTDIR, nDirection, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the angle of the font, in tenths of degree
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetTextAngle () AS LONG
   FUNCTION = .SendMessageW(this.m_hCtl, PGB_GETTXTANGLE, 0, 0)
END FUNCTION
' ========================================================================================
PRIVATE SUB CPgBar3D.SetTextAngle (BYVAL nAngle AS LONG, BYVAL fRebuild AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETTXTANGLE, nAngle, fRebuild)
END SUB
' ========================================================================================

' ========================================================================================
' Increases step while < max
' ========================================================================================
PRIVATE SUB CPgBar3D.StepUp
   .SendMessageW(this.m_hCtl, PGB_STEPUP, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Decreases step while > 0
' ========================================================================================
PRIVATE SUB CPgBar3D.StepDown
   .SendMessageW(this.m_hCtl, PGB_STEPDN, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Builds the progress bars
' ========================================================================================
PRIVATE SUB CPgBar3D.BuildBar (BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_BUILDBARS, 0, fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Set the bar text
' ========================================================================================
PRIVATE SUB CPgBar3D.SetBarText (BYREF wszText AS WSTRING, BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETTXTBAR, CAST(WPARAM, @wszText), fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Set the background text
' ========================================================================================
PRIVATE SUB CPgBar3D.SetBackText (BYREF wszText AS WSTRING, BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETTXTBKG, CAST(WPARAM, @wszText), fRedraw)
END SUB
' ========================================================================================

' ========================================================================================
' Gets/sets the normal font used by the control
' ========================================================================================
PRIVATE SUB CPgBar3D.SetFont (BYVAL hFont AS HFONT, BYVAL fRedraw AS LONG = FALSE)
   .SendMessageW(this.m_hCtl, PGB_SETFONT, CAST(WPARAM, hFont), fRedraw)
END SUB
' ========================================================================================
PRIVATE FUNCTION CPgBar3D.GetFont () AS HFONT
   FUNCTION = CAST(HFONT, .SendMessageW(this.m_hCtl, PGB_GETFONT, 0, 0))
END FUNCTION
' ========================================================================================

END NAMESPACE
