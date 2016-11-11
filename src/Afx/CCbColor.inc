' ########################################################################################
' Microsoft Windows
' File: CbColor.inc
' Contents: Combobox color list
' Based on the CBCOLOR control originally written by Börje Hagsten in 2001.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/CWindow.inc"
#include once "win/commdlg.bi"

NAMESPACE Afx

CONST CBCOL_SETAUTOCOLOR = WM_USER + 100
CONST CBCOL_SETUSERCOLOR = WM_USER + 101
CONST CBCOL_SETCOLORNAME = WM_USER + 102
CONST CBCOL_GETAUTOCOLOR = WM_USER + 200
CONST CBCOL_GETUSERCOLOR = WM_USER + 201
CONST CBCOL_GETSELCOLOR  = WM_USER + 202

' ========================================================================================
' CCbColor class
' ========================================================================================
TYPE CCbColor

   Private:
      m_hCtl  AS HWND           ' // Button handle
      m_hFont AS HFONT          ' // Font handle
      m_oldProc AS WNDPROC      ' // Subclass procedure
      m_autoColor AS COLORREF   ' // Auto color
      m_userColor AS COLORREF   ' // User color
      m_ratio AS SINGLE         ' // DPI ratio - width
      DIM m_rgColors(15) AS COLORREF
      DIM m_rgNames (17) AS WSTRING * 20 = {"Auto", "Black", "Blue", "Green", "Cyan", "Red", _
         "Magenta", "Brown", "Light Gray", "Gray", "Light Blue", "Light Green", "Light Cyan", _
         "Light Red", "Light Magenta", "Yellow", "Bright White", "User selected..."}

   Public:
      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER,  _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwAutoColor AS COLORREF = 0, BYVAL dwUserColor AS COLORREF = 0, _
         BYVAL prgNames AS WSTRING PTR = NULL, BYVAL numBytes AS DWORD = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hWindow () AS HWND
      DECLARE FUNCTION GetQBColor (BYVAL hWnd AS HWND, BYVAL c AS LONG, BYVAL dlg AS LONG) AS COLORREF
      DECLARE PROPERTY AutoColor (BYVAL nColor AS COLORREF)
      DECLARE PROPERTY AutoColor () AS COLORREF
      DECLARE PROPERTY UserColor (BYVAL nColor AS COLORREF)
      DECLARE PROPERTY UserColor () AS COLORREF
      DECLARE PROPERTY SelColor () AS COLORREF
      DECLARE PROPERTY ScalingRatio () AS SINGLE
      DECLARE PROPERTY ScalingRatio (BYVAL rxRatio AS SINGLE)
      DECLARE STATIC FUNCTION CCbColorProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
PRIVATE FUNCTION CCbColor.CCbColorProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM pCbColor AS CCbColor PTR

   SELECT CASE uMsg

      CASE WM_CREATE
         EXIT FUNCTION

      CASE WM_DESTROY
         ' // Remove the subclassing
         SetWindowLongPtrW hwnd, GWLP_WNDPROC, CAST(LONG_PTR, RemovePropW(hwnd, "OLDWNDPROC"))
         ' // Remove the property
         .RemovePropW(hwnd, "CBCOLORPTR")
         EXIT FUNCTION

      CASE WM_DRAWITEM
         ' // Get a pointer to the class
         pCbColor = CAST(CCbColor PTR, .GetPropW(hwnd, "CBCOLORPTR"))
         IF pCbColor = NULL THEN EXIT FUNCTION
         SCOPE
         DIM hBrush AS HBRUSH, wszText AS WSTRING * 80
         DIM lpdis AS DRAWITEMSTRUCT PTR, rc AS RECT
         lpdis = CAST(DRAWITEMSTRUCT PTR, lParam)
         IF lpdis->itemID = &hFFFFFFFF THEN EXIT FUNCTION
         SELECT CASE lpdis->itemAction
            CASE ODA_DRAWENTIRE, ODA_SELECT
               ' // Clear background
               FillRect lpdis->hDC, @lpdis->rcItem, .GetSysColorBrush(COLOR_WINDOW)
               ' // Get/draw text
               .SendMessageW hwnd, CB_GETLBTEXT, lpdis->itemID, CAST(.LPARAM, @wszText)
               .SetBkColor lpdis->hDC, GetSysColor(COLOR_WINDOW)
               .SetTextColor lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT)
               rc = lpdis->rcItem
               rc.Left = 28 * (pCbColor->m_ratio * 0.72)
               .DrawTextW lpdis->hDC, @wszText, LEN(wszText), @rc, DT_SINGLELINE OR DT_LEFT OR DT_VCENTER
               ' // If selected item...
               IF (lpdis->itemState AND ODS_SELECTED) THEN
                  ' // If not ODS_COMBOBOXEDIT...
                  IF (lpdis->itemState AND ODS_COMBOBOXEDIT) = 0 THEN
                     ' // Set coordinates
                     rc.Left = 26 * (pCbColor->m_ratio * 0.72)
                     rc.Right = lpdis->rcItem.Right
                     rc.Top = lpdis->rcItem.Top
                     rc.Bottom = lpdis->rcItem.Bottom
                     ' // Invert area around text only
                     .InvertRect lpdis->hDC, @rc
                  END IF
                  ' // and draw a focus rectangle around all
                  .DrawFocusRect lpdis->hDC, @lpdis->rcItem
               END IF
               ' // Paint color rectangle using RoundRect for nicer looks
               ' /// If ODS_COMBOBOXEDIT...
               IF (lpdis->itemState AND ODS_COMBOBOXEDIT) THEN
                  ' // Set coordinates
                  rc.Left = 4 * (pCbColor->m_ratio * 0.72)
                  rc.Right = 24 * (pCbColor->m_ratio * 0.72)
               ELSE
                  ' // A tiny bit to the left in list...
                  rc.Left = 3 * (pCbColor->m_ratio * 0.72)
                  rc.Right = 23 * (pCbColor->m_ratio * 0.72)
               END IF
               rc.Top    = (lpdis->rcItem.Top + (2 * pCbColor->m_ratio))
               rc.Bottom = (lpdis->rcItem.Bottom - (2 * pCbColor->m_ratio))
               ' // Create brush with the appropriate color
               hBrush = .CreateSolidBrush(pCbColor->GetQBColor(hWnd, lpdis->itemID, FALSE))
               ' // Select brush into device context
               hBrush = .SelectObject(lpdis->hDC, hBrush)
               ' // Draw the rectangle
               DIM nCorner AS LONG = pCbColor->m_ratio * 0.72
               .RoundRect(lpdis->hDC, rc.Left, rc.Top, rc.Right, rc.Bottom, nCorner, nCorner)
               ' // Select old brush back and delete new one
               .DeleteObject .SelectObject(lpdis->hDC, hBrush)
            END SELECT
         END SCOPE
         FUNCTION = CTRUE
         EXIT FUNCTION

      ' // Set auto color
      CASE CBCOL_SETAUTOCOLOR
         pCbColor = CAST(CCbColor PTR, .GetPropW(hwnd, "CBCOLORPTR"))
         IF pCbColor = NULL THEN EXIT FUNCTION
         pCbColor->m_AutoColor = wParam
         EXIT FUNCTION

      ' // Set auto color
      CASE CBCOL_GETAUTOCOLOR
         pCbColor = CAST(CCbColor PTR, .GetPropW(hwnd, "CBCOLORPTR"))
         IF pCbColor = NULL THEN EXIT FUNCTION
         FUNCTION = pCbColor->m_AutoColor
         EXIT FUNCTION

      ' // Set user color
      CASE CBCOL_SETUSERCOLOR
         pCbColor = CAST(CCbColor PTR, .GetPropW(hwnd, "CBCOLORPTR"))
         IF pCbColor = NULL THEN EXIT FUNCTION
         pCbColor->m_UserColor = wParam
         EXIT FUNCTION

      ' // Get user color
      CASE CBCOL_GETUSERCOLOR
         pCbColor = CAST(CCbColor PTR, .GetPropW(hwnd, "CBCOLORPTR"))
         IF pCbColor = NULL THEN EXIT FUNCTION
         FUNCTION = pCbColor->m_UserColor
         EXIT FUNCTION

      ' // Return selected color
      CASE CBCOL_GETSELCOLOR
         DIM nRes AS LONG
         nRes = SendMessageW(hwnd, CB_GETCURSEL, 0, 0)
         pCbColor = CAST(CCbColor PTR, .GetPropW(hwnd, "CBCOLORPTR"))
         IF pCbColor = NULL THEN EXIT FUNCTION
         IF nRes > CB_ERR THEN
            FUNCTION = pCbColor->GetQBColor(hwnd, nRes, CTRUE)
         END IF
         EXIT FUNCTION


   END SELECT

   ' // Default processing for other messages.
   FUNCTION = CallWindowProcW(GetPropW(hwnd, "OLDWNDPROC"), hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' CCbColor class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CCbColor (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER,  _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwAutoColor AS COLORREF = 0, BYVAL dwUserColor AS COLORREF = 0, _
   BYVAL prgNames AS WSTRING PTR = NULL, BYVAL numBytes AS DWORD = 0)

   ' // Color names
   IF prgNames <> NULL AND numBytes < (18 * 20 * 2) + 1 THEN
      .memcpy @m_rgNames(0), prgNames, numBytes
   END IF

   ' // Create the control
   IF pWindow THEN
      m_hCtl = pWindow->AddControl("COMBOBOX", pWindow->hWindow, cID, "", _
      x, y, nWidth, nHeight, WS_CHILD OR WS_VISIBLE OR CBS_OWNERDRAWFIXED OR CBS_HASSTRINGS OR _
      CBS_DROPDOWNLIST OR WS_TABSTOP OR CBS_DISABLENOSCROLL OR WS_VSCROLL, WS_EX_CLIENTEDGE, _
      NULL, CAST(WNDPROC, @CCbColorProc))
      m_ratio = pWindow->rxRatio
   END IF
   IF m_hCtl THEN
      ' // Set the same font used by the parent
      DIM lfw AS LOGFONTW
      IF pWindow->Font THEN
         IF .GetObjectW(pWindow->Font, SIZEOF(lfw), @lfw) THEN m_hFont = CreateFontIndirectW(@lfw)
      END IF
      ' // Store a pointer to the class as a property
      .SetPropW(m_hCtl, "CBCOLORPTR", CAST(HANDLE, @this))
      DIM i AS LONG
      FOR i = 0 TO UBOUND(m_rgNames)
         .SendMessageW m_hCtl, CB_ADDSTRING, 0, CAST(LPARAM, @m_rgNames(i))
      NEXT
      ' // Store default values (usually COLOR_WINDOW or COLOR_WINDOWTEXT)
      m_AutoColor = dwAutoColor
      m_UserColor = dwUserColor
      ' // Create initial colormap for the ChooseColor dialog
      DIM cl AS LONG
      FOR i = 1 TO 16
         cl = i * 16 - 1
         m_rgColors(16 - i) = BGR(cl, cl, cl)
      NEXT
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CCbColor class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CCbColor
   ' // Free resources
   IF m_hFont THEN DeleteObject m_hFont
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the button
' ========================================================================================
PRIVATE FUNCTION CCbColor.hWindow () AS HWND
   FUNCTION = m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Basic QB color function.
' ========================================================================================
PRIVATE FUNCTION CCbColor.GetQBColor (BYVAL hwnd AS HWND, BYVAL c AS LONG, BYVAL dlg AS LONG) AS COLORREF

   SELECT CASE c
      CASE  0 : FUNCTION = m_AutoColor          ' // Pre-set system color, like COLOR_WINDOW or COLOR_WINDOWTEXT
      CASE  1 : FUNCTION = BGR(0,     0,   0)   ' // Black
      CASE  2 : FUNCTION = BGR(0,     0, 128)   ' // Blue
      CASE  3 : FUNCTION = BGR(0,   128,   0)   ' // Green
      CASE  4 : FUNCTION = BGR(0,   128, 128)   ' // Cyan
      CASE  5 : FUNCTION = BGR(196,   0,   0)   ' // Red
      CASE  6 : FUNCTION = BGR(128,   0, 128)   ' // Magenta
      CASE  7 : FUNCTION = BGR(128,  64,   0)   ' // Brown
      CASE  8 : FUNCTION = BGR(196, 196, 196)   ' // Light Gray
      CASE  9 : FUNCTION = BGR(128, 128, 128)   ' // Gray
      CASE 10 : FUNCTION = BGR(0,     0, 255)   ' // Light Blue
      CASE 11 : FUNCTION = BGR(0,   255,   0)   ' // Light Green
      CASE 12 : FUNCTION = BGR(0,   255, 255)   ' // Light Cyan
      CASE 13 : FUNCTION = BGR(255,   0,   0)   ' // Light Red
      CASE 14 : FUNCTION = BGR(255,   0, 255)   ' // Light Magenta
      CASE 15 : FUNCTION = BGR(255, 255,   0)   ' // Yellow
      CASE 16 : FUNCTION = BGR(255, 255, 255)   ' // White
      CASE 17
         ' // Display the ChooseColor dialog
         IF dlg THEN
            DIM cc AS CHOOSECOLORW
            ' // Size of the structure
            cc.lStructSize  = SIZEOF(cc)
            ' // Pointer to the custom defined gray color array
            cc.lpCustColors = @m_rgColors(0)
            cc.Flags        = CC_RGBINIT OR CC_FULLOPEN
            cc.hwndowner    = hwnd
            ' // Allow the dialog to "auto-select" previously selected color
            ' // (only works for base colors, but still better than nothing)
            cc.rgbResult    = m_UserColor
            IF .ChooseColorW(@cc) THEN
               m_UserColor = cc.rgbResult
              .InvalidateRect hwnd, NULL, 0
              .UpdateWindow hWnd
            END IF
         END IF
         ' // Return the user selected color
         FUNCTION = m_UserColor
   END SELECT

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets/sets the RGB auto color
' ========================================================================================
PRIVATE PROPERTY CCbColor.AutoColor (BYVAL nColor AS COLORREF)
   m_autoColor = nColor
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CCbColor.AutoColor () AS COLORREF
   PROPERTY = m_autoColor
END PROPERTY

' ========================================================================================
' Gets/sets the RGB user color
' ========================================================================================
PRIVATE PROPERTY CCbColor.UserColor (BYVAL nColor AS COLORREF)
   m_autoColor = nColor
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CCbColor.UserColor () AS COLORREF
   PROPERTY = m_autoColor
END PROPERTY
' ========================================================================================

' ========================================================================================
' Get the selected color
' ========================================================================================
PRIVATE PROPERTY CCbColor.SelColor () AS COLORREF
   DIM nSel AS LRESULT
   nSel = SendMessageW(m_hCtl, CB_GETCURSEL, 0, 0)
   IF nSel > CB_ERR THEN PROPERTY = this.GetQBColor(m_hCtl, nSel, 1)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets/sets the width scaling ratio
' ========================================================================================
PRIVATE PROPERTY CCbColor.ScalingRatio (BYVAL rxRatio AS SINGLE)
   m_ratio = rxRatio
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CCbColor.ScalingRatio () AS SINGLE
   PROPERTY = m_ratio
END PROPERTY
' ========================================================================================

END NAMESPACE
