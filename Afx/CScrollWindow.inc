' ########################################################################################
' Microsoft Windows
' File: CScrollWindow.inc
' Contents: Scrolls the contents of a window client area.
' Compiler: Free Basic
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"

NAMESPACE Afx.CScrollWindowClass

' ========================================================================================
' CScrollWindow
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
' CScrollWindow class constructor
' ========================================================================================
CONSTRUCTOR CScrollWindow (BYVAL hwnd AS HWND)
   m_hwnd = hwnd
   m_HorzUnits = 10
   m_VertUnits = 10
   GetClientRect(m_hwnd, @m_ClientRect)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CScrollWindow class destructor
' ========================================================================================
DESTRUCTOR CScrollWindow
'   OutputDebugStringW "CScrollWindow.Destructor"
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Sets the amount, in device units, of horizontal and vertical scrolling
' ========================================================================================
SUB CScrollWindow.SetClientRect (BYVAL pRect AS RECT PTR)
   IF pRect = NULL THEN EXIT SUB
   m_CLientRect = *pRect
END SUB
' ========================================================================================

' ========================================================================================
' Sets the amount, in device units, of horizontal and vertical scrolling
' ========================================================================================
SUB CScrollWindow.SetScrollingUnits (BYVAL nHorzUnits AS LONG, BYVAL nVertUnits AS LONG)
   m_HorzUnits = nHorzUnits
   m_VertUnits = nVertUnits
END SUB
' ========================================================================================

' ========================================================================================
' Resets the window scrolling information
' ========================================================================================
SUB CScrollWindow.ResetScrollbars
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
SUB CScrollWindow.SetupScrollbars

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
SUB CScrollWindow.OnVScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

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
SUB CScrollWindow.OnHScroll (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

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
SUB CScrollWindow.OnSize (BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
   this.ResetScrollbars
   this.SetupScrollbars
END SUB
' ========================================================================================

END NAMESPACE
