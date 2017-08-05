' ########################################################################################
' Microsoft Windows
' File: CLayout.inc
' Contents: Layout manager
' Purpose: When a parent window is resized, this class manages the location and size
' of the child windows.
' Based on code originally written by Patrice Terrier and used with his permission.
' Compiler: FreeBasic 32 & 64-bit
' Copyright for the FreeBASIC adaptation (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/CWindow.inc"

NAMESPACE Afx

ENUM AFX_ANCHORPOINT
   AFX_ANCHOR_NONE =  0
   AFX_ANCHOR_WIDTH
   AFX_ANCHOR_RIGHT
   AFX_ANCHOR_CENTER_HORZ
   AFX_ANCHOR_HEIGHT
   AFX_ANCHOR_HEIGHT_WIDTH
   AFX_ANCHOR_HEIGHT_RIGHT
   AFX_ANCHOR_BOTTOM
   AFX_ANCHOR_BOTTOM_WIDTH
   AFX_ANCHOR_BOTTOM_RIGHT
   AFX_ANCHOR_CENTER_HORZ_BOTTOM
   AFX_ANCHOR_CENTER_VERT
   AFX_ANCHOR_CENTER_VERT_RIGHT
   AFX_ANCHOR_CENTER
END ENUM

TYPE AFX_ANCHORPROPERTY
   hwndCtl AS HWND
   anchor  AS LONG
   rc      AS RECT
   centerX AS LONG
   centerY AS LONG
END TYPE

' ========================================================================================
' CLayout class
' ========================================================================================
TYPE CLayout

   Private:
      m_hwndParent AS HWND                       ' // Handle of the parent window
      DIM m_rgProps(ANY) AS AFX_ANCHORPROPERTY   ' // Array of anchor property structures

   Public:
      DECLARE CONSTRUCTOR (BYVAL hwndParent AS HWND)
      DECLARE FUNCTION GetAnchorItem (BYVAL hwndCtl AS HWND) AS LONG
      DECLARE FUNCTION AnchorControl OVERLOAD (BYVAL hwndCtl AS HWND, BYVAL nAnchorMode AS LONG) AS BOOLEAN
      DECLARE FUNCTION AnchorControl OVERLOAD (BYVAL cID AS LONG, BYVAL nAnchorMode AS LONG) AS BOOLEAN
      DECLARE SUB AdjustControls
      ' // Callback function to enumerate child windows
      DECLARE STATIC FUNCTION AfxLayoutEnumChildProc (BYVAL hwnd AS HWND, BYVAL lParam AS LPARAM) AS LONG

END TYPE
' ========================================================================================

' ========================================================================================
' CLayout class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CLayout (BYVAL hwndParent AS HWND)
   m_hwndParent = hwndParent
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Get the anchor item
' Parameter:
' - hwndCtl = The control's window handle
' Return Value:
'   The array index on success, or -1 on failure.
' ========================================================================================
PRIVATE FUNCTION CLayout.GetAnchorItem (BYVAL hwndCtl AS HWND) AS LONG
   FUNCTION = -1
   DIM i AS LONG
   FOR i = 0 TO UBOUND(m_rgProps)
      IF m_rgProps(i).hwndCtl = hwndCtl THEN FUNCTION = i : EXIT FOR
   NEXT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Anchor the control
' ========================================================================================
PRIVATE FUNCTION CLayout.AnchorControl OVERLOAD (BYVAL hwndCtl AS HWND, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   IF .IsWindow(hwndCtl) = FALSE THEN EXIT FUNCTION
   ' // Chech if the item already exists
   DIM nItem AS LONG = this.GetAnchorItem(hwndCtl)
   ' // If it does not exist, make room for it
   IF nItem = -1 THEN
      REDIM PRESERVE m_rgProps(UBOUND(m_rgProps) + 1) AS AFX_ANCHORPROPERTY
      nItem = UBOUND(m_rgProps)
   END IF
   m_rgProps(nItem).hwndCtl = hwndCtl
   ' // Get the coordinates of the control
   DIM rcCtl AS RECT, pt AS POINT
   .GetWindowRect(hwndCtl, @rcCtl)
   pt.x = rcCtl.Left
   pt.y = rcCtl.Top
   ' // Convert the screen coordinates to screen to client-area coordinates
   .ScreenToClient(GetParent(hwndCtl), @pt)
   ' // Get the client coordinates of the parent window
   DIm rcParent AS RECT
   .GetClientRect(GetParent(hwndCtl), @rcParent)
   m_rgProps(nItem).anchor    = MIN(MAX(nAnchorMode, AFX_ANCHOR_NONE), AFX_ANCHOR_CENTER)
   m_rgProps(nItem).rc.Left   = pt.x
   m_rgProps(nItem).rc.Top    = pt.y
   m_rgProps(nItem).rc.Right  = rcParent.Right - (rcCtl.Right - rcCtl.Left + pt.x)
   m_rgProps(nItem).rc.Bottom = rcParent.Bottom - (rcCtl.Bottom - rcCtl.Top + pt.y)
   m_rgProps(nItem).centerX   = pt.x - (rcParent.Right \ 2)
   m_rgProps(nItem).centerY   = pt.y - (rcParent.Bottom \ 2)
   FUNCTION = TRUE
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CLayout.AnchorControl OVERLOAD (BYVAL cID AS LONG, BYVAL nAnchorMode AS LONG) AS BOOLEAN
   DIM hwndCtl AS HWND = GetDlgItem(m_hwndParent, cID)
   IF hwndCtl THEN FUNCTION = this.AnchorControl(hwndCtl, nAnchorMode)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Callback function to enumerate child windows
' ========================================================================================
PRIVATE FUNCTION CLayout.AfxLayoutEnumChildProc (BYVAL hwndChild AS HWND, BYVAL lParam AS LPARAM) AS LONG

   DIM rcParent AS RECT, rcChild AS RECT
   DIM nWidth AS LONG, nHeight AS LONG, x AS LONG, y AS LONG, xWidth AS LONG, yHeight AS LONG
   DIM pLayout AS CLayout PTR = CAST(CLayout PTR, lParam)
   IF pLayout THEN
      DIM nItem AS LONG = pLayout->GetAnchorItem(hwndChild)
      IF nItem > -1 THEN
         IF pLayout->m_rgProps(nItem).anchor > AFX_ANCHOR_NONE THEN
            GetWindowRect(hwndChild, @rcChild)
            nWidth = rcChild.Right - rcChild.Left
            nHeight = rcChild.Bottom - rcChild.Top
            GetClientRect(GetParent(hwndChild), @rcParent)
            x = 0: y = 0: xWidth = 0: yHeight = 0
            SELECT CASE pLayout->m_rgProps(nItem).anchor
               CASE AFX_ANCHOR_NONE
                  x = pLayout->m_rgProps(nItem).rc.Left
                  y = pLayout->m_rgProps(nItem).rc.Top
                  xWidth = MAX(rcParent.Right - pLayout->m_rgProps(nItem).rc.Left - pLayout->m_rgProps(nItem).rc.Right, 0)
                  yHeight = nHeight
               CASE AFX_ANCHOR_WIDTH
                  x = pLayout->m_rgProps(nItem).rc.Left
                  y = pLayout->m_rgProps(nItem).rc.Top
                  xWidth = MAX(rcParent.Right - pLayout->m_rgProps(nItem).rc.Left - pLayout->m_rgProps(nItem).rc.Right, 0)
                  yHeight = nHeight
               CASE AFX_ANCHOR_RIGHT
                  x = rcParent.Right - nWidth - pLayout->m_rgProps(nItem).rc.Right
                  y = pLayout->m_rgProps(nItem).rc.Top
                  xWidth = nWidth
                  yHeight = nHeight
               CASE AFX_ANCHOR_CENTER_HORZ
                  x = (rcParent.Right \ 2) + pLayout->m_rgProps(nItem).centerX
                  y = pLayout->m_rgProps(nItem).rc.Top
                  xWidth = nWidth
                  yHeight = nHeight
               CASE AFX_ANCHOR_HEIGHT
                  x = pLayout->m_rgProps(nItem).rc.Left
                  y = pLayout->m_rgProps(nItem).rc.Top
                  xWidth = nWidth
                  yHeight = MAX(rcParent.Bottom - pLayout->m_rgProps(nItem).rc.Top - pLayout->m_rgProps(nItem).rc.Bottom, 0)
               CASE AFX_ANCHOR_HEIGHT_WIDTH
                  x = pLayout->m_rgProps(nItem).rc.Left
                  y = pLayout->m_rgProps(nItem).rc.Top
                  xWidth = MAX(rcParent.Right - pLayout->m_rgProps(nItem).rc.Left - pLayout->m_rgProps(nItem).rc.Right, 0)
                  yHeight = MAX(rcParent.Bottom - pLayout->m_rgProps(nItem).rc.Top - pLayout->m_rgProps(nItem).rc.Bottom, 0)
               CASE AFX_ANCHOR_HEIGHT_RIGHT
                  x = rcParent.Right - nWidth - pLayout->m_rgProps(nItem).rc.Right
                  y = pLayout->m_rgProps(nItem).rc.Top
                  xWidth = nWidth
                  yHeight = MAX(rcParent.Bottom - pLayout->m_rgProps(nItem).rc.Top - pLayout->m_rgProps(nItem).rc.Bottom, 0)
               CASE AFX_ANCHOR_BOTTOM
                  x = pLayout->m_rgProps(nItem).rc.Left
                  y = rcParent.Bottom - pLayout->m_rgProps(nItem).rc.Bottom - nHeight
                  xWidth = nWidth
                  yHeight = nHeight
               CASE AFX_ANCHOR_BOTTOM_WIDTH
                  x = pLayout->m_rgProps(nItem).rc.Left
                  y = rcParent.Bottom - pLayout->m_rgProps(nItem).rc.Bottom - nHeight
                  xWidth = MAX(rcParent.Right - pLayout->m_rgProps(nItem).rc.Left - pLayout->m_rgProps(nItem).rc.Right, 0)
                  yHeight = nHeight
               CASE AFX_ANCHOR_BOTTOM_RIGHT
                  x = rcParent.Right - nWidth - pLayout->m_rgProps(nItem).rc.Right
                  y = rcParent.Bottom - pLayout->m_rgProps(nItem).rc.Bottom - nHeight
                  xWidth = nWidth
                  yHeight = nHeight
               CASE AFX_ANCHOR_CENTER_HORZ_BOTTOM
                  x = (rcParent.Right \ 2) + pLayout->m_rgProps(nItem).centerX
                  y = rcParent.Bottom - pLayout->m_rgProps(nItem).rc.Bottom - nHeight
                  xWidth = nWidth
                  yHeight = nHeight
               CASE AFX_ANCHOR_CENTER_VERT
                  x = pLayout->m_rgProps(nItem).rc.Left
                  y = (rcParent.Bottom - nHeight) \ 2
                  xWidth = nWidth
                  yHeight = nHeight
               CASE AFX_ANCHOR_CENTER_VERT_RIGHT
                  x = rcParent.Right - nWidth - pLayout->m_rgProps(nItem).rc.Right
                  y = (rcParent.Bottom - nHeight) \ 2
                  xWidth = nWidth
                  yHeight = nHeight
               CASE AFX_ANCHOR_CENTER
                  x = (rcParent.Right \ 2) + pLayout->m_rgProps(nItem).centerX
                  y = (rcParent.Bottom \ 2) + pLayout->m_rgProps(nItem).centerY
                  xWidth = nWidth
                  yHeight = nHeight
            END SELECT
            MoveWindow(hwndChild, x, y, xWidth, yHeight, TRUE)
         END IF
      END IF
   END IF

   FUNCTION = CTRUE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Adjust the controls size and/or position
' ========================================================================================
PRIVATE SUB CLayout.AdjustControls
   EnumChildWindows(m_hwndParent, CAST(WNDENUMPROC, @AfxLayoutEnumChildProc), CAST(LPARAM, @this))
END SUB
' ========================================================================================

END NAMESPACE
