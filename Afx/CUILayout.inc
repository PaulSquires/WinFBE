' ########################################################################################
' Microsoft Windows
' File: CUILayout.inc
' Contents: Layout manager
' Purpose: When a parent window is resized, this class manages the location and size
' of the child windows
' Based in UILayout.h (c) 2000 Jeffrey Richter
' Compiler: FreeBasic 32 & 64-bit
' Copyright for the FreeBASIC adaptation (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

#include once "windows.bi"
#include once "Afx/CWindow.inc"

NAMESPACE Afx.CUILayoutClass

ENUM UIL_ANCHORPOINT
   AP_TOPLEFT = 0
   AP_TOPMIDDLE
   AP_TOPRIGHT
   AP_MIDDLERIGHT
   AP_BOTTOMRIGHT
   AP_BOTTOMMIDDLE
   AP_BOTTOMLEFT
   AP_MIDDLELEFT
   AP_CENTER
END ENUM

TYPE UIL_CONTROL
   m_nID AS LONG
   m_fRedraw AS BOOLEAN
   m_apUpperLeft AS UIL_ANCHORPOINT
   m_apLowerRight AS UIL_ANCHORPOINT
   m_ptULDelta AS POINT
   m_ptLRDelta AS POINT
END TYPE

' ========================================================================================
' CUILayout class
' ========================================================================================
TYPE CUILayout

   Private:
      DIM m_CtrlInfo(255) AS UIL_CONTROL   ' // allow the maximum controls for a dialog box template
      m_nNumControls AS LONG               ' // Number of controls
      m_hwndParent AS HWND                 ' // Handle of the parent window

      DECLARE SUB PixelFromAnchorPoint (BYVAL ap AS LONG, BYVAL cxParent AS LONG, BYVAL cyParent AS LONG, BYVAL ppt AS POINT PTR)

   Public:

      DECLARE CONSTRUCTOR (BYVAL hwndParent AS HWND)
      DECLARE FUNCTION AnchorControl (BYVAL apUpperLeft AS LONG, BYVAL apLowerRight AS LONG, BYVAL nID AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS BOOLEAN
      DECLARE FUNCTION AdjustControls (BYVAL cx AS LONG, BYVAL cy AS LONG) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
' CUILayout class constructor
' ========================================================================================
CONSTRUCTOR CUILayout (BYVAL hwndParent AS HWND)
   m_hwndParent = hwndParent
   m_nNumControls = 0
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Calculate coordinate from anchor point
' ========================================================================================
SUB CUILayout.PixelFromAnchorPoint (BYVAL ap AS LONG, BYVAL cxParent AS LONG, BYVAL cyParent AS LONG, BYVAL ppt AS POINT PTR)

   ppt->x = 0
   ppt->y = 0

   SELECT CASE ap
      CASE AP_TOPMIDDLE, AP_CENTER, AP_BOTTOMMIDDLE
         ppt->x = cxParent / 2
      CASE AP_TOPRIGHT, AP_MIDDLERIGHT, AP_BOTTOMRIGHT
         ppt->x = cxParent
   END SELECT

   SELECT CASE ap
      CASE AP_MIDDLELEFT, AP_CENTER, AP_MIDDLERIGHT
         ppt->y = cyParent / 2
      CASE AP_BOTTOMLEFT, AP_BOTTOMMIDDLE, AP_BOTTOMRIGHT
         ppt->y = cyParent
   END SELECT

END SUB
' ========================================================================================

' ========================================================================================
' Anchors a control
' ========================================================================================
FUNCTION CUILayout.AnchorControl (BYVAL apUpperLeft AS LONG, BYVAL apLowerRight AS LONG, BYVAL nID AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS BOOLEAN

   DIM fOk AS BOOLEAN = FALSE

   ' // Fake DO.. LOOP to avoid GOTOs
   ' // There is an EXIT LOOP inmediately before LOOP, so the loop only executes once
   DO
      DIM hwndControl AS HWND = GetDlgItem(m_hwndParent, nID)
      IF hwndControl = NULL THEN EXIT DO
      IF m_nNumControls >= UBOUND(m_CtrlInfo) THEN EXIT DO
      m_CtrlInfo(m_nNumControls).m_nID = nID
      m_CtrlInfo(m_nNumControls).m_fRedraw = fRedraw
      m_CtrlInfo(m_nNumControls).m_apUpperLeft = apUpperLeft
      m_CtrlInfo(m_nNumControls).m_apLowerRight = apLowerRight
      ' // Control item coordinates
      DIM rcControl AS RECT
      GetWindowRect(hwndControl, @rcControl)
      ' // convert coords to the relative parent window
      MapWindowPoints(HWND_DESKTOP, m_hwndParent, CAST(POINT PTR, @rcControl), 2)
      DIM rcParent AS RECT
      GetClientRect(m_hwndParent, @rcParent)
      ' //
      DIM pt AS POINT
      PixelFromAnchorPoint(apUpperLeft, rcParent.right, rcParent.bottom, @pt)
      m_CtrlInfo(m_nNumControls).m_ptULDelta.x = pt.x - rcControl.left
      m_CtrlInfo(m_nNumControls).m_ptULDelta.y = pt.y - rcControl.top
      PixelFromAnchorPoint(apLowerRight, rcParent.right, rcParent.bottom, @pt)
      m_CtrlInfo(m_nNumControls).m_ptLRDelta.x = pt.x - rcControl.right
      m_CtrlInfo(m_nNumControls).m_ptLRDelta.y = pt.y - rcControl.bottom
      m_nNumControls += 1
      fOk = TRUE
      EXIT DO
   LOOP

   FUNCTION = fOk

END FUNCTION
' ========================================================================================

' ========================================================================================
' Adjust the controls
' ========================================================================================
FUNCTION CUILayout.AdjustControls (BYVAL cx AS LONG, BYVAL cy AS LONG) AS BOOLEAN

   DIM fOk AS BOOLEAN = FALSE

   ' // Created by composition of presence of all of the control items in the region
   DIM hrgnPaint AS HRGN = CreateRectRgn(0, 0, 0, 0)
   DIM n AS LONG
   FOR n = 0 TO m_nNumControls - 1
      DIM hwndControl AS HWND = GetDlgItem(m_hwndParent, m_CtrlInfo(n).m_nID)
      ' // Screen coordinates of the control
      DIM rcControl AS RECT
      GetWindowRect(hwndControl, @rcControl)
      ' // Convert coordinates relative to the parent window
      MapWindowPoints(HWND_DESKTOP, m_hwndParent, CAST(POINT PTR, @rcControl), 2)
      DIM hrgnTemp AS HRGN = CreateRectRgnIndirect(@rcControl)
      CombineRgn(hrgnPaint, hrgnPaint, hrgnTemp, RGN_OR)
      DeleteObject(hrgnTemp)
   NEXT

   FOR n = 0 TO m_nNumControls - 1
      ' // Obtain control on left margin, and use the parent window's width / height
      DIM rcControl AS RECT
      PixelFromAnchorPoint(m_CtrlInfo(n).m_apUpperLeft, cx, cy, CAST(POINT PTR, @rcControl))
      rcControl.left -= m_CtrlInfo(n).m_ptULDelta.x
      rcControl.top  -= m_CtrlInfo(n).m_ptULDelta.y
      PixelFromAnchorPoint(m_CtrlInfo(n).m_apLowerRight, cx, cy, CAST(POINT PTR, @rcControl.right))
      rcControl.right   -= m_CtrlInfo(n).m_ptLRDelta.x
      rcControl.bottom  -= m_CtrlInfo(n).m_ptLRDelta.y
      ' //
      DIM hwndControl AS HWND = GetDlgItem(m_hwndParent, m_CtrlInfo(n).m_nID)
      MoveWindow(hwndControl, rcControl.left, rcControl.top, _
         rcControl.right - rcControl.left, _
         rcControl.bottom - rcControl.top, FALSE)
      IF m_CtrlInfo(n).m_fRedraw THEN
         InvalidateRect(hwndControl, NULL, FALSE)
      ELSE
         DIM hrgnTemp AS HRGN = CreateRectRgnIndirect(@rcControl)
         CombineRgn(hrgnPaint, hrgnPaint, hrgnTemp, RGN_DIFF)
         DeleteObject(hrgnTemp)
         InvalidateRect(hwndControl, NULL, TRUE)
         SendMessage(hwndControl, WM_NCPAINT, 1, 0)
         UpdateWindow(hwndControl)
      END IF
   NEXT

   DIM hdc AS HDC = GetDC(m_hwndParent)
   FillRgn(hdc, hrgnPaint, CAST(HBRUSH, GetClassLongPtr(GetParent(m_hwndParent), GCLP_HBRBACKGROUND)))
   ReleaseDC(m_hwndParent, hdc)
   DeleteObject(hrgnPaint)
   InvalidateRect(m_hwndParent, NULL, TRUE)
   UpdateWindow(m_hwndParent)
   FUNCTION = fOk

END FUNCTION
' ========================================================================================

END NAMESPACE
