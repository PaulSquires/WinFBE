'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2022 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

#include once "modVDAnchors.bi"

dim shared bPreventItemChanged as boolean

' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmVDAnchors
' ========================================================================================
function frmVDAnchors_OnNotify( _
         ByVal HWnd As HWnd, _
         ByVal id As Long, _
         ByVal pNMHDR As NMHDR Ptr _
         ) As LRESULT

   select case id
      case IDC_FRMVDANCHORS_LISTVIEW
         select case pNMHDR->code
            case LVN_ITEMCHANGED
               if bPreventItemChanged then exit function
               dim as CWSTR wszValue
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
               dim pProp as clsProperty ptr = GetActivePropertyPtr()
               if pProp then
                  dim as HWND hLV = pNMHDR->hwndFrom
                  if ListView_GetCheckState(hLV, 0) then wszValue = wszValue & "Left,"
                  if ListView_GetCheckState(hLV, 1) then wszValue = wszValue & "Top,"
                  if ListView_GetCheckState(hLV, 2) then wszValue = wszValue & "Right,"
                  if ListView_GetCheckState(hLV, 3) then wszValue = wszValue & "Bottom,"
                  wszValue = rtrim(wszValue, ",")
                  pProp->wszPropValuePrev = pProp->wszPropValue 
                  pProp->wszPropValue = wszValue
                  pDoc->UserModified = true
                  pDoc->bRegenerateCode = true
               end if
         end select

   end select

   function = 0
end function

   
' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmVDAnchors
' ========================================================================================
function frmVDAnchors_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ' Never close the window; simply hide it.
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmVDAnchors
' ========================================================================================
function frmVDAnchors_OnDestroy( byval HWnd As HWnd ) As LRESULT
   HWND_FRMVDANCHORS = 0
   Function = 0
End Function


' ========================================================================================
' frmVDAnchors Window procedure
' ========================================================================================
function frmVDAnchors_WndProc( _
         ByVal HWnd   As HWnd, _
         ByVal uMsg   As UINT, _
         ByVal wParam As WPARAM, _
         ByVal lParam As LPARAM _
         ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmVDAnchors_OnNotify)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmVDAnchors_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmVDAnchors_OnDestroy)

   case WM_ACTIVATE
      if wParam = WA_INACTIVE then
         ShowWindow(HWND_FRMVDANCHORS, SW_HIDE)
      end if   

   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmVDAnchors_SetListViewSelections
' ========================================================================================
function frmVDAnchors_SetListViewSelections( byval hLV as HWND, byref wszPropValue as wstring ) as Long

   ' wszPropValue is a comma separated list of parts
   ' eg.  Left,Top
   '      Right,Bottom

   ' Clear any existing checkmarks
   for i as long = 0 to 3
      ListView_SetCheckState(hLV, i, FALSE)
   next

   dim as long numParts = AfxStrParseCount(wszPropValue, ",")
   for i as long = 1 to numParts
      dim as CWSTR wszPart = ucase(trim(AfxStrParse(wszPropValue, i, ",")))
      select case wszPart
         case "LEFT":   ListView_SetCheckState(hLV, 0, CTRUE)
         case "TOP":    ListView_SetCheckState(hLV, 1, CTRUE)
         case "RIGHT":  ListView_SetCheckState(hLV, 2, CTRUE)
         case "BOTTOM": ListView_SetCheckState(hLV, 3, CTRUE)
      end select
   next
      
   bPreventItemChanged = FALSE
   function = 0
end function


' ========================================================================================
' frmVDAnchors_Show
' ========================================================================================
function frmVDAnchors_Show( _
         ByVal hWndParent As HWnd, _
         byref wszPropValue as wstring _
         ) as LRESULT

   Dim As HWnd hLV

   bPreventItemChanged = TRUE

   ' If the Anchors popup already exists then no need to recreate it.
   If IsWindow(HWND_FRMVDANCHORS) Then 
      hLV = GetDlgItem(HWND_FRMVDANCHORS, IDC_FRMVDANCHORS_LISTVIEW)
      frmVDAnchors_SetListViewSelections(hLV, wszPropValue)
      exit function
   end if

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMVDANCHORS = _
   pWindow->Create( hWndParent, "", @frmVDAnchors_WndProc, 0, 0, 140, 78, _
        WS_POPUP or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN or WS_BORDER, 0)
   
   hLV = _ 
        pWindow->AddControl("LISTVIEW", , IDC_FRMVDANCHORS_LISTVIEW, "", 0, 0, 140, 78, _
        WS_CHILD or WS_VISIBLE OR WS_CLIPCHILDREN OR WS_TABSTOP OR LVS_SINGLESEL or LVS_REPORT OR LVS_ALIGNTOP or LVS_NOCOLUMNHEADER, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

   ' Configure the ListView
   dim as long dwExStyle = ListView_GetExtendedListViewStyle(hLV)
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_DOUBLEBUFFER Or LVS_EX_CHECKBOXES
   ListView_SetExtendedListViewStyle(hLV, dwExStyle)
   ListView_AddColumn( hLV, 0, "", pWindow->ScaleX(20) )
   ListView_AddColumn( hLV, 1, "", pWindow->ScaleX(100) )

   FF_ListView_InsertItem( hLV, 0, 0, "" ) 
   FF_ListView_InsertItem( hLV, 0, 1, "Left") 
   FF_ListView_InsertItem( hLV, 1, 0, "" ) 
   FF_ListView_InsertItem( hLV, 1, 1, "Top") 
   FF_ListView_InsertItem( hLV, 2, 0, "" ) 
   FF_ListView_InsertItem( hLV, 2, 1, "Right") 
   FF_ListView_InsertItem( hLV, 3, 0, "" ) 
   FF_ListView_InsertItem( hLV, 3, 1, "Bottom") 

   frmVDAnchors_SetListViewSelections(hLV, wszPropValue)

   Function = 0
End Function


