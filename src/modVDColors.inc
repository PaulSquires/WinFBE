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

#include once "modVDColors.bi"
#include once "clsDocument.bi"


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmVDColors
' ========================================================================================
private Function frmVDColors_OnCommand( ByVal HWnd As HWnd, _
                                        ByVal id As Long, _
                                        ByVal hwndCtl As HWnd, _
                                        ByVal codeNotify As UINT _
                                        ) As LRESULT
                                        
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr()
   dim pProp as clsProperty ptr = GetActivePropertyPtr()
   dim as long nCurSel, idx
   
   select case id
      CASE IDC_FRMVDCOLORS_LSTCUSTOM
         if codeNotify = LBN_SELCHANGE then
            if pProp then 
               nCurSel = ListBox_GetCurSel(hwndCtl)
               if nCurSel > -1 then 
                  dim lCustomColor(15) AS LONG
                  dim as COLORREF clrResult
                  dim pProp as clsProperty ptr = GetActivePropertyPtr()
                  if pProp then
                     ' Set the default color to be whatever the current color is for the
                     ' property that is currently selected.
                     lCustomColor(0) = Val(AfxStrParse(pProp->wszPropValue, 2, "|"))
                  end if
                  clrResult = AfxChooseColorDialog( hwnd, lCustomColor(0), @lCustomColor(0))
                  if clrResult <> -1 then 
                     pProp->wszPropValuePrev = pProp->wszPropValue 
                     pProp->wszPropValue = "CUSTOM|" & clrResult
                     AfxRedrawWindow(GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES))
                     pDoc->UserModified = true
                     pDoc->bRegenerateCode = true
                     frmMain_SetStatusbar
                     PostMessage(HWND_FRMVDCOLORS, WM_ACTIVATE, WA_INACTIVE, 0)  ' to apply properties
                  end if   
               end if
            END IF
         END IF
      case IDC_FRMVDCOLORS_LSTCOLORS, IDC_FRMVDCOLORS_LSTSYSTEM 
         if codeNotify = LBN_SELCHANGE then
            if pProp then 
               nCurSel = ListBox_GetCurSel(hwndCtl)
               if nCurSel > -1 then 
                  idx = ListBox_GetItemData(hwndCtl, nCurSel)
                  pProp->wszPropValuePrev = pProp->wszPropValue 
                  pProp->wszPropValue = _
                     iif(id = IDC_FRMVDCOLORS_LSTSYSTEM, "SYSTEM|", "COLORS|")
                  pProp->wszPropValue = pProp->wszPropValue & gColors(idx).wszColorName
                  pDoc->UserModified = true
                  pDoc->bRegenerateCode = true
                  frmMain_SetStatusbar
                  ShowWindow(HWND_FRMVDCOLORS, SW_HIDE)
                  AfxRedrawWindow(GetDlgItem(HWND_FRMVDTOOLBOX, IDC_FRMVDTOOLBOX_LSTPROPERTIES))
               end if   
            end if
         end if
   end select
          
   function = 0
end function

      
' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmVDColors
' ========================================================================================
private Function frmVDColors_OnNotify( ByVal HWnd As HWnd, _
                                       ByVal id As Long, _
                                       ByVal pNMHDR As NMHDR Ptr _
                                       ) As LRESULT

   SELECT CASE id
      CASE IDC_FRMVDCOLORS_TABCONTROL
         dim as long iPage = TabCtrl_GetCurSel(pNMHDR->hwndFrom)
         SELECT CASE pNMHDR->code
            CASE TCN_SELCHANGE
               ' Show the selected page controls
               if iPage = 0 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDCOLORS_LSTCUSTOM), SW_SHOW)
               if iPage = 1 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDCOLORS_LSTCOLORS), SW_SHOW)
               if iPage = 2 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDCOLORS_LSTSYSTEM), SW_SHOW)
                                
            CASE TCN_SELCHANGING
               if iPage = 0 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDCOLORS_LSTCUSTOM), SW_HIDE)
               if iPage = 1 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDCOLORS_LSTCOLORS), SW_HIDE)
               if iPage = 2 then ShowWindow( GetDlgItem(HWND, IDC_FRMVDCOLORS_LSTSYSTEM), SW_HIDE)
         END SELECT

   END SELECT

   function = 0
end function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
private Function frmVDColors_PositionWindows() As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMVDCOLORS)
   if pWindow = 0 THEN exit function
      
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWND_FRMVDCOLORS, @rc)
   
   Dim As HWnd hTabCtl = GetDlgItem(HWND_FRMVDCOLORS, IDC_FRMVDCOLORS_TABCONTROL)
   Dim As HWnd hList1  = GetDlgItem(HWND_FRMVDCOLORS, IDC_FRMVDCOLORS_LSTCUSTOM)
   Dim As HWnd hList2  = GetDlgItem(HWND_FRMVDCOLORS, IDC_FRMVDCOLORS_LSTCOLORS)
   Dim As HWnd hList3  = GetDlgItem(HWND_FRMVDCOLORS, IDC_FRMVDCOLORS_LSTSYSTEM)
   
   SetWindowPos( hTabCtl, 0, 0, 0, rc.Right-rc.Left, pWindow->ScaleY(24), SWP_SHOWWINDOW Or SWP_NOZORDER )
   SetWindowPos( hList1, 0, 0, pWindow->ScaleY(24), rc.Right-rc.Left, rc.Bottom-rc.top-pWindow->ScaleY(24), SWP_NOZORDER )
   SetWindowPos( hList2, 0, 0, pWindow->ScaleY(24), rc.Right-rc.Left, rc.Bottom-rc.top-pWindow->ScaleY(24), SWP_NOZORDER )
   SetWindowPos( hList3, 0, 0, pWindow->ScaleY(24), rc.Right-rc.Left, rc.Bottom-rc.top-pWindow->ScaleY(24), SWP_NOZORDER )
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmVDColors
' ========================================================================================
private Function frmVDColors_OnSize( ByVal HWnd As HWnd, _
                                     ByVal state As UINT, _
                                     ByVal cx As Long, _
                                     ByVal cy As Long _
                                     ) As LRESULT
   If state <> SIZE_MINIMIZED Then 
      frmVDColors_PositionWindows
   End If   
   Function = 0
End Function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmVDColors
' ========================================================================================
private Function frmVDColors_OnMeasureItem( ByVal HWnd As HWnd, _
                                            ByVal lpmis As MEASUREITEMSTRUCT Ptr _
                                            ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(FRMVDCOLORS_LISTBOX_LINEHEIGHT)
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmVDColors
' ========================================================================================
private Function frmVDColors_OnDrawItem( ByVal HWnd As HWnd, _
                                         ByVal lpdis As Const DRAWITEMSTRUCT Ptr _
                                         ) As Long

   Dim As HBRUSH hBrush, hBrushOld
   Dim As RECT rc1, rc2
   Dim wzText As WString * MAX_PATH
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   If lpdis->itemID = -1 Then Exit Function    ' no selected row
   
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT
         
         SaveDC(lpdis->hDC)
         
         ' CLEAR BACKGROUND
         If (lpdis->itemState And ODS_SELECTED) Then     
            SetBkColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT))   
            SetTextColor(lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
            hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
         else
            SetBkColor(lpdis->hDC, GetSysColor(COLOR_WINDOW))   
            SetTextColor(lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT))
            hBrush = GetSysColorBrush(COLOR_WINDOW) 
         end if
         SelectObject(lpdis->hDC, hBrush)      
         FillRect(lpdis->hDC, @lpdis->rcItem, hBrush)  

         If cast(long, lpdis->itemData) = Cast(long, -1) Then
            wzText = "Custom Color"
            dim pProp as clsProperty ptr = GetActivePropertyPtr()
            if pProp then
               dim as CWSTR wszColorValue = AfxStrParse(pProp->wszPropValue, 2, "|")
               hBrush = CreateSolidBrush(val(wszColorValue))
            else   
               hBrush = CreateSolidBrush(BGR(255,255,255))
            end if
         else
            wzText = gColors(lpdis->itemData).wszColorName
            hBrush = CreateSolidBrush(gColors(lpdis->itemData).ColorValue)
         end if
         
         rc1 = lpdis->rcItem
         rc1.left = pWindow->ScaleX(2): rc1.right = rc1.left + pWindow->ScaleX(16)
         rc1.top = rc1.top + pWindow->ScaleY(2): rc1.bottom = rc1.bottom - pWindow->ScaleY(2)
         
         rc2 = lpdis->rcItem
         rc2.left = rc2.left + pWindow->ScaleX(24)
         
         ' DRAW COLOR RECT
         hBrushOld = SelectObject( Cast(HDC, lpdis->hDC), hBrush)
         RoundRect( lpdis->hDC, rc1.Left, rc1.Top, rc1.Right, rc1.Bottom, pWindow->ScaleX(3), pWindow->ScaleY(3))
         SelectObject( Cast(HDC, lpdis->hDC), hBrushOld)
         
         ' DRAW TEXT
         DrawText( lpdis->hDC, wzText, -1, Cast(lpRect, @rc2), _
                   DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )

         RestoreDC(lpdis->hDC, -1)
         DeleteObject(hBrush)

         Function = True : Exit Function
   
   End Select

   Function = 0
End Function

   
' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmVDColors
' ========================================================================================
private Function frmVDColors_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ' Never close the window; simply hide it.
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmVDColors
' ========================================================================================
private Function frmVDColors_OnDestroy( byval HWnd As HWnd ) As LRESULT
   HWND_FRMVDCOLORS = 0
   Function = 0
End Function


' ========================================================================================
' frmVDColors Window procedure
' ========================================================================================
private Function frmVDColors_WndProc( ByVal HWnd   As HWnd, _
                                      ByVal uMsg   As UINT, _
                                      ByVal wParam As WPARAM, _
                                      ByVal lParam As LPARAM _
                                      ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,     frmVDcolors_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmVDColors_OnNotify)
      HANDLE_MSG (HWnd, WM_SIZE,        frmVDColors_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmVDColors_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmVDColors_OnDestroy)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmVDColors_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmVDColors_OnDrawItem)

   case WM_ACTIVATE
      if wParam = WA_INACTIVE then
         ShowWindow(HWND_FRMVDCOLORS, SW_HIDE)
         ' This function allows for two use cases: (1) is to be able to choose colors
         ' in the PropertyList for various controls, and (2) is to be able to select
         ' colors for StatusBar Panels. If the StatusBar Editor is active then we
         ' must be looking for the pPropColor related to that panel, otherwise it
         ' must be the PropertyList that is being used.
         if IsWindowVisible(HWND_FRMSTATUSBAREDITOR) then
            dim as hwnd hList1 = GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
            dim as long nCurSel = ListBox_GetCurSel(hList1)
            if nCurSel = -1 then exit function

            select case gPanelItems(nCurSel).idColorCombo
               case IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLOR
                  gPanelItems(nCurSel).wszBackColor = gPanelItems(nCurSel).pPropColor.wszPropValue
               case IDC_FRMSTATUSBAREDITOR_COMBOBACKCOLORHOT
                  gPanelItems(nCurSel).wszBackColorHot = gPanelItems(nCurSel).pPropColor.wszPropValue 
               case IDC_FRMSTATUSBAREDITOR_COMBOFORECOLOR
                  gPanelItems(nCurSel).wszForeColor = gPanelItems(nCurSel).pPropColor.wszPropValue 
               case IDC_FRMSTATUSBAREDITOR_COMBOFORECOLORHOT
                  gPanelItems(nCurSel).wszForeColorHot = gPanelItems(nCurSel).pPropColor.wszPropValue 
            end select
            AfxRedrawWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, gPanelItems(nCurSel).idColorCombo ))

         else
            ' Regular PropertyList case   
            dim pCtrl as clsControl ptr
            dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
            if pDoc then pCtrl = pDoc->Controls.GetActiveControl
            if pCtrl then ApplyControlProperties(pDoc, pCtrl)
         end if
      end if   

   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmVDColors_SetColorListBoxSelection
' ========================================================================================
private Function frmVDColors_SetColorListBoxSelection( byref wszPropValue as wstring ) as Long

   ' wszPropValue is in two parts separated by "|"
   ' eg.  SYSTEM|ButtonFace
   '      COLORS|Red
   '      CUSTOM|<rgbvalue>
   
   dim as HWND hTabCtl = GetDlgItem(HWND_FRMVDCOLORS, IDC_FRMVDCOLORS_TABCONTROL)
   dim as long nCurSel 
   dim as hwnd hList
   
   dim as CWSTR wszList, wszValue
   wszList = AfxStrParse(wszPropValue, 1, "|")
   wszValue = AfxStrParse(wszPropValue, 2, "|")
   
   dim as hwnd hList1 = GetDlgItem(HWND_FRMVDCOLORS, IDC_FRMVDCOLORS_LSTCUSTOM)
   dim as hwnd hList2 = GetDlgItem(HWND_FRMVDCOLORS, IDC_FRMVDCOLORS_LSTCOLORS)
   dim as hwnd hList3 = GetDlgItem(HWND_FRMVDCOLORS, IDC_FRMVDCOLORS_LSTSYSTEM)
   
   ' Ensure that any previously selected listbox row is now deselected
   ListBox_SetCurSel(hList1, -1)
   ListBox_SetCurSel(hList2, -1)
   ListBox_SetCurSel(hList3, -1)
   
   ' Hide all existing pages
   ShowWindow( hList1, SW_HIDE )
   ShowWindow( hList2, SW_HIDE )
   ShowWindow( hList3, SW_HIDE )
   
   select case wszList
      CASE "CUSTOM"
         TabCtrl_SetCurSel(hTabCtl, 0)
         wszValue = "Custom Color"
         hList = hList1
      case "COLORS"
         TabCtrl_SetCurSel(hTabCtl, 1)
         hList = hList2
      case "SYSTEM"      
         TabCtrl_SetCurSel(hTabCtl, 2)
         hList = hList3
   END SELECT

   nCurSel = ListBox_FindStringExact(hList, -1, wszValue.sptr)
   ShowWindow(hList, SW_SHOW)
   ListBox_SetCurSel(hList, nCurSel)
   SetFocus(hList)

   function = 0
end function


' ========================================================================================
' frmVDColors_Show
' ========================================================================================
public Function frmVDColors_Show( ByVal hWndParent As HWnd, _
                                  byref wszPropValue as wstring _
                                  ) as LRESULT

   ' If the colors popup already exists then no need to recreate it.
   If IsWindow(HWND_FRMVDCOLORS) Then 
      frmVDColors_SetColorListBoxSelection(wszPropValue)
      exit function
   END IF

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMVDCOLORS = _
   pWindow->Create( hWndParent, "", @frmVDColors_WndProc, 0, 0, 200, 240, _
        WS_POPUP or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN or WS_BORDER, 0)
   
   Dim As HWnd hTabCtl = _ 
        pWindow->AddControl("TAB", , IDC_FRMVDCOLORS_TABCONTROL, "", 0, 0, 0, 0)
   TabCtrl_AddTab(hTabCtl, 0, "Custom")
   TabCtrl_AddTab(hTabCtl, 0, "Colors")
   TabCtrl_AddTab(hTabCtl, 0, "System")
   
   Dim As HWnd hListCustom = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMVDCOLORS_LSTCUSTOM, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT Or WS_VSCROLL or  _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

   Dim As HWnd hListColors = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMVDCOLORS_LSTCOLORS, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT Or WS_VSCROLL or  _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

   Dim As HWnd hListSystem = _ 
        pWindow->AddControl("LISTBOX", , IDC_FRMVDCOLORS_LSTSYSTEM, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_TABSTOP Or LBS_NOINTEGRALHEIGHT Or WS_VSCROLL or  _
        LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS or LBS_NOTIFY, WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

   dim as long idx
   
   ' CUSTOM COLOR
   idx = ListBox_AddString(hListCustom, @WSTR("Custom Color"))
   ListBox_SetItemData(hListCustom, idx, -1)   ' special value to signify custom color
   
   ' Add the Colors to the listboxes
   for i as long = lbound(gColors) to ubound(gColors)
      select case gColors(i).ColorType
         CASE COLOR_COLORS
            idx = ListBox_AddString(hListColors, gColors(i).wszColorName.sptr)
            ListBox_SetItemData(hListColors, idx, i)   ' store the array index
         CASE COLOR_SYSTEM
            idx = ListBox_AddString(hListSystem, gColors(i).wszColorName.sptr)
            ListBox_SetItemData(hListSystem, idx, i)   ' store the array index
      END SELECT
   NEXT

   frmVDColors_PositionWindows
   frmVDColors_SetColorListBoxSelection(wszPropValue)

   Function = 0
End Function


