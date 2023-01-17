'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

#include once "frmCategories.bi"
#include once "clsConfig.bi"


' ========================================================================================
' Enable/disable controls based on current selected listbox line
' ========================================================================================
function frmCategories_UpdateControlStates() as long
   dim as hwnd hList = GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_LIST1)
   dim as long nCurSel = ListBox_GetCurSel(hList)
   dim as long idx = ListBox_GetItemData(hList, nCurSel)
   
   dim as hwnd hAdd = GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_CMDADD)
   dim as hwnd hEdit = GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_CMDEDIT)
   dim as hwnd hDelete = GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_CMDDELETE)
   
   ' Always allow Add
   EnableWindow( hAdd, true )
   
   if idx <> -1 then
      ' Allow Edit, Delete
      EnableWindow( hEdit, true )
      EnableWindow( hDelete, true )
   else
      ' do not allow edit, delete for default items or headers
      EnableWindow( hEdit, false )
      EnableWindow( hDelete, false )
   end if
   function = 0
end function


' ========================================================================================
' Move the description textbox into place
' ========================================================================================
function frmCategories_StartEdit( byval nCurSel as long ) as long            
   ' Move the description textbox into place
   if nCurSel = -1 then exit function
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMCATEGORIES)
   if pWindow = 0 THEN exit function
   
   dim as hwnd hList = GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_LIST1)

   dim as long idx = ListBox_GetItemData( hList, nCurSel )
   if idx <> -1 then
      dim as RECT rc: ListBox_GetItemRect( hList, nCurSel, @rc )
      rc.left = rc.left + pWindow->ScaleX(17)
      rc.top = rc.top + pWindow->ScaleY(2)
      MapWindowPoints( hList, HWND_FRMCATEGORIES, cast(POINT ptr, @rc), 2 )
      dim as HWND hCtrl = GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_TXTDESCRIPTION )
      dim as CWSTR wszText = AfxGetListBoxText( hList, nCurSel )
      AfxSetWindowText( hCtrl, wszText )
      SetWindowPos( hCtrl, HWND_TOP, rc.left, rc.top, rc.right-rc.left, rc.bottom-rc.top, SWP_SHOWWINDOW )
      SetFocus( hCtrl )
   end if
   function = 0
end function


' ========================================================================================
' Load all of the build descriptions into the listbox
' ========================================================================================
function frmCategories_LoadBuildListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMCATEGORIES_LIST1)

   dim as CWSTR wszText
   dim as long idx
   
   ListBox_ResetContent(hList1)
   
   ' Default system nodes
   wszText = "%%" & L(438, "Default Explorer Categories")
   idx = ListBox_AddString( hList1, wszText.sptr )
   ListBox_SetItemData( hList1, idx, -1 )
   for i as long = lbound(gConfig.CatTemp) to ubound(gConfig.CatTemp)
      if left(gConfig.CatTemp(i).idFileType, 1) <> "{" then
         idx = ListBox_AddString( hList1, gConfig.CatTemp(i).wszDescription.sptr )
         ListBox_SetItemData( hList1, idx, -1 )
      end if
   next

   wszText = "%%" & L(439, "User Defined Explorer Categories")
   idx = ListBox_AddString( hList1, wszText.sptr )
   ListBox_SetItemData( hList1, idx, -1 )
   for i as long = lbound(gConfig.CatTemp) to ubound(gConfig.CatTemp)
      ' User defined categories will have an GUID id. Starts with an {
      if left(gConfig.CatTemp(i).idFileType, 1) = "{" then
         idx = ListBox_AddString( hList1, gConfig.CatTemp(i).wszDescription.sptr )
         ListBox_SetItemData( hList1, idx, i )
      end if
   next

   dim as HWND hCtrl = GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_TXTDESCRIPTION )
   if IsWindowVisible( hCtrl ) then
      SetWindowPos( hCtrl, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE )
   end if
   
   frmCategories_UpdateControlStates()

   function = 0
end function


' ========================================================================================
' Swap two entries in the Listbox
' ========================================================================================
function frmCategories_SwapListBoxItems( _
            byval Item1 as long, _
            Byval Item2 as long _
            ) as Long
   dim as hwnd hList1 = GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_LIST1)

   dim as long idx1 = ListBox_GetItemData( hList1, Item1 )
   dim as long idx2 = ListBox_GetItemData( hList1, Item2 )
   if idx1 = -1 then exit function
   if idx2 = -1 then exit function

   ' We can not swap with an item that is a header (starts with %%)
   if left(gConfig.CatTemp(idx1).wszDescription, 2) = "%%" then exit function
   if left(gConfig.CatTemp(idx2).wszDescription, 2) = "%%" then exit function

   ' Swap the array values
   swap gConfig.CatTemp(idx1), gConfig.CatTemp(idx2)
   
   ListBox_ReplaceString( hList1, Item1, gConfig.CatTemp(idx1).wszDescription, idx1 )
   ListBox_ReplaceString( hList1, Item2, gConfig.CatTemp(idx2).wszDescription, idx2 )
 
   function = 0
end function


' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmCategories
' ========================================================================================
Function frmCategories_OnMeasureItem( _
            ByVal HWnd As HWnd, _
            ByVal lpmis As MEASUREITEMSTRUCT Ptr _
            ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY( FRMCATEGORIES_LISTBOX_LINEHEIGHT )
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmCategories
' ========================================================================================
Function frmCategories_OnDrawItem( _
            ByVal HWnd As HWnd, _
            ByVal lpdis As Const DRAWITEMSTRUCT Ptr _
            ) As Long

   Dim memDC as HDC                 ' Double buffering
   Dim hbit  As HBITMAP             ' Double buffering
   
   Dim As HBRUSH hBrush 
   Dim As RECT rc
   dim as long nWidth, nHeight
   dim as CWSTR wszText
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   If lpdis->itemID = -1 Then Exit Function
   
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT

         SaveDC( lpdis->hDC )

         wszText = AfxGetListBoxText( lpdis->hwndItem, lpdis->itemID)
         
         nWidth  = lpdis->rcItem.right - lpdis->rcItem.left
         nHeight = lpdis->rcItem.bottom - lpdis->rcItem.top 
         
         memDC = CreateCompatibleDC( lpdis->hDC )
         hbit  = CreateCompatibleBitmap( lpdis->hDC, nWidth, nHeight )
                                         
         If hbit Then hbit = SelectObject( memDC, hbit )

         ' Create our rect that works with the entire line
         SetRect( @rc, 0, 0, nWidth, nHeight )
         FillRect( memDC, @rc, GetSysColorBrush(COLOR_WINDOW) )  
            
         ' DETERMINE TEXT BACKGROUND
         If (lpdis->itemState And ODS_SELECTED) and _
            IsWindowVisible( GetDlgItem(HWnd, IDC_FRMCATEGORIES_TXTDESCRIPTION)) = false Then     
            SetBkColor( memDC, GetSysColor(COLOR_HIGHLIGHT) )   
            SetTextColor( memDC, GetSysColor(COLOR_HIGHLIGHTTEXT) )
            hBrush = GetSysColorBrush( COLOR_HIGHLIGHT ) 
         else
            SetBkColor( memDC, GetSysColor(COLOR_WINDOW) )   
            SetTextColor( memDC, GetSysColor(COLOR_WINDOWTEXT) )
            hBrush = GetSysColorBrush( COLOR_WINDOW ) 
         end if
                 
         ' Output the text
         SelectObject( memDC, AfxGetWindowFont(lpdis->hwndItem) )
         SelectObject( memDC, hBrush )      
         FillRect( memDC, @rc, hBrush )  

         dim as RECT rcText = rc
         
         if left(wszText, 2) = "%%" then
            SelectObject( memDC, ghStatusBar.hFontStatusBarBold )
            wszText = mid(wszText, 3)
            rcText.Left = rcText.Left + pWindow->ScaleX(4)
            DrawText( memDC, wszText, -1, Cast(lpRect, @rcText), _
                      DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )
         else   
            if IsWindowVisible( GetDlgItem(HWnd, IDC_FRMCATEGORIES_TXTDESCRIPTION)) Then     
               if lpdis->itemID = ListBox_GetCurSel( lpdis->hwndItem ) then
                  SelectObject( memDC, ghMenuBar.hFontSymbolSmall )
                  DrawText( memDC, wszChevronRight, -1, Cast(lpRect, @rcText), _
                            DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )
               end if
            end if
            SelectObject( memDC, ghStatusBar.hFontStatusBar )
            rcText.Left = rcText.Left + pWindow->ScaleX(20)
            DrawText( memDC, wszText, -1, Cast(lpRect, @rcText), _
                      DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )
         end if

         ' Draw the border edges
         SetBkMode( memDC, TRANSPARENT )   
         DrawEdge( memDC, @rc, EDGE_SUNKEN, BF_BOTTOMRIGHT )

         ' Copy the entire memory bitmap over to the visual display
         BitBlt( lpdis->hDC, lpdis->rcItem.left, lpdis->rcItem.top, nWidth, nHeight, memDC, 0, 0, SRCCOPY )

         ' Cleanup
         If hbit  Then DeleteObject( SelectObject(memDC, hbit) )
         If memDC Then DeleteDC( memDC )
   
         RestoreDC( lpdis->hDC, -1 )

         Function = True : Exit Function
   
   End Select

   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
Function frmCategories_ListBox_SubclassProc ( _
            ByVal HWnd   As HWnd, _                 ' // Control window handle
            ByVal uMsg   As UINT, _                 ' // Type of message
            ByVal wParam As WPARAM, _               ' // First message parameter
            ByVal lParam As LPARAM, _               ' // Second message parameter
            ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
            ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
            ) As LRESULT
   
   Select Case uMsg

      case WM_ERASEBKGND
         ' Only erase the bottom portion of the listbox that extends from the last item
         ' to the bottom edge of the listbox. All other lines are already drawn. This helps
         ' reduce screen flicker.
         dim as RECT rc = GetListBoxEmptyClientArea( HWND, FRMCATEGORIES_LISTBOX_LINEHEIGHT )
         if rc.top < rc.bottom then
            dim as HDC hDC = cast(HDC, wParam)
            FillRect( hDC, @rc, GetSysColorBrush(COLOR_WINDOW) )
         end if
         return true
         
      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_SPACE, VK_RETURN
               dim as long nCurSel = ListBox_GetCurSel(hwnd)
               frmCategories_StartEdit( nCurSel )
               Exit Function
            case VK_ESCAPE
               PostMessage( GetParent(HWnd), WM_CLOSE, 0, 0 )
               Exit Function
         End Select

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( HWnd, @frmCategories_ListBox_SubclassProc, uIdSubclass )

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc( HWnd, uMsg, wParam, lParam )

End Function


' ========================================================================================
' Processes messages for the subclassed TextBox window.
' ========================================================================================
Function frmCategories_Textbox_SubclassProc ( _
            ByVal HWnd   As HWnd, _                 ' // Control window handle
            ByVal uMsg   As UINT, _                 ' // Type of message
            ByVal wParam As WPARAM, _               ' // First message parameter
            ByVal lParam As LPARAM, _               ' // Second message parameter
            ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
            ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
            ) As LRESULT
   
   Select Case uMsg

      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN, VK_ESCAPE
               ShowWindow( HWnd, SW_HIDE )
               SetFocus( GetDlgItem( HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_LIST1) )
               Exit Function
         End Select

      case WM_CHAR
         ' prevent beep on Enter/Esc
         if wParam = 13 then return 0   
         if wParam = 27 then return 0   
               
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass( HWnd, @frmCategories_Textbox_SubclassProc, uIdSubclass )

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc( HWnd, uMsg, wParam, lParam )

End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmCategories
' ========================================================================================
Function frmCategories_OnCreate( _
            ByVal HWnd As HWnd, _
            ByVal lpCreateStructPtr As LPCREATESTRUCT _
            ) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmCategories
' ========================================================================================
Function frmCategories_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMCATEGORIES_LIST1)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   
   Select Case id
      case IDC_FRMCATEGORIES_LIST1
         if codeNotify = LBN_SELCHANGE then
            frmCategories_UpdateControlStates()
         elseif codeNotify = LBN_DBLCLK then
            frmCategories_StartEdit( nCurSel )
         end if
      
      case IDC_FRMCATEGORIES_TXTDESCRIPTION
         if codeNotify = EN_CHANGE then
            ' Update the temp array and the Listbox
            if nCurSel > -1 then
               dim as long idx = ListBox_GetItemData( hList1, nCurSel )
               if idx <> -1 then
                  gConfig.CatTemp(idx).wszDescription = AfxGetWindowText(hwndCtl)
                  ListBox_ReplaceString( hList1, nCurSel, gConfig.CatTemp(idx).wszDescription, idx )
               end if
            end if
         elseif codeNotify = EN_KILLFOCUS then
            ShowWindow( hwndCtl, SW_HIDE )         
         end if
         
      case IDC_FRMCATEGORIES_CMDUP
         if codeNotify = BN_CLICKED then
            if nCurSel > 0 then
               frmCategories_SwapListboxItems( nCurSel, nCurSel - 1 )
            end if
         end if

      case IDC_FRMCATEGORIES_CMDDOWN
         if codeNotify = BN_CLICKED then
            if nCurSel < ListBox_GetCount( hList1 ) - 1 then
               frmCategories_SwapListboxItems( nCurSel, nCurSel + 1 )
            end if
         end if

      case IDC_FRMCATEGORIES_CMDADD
         if codeNotify = BN_CLICKED THEN
            dim as long ub = ubound(gConfig.CatTemp) + 1
            redim preserve gConfig.CatTemp(ub)
            gConfig.CatTemp(ub).idFileType = AfxGuidText(AfxGuid())
            gConfig.CatTemp(ub).wszDescription = ""
            ' reload the listbox
            frmCategories_LoadBuildListBox(HWND)
            nCurSel = ListBox_GetCount(hList1) - 1
            ListBox_SetCurSel( hList1, nCurSel )
            frmCategories_StartEdit( nCurSel )
         end if

      case IDC_FRMCATEGORIES_CMDEDIT
         if codeNotify = BN_CLICKED THEN
            ListBox_SetCurSel( hList1, nCurSel )
            frmCategories_StartEdit( nCurSel )
         end if

      case IDC_FRMCATEGORIES_CMDDELETE
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1  THEN
               if MessageBox( HWND, L(437, "Are you sure you want to delete this Explorer Category?"), L(276,"Confirm"), _
                                 MB_YESNOCANCEL Or MB_ICONINFORMATION Or MB_DEFBUTTON1 ) = IDYES then
                  if ubound(gConfig.CatTemp) = 0 THEN
                     erase gConfig.CatTemp
                     nCurSel = -1
                  else
                     ' remove the item from the internal array
                     for i as long = nCurSel to ubound(gConfig.CatTemp) - 1
                        dim as long idx = ListBox_GetItemData( hList1, i )
                        if idx <> 1 then
                           gConfig.CatTemp(idx) = gConfig.CatTemp(idx+1)
                        end if   
                     NEXT
                     redim preserve gConfig.CatTemp(ubound(gConfig.CatTemp)-1)
                  END IF
                  ' reload the listbox
                  frmCategories_LoadBuildListBox(HWND)
                  nCurSel = Min(nCurSel, ubound(gConfig.CatTemp))
                  ListBox_SetCurSel( hList1, nCurSel )
                  SetFocus( hList1 )
               end if
            end if
         end if

      Case IDC_FRMCATEGORIES_CMDOK
         If codeNotify = BN_CLICKED Then
            ' Copy the temporary items to the main array  
            redim gConfig.Cat(ubound(gConfig.CatTemp))
            for i as long = lbound(gConfig.CatTemp) to ubound(gConfig.CatTemp)
               gConfig.Cat(i) = gConfig.CatTemp(i)   
            NEXT
            erase gConfig.CatTemp
            gConfig.SaveConfigFile
            LoadExplorerFiles()
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         end if
            
      Case IDC_FRMCATEGORIES_CMDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmCategories
' ========================================================================================
Function frmCategories_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmCategories
' ========================================================================================
Function frmCategories_OnDestroy( byval HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmCategories Window procedure
' ========================================================================================
Function frmCategories_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,      frmCategories_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmCategories_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmCategories_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmCategories_OnCommand)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmCategories_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmCategories_OnDrawItem)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc( HWnd, uMsg, wParam, lParam )

End Function


' ========================================================================================
' frmCategories_Show
' ========================================================================================
Function frmCategories_Show( ByVal hWndParent As HWnd ) As LRESULT

   dim as HWND hCtrl, hList
   dim wszImage as wstring * 100
   dim hBitmap as HANDLE
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMCATEGORIES = _
   pWindow->Create(hWndParent, L(436,"Explorer Categories"), _
        @frmCategories_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(510, 436)
   pWindow->Center(pWindow->hWindow, hWndParent)

   hList = _
   pWindow->AddControl("LISTBOX", , IDC_FRMCATEGORIES_LIST1, "", 10, 10, 490, 362, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or _
        LBS_NOINTEGRALHEIGHT OR LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmCategories_ListBox_SubclassProc), IDC_FRMCATEGORIES_LIST1, Cast(DWORD_PTR, @pWindow))

   hCtrl = _
   pWindow->AddControl("TEXTBOX", , IDC_FRMCATEGORIES_TXTDESCRIPTION, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmCategories_Textbox_SubclassProc), IDC_FRMCATEGORIES_TXTDESCRIPTION, Cast(DWORD_PTR, @pWindow))
        AfxSetWindowFont( hCtrl, ghStatusBar.hFontStatusBar, false )

   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDADD, L(380, "Add"), 8, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDEDIT, L(14, "Edit"), 87, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDDELETE, L(282, "Delete"), 166, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   hCtrl = _
   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDUP, wszTriangleUp, 245, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghMenuBar.hFontSymbolLargeBold, false )     

   hCtrl = _
   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDDOWN, wszTriangleDown, 278, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghMenuBar.hFontSymbolLargeBold, false )     

   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDOK, L(0,"OK"), 346, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDCANCEL, L(1,"Cancel"), 425, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   
   ' Copy all of the Categories to the temp array because we will work with 
   ' temporary copies until the user hits OK.
   redim gConfig.CatTemp(ubound(gConfig.Cat))
   for i as long = lbound(gConfig.Cat) to ubound(gConfig.Cat)
      gConfig.CatTemp(i) = gConfig.Cat(i)   
   NEXT
   frmCategories_LoadBuildListBox( HWND_FRMCATEGORIES )
   
   dim as long nCurSel = ListBox_GetCount(hList) - 1
   ListBox_SetCurSel( hList, nCurSel )
   frmCategories_UpdateControlStates()
   SetFocus( hList )
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents( SW_SHOW )
   
   ' Delete the CWindow class manually allocated memory 
   Delete pWindow

End Function

