'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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


' ========================================================================================
' Load all of the build descriptions into the listbox
' ========================================================================================
function LoadBuildListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMBUILDCONFIG_LIST1)

   ListBox_ResetContent(hList1)
   for i as long = lbound(gConfig.BuildsTemp) to ubound(gConfig.BuildsTemp)
      ListBox_AddString(hList1, gConfig.BuildsTemp(i).wszDescription)
   NEXT

   function = 0
end function

' ========================================================================================
' Return the string GUID of the currently selected build configuration
' ========================================================================================
function GetSelectedBuildGUID() as String

   dim as long nCurSel = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
   if nCurSel > -1 THEN
      function = gConfig.Builds(nCurSel).Id
   END IF

end function

' ========================================================================================
' Return the string GUID of the default (if any) build configuration
' ========================================================================================
function GetDefaultBuildGUID() as String

   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      if gConfig.Builds(i).IsDefault then 
         return gConfig.Builds(i).Id
      end if   
   NEXT
   function = ""
   
end function

' ========================================================================================
' Load all of the build descriptions into the toolbar combobox and set to selected
' ========================================================================================
function LoadBuildComboBox() as Long

   Dim pDocMain As clsDocument Ptr 
   dim as long nCurSel = 0
   dim as string sBuildId

   ComboBox_ResetContent(HWND_FRMMAIN_COMBOBUILDS)
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      ComboBox_AddString(HWND_FRMMAIN_COMBOBUILDS, gConfig.Builds(i).wszDescription)
   NEXT

   ' 1) If project is active then set using the selected build for the project.
   If gApp.IsProjectActive Then
      sBuildId = gApp.ProjectBuild
   Else
   ' 2) If no project is active, set using the selected build for the active document.
      pDocMain = gTTabCtl.GetActiveDocumentPtr()
      If pDocMain Then sBuildId = pDocMain->DocumentBuild
   End If

   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      if gConfig.Builds(i).Id = sBuildId then
         nCurSel = i: exit for
      end if   
      if gConfig.Builds(i).IsDefault then nCurSel = i
   NEXT

   ComboBox_SetCurSel(HWND_FRMMAIN_COMBOBUILDS, nCurSel)
   
   function = 0
end function


' ========================================================================================
' Swap two entries in the build Listbox
' ========================================================================================
function SwapBuildListBoxItems( byval Item1 as long, Byval Item2 as long) as Long
   dim as hwnd hList1 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_LIST1)
   
   ' Swap the array values
   swap gConfig.BuildsTemp(Item1), gConfig.BuildsTemp(Item2)
   
   ListBox_ReplaceString(hList1, Item1, gConfig.BuildsTemp(Item1).wszDescription)
   ListBox_ReplaceString(hList1, Item2, gConfig.BuildsTemp(Item2).wszDescription)

   function = 0
end function


' ========================================================================================
' Set the build description and compiler options depending on what listbox entry is selected
' ========================================================================================
function SetCompilerConfigTextboxes() as long
   dim as hwnd hList1 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_LIST1)
   dim as hwnd hText1 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_TXTDESCRIPTION)
   dim as hwnd hText2 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_TXTOPTIONS)
   dim as hwnd hOpt32 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT32)
   dim as hwnd hOpt64 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT64)
   
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   if nCurSel < 0 THEN 
      AfxSetWindowText( hText1, "")
      AfxSetWindowText( hText2, "")
      CheckRadioButton(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT32, IDC_FRMBUILDCONFIG_OPT64, IDC_FRMBUILDCONFIG_OPT32)
      EnableWindow(hText1, false)
      EnableWindow(hText2, false)
      EnableWindow(hOpt32, false)
      EnableWindow(hOpt64, false)
   else
      AfxSetWindowText( hText1, gConfig.BuildsTemp(nCurSel).wszDescription)
      AfxSetWindowText( hText2, gConfig.BuildsTemp(nCurSel).wszOptions)
      if gConfig.BuildsTemp(nCurSel).Is32bit THEN
         CheckRadioButton(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT32, IDC_FRMBUILDCONFIG_OPT64, IDC_FRMBUILDCONFIG_OPT32)
      else   
         CheckRadioButton(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT32, IDC_FRMBUILDCONFIG_OPT64, IDC_FRMBUILDCONFIG_OPT64)
      END IF
      EnableWindow(hText1, true)
      EnableWindow(hText2, true)
      EnableWindow(hOpt32, true)
      EnableWindow(hOpt64, true)
   end if
   
   function = 0
end function
               

' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmBuildConfig
' ========================================================================================
Function frmBuildConfig_OnMeasureItem( ByVal HWnd As HWnd, ByVal lpmis As MEASUREITEMSTRUCT Ptr ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(LISTBOX_LINEHEIGHT)
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmBuildConfig
' ========================================================================================
Function frmBuildConfig_OnDrawItem( ByVal HWnd As HWnd, ByVal lpdis As Const DRAWITEMSTRUCT Ptr ) As Long

   Dim memDC as HDC                 ' Double buffering
   Dim hbit  As HBITMAP             ' Double buffering
   
   Dim As HBRUSH hBrush 
   Dim As RECT rc, rc2
   dim as long nWidth, nHeight
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   If lpdis->itemID = -1 Then Exit Function
   
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT

         if lpdis->CtlID = IDC_FRMBUILDCONFIG_LIST1 then
            SaveDC(lpdis->hDC)

            nWidth  = lpdis->rcItem.right-lpdis->rcItem.left
            nHeight = lpdis->rcItem.bottom-lpdis->rcItem.top 
            
            memDC = CreateCompatibleDC( lpdis->hDC )
            hbit  = CreateCompatibleBitmap( lpdis->hDC, nWidth, nHeight )
                                            
            If hbit Then hbit = SelectObject( memDC, hbit )

            ' Create our rect that works with the entire line
            SetRect(@rc, 0, 0, nWidth, nHeight)
            FillRect(memDC, @rc, GetSysColorBrush(COLOR_WINDOW))  
            
            ' Draw the tick/untick image
            ' The line height is LISTBOX_LINEHEIGHT so center the 16x16 icon vertically and horizontally
            DrawIconEx( memDC, _
                        rc.Left + pWindow->ScaleX(2), _
                        rc.Top + pWindow->ScaleY(2), _
                        iIf( gConfig.BuildsTemp(lpdis->itemID).IsDefault, ghIconTick, ghIconNoTick), _
                        pWindow->ScaleX(16), pWindow->ScaleY(16), 0, 0, DI_NORMAL) 
            
            ' DETERMINE TEXT BACKGROUND
            If (lpdis->itemState And ODS_SELECTED) Then     
               SetBkColor(memDC, GetSysColor(COLOR_HIGHLIGHT))   
               SetTextColor(memDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
               hBrush = GetSysColorBrush(COLOR_HIGHLIGHT) 
            else
               SetBkColor(memDC, GetSysColor(COLOR_WINDOW))   
               SetTextColor(memDC, GetSysColor(COLOR_WINDOWTEXT))
               hBrush = GetSysColorBrush(COLOR_WINDOW) 
            end if
                    
            ' Draw the build name
            rc2 = rc: rc2.Left = pWindow->ScaleX(20)
            SelectObject(memDC, AfxGetWindowFont(lpdis->hwndItem))
            SelectObject(memDC, hBrush)      
            FillRect(memDC, @rc2, hBrush)  
            DrawText( memDC, _
                      **gConfig.BuildsTemp(lpdis->itemID).wszDescription, _
                      -1, Cast(lpRect, @rc2), _
                      DT_LEFT Or DT_SINGLELINE Or DT_VCENTER )

            ' Draw the border edges
            SetBkMode(memDC, TRANSPARENT)   
            DrawEdge( memDC, @rc, EDGE_SUNKEN, BF_BOTTOMRIGHT)

            ' Copy the entire memory bitmap over to the visual display
            BitBlt lpdis->hDC, lpdis->rcItem.left, lpdis->rcItem.top, nWidth, nHeight, memDC, 0, 0, SRCCOPY 

            ' Cleanup
            If hbit  Then DeleteObject SelectObject(memDC, hbit)
            If memDC Then DeleteDC memDC
   
            RestoreDC(lpdis->hDC, -1)
         end if

         Function = True : Exit Function
   
   End Select

   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
Function frmBuildConfig_ListBox_SubclassProc ( _
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
         ' Determine the number of items in the listbox and multiply by the item height
         dim as RECT rc: GetClientRect(hwnd, @rc)
         rc.top = rc.top + (ListBox_GetCount(hwnd) * ListBox_GetItemHeight(hwnd, 0))
         if rc.top < rc.bottom then
            dim as HDC hDC = cast(HDC, wParam)
            FillRect(hDC, @rc, cast(HBRUSH, GetCurrentObject(hDC, OBJ_BRUSH)))
         end if
         return TRUE
         
      case WM_COMMAND
            
      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN
               ' Simulate the sending of a LBN_DBLCLK to the control.
               SendMessage( GetParent(HWnd), WM_COMMAND, MAKEWPARAM(IDC_FRMBUILDCONFIG_LIST1, LBN_DBLCLK), Cast(LPARAM,HWnd) )
               Exit Function
         End Select

      case WM_MOUSEMOVE
      case WM_LBUTTONDOWN
         
      case WM_LBUTTONUP
         if hwnd = GetDlgItem(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_LIST1) THEN
            ' Handle if the checkbox is clicked
            dim as long nCurSel = ListBox_GetCurSel(hwnd)
            if nCurSel = -1 then exit function
            Dim pWindow As CWindow Ptr = AfxCWindowPtr(GetParent(HWnd))
            if pWindow = 0 THEN exit function
            dim as RECT rc: SendMessage(hwnd, LB_GETITEMRECT, nCurSel, cast(LPARAM, @rc))
            ' The checkbox is the first 20x20 area (16x16 icon)
            rc.Left   = rc.Left + pWindow->ScaleX(2)
            rc.Top    = rc.Top  + pWindow->ScaleY(2)
            rc.Right  = rc.Left + pWindow->ScaleX(16)
            rc.Bottom = rc.Top  + pWindow->ScaleY(16)
            dim as POINT pt = (loword(lParam), Hiword(lParam))
            if PtInRect(@rc, pt) then
               gConfig.BuildsTemp(nCurSel).IsDefault = not gConfig.BuildsTemp(nCurSel).IsDefault
               ' Can only have 1 entry as the default so ensure all others are reset
               if gConfig.BuildsTemp(nCurSel).IsDefault THEN
                  for i as long = lbound(gConfig.BuildsTemp) to ubound(gConfig.BuildsTemp)
                     if i <> nCurSel THEN gConfig.BuildsTemp(i).IsDefault = 0
                  NEXT
               END IF
               InvalidateRect(hwnd, @rc, true): UpdateWindow(hwnd)
            END IF
         end if
         
         
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmBuildConfig_ListBox_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmBuildConfig
' ========================================================================================
Function frmBuildConfig_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmBuildConfig
' ========================================================================================
Function frmBuildConfig_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMBUILDCONFIG_LIST1)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   ' Convert a ListBox DoubleClick into pressing OK button action
   if (id = IDC_FRMBUILDCONFIG_LIST1) AndAlso (codeNotify = LBN_DBLCLK) THEN
      id = IDOK: codeNotify = BN_CLICKED
   END IF
   
   Select Case id
      case IDC_FRMBUILDCONFIG_LIST1
         if codeNotify = LBN_SELCHANGE THEN
            SetCompilerConfigTextboxes()
         END IF
          
      case IDC_FRMBUILDCONFIG_TXTDESCRIPTION
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array and the Listbox
            if nCurSel > -1 THEN
               gConfig.BuildsTemp(nCurSel).wszDescription = AfxGetWindowText(hwndCtl)
               ListBox_ReplaceString(hList1, nCurSel, gConfig.BuildsTemp(nCurSel).wszDescription)
            END IF
         end if
         
      CASE IDC_FRMBUILDCONFIG_TXTOPTIONS
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array 
            if nCurSel > -1 THEN
               gConfig.BuildsTemp(nCurSel).wszOptions = AfxGetWindowText(hwndCtl)
            END IF
         end if
         
      case IDC_FRMBUILDCONFIG_OPT32
         if codeNotify = BN_CLICKED THEN
            ' Update the temp array 
            if nCurSel > -1 THEN
               if Button_GetCheck(hwndCtl) THEN
                  gConfig.BuildsTemp(nCurSel).Is32bit = 1
                  gConfig.BuildsTemp(nCurSel).Is64bit = 0
               else   
                  gConfig.BuildsTemp(nCurSel).Is32bit = 0
                  gConfig.BuildsTemp(nCurSel).Is64bit = 1
               END IF
            END IF
         end if

      case IDC_FRMBUILDCONFIG_OPT64
         if codeNotify = BN_CLICKED THEN
            ' Update the temp array 
            if nCurSel > -1 THEN
               if Button_GetCheck(hwndCtl) THEN
                  gConfig.BuildsTemp(nCurSel).Is32bit = 0
                  gConfig.BuildsTemp(nCurSel).Is64bit = 1
               else   
                  gConfig.BuildsTemp(nCurSel).Is32bit = 1
                  gConfig.BuildsTemp(nCurSel).Is64bit = 0
               END IF
            END IF
         end if

      case IDC_FRMBUILDCONFIG_CMDUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               SwapBuildListboxItems(nCurSel, nCurSel-1)
            END IF
         end if

      case IDC_FRMBUILDCONFIG_CMDDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1)-1 THEN
               SwapBuildListboxItems(nCurSel, nCurSel+1)
            END IF
         end if

      case IDC_FRMBUILDCONFIG_CMDINSERT
         if codeNotify = BN_CLICKED THEN
            if ubound(gConfig.BuildsTemp) = -1 THEN
               redim gConfig.BuildsTemp(0)
               nCurSel = 0
            Else
               redim preserve gConfig.BuildsTemp(ubound(gConfig.BuildsTemp)+1)
               if nCurSel = -1 THEN nCurSel = 0
               ' insert the item above current entry in the internal array
               for i as long = ubound(gConfig.BuildsTemp) to nCurSel + 1 step -1
                  gConfig.BuildsTemp(i) = gConfig.BuildsTemp(i-1)
               NEXT
            END IF
            gConfig.BuildsTemp(nCurSel).Id = AfxGuidText(AfxGuid())
            gConfig.BuildsTemp(nCurSel).wszDescription = ""
            gConfig.BuildsTemp(nCurSel).wszOptions = ""
            gConfig.BuildsTemp(nCurSel).IsDefault = 0
            gConfig.BuildsTemp(nCurSel).Is32bit = 1
            gConfig.BuildsTemp(nCurSel).Is64bit = 0
            ' reload the listbox
            LoadBuildListBox(HWND)
            nCurSel = Min(nCurSel, ubound(gConfig.BuildsTemp))
            ListBox_SetCurSel(hList1, nCurSel)
            SetCompilerConfigTextboxes()
            SetFocus hList1
         end if

      case IDC_FRMBUILDCONFIG_CMDDELETE
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1  THEN
               if MessageBox( HWND, L(287, "Are you sure you want to delete this build configuration?"), L(276,"Confirm"), _
                                 MB_YESNOCANCEL Or MB_ICONINFORMATION Or MB_DEFBUTTON1 ) = IDYES then
                  if ubound(gConfig.BuildsTemp) = 0 THEN
                     erase gConfig.BuildsTemp
                     nCurSel = -1
                  else
                     ' remove the item from the internal array
                     for i as long = nCurSel to ubound(gConfig.BuildsTemp) - 1
                        gConfig.BuildsTemp(i) = gConfig.BuildsTemp(i+1)
                     NEXT
                     redim preserve gConfig.BuildsTemp(ubound(gConfig.BuildsTemp)-1)
                  END IF
                  ' reload the listbox
                  LoadBuildListBox(HWND)
                  nCurSel = Min(nCurSel, ubound(gConfig.BuildsTemp))
                  ListBox_SetCurSel(hList1, nCurSel)
                  SetCompilerConfigTextboxes()
                  SetFocus hList1
               end if
            END IF
         end if

      Case IDOK
         If codeNotify = BN_CLICKED Then
            ' Copy the temporary items to the main array  
            redim gConfig.Builds(ubound(gConfig.BuildsTemp))
            for i as long = lbound(gConfig.BuildsTemp) to ubound(gConfig.BuildsTemp)
               gConfig.Builds(i) = gConfig.BuildsTemp(i)   
            NEXT
            erase gConfig.BuildsTemp
            
            ' Set the active project or the active document to the selected build
            if nCurSel = -1 THEN exit function
            If gApp.IsProjectActive Then
               gApp.ProjectBuild = gConfig.Builds(nCurSel).Id
            Else
               Dim pDocMain As clsDocument Ptr 
               pDocMain = gTTabCtl.GetActiveDocumentPtr()
               If pDocMain Then pDocMain->DocumentBuild = gConfig.Builds(nCurSel).Id
            end if
            LoadBuildListBox(HWND)
            frmMain_UpdateLineCol(HWND_FRMMAIN)
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         end if
            
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmBuildConfig
' ========================================================================================
Function frmBuildConfig_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmBuildConfig
' ========================================================================================
Function frmBuildConfig_OnDestroy(HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmBuildConfig Window procedure
' ========================================================================================
Function frmBuildConfig_WndProc( ByVal HWnd   As HWnd, _
                                 ByVal uMsg   As UINT, _
                                 ByVal wParam As WPARAM, _
                                 ByVal lParam As LPARAM _
                                 ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,      frmBuildConfig_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmBuildConfig_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmBuildConfig_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmBuildConfig_OnCommand)
      HANDLE_MSG (HWnd, WM_MEASUREITEM, frmBuildConfig_OnMeasureItem)
      HANDLE_MSG (HWnd, WM_DRAWITEM,    frmBuildConfig_OnDrawItem)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmBuildConfig_Show
' ========================================================================================
Function frmBuildConfig_Show( ByVal hWndParent As HWnd, _
                              ByVal nCmdShow   As Long = 0 _
                              ) As Long

   DIM hBitmap AS HBITMAP 
   dim hCtrl as HWnd
   dim wszImage as wstring * 100

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create(hWndParent, L(277,"Build Configurations"), _
        @frmBuildConfig_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(677, 436)
   pWindow->Center(pWindow->hWindow, hWndParent)

   pWindow->AddControl("LISTBOX", , IDC_FRMBUILDCONFIG_LIST1, "", 10, 10, 218, 362, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or _
        LBS_NOINTEGRALHEIGHT OR LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmBuildConfig_ListBox_SubclassProc), IDC_FRMBUILDCONFIG_LIST1, Cast(DWORD_PTR, @pWindow))

   pWindow->AddControl("LABEL", , IDC_FRMBUILDCONFIG_LABEL1, L(278,"Description") & ":", 240, 10, 91, 17, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMBUILDCONFIG_TXTDESCRIPTION, "", 240, 27, 224, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMBUILDCONFIG_LABEL2, L(279,"Compiler Options") & ":", 240, 50, 136, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMBUILDCONFIG_TXTOPTIONS, "", 240, 70, 372, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("RADIOBUTTON", , IDC_FRMBUILDCONFIG_OPT32, "FBC 32-bit", 240, 94, 100, 20, _
        WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER or WS_GROUP, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("RADIOBUTTON", , IDC_FRMBUILDCONFIG_OPT64, "FBC 64-bit", 342, 94, 100, 20, _
        WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("BUTTON", , IDC_FRMBUILDCONFIG_CMDINSERT, L(281, "Insert"), 8, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMBUILDCONFIG_CMDDELETE, L(282, "Delete"), 87, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMBUILDCONFIG_CMDUP, "", 166, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWUP", "IMAGE_ARROWUP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMBUILDCONFIG_CMDDOWN, "", 199, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("LABEL", , IDC_FRMBUILDCONFIG_LABEL9, "", 238, 370, 425, 3, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY Or SS_SUNKEN, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDOK, L(0,"OK"), 509, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 591, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   HWND_FRMBUILDCONFIG = hForm
   
   ' Copy all of the Builds to the BuildsTemp array because we will work with 
   ' temporary copies until the user hits OK.
   redim gConfig.BuildsTemp(ubound(gConfig.Builds))
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      gConfig.BuildsTemp(i) = gConfig.Builds(i)   
   NEXT
   LoadBuildListBox(hForm)
   
   ' Set the current selection to the same selection as the combobox builds in the toolbar
   dim as long nCurSel = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
   dim hList as hwnd = GetDlgItem(hForm, IDC_FRMBUILDCONFIG_LIST1)
   if nCurSel = -1 then nCurSel = 0 
   ListBox_SetCurSel(hList, nCurSel)
   SetCompilerConfigTextboxes()
   
   SetFocus GetDlgItem(hForm, IDC_FRMBUILDCONFIG_LIST1)
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmBuildConfig CWindow class manually allocated memory 
   Delete pWindow

End Function





