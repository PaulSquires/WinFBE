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

#include once "frmCategories.bi"
#include once "clsConfig.bi"


' ========================================================================================
' Load all of the build descriptions into the listbox
' ========================================================================================
function frmCategories_LoadBuildListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMCATEGORIES_LIST1)

   ListBox_ResetContent(hList1)
   for i as long = lbound(gCategoriesTemp) to ubound(gCategoriesTemp)
      ListBox_AddString(hList1, gConfig.BuildsTemp(i).wszDescription.sptr)
   NEXT

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
   
   ' Swap the array values
   swap gConfig.BuildsTemp(Item1), gConfig.BuildsTemp(Item2)
   
   ListBox_ReplaceString(hList1, Item1, gConfig.BuildsTemp(Item1).wszDescription)
   ListBox_ReplaceString(hList1, Item2, gConfig.BuildsTemp(Item2).wszDescription)

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
   lpmis->itemHeight = pWindow->ScaleY(FRMBUILDCONFIG_LISTBOX_LINEHEIGHT)
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
   dim as HICON hCheckBox
   Dim As RECT rc, rc2
   dim as long nWidth, nHeight
   dim as CWSTR wszText, wszOptions, wszAction
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   if pWindow = 0 THEN exit function
      
   If lpdis->itemID = -1 Then Exit Function
   
   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT

         SaveDC(lpdis->hDC)

         if lpdis->CtlID = IDC_FRMCATEGORIES_LIST1 then
            wszText = gConfig.BuildsTemp(lpdis->itemID).wszDescription
            hCheckBox = 0
         elseif lpdis->CtlID = IDC_FRMCATEGORIES_LSTOPTIONS then
            wszText = gBuildOptions(lpdis->itemID)
            ' Get the selected item in the Builds listbox and then look at the OptionList for
            ' that build for each of the individual OptionList build options.
            dim as long nSelBuild = ListBox_GetCurSel(GetDlgItem(hwnd, IDC_FRMCATEGORIES_LIST1))
            if nSelBuild > -1 then 
               wszOptions = " " & gConfig.BuildsTemp(nSelBuild).wszOptions & " "
               wszAction  = " " & AfxStrExtract(gBuildOptions(lpdis->itemID), "(", ")") & " "
               hCheckBox = iif( instr(wszOptions, wszAction), ghIconTick, ghIconNoTick)
            end if
         end if
         
         nWidth  = lpdis->rcItem.right-lpdis->rcItem.left
         nHeight = lpdis->rcItem.bottom-lpdis->rcItem.top 
         
         memDC = CreateCompatibleDC( lpdis->hDC )
         hbit  = CreateCompatibleBitmap( lpdis->hDC, nWidth, nHeight )
                                         
         If hbit Then hbit = SelectObject( memDC, hbit )

         ' Create our rect that works with the entire line
         SetRect(@rc, 0, 0, nWidth, nHeight)
         FillRect(memDC, @rc, GetSysColorBrush(COLOR_WINDOW))  
            
         ' Draw the tick/untick image
         ' The line height is FRMBUILDCONFIG_LISTBOX_LINEHEIGHT so center the 16x16 icon vertically and horizontally
         rc2 = rc
         if hCheckBox then
            DrawIconEx( memDC, _
                        rc.Left + pWindow->ScaleX(2), _
                        rc.Top + pWindow->ScaleY(2), _
                        hCheckBox, _
                        pWindow->ScaleX(16), pWindow->ScaleY(16), 0, 0, DI_NORMAL) 
            rc2.Left = pWindow->ScaleX(20)
         end if
            
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
                 
         ' Output the text
         SelectObject(memDC, AfxGetWindowFont(lpdis->hwndItem))
         SelectObject(memDC, hBrush)      
         FillRect(memDC, @rc2, hBrush)  

         rc2.Left = rc2.Left + pWindow->ScaleX(4)
         DrawText( memDC, wszText, _
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


         Function = True : Exit Function
   
   End Select

   Function = 0
End Function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
function frmCategories_ProcessListboxClick( _
            byval hwnd as HWND, _
            byval nCurSel as long _
            ) as long
   
   if nCurSel = -1 then exit function
   
   ' Get the action text for this listbox item. If it exists in the Options
   ' textbox then remove it
   dim as CWSTR wszText = " " & AfxGetWindowText( GetDlgItem(HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_TXTOPTIONS)) & " "
   dim as CWSTR wszAction = " " & AfxStrExtract(gBuildOptions(nCurSel), "(", ")") & " "
   ' If the action exists, remove it. If it does not exist then add it.
   if instr(wszText, wszAction) then
      wszText = trim(AfxStrRemove(wszText, wszAction))
   else
      wszText = trim(wszText) & " " & trim(wszAction)
   end if   
   AfxSetWindowText( GetDlgItem(HWND_FRMCATEGORIES, IDC_FRMCATEGORIES_TXTOPTIONS), wszText)

   function = 0

end function


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
         dim as RECT rc = GetListBoxEmptyClientArea( HWND, FRMBUILDCONFIG_LISTBOX_LINEHEIGHT )
         if rc.top < rc.bottom then
            dim as HDC hDC = cast(HDC, wParam)
            FillRect(hDC, @rc, GetSysColorBrush(COLOR_WINDOW))
         end if
         return TRUE
         
            
      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_SPACE
               dim as long nCurSel = ListBox_GetCurSel(hwnd)
               frmCategories_ProcessListboxClick(hwnd, nCurSel)
               Exit Function
            case VK_ESCAPE
               PostMessage( GetParent(HWnd), WM_CLOSE, 0, 0 )
               Exit Function
         End Select

      case WM_MOUSEMOVE
      case WM_LBUTTONDOWN
         
      case WM_LBUTTONUP
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
            frmCategories_ProcessListboxClick(hwnd, nCurSel)
            InvalidateRect(hwnd, @rc, true): UpdateWindow(hwnd)
         END IF
         
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmCategories_ListBox_SubclassProc, uIdSubclass

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

   ' Convert a ListBox DoubleClick into pressing OK button action
   if (id = IDC_FRMCATEGORIES_LIST1) AndAlso (codeNotify = LBN_DBLCLK) THEN
      id = IDOK: codeNotify = BN_CLICKED
   END IF
   
   Select Case id
      case IDC_FRMCATEGORIES_LIST1
         if codeNotify = LBN_SELCHANGE THEN
            frmBuildConfig_SetBuildConfigTextboxes()
         END IF
          
      case IDC_FRMCATEGORIES_TXTDESCRIPTION
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array and the Listbox
            if nCurSel > -1 THEN
               gConfig.BuildsTemp(nCurSel).wszDescription = AfxGetWindowText(hwndCtl)
               ListBox_ReplaceString(hList1, nCurSel, gConfig.BuildsTemp(nCurSel).wszDescription)
            END IF
         end if
         
      CASE IDC_FRMCATEGORIES_TXTOPTIONS
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array 
            if nCurSel > -1 THEN
               gConfig.BuildsTemp(nCurSel).wszOptions = AfxGetWindowText(hwndCtl)
               ' Refresh the Options listbox so any manually entered options will
               ' now be checked in the listbox.
               AfxRedrawWindow(GetDlgItem(hwnd, IDC_FRMCATEGORIES_LSTOPTIONS))
            END IF
         end if
         
      case IDC_FRMCATEGORIES_CMDUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               frmBuildConfig_SwapListboxItems(nCurSel, nCurSel-1)
            END IF
         end if

      case IDC_FRMCATEGORIES_CMDDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1)-1 THEN
               frmBuildConfig_SwapListboxItems(nCurSel, nCurSel+1)
            END IF
         end if

      case IDC_FRMCATEGORIES_CMDINSERT
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
            frmBuildConfig_LoadBuildListBox(HWND)
            nCurSel = Min(nCurSel, ubound(gConfig.BuildsTemp))
            ListBox_SetCurSel(hList1, nCurSel)
            frmBuildConfig_SetBuildConfigTextboxes()
            SetFocus hList1
         end if

      case IDC_FRMCATEGORIES_CMDDELETE
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
                  frmBuildConfig_LoadBuildListBox(HWND)
                  nCurSel = Min(nCurSel, ubound(gConfig.BuildsTemp))
                  ListBox_SetCurSel(hList1, nCurSel)
                  frmBuildConfig_SetBuildConfigTextboxes()
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
            if nCurSel > -1 THEN 
               If gApp.IsProjectActive Then
                  gApp.ProjectBuild = gConfig.Builds(nCurSel).Id
               Else
                  Dim pDocMain As clsDocument Ptr 
                  pDocMain = gTTabCtl.GetActiveDocumentPtr()
                  If pDocMain Then pDocMain->DocumentBuild = gConfig.Builds(nCurSel).Id
                  gConfig.CompilerBuild = gConfig.Builds(nCurSel).Id
               end if
               frmMain_SetStatusbar
            end if
            
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

   dim hCtrl as HWnd
   dim wszImage as wstring * 100
   
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create(hWndParent, L(436,"Explorer Categories"), _
        @frmBuildConfig_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(622, 436)
   pWindow->Center(pWindow->hWindow, hWndParent)

   pWindow->AddControl("LISTBOX", , IDC_FRMCATEGORIES_LIST1, "", 10, 10, 218, 362, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or _
        LBS_NOINTEGRALHEIGHT OR LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmBuildConfig_ListBox_SubclassProc), IDC_FRMCATEGORIES_LIST1, Cast(DWORD_PTR, @pWindow))

   pWindow->AddControl("LABEL", , IDC_FRMCATEGORIES_LABEL1, L(278,"Description") & ":", 240, 11, 91, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMCATEGORIES_TXTDESCRIPTION, "", 240, 32, 372, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDINSERT, L(281, "Insert"), 8, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMCATEGORIES_CMDDELETE, L(282, "Delete"), 87, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMCATEGORIES_CMDUP, "", 166, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWUP", "IMAGE_ARROWUP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMCATEGORIES_CMDDOWN, "", 199, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("BUTTON", , IDOK, L(0,"OK"), 454, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 536, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   HWND_FRMCATEGORIES = hForm
   
   ' Copy all of the Categories to the temp array because we will work with 
   ' temporary copies until the user hits OK.
   redim gCategoriesTemp(ubound(gCategories))
   for i as long = lbound(gCategories) to ubound(gCategories)
      gCategoriesTemp(i) = gCategories(i)   
   NEXT
   frmCategories_LoadBuildListBox(hForm)
   
   SetFocus GetDlgItem( hForm, IDC_FRMCATEGORIES_LIST1 )
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmBuildConfig CWindow class manually allocated memory 
   Delete pWindow

End Function

