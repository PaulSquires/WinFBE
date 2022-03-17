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

#include once "frmBuildConfig.bi"
#include once "clsConfig.bi"
#include once "clsTopTabCtl.bi"
#include once "clsDocument.bi"

' Create an array that holds all options in the Build compiler options listbox. The
' description contains a parenthesis enclosed action.
Dim shared gBuildOptions(...) as CWSTR => _
   {  "Language compatibility FreeBasic (-lang fb)", _
      "Language compatibility FreeBasic Lite (-lang fblite)", _
      "Language compatibility QuickBasic (-lang qb)", _
      "Subsystem to console (-s console)", _
      "Subsystem to GUI (-s gui)", _
      "Compiler 32-bit assembler backend (-gen gas)", _
      "Compiler 64-bit assembler backend (-gen gas64)", _
      "Compiler GCC backend (-gen gcc)", _
      "Compiler LLVM backend (-gen llvm)", _
      "Create DLL and import library (-dll)", _
      "Create static library (-lib)", _
      "Add error checking (-e)", _
      "Add error checking with RESUME support (-ex)", _
      "Same as -ex with array bounds and null pointer (-exx)", _
      "Add debug information (-g)", _
      "Compile only, do not link (-c)", _
      "Do not delete the object files (-C)", _
      "Emit preprocessed output, do not compile (-pp)" _
   }

' ========================================================================================
' Load all of the build descriptions into the listbox
' ========================================================================================
private function frmBuildConfig_LoadBuildListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMBUILDCONFIG_LIST1)

   ListBox_ResetContent(hList1)
   for i as long = lbound(gConfig.BuildsTemp) to ubound(gConfig.BuildsTemp)
      ListBox_AddString(hList1, gConfig.BuildsTemp(i).wszDescription.sptr)
   NEXT

   function = 0
end function


' ========================================================================================
' Determine the current active build index
' ========================================================================================
public function frmBuildConfig_getActiveBuildIndex() as long
   ' determine the default build in case no other build overrides it
   dim as long nDefault
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      if gConfig.Builds(i).IsDefault then
         nDefault = i: exit for
      end if
   next

   ' determine if a project or previous config selection overrides the default
   dim as long idx = -1
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      if gApp.IsProjectActive then
         if gConfig.Builds(i).id = gApp.ProjectBuild then
            idx = i: exit for
         end if
      else
         if gConfig.CompilerBuild = gConfig.Builds(i).id then
            idx = i: exit for
         end if
      end if
   next

   if idx = -1 then idx = nDefault
   function = idx
end function


' ========================================================================================
' Return the string description of the currently selected build configuration
' ========================================================================================
public function frmBuildConfig_GetSelectedBuildDescription() as CWSTR
   dim as long nCurSel = frmBuildConfig_getActiveBuildIndex()
   if nCurSel > -1 THEN
      function = gConfig.Builds(nCurSel).wszDescription
   END IF
end function


' ========================================================================================
' Return the string GUID of the currently selected build configuration
' ========================================================================================
public function frmBuildConfig_GetSelectedBuildGUID() as String
   dim as long nCurSel = frmBuildConfig_getActiveBuildIndex()
   if nCurSel > -1 THEN
      function = gConfig.Builds(nCurSel).Id
   END IF
end function


' ========================================================================================
' Return the string GUID of the default (if any) build configuration
' ========================================================================================
private function frmBuildConfig_GetDefaultBuildGUID() as String
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      if gConfig.Builds(i).IsDefault then 
         return gConfig.Builds(i).Id
      end if   
   NEXT
   function = ""
end function


' ========================================================================================
' Swap two entries in the build Listbox
' ========================================================================================
private function frmBuildConfig_SwapListBoxItems( byval Item1 as long, _
                                                  Byval Item2 as long _
                                                  ) as Long
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
private function frmBuildConfig_SetBuildConfigTextboxes() as long
   dim as hwnd hList1 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_LIST1)
   dim as hwnd hList2 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_LSTOPTIONS)
   dim as hwnd hCheck = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_CHKISDEFAULT)
   dim as hwnd hText1 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_TXTDESCRIPTION)
   dim as hwnd hText2 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_TXTOPTIONS)
   dim as hwnd hOpt32 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT32)
   dim as hwnd hOpt64 = GetDlgItem( HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT64)
   
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   if nCurSel < 0 THEN 
      AfxSetWindowText( hText1, "")
      AfxSetWindowText( hText2, "")
      Button_SetCheck( hCheck, 0)
      CheckRadioButton(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT32, IDC_FRMBUILDCONFIG_OPT64, IDC_FRMBUILDCONFIG_OPT32)
      EnableWindow(hText1, false)
      EnableWindow(hText2, false)
      EnableWindow(hCheck, false)
      EnableWindow(hOpt32, false)
      EnableWindow(hOpt64, false)
      EnableWindow(hList2, false)
   else
      AfxSetWindowText( hText1, gConfig.BuildsTemp(nCurSel).wszDescription)
      AfxSetWindowText( hText2, gConfig.BuildsTemp(nCurSel).wszOptions)
      Button_SetCheck( hCheck, gConfig.BuildsTemp(nCurSel).IsDefault)
      if gConfig.BuildsTemp(nCurSel).Is32bit THEN
         CheckRadioButton(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT32, IDC_FRMBUILDCONFIG_OPT64, IDC_FRMBUILDCONFIG_OPT32)
      else   
         CheckRadioButton(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_OPT32, IDC_FRMBUILDCONFIG_OPT64, IDC_FRMBUILDCONFIG_OPT64)
      END IF
      EnableWindow(hText1, true)
      EnableWindow(hText2, true)
      EnableWindow(hCheck, true)
      EnableWindow(hOpt32, true)
      EnableWindow(hOpt64, true)
      EnableWindow(hList2, true)
   end if

   ' Set the checkmarks in the OptionList Listbox by redrawing the Listbox
   AfxRedrawWindow(hList2)
      
   function = 0
end function
               

' ========================================================================================
' Process WM_MEASUREITEM message for window/dialog: frmBuildConfig
' ========================================================================================
private Function frmBuildConfig_OnMeasureItem( ByVal HWnd As HWnd, _
                                               ByVal lpmis As MEASUREITEMSTRUCT Ptr _
                                               ) As Long
   ' Set the height of the List box items. 
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   lpmis->itemHeight = pWindow->ScaleY(FRMBUILDCONFIG_LISTBOX_LINEHEIGHT)
   Function = 0
End Function


' ========================================================================================
' Process WM_DRAWITEM message for window/dialog: frmBuildConfig
' ========================================================================================
private Function frmBuildConfig_OnDrawItem( ByVal HWnd As HWnd, _
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

         if lpdis->CtlID = IDC_FRMBUILDCONFIG_LIST1 then
            wszText = gConfig.BuildsTemp(lpdis->itemID).wszDescription
            hCheckBox = 0
         elseif lpdis->CtlID = IDC_FRMBUILDCONFIG_LSTOPTIONS then
            wszText = gBuildOptions(lpdis->itemID)
            ' Get the selected item in the Builds listbox and then look at the OptionList for
            ' that build for each of the individual OptionList build options.
            dim as long nSelBuild = ListBox_GetCurSel(GetDlgItem(hwnd, IDC_FRMBUILDCONFIG_LIST1))
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
private function frmBuildConfig_ProcessListboxClick( byval hwnd as HWND, _
                                                     byval nCurSel as long _
                                                     ) as long
   if nCurSel = -1 then exit function
   
   ' Get the action text for this listbox item. If it exists in the Options
   ' textbox then remove it
   dim as CWSTR wszText = " " & AfxGetWindowText( GetDlgItem(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_TXTOPTIONS)) & " "
   dim as CWSTR wszAction = " " & AfxStrExtract(gBuildOptions(nCurSel), "(", ")") & " "
   ' If the action exists, remove it. If it does not exist then add it.
   if instr(wszText, wszAction) then
      wszText = trim(AfxStrRemove(wszText, wszAction))
   else
      wszText = trim(wszText) & " " & trim(wszAction)
   end if   
   AfxSetWindowText( GetDlgItem(HWND_FRMBUILDCONFIG, IDC_FRMBUILDCONFIG_TXTOPTIONS), wszText)

   function = 0

end function


' ========================================================================================
' Processes messages for the subclassed ListBox window.
' ========================================================================================
private Function frmBuildConfig_ListBox_SubclassProc ( _
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
               frmBuildConfig_ProcessListboxClick(hwnd, nCurSel)
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
            frmBuildConfig_ProcessListboxClick(hwnd, nCurSel)
            InvalidateRect(hwnd, @rc, true): UpdateWindow(hwnd)
         END IF
         
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
private Function frmBuildConfig_OnCreate( ByVal HWnd As HWnd, _
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
' Process WM_COMMAND message for window/dialog: frmBuildConfig
' ========================================================================================
private Function frmBuildConfig_OnCommand( ByVal HWnd As HWnd, _
                                           ByVal id As Long, _
                                           ByVal hwndCtl As HWnd, _
                                           ByVal codeNotify As UINT _
                                           ) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMBUILDCONFIG_LIST1)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   ' Convert a ListBox DoubleClick into pressing OK button action
   if (id = IDC_FRMBUILDCONFIG_LIST1) AndAlso (codeNotify = LBN_DBLCLK) THEN
      id = IDOK: codeNotify = BN_CLICKED
   END IF
   
   Select Case id
      case IDC_FRMBUILDCONFIG_LIST1
         if codeNotify = LBN_SELCHANGE THEN
            frmBuildConfig_SetBuildConfigTextboxes()
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
               ' Refresh the Options listbox so any manually entered options will
               ' now be checked in the listbox.
               AfxRedrawWindow(GetDlgItem(hwnd, IDC_FRMBUILDCONFIG_LSTOPTIONS))
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

      case IDC_FRMBUILDCONFIG_CHKISDEFAULT
         if codeNotify = BN_CLICKED THEN
            ' Update the temp array 
            if nCurSel > -1 THEN
               gConfig.BuildsTemp(nCurSel).IsDefault = Button_GetCheck(hwndCtl)
               ' Can only have 1 entry as the default so ensure all others are reset
               if Button_GetCheck(hwndCtl) THEN
                  for i as long = lbound(gConfig.BuildsTemp) to ubound(gConfig.BuildsTemp)
                     if i <> nCurSel THEN gConfig.BuildsTemp(i).IsDefault = 0
                  NEXT
               END IF
            END IF
         end if

      case IDC_FRMBUILDCONFIG_CMDUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               frmBuildConfig_SwapListboxItems(nCurSel, nCurSel-1)
            END IF
         end if

      case IDC_FRMBUILDCONFIG_CMDDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1)-1 THEN
               frmBuildConfig_SwapListboxItems(nCurSel, nCurSel+1)
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
            frmBuildConfig_LoadBuildListBox(HWND)
            nCurSel = Min(nCurSel, ubound(gConfig.BuildsTemp))
            ListBox_SetCurSel(hList1, nCurSel)
            frmBuildConfig_SetBuildConfigTextboxes()
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
            
            gConfig.SaveConfigFile
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
private Function frmBuildConfig_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmBuildConfig
' ========================================================================================
private Function frmBuildConfig_OnDestroy( byval HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmBuildConfig Window procedure
' ========================================================================================
private Function frmBuildConfig_WndProc( ByVal HWnd   As HWnd, _
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
public Function frmBuildConfig_Show( ByVal hWndParent As HWnd ) As LRESULT

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
   pWindow->SetClientSize(622, 436)
   pWindow->Center(pWindow->hWindow, hWndParent)

   pWindow->AddControl("LISTBOX", , IDC_FRMBUILDCONFIG_LIST1, "", 10, 10, 218, 362, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or _
        LBS_NOINTEGRALHEIGHT OR LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmBuildConfig_ListBox_SubclassProc), IDC_FRMBUILDCONFIG_LIST1, Cast(DWORD_PTR, @pWindow))

   pWindow->AddControl("LABEL", , IDC_FRMBUILDCONFIG_LABEL1, L(278,"Description") & ":", 240, 11, 91, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMBUILDCONFIG_CHKISDEFAULT, L(280,"Set as default"), 380, 10, 240, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMBUILDCONFIG_TXTDESCRIPTION, "", 240, 32, 372, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("RADIOBUTTON", , IDC_FRMBUILDCONFIG_OPT32, "FBC 32-bit", 240, 54, 100, 20, _
        WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER or WS_GROUP, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("RADIOBUTTON", , IDC_FRMBUILDCONFIG_OPT64, "FBC 64-bit", 342, 54, 100, 20, _
        WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMBUILDCONFIG_LABEL2, L(279,"Compiler Options") & ":", 240, 82, 136, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMBUILDCONFIG_TXTOPTIONS, "", 240, 103, 372, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LISTBOX", , IDC_FRMBUILDCONFIG_LSTOPTIONS, "", 240, 129, 372, 243, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or _
        LBS_NOINTEGRALHEIGHT OR LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmBuildConfig_ListBox_SubclassProc), IDC_FRMBUILDCONFIG_LSTOPTIONS, Cast(DWORD_PTR, @pWindow))


   pWindow->AddControl("BUTTON", , IDC_FRMBUILDCONFIG_CMDINSERT, L(281, "Insert"), 8, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMBUILDCONFIG_CMDDELETE, L(282, "Delete"), 87, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   hCtrl = _
   pWindow->AddControl("BUTTON", , IDC_FRMBUILDCONFIG_CMDUP, wszTriangleUp, 166, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghMenuBar.hFontSymbolLargeBold, false )     

   hCtrl = _
   pWindow->AddControl("BUTTON", , IDC_FRMBUILDCONFIG_CMDDOWN, wszTriangleDown, 199, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghMenuBar.hFontSymbolLargeBold, false )     

   pWindow->AddControl("BUTTON", , IDOK, L(0,"OK"), 454, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 536, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   HWND_FRMBUILDCONFIG = hForm
   
   ' Copy all of the Builds to the BuildsTemp array because we will work with 
   ' temporary copies until the user hits OK.
   redim gConfig.BuildsTemp(ubound(gConfig.Builds))
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      gConfig.BuildsTemp(i) = gConfig.Builds(i)   
   NEXT
   frmBuildConfig_LoadBuildListBox(hForm)
   
   ' Load the Compiler Options Listbox
   hCtrl = GetDlgItem(hForm, IDC_FRMBUILDCONFIG_LSTOPTIONS)
   for i as long = lbound(gBuildOptions) to ubound(gBuildOptions)
      ListBox_AddString(hCtrl, gBuildOptions(i).sptr)
   next

   
   ' Set the current build selection 
   dim as long nCurSel = frmBuildConfig_getActiveBuildIndex()
   dim hList as hwnd = GetDlgItem(hForm, IDC_FRMBUILDCONFIG_LIST1)
   if nCurSel = -1 then nCurSel = 0 
   ListBox_SetCurSel(hList, nCurSel)
   frmBuildConfig_SetBuildConfigTextboxes()
   
   SetFocus GetDlgItem(hForm, IDC_FRMBUILDCONFIG_LIST1)
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmBuildConfig CWindow class manually allocated memory 
   Delete pWindow

End Function

