'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMCOMPILECONFIG_LIST1)

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
   dim as long idx = gApp.GetActiveProjectIndex()
   dim as string sBuildId

   ComboBox_ResetContent(HWND_FRMMAIN_COMBOBUILDS)
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      ComboBox_AddString(HWND_FRMMAIN_COMBOBUILDS, gConfig.Builds(i).wszDescription)
   NEXT

   ' 1) If project is active then set using the selected build for the project.
   If gApp.IsProjectActive Then
      sBuildId = gApp.Projects(idx).ProjectBuild
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
   dim as hwnd hList1 = GetDlgItem( HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_LIST1)
   
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
   dim as hwnd hList1 = GetDlgItem( HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_LIST1)
   dim as hwnd hText1 = GetDlgItem( HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_TXTDESCRIPTION)
   dim as hwnd hText2 = GetDlgItem( HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_TXTOPTIONS)
   dim as hwnd hCheck = GetDlgItem( HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_CHKISDEFAULT)
   dim as hwnd hOpt32 = GetDlgItem( HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_OPT32)
   dim as hwnd hOpt64 = GetDlgItem( HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_OPT64)
   
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   if nCurSel < 0 THEN 
      AfxSetWindowText( hText1, "")
      AfxSetWindowText( hText2, "")
      Button_SetCheck( hCheck, 0)
      CheckRadioButton(HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_OPT32, IDC_FRMCOMPILECONFIG_OPT64, IDC_FRMCOMPILECONFIG_OPT32)
      EnableWindow(hText1, false)
      EnableWindow(hText2, false)
      EnableWindow(hCheck, false)
      EnableWindow(hOpt32, false)
      EnableWindow(hOpt64, false)
   else
      AfxSetWindowText( hText1, gConfig.BuildsTemp(nCurSel).wszDescription)
      AfxSetWindowText( hText2, gConfig.BuildsTemp(nCurSel).wszOptions)
      Button_SetCheck( hCheck, gConfig.BuildsTemp(nCurSel).IsDefault)
      if gConfig.BuildsTemp(nCurSel).Is32bit THEN
         CheckRadioButton(HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_OPT32, IDC_FRMCOMPILECONFIG_OPT64, IDC_FRMCOMPILECONFIG_OPT32)
      else   
         CheckRadioButton(HWND_FRMCOMPILECONFIG, IDC_FRMCOMPILECONFIG_OPT32, IDC_FRMCOMPILECONFIG_OPT64, IDC_FRMCOMPILECONFIG_OPT64)
      END IF
      EnableWindow(hText1, true)
      EnableWindow(hText2, true)
      EnableWindow(hCheck, true)
      EnableWindow(hOpt32, true)
      EnableWindow(hOpt64, true)
   end if
   
   function = 0
end function
               

' ========================================================================================
' Process WM_CREATE message for window/dialog: frmCompileConfig
' ========================================================================================
Function frmCompileConfig_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' Retrieve a reference to the CWindow class from the CREATESTRUCT structure
   'Dim pCreateStruct As CREATESTRUCT Ptr = lpCreateStructPtr
   'Dim pWindow As CWindow Ptr = Cast(CWindow Ptr, pCreateStruct->lpCreateParams)

   ' This is a modal popup window so disable the parent window
   EnableWindow GetParent(HWnd), False

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmCompileConfig
' ========================================================================================
Function frmCompileConfig_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMCOMPILECONFIG_LIST1)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   ' Convert a ListBox DoubleClick into pressing OK button action
   if (id = IDC_FRMCOMPILECONFIG_LIST1) AndAlso (codeNotify = LBN_DBLCLK) THEN
      id = IDOK: codeNotify = BN_CLICKED
   END IF
   
   Select Case id
      case IDC_FRMCOMPILECONFIG_LIST1
         if codeNotify = LBN_SELCHANGE THEN
            SetCompilerConfigTextboxes()
         END IF
          
      case IDC_FRMCOMPILECONFIG_TXTDESCRIPTION
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array and the Listbox
            if nCurSel > -1 THEN
               gConfig.BuildsTemp(nCurSel).wszDescription = AfxGetWindowText(hwndCtl)
               ListBox_ReplaceString(hList1, nCurSel, gConfig.BuildsTemp(nCurSel).wszDescription)
            END IF
         end if
         
      CASE IDC_FRMCOMPILECONFIG_TXTOPTIONS
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array 
            if nCurSel > -1 THEN
               gConfig.BuildsTemp(nCurSel).wszOptions = AfxGetWindowText(hwndCtl)
            END IF
         end if
         
      case IDC_FRMCOMPILECONFIG_CHKISDEFAULT
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
                  
      case IDC_FRMCOMPILECONFIG_OPT32
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

      case IDC_FRMCOMPILECONFIG_OPT64
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

      case IDC_FRMCOMPILECONFIG_CMDUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               SwapBuildListboxItems(nCurSel, nCurSel-1)
            END IF
         end if

      case IDC_FRMCOMPILECONFIG_CMDDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1)-1 THEN
               SwapBuildListboxItems(nCurSel, nCurSel+1)
            END IF
         end if

      case IDC_FRMCOMPILECONFIG_CMDINSERT
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

      case IDC_FRMCOMPILECONFIG_CMDDELETE
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
            dim as long idx = gApp.GetActiveProjectIndex()
            if nCurSel = -1 THEN exit function
            If gApp.IsProjectActive Then
               gApp.Projects(idx).ProjectBuild = gConfig.Builds(nCurSel).Id
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
' Process WM_CLOSE message for window/dialog: frmCompileConfig
' ========================================================================================
Function frmCompileConfig_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow( GetParent(HWnd), True )
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmCompileConfig
' ========================================================================================
Function frmCompileConfig_OnDestroy(HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmCompileConfig Window procedure
' ========================================================================================
Function frmCompileConfig_WndProc( ByVal HWnd   As HWnd, _
                                    ByVal uMsg   As UINT, _
                                    ByVal wParam As WPARAM, _
                                    ByVal lParam As LPARAM _
                                    ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmCompileConfig_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmCompileConfig_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmCompileConfig_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmCompileConfig_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmCompileConfig_Show
' ========================================================================================
Function frmCompileConfig_Show( ByVal hWndParent As HWnd, _
                                ByVal nCmdShow   As Long = 0 _
                                ) As Long

   DIM hBitmap AS HBITMAP 
   dim hCtrl as HWnd
   dim wszImage as wstring * 100

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create(hWndParent, rtrim(L(277,"Build Configurations..."),"."), _
        @frmCompileConfig_WndProc, 0, 0, 396, 323, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(393, 315)
   pWindow->Center(pWindow->hWindow, hWndParent)
      
   pWindow->AddControl("LISTBOX", , IDC_FRMCOMPILECONFIG_LIST1, "", 10, 10, 223, 181, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMCOMPILECONFIG_LABEL1, L(278,"Description:"), 10, 196, 91, 17, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMCOMPILECONFIG_TXTDESCRIPTION, "", 10, 215, 224, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("CHECKBOX", , IDC_FRMCOMPILECONFIG_CHKISDEFAULT, L(280,"Set as default"), 243, 216, 172, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMCOMPILECONFIG_LABEL2, L(279,"Compiler Options:"), 10, 240, 136, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMCOMPILECONFIG_TXTOPTIONS, "", 10, 258, 372, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("RADIOBUTTON", , IDC_FRMCOMPILECONFIG_OPT32, "FBC 32-bit", 10, 285, 100, 20, _
        WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER or WS_GROUP, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("RADIOBUTTON", , IDC_FRMCOMPILECONFIG_OPT64, "FBC 64-bit", 115, 285, 100, 20, _
        WS_VISIBLE OR WS_TABSTOP OR BS_AUTORADIOBUTTON OR BS_LEFT OR BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMCOMPILECONFIG_CMDUP, "", 243, 10, 32, 32, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_ICON Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_UP", "IMAGE_UP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMCOMPILECONFIG_CMDDOWN, "", 243, 159, 32, 32, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_ICON Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_DOWN", "IMAGE_DOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("BUTTON", , IDOK,  L(0,"&OK"), 300, 10, 81, 33, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"&Cancel"), 300, 48, 81, 33, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMCOMPILECONFIG_CMDINSERT, L(281,"Insert"), 300, 120, 81, 33, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMCOMPILECONFIG_CMDDELETE, L(282,"Delete"), 300, 158, 81, 33, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   HWND_FRMCOMPILECONFIG = hForm
   
   ' Copy all of the Builds to the BuildsTemp array because we will work with 
   ' temporary copies until the user hits OK.
   redim gConfig.BuildsTemp(ubound(gConfig.Builds))
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      gConfig.BuildsTemp(i) = gConfig.Builds(i)   
   NEXT
   LoadBuildListBox(hForm)
   
   ' Set the current selection to the same selection as the combobox builds in the toolbar
   dim as long nCurSel = ComboBox_GetCurSel(HWND_FRMMAIN_COMBOBUILDS)
   dim hList as hwnd = GetDlgItem(hForm, IDC_FRMCOMPILECONFIG_LIST1)
   if nCurSel = -1 then nCurSel = 0 
   ListBox_SetCurSel(hList, nCurSel)
   SetCompilerConfigTextboxes()
   

   SetFocus GetDlgItem(hForm, IDC_FRMCOMPILECONFIG_LIST1)
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(nCmdShow)
   
   ' Delete the frmCompileConfig CWindow class manually allocated memory 
   Delete pWindow

End Function





