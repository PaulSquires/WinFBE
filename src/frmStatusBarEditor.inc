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

#include once "frmStatusBarEditor.bi"
#include once "clsConfig.bi"


' ========================================================================================
' Load all of the Panel descriptions into the listbox
' ========================================================================================
private function frmStatusBarEditor_LoadListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMSTATUSBAREDITOR_LSTPANELS)

'   ListBox_ResetContent(hList1)
'   for i as long = lbound(gConfig.ToolsTemp) to ubound(gConfig.ToolsTemp)
'      ListBox_AddString(hList1, gConfig.ToolsTemp(i).wszDescription)
'   NEXT

   function = 0
end function


' ========================================================================================
' Swap two entries in the StatusBarEditor Listbox
' ========================================================================================
private function frmStatusBarEditor_SwapListBoxItems( byval Item1 as long, _
                                                      Byval Item2 as long _
                                                      ) as Long
   dim as hwnd hList1 = GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
   
'   ' Swap the array values
'   swap gConfig.ToolsTemp(Item1), gConfig.ToolsTemp(Item2)
    
'   ListBox_ReplaceString(hList1, Item1, gConfig.ToolsTemp(Item1).wszDescription)
'   ListBox_ReplaceString(hList1, Item2, gConfig.ToolsTemp(Item2).wszDescription)

   function = 0
end function


' ========================================================================================
' Position all child windows. Called manually.
' ========================================================================================
private Function frmStatusBarEditor_PositionWindows() As LRESULT
   
'   Dim As HWnd hTabCtl = GetDlgItem( HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TABCONTROL )
'   dim as long nCurSel = TabCtrl_GetCurSel( hTabCtl )

'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LABEL1), _
'               iif( nCurSel = 1, SW_HIDE, SW_SHOW ) )
'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LABEL2), _
'               iif( nCurSel = 1, SW_HIDE, SW_SHOW ) )
'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOSTYLE), _
'               iif( nCurSel = 1, SW_HIDE, SW_SHOW ) )
'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTSIMPLETEXT), _
'               iif( nCurSel = 1, SW_HIDE, SW_SHOW ) )

'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS), _
'               iif( nCurSel = 0, SW_HIDE, SW_SHOW ) )
'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CMDINSERT), _
'               iif( nCurSel = 0, SW_HIDE, SW_SHOW ) )
'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CMDDELETE), _
'               iif( nCurSel = 0, SW_HIDE, SW_SHOW ) )
'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CMDUP), _
'               iif( nCurSel = 0, SW_HIDE, SW_SHOW ) )
'   ShowWindow( GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CMDDOWN), _
'               iif( nCurSel = 0, SW_HIDE, SW_SHOW ) )
   
   Function = 0
End Function


' ========================================================================================
' Set the frmStatusBarEditor textboxes and options depending on what listbox entry is selected
' ========================================================================================
private function frmStatusBarEditor_SetTextboxes() as long
   dim as long nCurSel = ListBox_GetCurSel(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_LSTPANELS))
   dim as Boolean fEnabled = iif( nCurSel < 0, false, true)

'   AfxSetWindowText(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTTOOLNAME), iif(fEnabled, **gConfig.ToolsTemp(nCurSel).wszDescription, wstr("")))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTTOOLNAME), fEnabled)
'   AfxSetWindowText(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTCOMMAND), iif(fEnabled, **gConfig.ToolsTemp(nCurSel).wszCommand, wstr("")))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTCOMMAND), fEnabled)
'   AfxSetWindowText(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTPARAMETERS), iif(fEnabled, **gConfig.ToolsTemp(nCurSel).wszParameters, wstr("")))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTPARAMETERS), fEnabled)
'   AfxSetWindowText(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTKEY), iif(fEnabled, **gConfig.ToolsTemp(nCurSel).wszKey, wstr("")))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTKEY), fEnabled)
'   AfxSetWindowText(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTWORKINGFOLDER), iif(fEnabled, **gConfig.ToolsTemp(nCurSel).wszWorkingFolder, wstr("")))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_TXTWORKINGFOLDER), fEnabled)
'   Button_SetCheck(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKCTRL), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsCtrl, 0))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKCTRL), fEnabled)
'   Button_SetCheck(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKALT), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsAlt, 0))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKALT), fEnabled)
'   Button_SetCheck(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKSHIFT), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsShift, 0))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKSHIFT), fEnabled)
'   Button_SetCheck(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKPROMPT), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsPromptRun, 0))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKPROMPT), fEnabled)
'   Button_SetCheck(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKMINIMIZED), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsMinimized, 0))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKMINIMIZED), fEnabled)
'   Button_SetCheck(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKWAIT), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsWaitFinish, 0))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKWAIT), fEnabled)
'   Button_SetCheck(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKDISPLAYMENU), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsDisplayMenu, 0))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_CHKDISPLAYMENU), fEnabled)
'   ComboBox_SetCurSel(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOACTION), iif(fEnabled, gConfig.ToolsTemp(nCurSel).Action, 0))
'   EnableWindow(GetDlgItem(HWND_FRMSTATUSBAREDITOR, IDC_FRMSTATUSBAREDITOR_COMBOACTION), fEnabled)

   function = 0
end function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnCreate( ByVal HWnd As HWnd, _
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
' Process WM_COMMAND message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnCommand( ByVal HWnd As HWnd, _
                                               ByVal id As Long, _
                                               ByVal hwndCtl As HWnd, _
                                               ByVal codeNotify As UINT _
                                               ) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   Select Case id
      case IDC_FRMSTATUSBAREDITOR_LSTPANELS
         if codeNotify = LBN_SELCHANGE THEN
            frmStatusBarEditor_SetTextboxes()
         END IF
           
'      case IDC_FRMSTATUSBAREDITOR_TXTTOOLNAME
'         if codeNotify = EN_CHANGE THEN
'            ' Update the temp array and the Listbox
'            if nCurSel > -1 THEN
'               gConfig.ToolsTemp(nCurSel).wszDescription = AfxGetWindowText(hwndCtl)
'               ListBox_ReplaceString(hList1, nCurSel, gConfig.ToolsTemp(nCurSel).wszDescription)
'            END IF
'         end if
           
'      CASE IDC_FRMSTATUSBAREDITOR_TXTCOMMAND
'         if codeNotify = EN_CHANGE THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).wszCommand = AfxGetWindowText(hwndCtl)
'         end if
           
'      case IDC_FRMSTATUSBAREDITOR_CMDBROWSEEXE
'         if codeNotify = BN_CLICKED THEN
'            Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, IDC_FRMSTATUSBAREDITOR_CMDBROWSEEXE)
'            If pwszName Then
'               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMSTATUSBAREDITOR_TXTCOMMAND), pwszName )
'               CoTaskMemFree pwszName
'            End If
'         end if
          
'      CASE IDC_FRMSTATUSBAREDITOR_TXTPARAMETERS
'         if codeNotify = EN_CHANGE THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).wszParameters = AfxGetWindowText(hwndCtl)
'         end if

'      CASE IDC_FRMSTATUSBAREDITOR_TXTKEY
'         if codeNotify = EN_CHANGE THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).wszKey = AfxGetWindowText(hwndCtl)
'         end if

'      CASE IDC_FRMSTATUSBAREDITOR_TXTWORKINGFOLDER
'         if codeNotify = EN_CHANGE THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).wszWorkingFolder = AfxGetWindowText(hwndCtl)
'         end if
                    
'      case IDC_FRMSTATUSBAREDITOR_CMDBROWSEFOLDER
'         if codeNotify = BN_CLICKED THEN
'            dim as CWSTR cwsFolder = AfxBrowseForFolder(HWND, L(293,"Working Folder:"), curdir)
'            if len(cwsFolder) THEN
'               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMSTATUSBAREDITOR_TXTWORKINGFOLDER), cwsFolder)
'            END IF
'         end if
          
'      case IDC_FRMSTATUSBAREDITOR_CHKCTRL
'         if codeNotify = BN_CLICKED THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsCtrl = Button_GetCheck(hwndCtl) 
'         end if

'      case IDC_FRMSTATUSBAREDITOR_CHKALT
'         if codeNotify = BN_CLICKED THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsAlt = Button_GetCheck(hwndCtl) 
'         end if

'      case IDC_FRMSTATUSBAREDITOR_CHKSHIFT
'         if codeNotify = BN_CLICKED THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsShift = Button_GetCheck(hwndCtl) 
'         end if

'      case IDC_FRMSTATUSBAREDITOR_CHKPROMPT
'         if codeNotify = BN_CLICKED THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsPromptRun = Button_GetCheck(hwndCtl) 
'         end if

'      case IDC_FRMSTATUSBAREDITOR_CHKMINIMIZED
'         if codeNotify = BN_CLICKED THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsMinimized = Button_GetCheck(hwndCtl) 
'         end if

'      case IDC_FRMSTATUSBAREDITOR_CHKWAIT
'         if codeNotify = BN_CLICKED THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsWaitFinish = Button_GetCheck(hwndCtl) 
'         end if

'      case IDC_FRMSTATUSBAREDITOR_CHKDISPLAYMENU
'         if codeNotify = BN_CLICKED THEN
'            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsDisplayMenu = Button_GetCheck(hwndCtl) 
'         end if

'      case IDC_FRMSTATUSBAREDITOR_COMBOACTION
'         if codeNotify = CBN_SELCHANGE THEN
'            if nCurSel > -1 then gConfig.ToolsTemp(nCurSel).Action = ComboBox_GetCurSel(hwndCtl) 
'         END IF

      case IDC_FRMSTATUSBAREDITOR_CMDPANELUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               frmStatusBarEditor_SwapListboxItems(nCurSel, nCurSel-1)
            END IF
         end if

      case IDC_FRMSTATUSBAREDITOR_CMDPANELDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1)-1 THEN
               frmStatusBarEditor_SwapListboxItems(nCurSel, nCurSel+1)
            END IF
         end if

      case IDC_FRMSTATUSBAREDITOR_CMDPANELINSERT
         if codeNotify = BN_CLICKED THEN
'            if ubound(gConfig.ToolsTemp) = -1 THEN
'               redim gConfig.ToolsTemp(0)
'               nCurSel = 0
'            Else
'               redim preserve gConfig.ToolsTemp(ubound(gConfig.ToolsTemp)+1)
'               if nCurSel = -1 THEN nCurSel = 0
'               ' insert the item above current entry in the internal array
'               for i as long = ubound(gConfig.ToolsTemp) to nCurSel + 1 step -1
'                  gConfig.ToolsTemp(i) = gConfig.ToolsTemp(i-1)
'               NEXT
'            END IF
'            gConfig.ToolsTemp(nCurSel).wszDescription   = ""
'            gConfig.ToolsTemp(nCurSel).wszCommand       = ""
'            gConfig.ToolsTemp(nCurSel).wszParameters    = ""
'            gConfig.ToolsTemp(nCurSel).wszKey           = ""
'            gConfig.ToolsTemp(nCurSel).wszWorkingFolder = ""
'            gConfig.ToolsTemp(nCurSel).IsCtrl           = 0
'            gConfig.ToolsTemp(nCurSel).IsAlt            = 0
'            gConfig.ToolsTemp(nCurSel).IsShift          = 0
'            gConfig.ToolsTemp(nCurSel).IsPromptRun      = 0
'            gConfig.ToolsTemp(nCurSel).IsMinimized      = 0
'            gConfig.ToolsTemp(nCurSel).IsWaitFinish     = 0
'            gConfig.ToolsTemp(nCurSel).IsDisplayMenu    = 0
'            ' reload the listbox
'            frmUserTools_LoadListBox(HWND)
'            nCurSel = Min(nCurSel, ubound(gConfig.ToolsTemp))
'            ListBox_SetCurSel(hList1, nCurSel)
'            frmUserTools_SetTextboxes()
'            SetFocus GetDlgItem(HWND, IDC_FRMSTATUSBAREDITOR_TXTTOOLNAME)
         end if

      case IDC_FRMSTATUSBAREDITOR_CMDPANELDELETE
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1  THEN
'               if MessageBox( HWND, L(300, "Are you sure you want to delete this user tool?"), L(276,"Confirm"), _
'                                 MB_YESNOCANCEL Or MB_ICONINFORMATION Or MB_DEFBUTTON1 ) = IDYES then
'                  if ubound(gConfig.ToolsTemp) = 0 THEN
'                     erase gConfig.ToolsTemp
'                     nCurSel = -1
'                  else
'                     ' remove the item from the internal array
'                     for i as long = nCurSel to ubound(gConfig.ToolsTemp) - 1
'                        gConfig.ToolsTemp(i) = gConfig.ToolsTemp(i+1)
'                     NEXT
'                     redim preserve gConfig.ToolsTemp(ubound(gConfig.ToolsTemp)-1)
'                  END IF
'                  ' reload the listbox
'                  frmUserTools_LoadListBox(HWND)
'                  nCurSel = Min(nCurSel, ubound(gConfig.ToolsTemp))
'                  ListBox_SetCurSel(hList1, nCurSel)
'                  frmUserTools_SetTextboxes()
'                  SetFocus hList1
'               end if
            END IF
         end if

      Case IDC_FRMSTATUSBAREDITOR_CMDOK    
         If codeNotify = BN_CLICKED Then
'            ' Copy the temporary items to the main array  
'            redim gConfig.Tools(ubound(gConfig.ToolsTemp))
'            for i as long = lbound(gConfig.ToolsTemp) to ubound(gConfig.ToolsTemp)
'               gConfig.Tools(i) = gConfig.ToolsTemp(i)   
'            NEXT
'            erase gConfig.ToolsTemp
'            gConfig.SaveConfigFile
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
' Process WM_NOTIFY message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnNotify( ByVal HWnd As HWnd, _
                                              ByVal id As Long, _
                                              ByVal pNMHDR As NMHDR Ptr _
                                              ) As LRESULT

   Select Case pNMHDR->code 
      
      Case TCN_SELCHANGE
'         if id = IDC_FRMSTATUSBAREDITOR_TABCONTROL then
'            frmStatusBarEditor_PositionWindows
'         end if
   
   end select
   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmStatusBarEditor
' ========================================================================================
private Function frmStatusBarEditor_OnDestroy( byval HWnd As HWnd ) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmStatusBarEditor Window procedure
' ========================================================================================
private Function frmStatusBarEditor_WndProc( ByVal HWnd   As HWnd, _
                                             ByVal uMsg   As UINT, _
                                             ByVal wParam As WPARAM, _
                                             ByVal lParam As LPARAM _
                                             ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmStatusBarEditor_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmStatusBarEditor_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmStatusBarEditor_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmStatusBarEditor_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,   frmStatusBarEditor_OnNotify)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmStatusBarEditor_Show
' ========================================================================================
public Function frmStatusBarEditor_Show( ByVal hWndParent As HWnd ) as LRESULT

   DIM hBitmap AS HBITMAP
   dim hCtrl as HWnd
   dim wszImage as wstring * 100
   

   '  Create the main window and child controls
   Dim pWindow as CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI
   
   Dim As HWnd hForm = _
   pWindow->Create( hWndParent, L(314,"Statusbar Editor"), _
        @frmStatusBarEditor_WndProc, 0, 0, 408, 380, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(408, 380)
   pWindow->Center
        
   HWND_FRMSTATUSBAREDITOR = hForm

   pWindow->AddControl("LISTBOX", , IDC_FRMSTATUSBAREDITOR_LSTPANELS, "", 21, 14, 365, 84, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELUP, "", 21, 102, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWUP", "IMAGE_ARROWUP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELDOWN, "", 48, 102, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELINSERT, L(281, "Insert"), 77, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDPANELDELETE, L(282, "Delete"), 152, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL1, L(150,"Text") & ":", 21, 135, 93, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTTEXT, "", 114, 134, 272, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL2, L(235,"Tooltip Text") & ":", 21, 160, 93, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTTOOLTIP, "", 114, 158, 272, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL3, L(236,"Alignment") & ":", 21, 194, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOALIGNMENT, "", 114, 190, 107, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL4, L(237,"Style") & ":", 21, 220, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOSTYLE, "", 114, 217, 107, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL5, L(238,"Bevel") & ":", 21, 247, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOBEVEL, "", 114, 244, 107, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL6, L(239,"Autosize") & ":", 21, 275, 93, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMSTATUSBAREDITOR_COMBOAUTOSIZE, "", 114, 271, 107, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSTATUSBAREDITOR_LABEL7, L(245,"Minimum Width") & ":", 21, 301, 93, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSTATUSBAREDITOR_TXTMINWIDTH, "", 114, 298, 107, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)


   pWindow->AddControl("GROUPBOX", , IDC_FRMSTATUSBAREDITOR_FRAME1, L(246,"Image"), 242, 190, 144, 128, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)


'   pWindow->AddControl("BUTTON", pWindow->hWindow, IDC_STATUSBAREDITOR_CMDPICSELECT, "Select", 304, 200, 48, 24, _
'        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
'        WS_EX_LEFT Or WS_EX_LTRREADING)
'   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_STATUSBAREDITOR_PICICON, "", 250, 220, 32, 32, _
'        WS_CHILD Or WS_BORDER Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_ICON Or SS_NOTIFY Or LR_SHARED, _
'        )


   pWindow->AddControl("BUTTON", , IDC_FRMSTATUSBAREDITOR_CMDOK, L(0,"OK"), 231, 333, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 312, 333, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


'   dim hList as hwnd = GetDlgItem(hForm, IDC_FRMSTATUSBAREDITOR_LSTPANELS)
'   ListBox_SetCurSel( hList, 0 )
'   TabCtrl_SetCurSel( hTabCtl, 0 )
'   frmStatusBarEditor_PositionWindows()
'   frmStatusBarEditor_SetTextboxes()
'   SetFocus hList

   ' Process Windows messages
   Function = pWindow->DoEvents( SW_SHOW )
   
   ' Delete the frmStatusBarEditor CWindow class manually allocated memory 
   Delete pWindow

   function = 0
end function

