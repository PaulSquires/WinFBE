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

#include once "frmVDTabChild.bi"


' ========================================================================================
' Load all of the tab child pages into the global gTabPages array
' ========================================================================================
function frmVDTabChild_LoadTabPagesArray( byref wszPropValue as wstring ) as Long
   
   if len(trim(wszPropValue)) then
      dim as long NumRecs = AfxStrParseCount( wszPropValue, gTabRecordSep )
      redim gTabPages( NumRecs - 1 )
     
      for i as long = 0 to ubound(gTabPages)
         dim as CWSTR wszRecord = AfxStrParse( wszPropValue, i+1, gTabRecordSep )
         ' the record's fields are separated by gTabFieldSep
         gTabPages(i).wszText      = AfxStrParse( wszRecord, 1, gTabFieldSep )
         gTabPages(i).wszImage     = AfxStrParse( wszRecord, 2, gTabFieldSep )
         gTabPages(i).wszTabPage   = AfxStrParse( wszRecord, 3, gTabFieldSep )
         gTabPages(i).IsActiveTab  = val(AfxStrParse( wszRecord, 4, gTabFieldSep ))
         gTabPages(i).wszReserved1 = AfxStrParse( wszRecord, 5, gTabFieldSep )
         gTabPages(i).wszReserved2 = AfxStrParse( wszRecord, 6, gTabFieldSep )
         gTabPages(i).wszReserved3 = AfxStrParse( wszRecord, 7, gTabFieldSep )
         gTabPages(i).wszReserved4 = AfxStrParse( wszRecord, 8, gTabFieldSep )
      next
   end if

   function = 0
end function

' ========================================================================================
' Load all of the tab child pages into the listbox
' ========================================================================================
private function frmVDTabChild_LoadBuildListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMVDTABCHILD_LIST1)

   ListBox_ResetContent(hList1)
   for i as long = lbound(gTabPages) to ubound(gTabPages)
      ListBox_AddString(hList1, gTabPages(i).wszText.sptr)
   NEXT

   function = 0
end function


' ========================================================================================
' Swap two entries in the Listbox
' ========================================================================================
private function frmVDTabChild_SwapListBoxItems( byval Item1 as long, _
                                                 Byval Item2 as long _
                                                 ) as Long
   dim as hwnd hList1 = GetDlgItem( HWND_frmVDTabChild, IDC_FRMVDTABCHILD_LIST1)
   
   ' Swap the array values
   swap gTabPages(Item1), gTabPages(Item2)
    
   ListBox_ReplaceString(hList1, Item1, gTabPages(Item1).wszText)
   ListBox_ReplaceString(hList1, Item2, gTabPages(Item2).wszText)

   function = 0
end function


' ========================================================================================
' Set the tab children depending on what listbox entry is selected
' ========================================================================================
private function frmVDTabChild_SetTabChildTextboxes() as long
   dim as hwnd hList1 = GetDlgItem( HWND_frmVDTabChild, IDC_FRMVDTABCHILD_LIST1)
   dim as hwnd hText1 = GetDlgItem( HWND_frmVDTabChild, IDC_FRMVDTABCHILD_TXTDESCRIPTION)
   dim as hwnd hText2 = GetDlgItem( HWND_frmVDTabChild, IDC_FRMVDTABCHILD_TXTIMAGE)
   dim as hwnd hCombo = GetDlgItem( HWND_frmVDTabChild, IDC_FRMVDTABCHILD_COMBOTABPAGES)
   dim as hwnd hCheck = GetDlgItem( HWND_frmVDTabChild, IDC_FRMVDTABCHILD_CHKISDEFAULT)
    
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   if nCurSel < 0 THEN 
      AfxSetWindowText( hText1, "" )
      AfxSetWindowText( hText2, "" )
      ComboBox_SetCurSel( hCombo, 0 )
      Button_SetCheck( hCheck, 0 )
      EnableWindow(hText1, false)
      EnableWindow(hText2, false)
      EnableWindow(hCombo, false)
      EnableWindow(hCheck, false)
   else
      AfxSetWindowText( hText1, gTabPages(nCurSel).wszText)
      AfxSetWindowText( hText2, gTabPages(nCurSel).wszImage)
      Button_SetCheck( hCheck, gTabPages(nCurSel).IsActiveTab)
      EnableWindow(hText1, true)
      EnableWindow(hText2, true)
      EnableWindow(hCombo, true)
      EnableWindow(hCheck, true)

      ' Match the child form combobox
      ComboBox_SetCurSel( hCombo, 0 )
      dim as long nCount = ComboBox_GetCount(hCombo)
      for i as long = 0 to nCount - 1
         if ucase(gTabPages(nCurSel).wszTabPage) = ucase(AfxGetComboBoxText(hCombo, i)) then
            ComboBox_SetCurSel(hCombo, i)
            exit for
         END IF
      next
   end if

   function = 0
end function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmVDTabChild
' ========================================================================================
private Function frmVDTabChild_OnCreate( ByVal HWnd As HWnd, _
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
' Process WM_COMMAND message for window/dialog: frmVDTabChild
' ========================================================================================
private Function frmVDTabChild_OnCommand( ByVal HWnd As HWnd, _
                                          ByVal id As Long, _
                                          ByVal hwndCtl As HWnd, _
                                          ByVal codeNotify As UINT _
                                          ) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMVDTABCHILD_LIST1)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   Select Case id
      case IDC_FRMVDTABCHILD_LIST1
         if codeNotify = LBN_SELCHANGE THEN
            frmVDTabChild_SetTabChildTextboxes()
         END IF
           
      case IDC_FRMVDTABCHILD_TXTDESCRIPTION
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array and the Listbox
            if nCurSel > -1 THEN
               gTabPages(nCurSel).wszText = AfxGetWindowText(hwndCtl)
               ListBox_ReplaceString(hList1, nCurSel, gTabPages(nCurSel).wszText)
            END IF
         end if
          
      CASE IDC_FRMVDTABCHILD_TXTIMAGE
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array 
            if nCurSel > -1 THEN
               gTabPages(nCurSel).wszImage = AfxGetWindowText(hwndCtl)
            END IF
         end if
          
      case IDC_FRMVDTABCHILD_CMDIMAGE
         If codeNotify = BN_CLICKED Then
            if nCurSel > -1 then
               dim pProp as clsProperty ptr = new clsProperty
               pProp->wszPropValue = gTabPages(nCurSel).wszImage
               
               ' Need to save the gTabPages array because the frmImageManager will
               ' reset it during a pDoc Save.
               dim temp(ubound(gTabPages)) as TabPage
               for i as long = lbound(gTabPages) to ubound(gTabPages)
                  temp(i) = gTabPages(i)
               next
               
               frmImageManager_Show( HWND, pProp )
               
               ' Need to reload the gTabPages array that would have been reset.
               redim gTabPages(ubound(temp)) as TabPage
               for i as long = lbound(temp) to ubound(temp)
                  gTabPages(i) = temp(i)
               next
 
               gTabPages(nCurSel).wszImage = pProp->wszPropValue

               delete pProp
               frmVDTabChild_SetTabChildTextboxes
            end if
         end if
         
      case IDC_FRMVDTABCHILD_COMBOTABPAGES
         if codeNotify = CBN_SELCHANGE THEN
            if nCurSel > -1 then 
               dim as long nComboSel = ComboBox_GetCurSel(hwndCtl)
               gTabPages(nCurSel).wszTabPage = AfxGetComboBoxText(hwndCtl, nComboSel) 
            end if
         END IF

      case IDC_FRMVDTABCHILD_CHKISDEFAULT
         if codeNotify = BN_CLICKED THEN
            ' Update the temp array 
            if nCurSel > -1 THEN
               gTabPages(nCurSel).IsActiveTab = Button_GetCheck(hwndCtl)
               ' Can only have 1 entry as the default so ensure all others are reset
               if Button_GetCheck(hwndCtl) THEN
                  for i as long = lbound(gTabPages) to ubound(gTabPages)
                     if i <> nCurSel THEN gTabPages(i).IsActiveTab = 0
                  NEXT
               END IF
            END IF
         end if

      case IDC_FRMVDTABCHILD_CMDUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               frmVDTabChild_SwapListboxItems(nCurSel, nCurSel-1)
            END IF
         end if

      case IDC_FRMVDTABCHILD_CMDDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1)-1 THEN
               frmVDTabChild_SwapListboxItems(nCurSel, nCurSel+1)
            END IF
         end if

      case IDC_FRMVDTABCHILD_CMDINSERT, IDC_FRMVDTABCHILD_CMDADD
         if codeNotify = BN_CLICKED THEN
            static as long nNextTabNum 
            nNextTabNum = nNextTabNum + 1
            if ubound(gTabPages) = -1 THEN
               redim gTabPages(0)
               nCurSel = 0
            Else
               redim preserve gTabPages(ubound(gTabPages)+1)
               if nCurSel = -1 THEN nCurSel = 0
               if id = IDC_FRMVDTABCHILD_CMDINSERT then
                  ' insert the item above current entry in the internal array
                  for i as long = ubound(gTabPages) to nCurSel + 1 step -1
                     gTabPages(i) = gTabPages(i-1)
                  NEXT
               elseif id = IDC_FRMVDTABCHILD_CMDADD then
                  nCurSel = ubound(gTabPages)
               end if
            END IF
            gTabPages(nCurSel).wszText      = "Tab " & nNextTabNum
            gTabPages(nCurSel).wszImage     = ""
            gTabPages(nCurSel).wszTabPage   = ""
            gTabPages(nCurSel).IsActiveTab  = 0
            gTabPages(nCurSel).wszReserved1 = ""
            gTabPages(nCurSel).wszReserved2 = ""
            gTabPages(nCurSel).wszReserved3 = ""
            gTabPages(nCurSel).wszReserved4 = ""
            ' reload the listbox
            frmVDTabChild_LoadBuildListBox(HWND)
            nCurSel = Min(nCurSel, ubound(gTabPages))
            ListBox_SetCurSel(hList1, nCurSel)
            frmVDTabChild_SetTabChildTextboxes()
            dim as HWND hEdit = GetDlgItem( HWND_frmVDTabChild, IDC_FRMVDTABCHILD_TXTDESCRIPTION)
            Edit_SetSel( hEdit, 0, -1 )
            SetFocus( hEdit )
         end if

      case IDC_FRMVDTABCHILD_CMDDELETE
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1  THEN
               if MessageBox( HWND, L(366, "Are you sure you want to delete?"), L(276,"Confirm"), _
                                 MB_YESNOCANCEL Or MB_ICONINFORMATION Or MB_DEFBUTTON1 ) = IDYES then
                  if ubound(gTabPages) = 0 THEN
                     erase gTabPages
                     nCurSel = -1
                  else
                     ' remove the item from the internal array
                     for i as long = nCurSel to ubound(gTabPages) - 1
                        gTabPages(i) = gTabPages(i+1)
                     NEXT
                     redim preserve gTabPages(ubound(gTabPages)-1)
                  END IF
                  ' reload the listbox
                  frmVDTabChild_LoadBuildListBox(HWND)
                  nCurSel = Min(nCurSel, ubound(gTabPages))
                  ListBox_SetCurSel(hList1, nCurSel)
                  frmVDTabChild_SetTabChildTextboxes()
                  SetFocus hList1
               end if
            END IF
         end if

      Case IDC_FRMVDTABCHILD_CMDOK
         If codeNotify = BN_CLICKED Then
            ' Copy the TabPages back into the PropValue 
            dim pProp as clsProperty ptr = GetActivePropertyPtr()
            if pProp then
               pProp->wszPropValue = ""
               for i as long = 0 to ubound(gTabPages)
                  pProp->wszPropValue = pProp->wszPropValue & _
                     gTabPages(i).wszText      & gTabFieldSep & _
                     gTabPages(i).wszImage     & gTabFieldSep & _
                     gTabPages(i).wszTabPage   & gTabFieldSep & _
                     gTabPages(i).IsActiveTab  & gTabFieldSep & _
                     gTabPages(i).wszReserved1 & gTabFieldSep & _
                     gTabPages(i).wszReserved2 & gTabFieldSep & _
                     gTabPages(i).wszReserved3 & gTabFieldSep & _
                     gTabPages(i).wszReserved4 & gTabRecordSep
               next
               pProp->wszPropValue = trim(pProp->wszPropValue, gTabRecordSep)            
               dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
               if pDoc then 
                  pDoc->UserModified = true
                  dim pCtrl as clsControl ptr = pDoc->Controls.GetActiveControl
                  if pCtrl THEN
                     ' Recreate the control
                     ReCreateToolboxControl( pDoc, pCtrl )
                  end if
               end if      
            end if
             
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         end if
             
      Case IDC_FRMVDTABCHILD_CMDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmVDTabChild
' ========================================================================================
private Function frmVDTabChild_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmVDTabChild
' ========================================================================================
private Function frmVDTabChild_OnDestroy( byval HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmVDTabChild Window procedure
' ========================================================================================
private Function frmVDTabChild_WndProc( ByVal HWnd   As HWnd, _
                                        ByVal uMsg   As UINT, _
                                        ByVal wParam As WPARAM, _
                                        ByVal lParam As LPARAM _
                                        ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,      frmVDTabChild_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmVDTabChild_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmVDTabChild_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,     frmVDTabChild_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmVDTabChild_Show
' ========================================================================================
public Function frmVDTabChild_Show( ByVal hWndParent As HWnd, _
                                    byref wszPropValue as wstring _
                                    ) As LRESULT

   DIM hBitmap AS HBITMAP 
   dim hCtrl as HWnd
   dim wszImage as wstring * 100

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create(hWndParent, L(394,"TabControl Configuration"), _
        @frmVDTabChild_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(408, 360)
   pWindow->Center(pWindow->hWindow, hWndParent)

   pWindow->AddControl("LISTBOX", , IDC_FRMVDTABCHILD_LIST1, "", 21, 14, 365, 84, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMVDTABCHILD_CMDUP, "", 21, 102, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWUP", "IMAGE_ARROWUP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMVDTABCHILD_CMDDOWN, "", 48, 102, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_LOADMAP3DCOLORS Or LR_LOADTRANSPARENT Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("BUTTON", , IDC_FRMVDTABCHILD_CMDADD, L(380, "Add"), 77, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMVDTABCHILD_CMDINSERT, L(281, "Insert"), 152, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMVDTABCHILD_CMDDELETE, L(282, "Delete"), 227, 102, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_FLAT Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMVDTABCHILD_LABEL1, L(150,"Text") & ":", 20, 145, 218, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMVDTABCHILD_TXTDESCRIPTION, "", 20, 165, 365, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMVDTABCHILD_LABEL2, L(246,"Image") & ":", 20, 195, 218, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMVDTABCHILD_TXTIMAGE, "", 20, 215, 218, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMVDTABCHILD_CMDIMAGE, "...", 250, 214, 30, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMVDTABCHILD_LABEL2, L(395,"TabPage") & ":", 20, 245, 218, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMVDTABCHILD_COMBOTABPAGES, "", 20, 265, 218, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST or CBS_SORT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("CHECKBOX", , IDC_FRMVDTABCHILD_CHKISDEFAULT, L(280,"Set as default"), 250, 265, 240, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)


   pWindow->AddControl("BUTTON", , IDC_FRMVDTABCHILD_CMDOK, L(0,"OK"), 231, 320, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMVDTABCHILD_CMDCANCEL, L(1,"Cancel"), 312, 320, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   HWND_FRMVDTABCHILD = hForm
   

   ' Load the list of forms that have the FormChild = True
   dim as HWND hCombo = GetDlgItem(hForm, IDC_FRMVDTABCHILD_COMBOTABPAGES)
   dim as CWSTR wszFormName = "(" & L(396,"No Child Form Selected") & ")"
   ComboBox_AddString( hCombo, wszFormName.sptr )
   dim pDoc as clsDocument ptr = gApp.pDocList
   dim pCtrl as clsControl ptr
   do until pDoc = 0
      pCtrl = GetFormCtrlPtr(pDoc)
      if pCtrl then
         if GetControlProperty(pCtrl, "CHILDFORM") = "True" then
            wszFormName = GetControlProperty(pCtrl, "NAME")
            ComboBox_AddString( hCombo, wszFormName.sptr )
         end if
      end if
      pDoc = pDoc->pDocNext
   loop
   
   
   ' Copy all of the custom property values to the gTabPages array. We will work with
   ' temporary copies until the user hits OK.
   dim as HWND hList = GetDlgItem(hForm, IDC_FRMVDTABCHILD_LIST1)
   frmVDTabChild_LoadTabPagesArray(wszPropValue)
   frmVDTabChild_LoadBuildListBox(hForm)
   
   ListBox_SetCurSel( hList, 0 )
   frmVDTabChild_SetTabChildTextboxes   
   SetFocus hList
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the CWindow class manually allocated memory 
   Delete pWindow

   ' Reset the gTabPages global array
   erase gTabPages
   
End Function





