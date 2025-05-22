'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2025 Paul Squires, PlanetSquires Software
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

#include once "frmMenuEditor.bi"
#include once "clsMenuItem.bi"
#include once "clsDocument.bi"


' Temporary MenuItem array to hold items while they are being
' edited in the Menu Editor. 
dim shared gMenuItems(any) as clsMenuItem


' ========================================================================================
' Create the "fake" topmenu to display on the Form
' ========================================================================================
function frmMenuEditor_CreateFakeMainMenu( ByVal pDoc as clsDocument ptr ) As Long

   ' Rather than create a bunch of individual labels, create the one string
   ' and then only one label. If the label is clicked, then display the Menu Editor.
   ' Only create the mainmenu if menuitems actually exist.
   if pDoc->MainMenuExists = false then
      ShowWindow(pDoc->hWndFakeMenu, SW_HIDE)
      exit function
   end if
   
   dim as CWSTR wszMenu = wspace(2)
   for i as long = lbound(pDoc->MenuItems) to ubound(pDoc->MenuItems)
      if pDoc->MenuItems(i).nIndent = 0 then
         wszMenu = wszMenu & pDoc->MenuItems(i).wszCaption & wspace(4)
      end if
   next
   AfxSetWindowText( pDoc->hWndFakeMenu, wszMenu )
   
   dim ncm As NONCLIENTMETRICS
   ncm.cbSize = SizeOf(ncm)
   SystemParametersInfo(SPI_GETNONCLIENTMETRICS, SizeOf(ncm), @ncm, 0)
   ' ncm.iMenuHeight  ' holds the height of the mainmenu

   dim rc as RECT
   GetClientRect( pDoc->hWndForm, @rc)
   SetWindowPos( pDoc->hWndFakeMenu, 0, 0, 0, rc.right-rc.left, ncm.iMenuHeight, SWP_NOZORDER or SWP_SHOWWINDOW)
   
       
   function = 0 
End Function
                  

' ========================================================================================
' Display menu information for the specified line.
' ========================================================================================
function frmMenuEditor_ListBoxLine( byval nCurSel as long ) as CWSTR
   dim as long nBullet, nDash
   IF AfxWindowsVersion >= 600 AND CLNG(AfxIsProcessDpiAware) THEN 
      ' Segoe UI
      nBullet = 183: nDash = 8211 
   ELSE 
      ' Tahoma
      nBullet = 183: nDash = 8211 
   end if

   dim as CWSTR wszLine = wstring(gMenuItems(nCurSel).nIndent * 4, nBullet)
   dim as long nShow = SW_SHOW
   if trim(gMenuItems(nCurSel).wszCaption) = "-" then
      wszLine = wszLine & " " & wstring(4, nDash)  
      nShow = SW_HIDE
   else   
      wszLine = wszLine & gMenuItems(nCurSel).wszCaption
   end if
      
   ShowWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_LBLSHORTCUT), nShow )
   ShowWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_LBLSTATE), nShow )
   ShowWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CMBOSHORTCUT), nShow )
   ShowWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKCTRL), nShow )
   ShowWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKALT), nShow )
   ShowWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKSHIFT), nShow )
   ShowWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKCHECKED), nShow )
   ShowWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKGRAYED), nShow )

   function = wszLine
end function


' ========================================================================================
' Determine if a blank entry should be added to the end of the listbox.
' ========================================================================================
function frmMenuEditor_EnsureLastLine() As Long

   dim as HWND hListBox = GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_LSTDETAIL )
   dim as long NumItems = Listbox_GetCount(hListBox)

   If NumItems > 0 Then
      ' Check the Caption and Name values for the last menu item. If
      ' no value exists then no need to add a blank entry at the end.
      If ( gMenuItems(NumItems-1).wszName = "" ) AndAlso ( gMenuItems(NumItems-1).wszCaption = "" ) Then
         Exit Function
      End If   
   End If

   dim wszBlankLine as wstring * 10
   ListBox_AddString( hListBox, @wszBlankLine )
   redim preserve gMenuItems(NumItems + 1)
   
   function = 0
End Function


' ========================================================================================
' Display the details of this MenuItem
' ========================================================================================
function frmMenuEditor_DisplayMenuItem() As Long

   dim as long nCurSel 
   dim as Boolean bEnabled = true
   
   nCurSel    = Listbox_GetCurSel( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_LSTDETAIL) )
   If nCurSel = -1 Then Exit Function
   
   ' If this is a TopLevel Menu Item then disable the Shortcut key
   If gMenuItems(nCurSel).nIndent = 0 Then
      gMenuItems(nCurSel).wszShortcut = ""   ' this will display (None)
      gMenuItems(nCurSel).chkAlt      = BST_UNCHECKED
      gMenuItems(nCurSel).chkShift    = BST_UNCHECKED
      gMenuItems(nCurSel).chkCtrl     = BST_UNCHECKED
      gMenuItems(nCurSel).chkChecked  = BST_UNCHECKED
      bEnabled = false
   end if
      
   AfxSetWindowText( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_TXTCAPTION), gMenuItems(nCurSel).wszCaption)
   AfxSetWindowText( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_TXTNAME), gMenuItems(nCurSel).wszName)

   ' Shortcut
   dim as long nIndex = ComboBox_FindStringExact( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CMBOSHORTCUT), -1, gMenuItems(nCurSel).wszShortcut.sptr ) 
   ComboBox_SetCurSel( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CMBOSHORTCUT), nIndex)
   EnableWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CMBOSHORTCUT), bEnabled )
   Button_SetCheck( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKCTRL), gMenuItems(nCurSel).chkCtrl)
   EnableWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKCTRL), bEnabled )
   Button_SetCheck( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKALT), gMenuItems(nCurSel).chkAlt)
   EnableWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKALT), bEnabled )
   Button_SetCheck( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKSHIFT), gMenuItems(nCurSel).chkShift)
   EnableWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKSHIFT), bEnabled )

   ' Checked
   Button_SetCheck( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKCHECKED), gMenuItems(nCurSel).chkChecked )
   EnableWindow( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKCHECKED), bEnabled )
   
   ' Grayed
   Button_SetCheck( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKGRAYED), gMenuItems(nCurSel).chkGrayed )

   SetFocus( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_TXTCAPTION) )
   SendMessage GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_TXTCAPTION), EM_SETSEL, 0, cast(LPARAM, -1)

   function = 0
End Function


' ========================================================================================
' Generate a default MenuItem Name based on the MenuItem Caption
' ========================================================================================
function frmMenuEditor_GenerateMenuItemName( byval nCurSel as long ) As CWSTR
   
   dim as CWSTR wszCaption = gMenuItems(nCurSel).wszCaption
   dim as CWSTR wszName
   dim as Boolean bFound
   dim as long nextSeperator
      
   If Trim(wszCaption) = "-" Then
      ' Determine what separator number has not been used yet.
      Do                              
         bFound = false  ' default that it doesn't exist
         nextSeperator = nextSeperator + 1
         wszName = "mnuSep" & nextSeperator
         For i as long = 0 To ubound(gMenuItems) - 1
            If UCase(gMenuItems(i).wszName) = UCase(wszName) Then continue do
         Next                       
         if bFound = false then exit do
      Loop 
   Else
      ' The name can only be alphanumeric
      static wszRetain as CWSTR = "abcdefghijklmnopqrstuvwxyz0123456789"
      wszName = "mnu" & AfxStrRetainAnyI(wszCaption, wszRetain)
   End If
         
   function = wszName
end function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmMenuEditor
' ========================================================================================
function frmMenuEditor_OnCreate( _
            ByVal HWnd As HWnd, _
            ByVal lpCreateStructPtr As LPCREATESTRUCT _
            ) as boolean

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmMenuEditor
' ========================================================================================
function frmMenuEditor_OnCommand( _
            ByVal HWnd As HWnd, _
            ByVal id As Long, _
            ByVal hwndCtl As HWnd, _
            ByVal codeNotify As UINT _
            ) As LRESULT

   dim as HWND hListBox = GetDlgItem( HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_LSTDETAIL)
   dim as long nCurSel = ListBox_GetCurSel(hListBox)
   
   Select Case id
      Case IDC_FRMMENUEDITOR_CMDOK
         If codeNotify = BN_CLICKED Then
            ' Do error checking to ensure that the menu structure is valid
            for i as long = 0 to ubound(gMenuItems) - 1
               ' If a caption exists but does not have a name then present the error.
               If ( Len(gMenuItems(i).wszCaption) > 0 ) andalso ( Len(Trim(gMenuItems(i).wszName)) = 0 ) Then
                  ListBox_SetCurSel(hListBox, i)
                  frmMenuEditor_DisplayMenuItem
                  MessageBox( HWND, L(357,"A Name must exist for menu items with a Caption."), _
                              L(109,"Warning"), MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 Or MB_APPLMODAL )
                  exit function
               End If
            next

            ' Do error checking for duplicate menu names
            for i as long = 0 to ubound(gMenuItems) - 1
               ' If a caption exists but does not have a name then present the error.
               for ii as long = 0 to ubound(gMenuItems) - 1
                  if ii = i then continue for
                  If ucase(trim(gMenuItems(ii).wszName)) = ucase(trim(gMenuItems(i).wszName))  Then
                     ListBox_SetCurSel(hListBox, ii)
                     frmMenuEditor_DisplayMenuItem
                     MessageBox( HWND, L(367,"Duplicate menu name.") & vbcrlf & gMenuItems(i).wszName, _
                                 L(109,"Warning"), MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 Or MB_APPLMODAL )
                     exit function
                  end if
               next
            next

            ' Make sure the indent levels are consistent.
            for i as long = 1 to ubound(gMenuItems) - 1
               if gMenuItems(i).nIndent > gMenuItems(i-1).nIndent + 1 then
                  ListBox_SetCurSel(hListBox, i)
                  frmMenuEditor_DisplayMenuItem
                  MessageBox( HWND, L(358,"Menu item has skipped a level."), _
                              L(109,"Warning"), MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 Or MB_APPLMODAL )
                  exit function
               end if
            next

            ' Copy the temporary array back to the form's menu items 
            Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
            If pDoc Then 
               ' Copy all menuitems except for the last one because that is
               ' the last blank line in the listbox.
               dim as long numItems = ubound(gMenuItems) - 1
               if numItems > -1 then
                  redim pDoc->MenuItems(numItems)
                  for i as long = 0 to numItems
                     pDoc->MenuItems(i) = gMenuItems(i) 
                  next
               end if

               ' Save the value indicating that we want the menu to display in
               ' the generated code for the form.
               pDoc->GenerateMenu = _
                  iif(Button_GetCheck(GetDlgItem(HWND, IDC_FRMMENUEDITOR_CHKDISPLAYONFORM)) = BST_CHECKED, true, false)

               ' If there are no valid MenuItems then set the GenerateMenu to False.
               if numItems < 0 then pDoc->GenerateMenu = false
               
               pDoc->UserModified = true
            end if
            
            SendMessage(HWnd, WM_CLOSE, 0, 0)
            Exit Function
         End If
         
      Case IDC_FRMMENUEDITOR_CMDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage(HWnd, WM_CLOSE, 0, 0)
            Exit Function
         End If
   
      case IDC_FRMMENUEDITOR_LSTDETAIL
         If codeNotify = LBN_SELCHANGE Then
            frmMenuEditor_DisplayMenuItem
         end if

      case IDC_FRMMENUEDITOR_TXTCAPTION
         If codeNotify = EN_CHANGE Then
            if nCurSel = -1 then exit function
            ' Update the listbox with the new caption
            gMenuItems(nCurSel).wszCaption = AfxGetWindowText(hwndCtl)
            dim as CWSTR wszCaption = frmMenuEditor_ListBoxLine(nCurSel)
            ListBox_ReplaceString( GetDlgItem(HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_LSTDETAIL), nCurSel, wszCaption) 
            frmMenuEditor_EnsureLastLine
         end if
         
      case IDC_FRMMENUEDITOR_TXTNAME
         if nCurSel = -1 then exit function
         If codeNotify = EN_CHANGE Then
            gMenuItems(nCurSel).wszName = AfxGetWindowText(hwndCtl)
            frmMenuEditor_EnsureLastLine
         elseIf codeNotify = EN_SETFOCUS Then
            if len(gMenuItems(nCurSel).wszName) = 0 then
               gMenuItems(nCurSel).wszName = frmMenuEditor_GenerateMenuItemName(nCurSel)
               AfxSetWindowText(hwndCtl, gMenuItems(nCurSel).wszName)
               SendMessage( hwndCtl, EM_SETSEL, 0, cast(LPARAM, -1) )
            end if
         end if
         
      Case IDC_FRMMENUEDITOR_CMDNEXT
         If codeNotify = BN_CLICKED Then
            frmMenuEditor_EnsureLastLine
            ' move down one selection or wrap to start
            nCurSel = nCurSel + 1
            If nCurSel > Listbox_GetCount(hListBox) - 1 then nCurSel = 0
            Listbox_SetCurSel(hListBox, nCurSel )
            frmMenuEditor_DisplayMenuItem
         End If

      Case IDC_FRMMENUEDITOR_CHKDISPLAYONFORM
         If codeNotify = BN_CLICKED Then
            ' Don't need to store the value. We will check it when the 
            ' Menu Editor is closed.
         End If

      Case IDC_FRMMENUEDITOR_CHKALT
         if nCurSel = -1 then exit function
         If codeNotify = BN_CLICKED Then
            gMenuItems(nCurSel).chkAlt = Button_GetCheck(hwndCtl)
         End If

      Case IDC_FRMMENUEDITOR_CHKSHIFT
         if nCurSel = -1 then exit function
         If codeNotify = BN_CLICKED Then
            gMenuItems(nCurSel).chkShift = Button_GetCheck(hwndCtl)
         End If

      Case IDC_FRMMENUEDITOR_CHKCTRL
         if nCurSel = -1 then exit function
         If codeNotify = BN_CLICKED Then
            gMenuItems(nCurSel).chkCtrl = Button_GetCheck(hwndCtl)
         End If

      Case IDC_FRMMENUEDITOR_CHKCHECKED
         if nCurSel = -1 then exit function
         If codeNotify = BN_CLICKED Then
            gMenuItems(nCurSel).chkChecked = Button_GetCheck(hwndCtl)
         End If

      Case IDC_FRMMENUEDITOR_CHKGRAYED
         if nCurSel = -1 then exit function
         If codeNotify = BN_CLICKED Then
            gMenuItems(nCurSel).chkGrayed = Button_GetCheck(hwndCtl)
         End If

      Case IDC_FRMMENUEDITOR_CMBOSHORTCUT
         if nCurSel = -1 then exit function
         If codeNotify = CBN_SELCHANGE Then
            gMenuItems(nCurSel).wszShortcut = AfxGetWindowText(hwndCtl)
         End If

      Case IDC_FRMMENUEDITOR_CMDLEFT
         If codeNotify = BN_CLICKED Then
            if nCurSel = -1 then exit function  
            ' Can not move less than 0 position
            if gMenuItems(nCurSel).nIndent - 1 < 0 then exit function
            gMenuItems(nCurSel).nIndent = gMenuItems(nCurSel).nIndent - 1
            frmMenuEditor_DisplayMenuItem
         End If

      Case IDC_FRMMENUEDITOR_CMDRIGHT
         If codeNotify = BN_CLICKED Then
            ' You can only move the item one place to the right of the item's parent
            if nCurSel <= 0 then exit function   ' can not move the first item in the list
            dim as long nPrevIndent = gMenuItems(nCurSel-1).nIndent
            dim as long nCurIndent = gMenuItems(nCurSel).nIndent
            if nCurIndent + 1 > nPrevIndent + 1 then exit function
            gMenuItems(nCurSel).nIndent = nCurIndent + 1
            frmMenuEditor_DisplayMenuItem
         End If

      Case IDC_FRMMENUEDITOR_CMDUP
         If codeNotify = BN_CLICKED Then
            ' You can only move the item one place to the right of the item's parent
            if nCurSel <= 0 then exit function   ' can not move the first item in the list
            ' If the MenuItem was moved to position 0 then ensure that the indent is 0.
            if nCurSel = 1 then gMenuItems(nCurSel).nIndent = 0
            swap gMenuItems(nCurSel), gMenuItems(nCurSel-1)
            ' Display the two lines
            ListBox_ReplaceString(hListBox, nCurSel-1, frmMenuEditor_ListBoxLine(nCurSel-1))
            ListBox_ReplaceString(hListBox, nCurSel, frmMenuEditor_ListBoxLine(nCurSel))
            Listbox_SetCurSel( hListBox, nCurSel - 1)
            frmMenuEditor_DisplayMenuItem
         End If
         
      Case IDC_FRMMENUEDITOR_CMDDOWN
         If codeNotify = BN_CLICKED Then
            ' You can only move the item one place to the right of the item's parent
            if nCurSel = -1 then exit function 
            ' If the MenuItem was moved from position 0 then ensure that the indent is 0.
            if nCurSel = 0 then gMenuItems(nCurSel+1).nIndent = 0
            swap gMenuItems(nCurSel), gMenuItems(nCurSel+1)
            ' Display the two lines
            ListBox_ReplaceString(hListBox, nCurSel, frmMenuEditor_ListBoxLine(nCurSel))
            ListBox_ReplaceString(hListBox, nCurSel+1, frmMenuEditor_ListBoxLine(nCurSel+1))
            Listbox_SetCurSel( hListBox, nCurSel + 1)
            frmMenuEditor_DisplayMenuItem
         End If

      Case IDC_FRMMENUEDITOR_CMDDELETE
         If codeNotify = BN_CLICKED Then
            if nCurSel = -1 then exit function 
            for i as long = nCurSel to ubound(gMenuItems) - 1
               gMenuItems(i) = gMenuItems(i+1) 
            next
            Redim Preserve gMenuItems(Ubound(gMenuItems)-1)
            Listbox_DeleteString(hListBox, nCurSel)
            ' If position 0 is deleted then ensure the item replacing it is indent 0.
            if nCurSel = 0 then 
               gMenuItems(0).nIndent = 0
               ListBox_ReplaceString(hListBox, 0, frmMenuEditor_ListBoxLine(0))
            end if
            frmMenuEditor_EnsureLastLine
            dim as long NumItems = Listbox_GetCount(hListBox)
            Listbox_SetCurSel( hListBox, Min(nCurSel, NumItems-1) )
            frmMenuEditor_DisplayMenuItem
         end if

      Case IDC_FRMMENUEDITOR_CMDINSERT
         If codeNotify = BN_CLICKED Then
            if nCurSel = -1 then exit function 
            Redim Preserve gMenuItems(Ubound(gMenuItems)+1)
            For i as long = Ubound(gMenuItems) To nCurSel + 1 Step -1
               gMenuItems(i) = gMenuItems(i-1) 
            Next
            dim newMenuItem as clsMenuItem
            gMenuItems(nCurSel) = newMenuItem
            Listbox_InsertString(hListBox, nCurSel, gMenuItems(nCurSel).wszCaption.sptr)
            frmMenuEditor_EnsureLastLine
            dim as long NumItems = Listbox_GetCount(hListBox)
            Listbox_SetCurSel( hListBox, nCurSel )
            frmMenuEditor_DisplayMenuItem
         end if

   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmMenuEditor
' ========================================================================================
function frmMenuEditor_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Reset the global MenuItems array
   erase gMenuItems
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow HWnd
   Function = 0
end function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmMenuEditor
' ========================================================================================
function frmMenuEditor_OnDestroy( byval HWnd As HWnd ) As LRESULT
   PostQuitMessage(0)
   Function = 0
end function


' ========================================================================================
' frmMenuEditor Window procedure
' ========================================================================================
function frmMenuEditor_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT
   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmMenuEditor_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmMenuEditor_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmMenuEditor_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmMenuEditor_OnCommand)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)
end function


' ========================================================================================
' frmMenuEditor_Show
' ========================================================================================
function frmMenuEditor_Show( ByVal hWndParent As HWnd ) as LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   DIM hBitmap AS HBITMAP 
   dim hCtrl as HWnd
   dim wszImage as wstring * 100
   dim wszString  as WSTRING * 100
   
   HWND_FRMMENUEDITOR = _
   pWindow->Create(hWndParent, L(312,"Menu Editor"), @frmMenuEditor_WndProc, 0, 0, 411, 438, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->Center(pWindow->hWindow, hWndParent)

   pWindow->AddControl("LABEL", , IDC_FRMMENUEDITOR_LABEL3, L(150,"Text") & ":", 10, 16, 62, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMMENUEDITOR_LABEL1, L(364,"Name") & ":", 10, 40, 62, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMMENUEDITOR_TXTCAPTION, "", 74, 13, 233, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("TEXTBOX", , IDC_FRMMENUEDITOR_TXTNAME, "", 74, 37, 233, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMMENUEDITOR_LBLSHORTCUT, L(363,"Shortcut") & ":", 10, 65, 54, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMMENUEDITOR_CHKCTRL, "Ctrl", 74, 65, 48, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMMENUEDITOR_CHKALT, "Alt", 124, 65, 42, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMMENUEDITOR_CHKSHIFT, "Shift", 168, 65, 48, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMMENUEDITOR_CMBOSHORTCUT, "", 240, 65, 67, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMMENUEDITOR_LBLSTATE, L(360,"State") & ":", 10, 90, 63, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMMENUEDITOR_CHKCHECKED, L(361,"Checked"), 74, 91, 175, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMMENUEDITOR_CHKGRAYED, L(362,"Grayed"), 74, 110, 175, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMMENUEDITOR_CMDLEFT, "", 17, 141, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWLEFT", "IMAGE_ARROWLEFT16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMMENUEDITOR_CMDRIGHT, "", 43, 141, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWRIGHT", "IMAGE_ARROWRIGHT16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMMENUEDITOR_CMDUP, "", 69, 141, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWUP", "IMAGE_ARROWUP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMMENUEDITOR_CMDDOWN, "", 95, 141, 24, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("BUTTON", , IDC_FRMMENUEDITOR_CMDNEXT, L(359,"&Next"), 157, 141, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMMENUEDITOR_CMDINSERT, L(281,"&Insert"), 237, 141, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMMENUEDITOR_CMDDELETE, L(282,"&Delete"), 317, 141, 70, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LISTBOX", , IDC_FRMMENUEDITOR_LSTDETAIL, "", 13, 173, 384, 200, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMMENUEDITOR_CMDOK, L(0,"&OK"), 317, 13, 80, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMMENUEDITOR_CMDCANCEL, L(1,"&Cancel"), 317, 42, 80, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("CHECKBOX", , IDC_FRMMENUEDITOR_CHKDISPLAYONFORM, L(237,"Display on Form"), 13, 383, 175, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   hCtrl = GetDlgItem( HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CMBOSHORTCUT)
   ComboBox_ResetContent(hCtrl)

   ' Have a blank string at the beginning of the list so that an accelerator can be unselected
   wszString = ""
   ComboBox_AddString(hCtrl, @wszString)
   
   For i as long = 1 To 26  ' A-Z
      wszString = wchr(i + 64)
      ComboBox_AddString(hCtrl, @wszString)
   Next 
   For i as long = 0 To 9   ' 0-9
      wszString = wchr(i + 48)
      ComboBox_AddString(hCtrl, @wszString)
   Next 
   For i as long = 1 To 12  ' F1 to F12
      wszString = "F" & i
      ComboBox_AddString(hCtrl, @wszString)
   Next 

   ComboBox_AddString(hCtrl, @wstr("Ins"))
   ComboBox_AddString(hCtrl, @wstr("Del"))
   ComboBox_AddString(hCtrl, @wstr("TAB"))
   ComboBox_AddString(hCtrl, @wstr("Left"))
   ComboBox_AddString(hCtrl, @wstr("Right"))
   ComboBox_AddString(hCtrl, @wstr("Up"))
   ComboBox_AddString(hCtrl, @wstr("Down"))
   ComboBox_AddString(hCtrl, @wstr("Home"))
   ComboBox_AddString(hCtrl, @wstr("End"))
   ComboBox_AddString(hCtrl, @wstr("PgUp"))
   ComboBox_AddString(hCtrl, @wstr("PgDn"))
   ComboBox_AddString(hCtrl, @wstr("Space"))
   ComboBox_AddString(hCtrl, @wstr("Plus"))
   ComboBox_AddString(hCtrl, @wstr("Minus"))
   ComboBox_AddString(hCtrl, @wstr("Enter"))
    
   ComboBox_SetCurSel(hCtrl, 0)


   ' Copy the form's menu items to the temporary array for editing
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then
      hCtrl = GetDlgItem( HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_CHKDISPLAYONFORM )
      Button_SetCheck( hCtrl, iif(pDoc->GenerateMenu, BST_CHECKED, BST_UNCHECKED))
      dim as long numItems = ubound(pDoc->MenuItems)
      if numItems > -1 then
         redim gMenuItems(numItems)
         for i as long = 0 to numItems
            gMenuItems(i) = pDoc->MenuItems(i)
         next
      end if
   end if
   
   ' Display the menuitems in the listbox
   hCtrl = GetDlgItem( HWND_FRMMENUEDITOR, IDC_FRMMENUEDITOR_LSTDETAIL)
   for i as long = 0 to ubound(gMenuItems)
      ListBox_AddString(hCtrl, frmMenuEditor_ListBoxLine(i).sptr)
   next
   frmMenuEditor_EnsureLastLine
   ListBox_SetCurSel(hCtrl, 0)
   frmMenuEditor_DisplayMenuItem
   
   ' Process Windows messages
   Function = pWindow->DoEvents( SW_SHOW )
   
   ' Delete the frmMenuEditor CWindow class manually allocated memory 
   Delete pWindow
   
   ' Erase the temporary global MenuItems array
   erase gMenuItems

end function





