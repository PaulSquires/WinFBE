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
' Attempt to insert/expand a user snippet
' ========================================================================================
function DoInsertSnippet( byval pDoc as clsDocument ptr ) as Boolean
   if pDoc = 0 then exit function
   
   Dim As HWnd hEdit = pDoc->hWndActiveScintilla
   Dim As Long curPos = SendMessage( hEdit, SCI_GETCURRENTPOS, 0, 0)

   ' Get the current word at the cursor
   dim as string strWord = ucase( pDoc->GetWord )
   dim as string strReplaceText
   
   ' Search the snippets array for a matching trigger
   for i as long = lbound(gConfig.Snippets) to ubound(gConfig.Snippets)
      if ucase( gConfig.Snippets(i).wszTrigger ) = strWord then
         ' Insert the snippet by replacing the current word
         SendMessage( hEdit, SCI_SETSELECTIONSTART, curPos - len(strWord), 0)
         SendMessage( hEdit, SCI_SETSELECTIONEND, curPos, 0)
         ' Replace the selection (SCI_REPLACESEL fails if text is "" so use Cut instead for that scenario)                                
         If Len(strWord) = 0 Then
            SendMessage( hEdit, SCI_CUT, 0, 0 )
         Else
            strReplaceText = gConfig.Snippets(i).wszCode
            SendMessage( hEdit, SCI_REPLACESEL, 0, Cast(LPARAM, Strptr(strReplaceText)) )
         End If
         return true
      end if
   next
   
   return false
end function


' ========================================================================================
' Load all of the snippets descriptions into the listbox
' ========================================================================================
function LoadSnippetsListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMSNIPPETS_LIST1)

   ListBox_ResetContent(hList1)
   for i as long = lbound(gConfig.SnippetsTemp) to ubound(gConfig.SnippetsTemp)
      ListBox_AddString(hList1, gConfig.SnippetsTemp(i).wszDescription)
   NEXT

   function = 0
end function


' ========================================================================================
' Swap two entries in the snippets Listbox
' ========================================================================================
function SwapSnippetsListBoxItems( byval Item1 as long, Byval Item2 as long) as Long
   dim as hwnd hList1 = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_LIST1)
   
   ' Swap the array values
   swap gConfig.SnippetsTemp(Item1), gConfig.SnippetsTemp(Item2)
   
   ListBox_ReplaceString(hList1, Item1, gConfig.SnippetsTemp(Item1).wszDescription)
   ListBox_ReplaceString(hList1, Item2, gConfig.SnippetsTemp(Item2).wszDescription)

   function = 0
end function


' ========================================================================================
' Set the snippets information depending on what listbox entry is selected
' ========================================================================================
function SetSnippetsTextboxes() as long
   dim as hwnd hList1 = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_LIST1)
   dim as hwnd hText1 = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTDESCRIPTION)
   dim as hwnd hText2 = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTTRIGGER)
   dim as hwnd hText3 = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE)
   
   dim as long nCurSel = ListBox_GetCurSel(hList1)
   if nCurSel < 0 THEN 
      AfxSetWindowText( hText1, "")
      AfxSetWindowText( hText2, "")
      AfxSetWindowText( hText3, "")
      EnableWindow(hText1, false)
      EnableWindow(hText2, false)
      EnableWindow(hText3, false)
   else
      AfxSetWindowText( hText1, gConfig.SnippetsTemp(nCurSel).wszDescription)
      AfxSetWindowText( hText2, gConfig.SnippetsTemp(nCurSel).wszTrigger)
      AfxSetWindowText( hText3, gConfig.SnippetsTemp(nCurSel).wszCode)
      EnableWindow(hText1, true)
      EnableWindow(hText2, true)
      EnableWindow(hText3, true)
   end if

   function = 0
end function
               

' ========================================================================================
' Process WM_CREATE message for window/dialog: frmSnippets
' ========================================================================================
Function frmSnippets_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmSnippets
' ========================================================================================
Function frmSnippets_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMSNIPPETS_LIST1)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   ' Convert a ListBox DoubleClick into pressing OK button action
   if (id = IDC_FRMSNIPPETS_LIST1) AndAlso (codeNotify = LBN_DBLCLK) THEN
      id = IDOK: codeNotify = BN_CLICKED
   END IF
   
   Select Case id
      case IDC_FRMSNIPPETS_LIST1
         if codeNotify = LBN_SELCHANGE THEN
            SetSnippetsTextboxes()
         END IF
          
      case IDC_FRMSNIPPETS_TXTDESCRIPTION
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array and the Listbox
            if nCurSel > -1 THEN
               gConfig.SnippetsTemp(nCurSel).wszDescription = AfxGetWindowText(hwndCtl)
               ListBox_ReplaceString(hList1, nCurSel, gConfig.SnippetsTemp(nCurSel).wszDescription)
            END IF
         end if
         
      case IDC_FRMSNIPPETS_TXTTRIGGER
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array
            if nCurSel > -1 THEN
               gConfig.SnippetsTemp(nCurSel).wszTrigger = AfxGetWindowText(hwndCtl)
            END IF
         end if

      case IDC_FRMSNIPPETS_TXTCODE
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array
            if nCurSel > -1 THEN
               gConfig.SnippetsTemp(nCurSel).wszCode = AfxGetWindowText(hwndCtl)
            END IF
         end if

      case IDC_FRMSNIPPETS_CMDUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               SwapSnippetsListboxItems( nCurSel, nCurSel - 1 )
            END IF
         end if

      case IDC_FRMSNIPPETS_CMDDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1) - 1 THEN
               SwapSnippetsListboxItems( nCurSel, nCurSel + 1 )
            END IF
         end if

      case IDC_FRMSNIPPETS_CMDINSERT
         if codeNotify = BN_CLICKED THEN
            if ubound(gConfig.SnippetsTemp) = -1 THEN
               redim gConfig.SnippetsTemp(0)
               nCurSel = 0
            Else
               redim preserve gConfig.SnippetsTemp(ubound(gConfig.SnippetsTemp)+1)
               if nCurSel = -1 THEN nCurSel = 0
               ' insert the item above current entry in the internal array
               for i as long = ubound(gConfig.SnippetsTemp) to nCurSel + 1 step -1
                  gConfig.SnippetsTemp(i) = gConfig.SnippetsTemp(i-1)
               NEXT
            END IF
            gConfig.SnippetsTemp(nCurSel).wszDescription = ""
            gConfig.SnippetsTemp(nCurSel).wszTrigger = ""
            gConfig.SnippetsTemp(nCurSel).wszCode = ""
            ' reload the listbox
            LoadSnippetsListBox(HWND)
            nCurSel = Min(nCurSel, ubound(gConfig.SnippetsTemp))
            ListBox_SetCurSel(hList1, nCurSel)
            SetSnippetsTextboxes()
            SetFocus GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTDESCRIPTION )
         end if

      case IDC_FRMSNIPPETS_CMDDELETE
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1  THEN
               if MessageBox( HWND, L(87, "Are you sure you want to delete this Snippet?"), L(276,"Confirm"), _
                                 MB_YESNOCANCEL Or MB_ICONINFORMATION Or MB_DEFBUTTON1 ) = IDYES then
                  if ubound(gConfig.SnippetsTemp) = 0 THEN
                     erase gConfig.SnippetsTemp
                     nCurSel = -1
                  else
                     ' remove the item from the internal array
                     for i as long = nCurSel to ubound(gConfig.SnippetsTemp) - 1
                        gConfig.SnippetsTemp(i) = gConfig.SnippetsTemp( i + 1 )
                     NEXT
                     redim preserve gConfig.BuildsTemp( ubound(gConfig.SnippetsTemp) - 1 )
                  END IF
                  ' reload the listbox
                  LoadSnippetsListBox(HWND)
                  nCurSel = Min( nCurSel, ubound(gConfig.SnippetsTemp) )
                  ListBox_SetCurSel(hList1, nCurSel)
                  SetSnippetsTextboxes()
                  SetFocus hList1
               end if
            END IF
         end if

      Case IDC_FRMSNIPPETS_CMDOK
         If codeNotify = BN_CLICKED Then
            ' Copy the temporary items to the main array  
            redim gConfig.Snippets( ubound(gConfig.SnippetsTemp) )
            for i as long = lbound(gConfig.SnippetsTemp) to ubound(gConfig.SnippetsTemp)
               gConfig.Snippets(i) = gConfig.SnippetsTemp(i)   
            NEXT
            erase gConfig.SnippetsTemp
            gConfig.SaveSnippets
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
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmSnippets_PositionWindows( ByVal HWnd As HWnd ) As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND)
   if pWindow = 0 then exit function
   
   dim as hwnd hCtrl
   dim as long nLeft, nTop, nWidth, nHeight, cx, cy
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect( HWnd, @rc )
  
   ' The command buttons are initially positioned at 48 pixels from the bottom of the form
   cy = pWindow->ScaleY(48)
   
   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_CMDUP )
   AfxGetWindowLocation( hCtrl, nLeft, nTop )
   AfxSetWindowLocation( hCtrl, nLeft, rc.bottom - cy )

   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_CMDDOWN )
   AfxGetWindowLocation( hCtrl, nLeft, nTop )
   AfxSetWindowLocation( hCtrl, nLeft, rc.bottom - cy )

   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_CMDINSERT )
   AfxGetWindowLocation( hCtrl, nLeft, nTop )
   AfxSetWindowLocation( hCtrl, nLeft, rc.bottom - cy )

   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_CMDDELETE )
   AfxGetWindowLocation( hCtrl, nLeft, nTop )
   AfxSetWindowLocation( hCtrl, nLeft, rc.bottom - cy )

   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_CMDOK )
   cx = pWindow->ScaleX(168)
   AfxSetWindowLocation( hCtrl, rc.right - cx, rc.bottom - cy )

   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDCANCEL )
   cx = pWindow->ScaleX(86)
   AfxSetWindowLocation( hCtrl, rc.right - cx, rc.bottom - cy )

   ' The listbox is positioned 60 pixels from the bottom
   cy = pWindow->ScaleY(60)
   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_LIST1 )
   AfxGetWindowLocation( hCtrl, nLeft, nTop )
   nWidth = AfxGetWindowWidth( hCtrl )
   AfxSetWindowSize( hCtrl, nWidth, rc.bottom - nTop - cy )
   
   ' The textboxes are positioned 10 pixels from the right
   cx = pWindow->ScaleX(10)
   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTDESCRIPTION )
   AfxGetWindowLocation( hCtrl, nLeft, nTop )
   nHeight = AfxGetWindowHeight( hCtrl )
   AfxSetWindowSize( hCtrl, rc.right - nLeft - cx, nHeight )

   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_TXTCODE )
   AfxGetWindowLocation( hCtrl, nLeft, nTop )
   nHeight = AfxGetWindowHeight( hCtrl )
   AfxSetWindowSize( hCtrl, rc.right - nLeft - cx, rc.bottom - nTop - cy )

   function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmSnippets
' ========================================================================================
Function frmSnippets_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      frmSnippets_PositionWindows(HWnd)
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmSnippets
' ========================================================================================
Function frmSnippets_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmSnippets
' ========================================================================================
Function frmSnippets_OnDestroy(HWnd As HWnd) As LRESULT
   ' Delete the font we created and applied to the multiline textbox
   DeleteObject( AfxGetWindowFont( GetDlgItem(hwnd, IDC_FRMSNIPPETS_TXTCODE) ))
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmSnippets Window procedure
' ========================================================================================
Function frmSnippets_WndProc( ByVal HWnd   As HWnd, _
                              ByVal uMsg   As UINT, _
                              ByVal wParam As WPARAM, _
                              ByVal lParam As LPARAM _
                              ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,  frmSnippets_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,   frmSnippets_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY, frmSnippets_OnDestroy)
      HANDLE_MSG (HWnd, WM_SIZE,    frmSnippets_OnSize)
      HANDLE_MSG (HWnd, WM_COMMAND, frmSnippets_OnCommand)

      Case WM_GETMINMAXINFO
         Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND)
         if pWindow = 0 then exit function
         DefWindowProc(hwnd, uMsg, wParam, lParam)
         Dim pMinMaxInfo As MINMAXINFO Ptr
         pMinMaxInfo = Cast(MINMAXINFO Ptr,lParam)
         pMinMaxInfo->ptMinTrackSize.x = pWindow->ScaleX( 622 )
         pMinMaxInfo->ptMinTrackSize.y = pWindow->ScaleY( 436 )
         pMinMaxInfo->ptMaxTrackSize.x = pMinMaxInfo->ptMaxTrackSize.x
         pMinMaxInfo->ptMaxTrackSize.y = pMinMaxInfo->ptMaxTrackSize.y
         return 0

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmSnippets_Show
' ========================================================================================
Function frmSnippets_Show( ByVal hWndParent As HWnd, _
                           ByVal nCmdShow   As Long = 0 _
                           ) As Long

   DIM hBitmap AS HBITMAP 
   dim hCtrl as HWnd
   dim wszImage as wstring * 100

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create(hWndParent, L(88,"User Snippets"), _
        @frmSnippets_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN or WS_THICKFRAME, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   pWindow->SetClientSize(622, 436)
   pWindow->Center(pWindow->hWindow, hWndParent)

   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   pWindow->AddControl("LISTBOX", , IDC_FRMSNIPPETS_LIST1, "", 10, 10, 218, 362, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or _
        LBS_NOINTEGRALHEIGHT Or LBS_HASSTRINGS, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR )

   pWindow->AddControl("LABEL", , IDC_FRMSNIPPETS_LABEL1, L(278,"Description") & ":", 240, 11, 91, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSNIPPETS_TXTDESCRIPTION, "", 240, 32, 372, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMSNIPPETS_LABEL2, L(89,"Trigger") & ":", 240, 56, 100, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMSNIPPETS_TXTTRIGGER, "", 240, 77, 100, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMSNIPPETS_LABEL4, "(" & L(90,"Press TAB in code editor to activate") & ")", 346, 77, 250, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMSNIPPETS_LABEL3, L(328,"Code") & ":", 240, 101, 91, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   hCtrl = _
   pWindow->AddControl("MULTILINETEXTBOX", , IDC_FRMSNIPPETS_TXTCODE, "", 240, 122, 372, 270, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or _  'WS_HSCROLL or WS_VSCROLL or _
        ES_LEFT Or ES_MULTILINE or ES_WANTRETURN or ES_AUTOHSCROLL or ES_AUTOVSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   dim as HFONT _hFont = pWindow->CreateFont( gConfig.EditorFontname, Val(**gConfig.EditorFontsize) )
   AfxSetWindowFont( hCtrl, _hFont )
   
   pWindow->AddControl("BUTTON", , IDC_FRMSNIPPETS_CMDINSERT, L(281, "Insert"), 8, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMSNIPPETS_CMDDELETE, L(282, "Delete"), 87, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMSNIPPETS_CMDUP, "", 166, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWUP", "IMAGE_ARROWUP16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   hCtrl = _
   pWindow->AddControl("BITMAPBUTTON", , IDC_FRMSNIPPETS_CMDDOWN, "", 199, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or BS_BITMAP Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER Or LR_DEFAULTCOLOR Or LR_SHARED, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   wszImage = iif(pWindow->DPI > 96, "IMAGE_ARROWDOWN", "IMAGE_ARROWDOWN16")
         hBitmap = AfxGdipImageFromRes(pWindow->InstanceHandle, wszImage, 0, false, IMAGE_BITMAP, 0)
         SendMessage(hCtrl, BM_SETIMAGE, IMAGE_BITMAP, cast(LPARAM, hBitmap))
         IF hBitmap THEN DeleteObject(hBitmap)

   pWindow->AddControl("BUTTON", , IDC_FRMSNIPPETS_CMDOK, L(0,"OK"), 454, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 536, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   HWND_FRMSNIPPETS = hForm
   
   ' Copy all of the Snippets to the SnippetsTemp array because we will work with 
   ' temporary copies until the user hits OK.
   redim gConfig.SnippetsTemp(ubound(gConfig.Snippets))
   for i as long = lbound(gConfig.Snippets) to ubound(gConfig.Snippets)
      gConfig.SnippetsTemp(i) = gConfig.Snippets(i)   
   NEXT
   LoadSnippetsListBox(hForm)
   
   ListBox_SetCurSel( GetDlgItem(hForm, IDC_FRMSNIPPETS_LIST1), 0 )
   SetSnippetsTextboxes()
   
   frmSnippets_PositionWindows(hForm)
   
   SetFocus GetDlgItem(hForm, IDC_FRMSNIPPETS_LIST1)
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(SW_SHOW)

   ' Delete the frmSnippets CWindow class manually allocated memory 
   Delete pWindow

End Function

