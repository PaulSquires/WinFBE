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

#include once "frmSnippets.bi"

' Create an array that holds all options in the replaceable parameters for snippets.
Dim shared gSnippetParameter(...) as CWSTR => _
   {  "CLIPBOARD", _
      "CURSOR_POSITION", _
      "FILENAME", _
      "FILENAME_SHORT", _
      "CURRENT_YEAR", _
      "CURRENT_YEAR_SHORT", _
      "CURRENT_MONTH", _
      "CURRENT_MONTH_NAME", _
      "CURRENT_MONTH_NAME_SHORT", _
      "CURRENT_DAY", _
      "CURRENT_DAY_NAME", _
      "CURRENT_DAY_NAME_SHORT", _
      "CURRENT_HOUR", _
      "CURRENT_MINUTE", _
      "CURRENT_SECOND" _
   }


' ========================================================================================
' Create the right-click popup Snippets Parameters menu 
' ========================================================================================
private Function frmSnippets_CreateSnippetsParametersMenu() As HMENU
   
   Dim hPopUpMenu As HMENU = CreatePopupMenu()
   for i as long = lbound(gSnippetParameter) to ubound(gSnippetParameter)
      AppendMenu( hPopUpMenu, MF_ENABLED or MF_STRING, _
                  IDM_FRMSNIPPETS_PARAMETERBASE + i, gSnippetParameter(i) )
   next
   
   Function = hPopupMenu
 
End Function


' ========================================================================================
' Replace any embedded Parameters in the code text
' ========================================================================================
private function frmSnippets_DoReplaceParameters( byval wszInsertText as CWSTR ) as CWSTR
   dim pDoc as clsDocument ptr = gTTabCtl.GetActiveDocumentPtr
   
   dim as CWSTR wszText = wszInsertText
   
   ' Contents of the clipboard
   wszText = AfxStrReplace( wszText, "{CLIPBOARD}", AfxGetClipboardText )

   ' Filename (full path and name)
   if pDoc then
      wszText = AfxStrReplace( wszText, "{FILENAME}", pDoc->DiskFilename )
      wszText = AfxStrReplace( wszText, "{FILENAME_SHORT}", AfxStrPathname( "NAMEX", pDoc->DiskFilename ))
   end if   
   
   ' The current year
   wszText = AfxStrReplace( wszText, "{CURRENT_YEAR}", wstr(AfxLocalYear) )
   
   ' The current year's last two digits
   wszText = AfxStrReplace( wszText, "{CURRENT_YEAR_SHORT}", right( wstr(AfxLocalYear), 2) )

   ' The month as two digits (example '02')
   wszText = AfxStrReplace( wszText, "{CURRENT_MONTH}", AfxStrRSet(wstr(AfxLocalMonth), 2, "0") )
     
   ' The full name of the month (example 'July')
   wszText = AfxStrReplace( wszText, "{CURRENT_MONTH_NAME}", AfxLocalMonthName )

   ' The short name of the month (example 'Jul')
   wszText = AfxStrReplace( wszText, "{CURRENT_MONTH_NAME_SHORT}", AfxLocalShortMonthName )

   ' The day of the month as two digits (example '09')
   wszText = AfxStrReplace( wszText, "{CURRENT_DAY}", AfxStrRSet(wstr(AfxLocalDay), 2, "0") )

   ' The name of day (example 'Monday')
   wszText = AfxStrReplace( wszText, "{CURRENT_DAY_NAME}", AfxLocalDayName )
   
   ' The short name of the day (example 'Mon')
   wszText = AfxStrReplace( wszText, "{CURRENT_DAY_NAME_SHORT}", AfxLocalDayShortName )
   
   dim as CWSTR wszCurTime = AfxLocalTimeStr( "HHmmss" )
   
   ' The current hour in 24-hour clock format
   wszText = AfxStrReplace( wszText, "{CURRENT_HOUR}", left(wszCurTime, 2) )
   
   ' The current minute
   wszText = AfxStrReplace( wszText, "{CURRENT_MINUTE}", mid(wszCurTime, 3, 2) )
   
   ' The current second
   wszText = AfxStrReplace( wszText, "{CURRENT_SECOND}", right(wszCurTime, 2) )

   function = wszText
end function


' ========================================================================================
' Attempt to insert/expand a user snippet
' ========================================================================================
public function frmSnippets_DoInsertSnippet( byval pDoc as clsDocument ptr ) as Boolean
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
         curPos = curPos - len(strWord)
         ' Replace the selection (SCI_REPLACESEL fails if text is "" so use Cut instead for that scenario)                                
         If Len(strWord) = 0 Then
            SendMessage( hEdit, SCI_CUT, 0, 0 )
         Else
            strReplaceText = frmSnippets_DoReplaceParameters( gConfig.Snippets(i).wszCode )
            
            ' Check of the {CURSOR_POSITION} parameter exists. This is where we want to 
            ' position our cursor after the edit it made.
            dim as string strParam = "{CURSOR_POSITION}"
            dim as long iPos = instr( strReplaceText, strParam )
            if iPos then strReplaceText = AfxStrRemove( strReplaceText, strParam )
            SciExec( hEdit, SCI_REPLACESEL, 0, Strptr(strReplaceText) )
            if iPos then SciExec( hEdit, SCI_GOTOPOS, curPos + iPos - 1, 0 )

         End If
         return true
      end if
   next
   
   return false
end function


' ========================================================================================
' Load all of the snippets descriptions into the listbox
' ========================================================================================
private function frmSnippets_LoadListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMSNIPPETS_LIST1)

   ListBox_ResetContent(hList1)
   for i as long = lbound(gConfig.SnippetsTemp) to ubound(gConfig.SnippetsTemp)
      ListBox_AddString(hList1, gConfig.SnippetsTemp(i).wszDescription.sptr)
   NEXT

   function = 0
end function


' ========================================================================================
' Swap two entries in the snippets Listbox
' ========================================================================================
private function frmSnippets_SwapListBoxItems( byval Item1 as long, _
                                               Byval Item2 as long _
                                               ) as Long
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
private function frmSnippets_SetTextboxes() as long
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
' Processes messages for the subclassed code textbox window.
' ========================================================================================
private Function frmSnippets_TextBox_SubclassProc ( _
                           ByVal HWnd   As HWnd, _                 ' // Control window handle
                           ByVal uMsg   As UINT, _                 ' // Type of message
                           ByVal wParam As WPARAM, _               ' // First message parameter
                           ByVal lParam As LPARAM, _               ' // Second message parameter
                           ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                           ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                           ) As LRESULT

   
   Select Case uMsg

      case WM_CONTEXTMENU
         ' Override the default edit control menu with our own to show a list
         ' of replaceable parameters.
         dim pt as POINT
         Dim hPopUpMenu As HMENU = frmSnippets_CreateSnippetsParametersMenu()
         GetCursorPos @pt
         TrackPopupMenu( hPopUpMenu, 0, pt.x, pt.y, 0, HWND_FRMSNIPPETS, ByVal Null )
         DestroyMenu hPopUpMenu
         return true
            
      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmSnippets_TextBox_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmSnippets
' ========================================================================================
private Function frmSnippets_OnCreate( ByVal HWnd As HWnd, _
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
' Process WM_COMMAND message for window/dialog: frmSnippets
' ========================================================================================
private Function frmSnippets_OnCommand( ByVal HWnd As HWnd, _
                                        ByVal id As Long, _
                                        ByVal hwndCtl As HWnd, _
                                        ByVal codeNotify As UINT _
                                        ) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMSNIPPETS_LIST1)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   ' Convert a ListBox DoubleClick into pressing OK button action
   if (id = IDC_FRMSNIPPETS_LIST1) AndAlso (codeNotify = LBN_DBLCLK) THEN
      id = IDOK: codeNotify = BN_CLICKED
   END IF
   
   Select Case id
      case IDM_FRMSNIPPETS_PARAMETERBASE + lbound(gSnippetParameter) to _
           IDM_FRMSNIPPETS_PARAMETERBASE + ubound(gSnippetParameter)
         
         ' Create the parameter to insert
         dim wszText as wstring * MAX_PATH
         wszText = "{" & gSnippetParameter( id-IDM_FRMSNIPPETS_PARAMETERBASE ) & "}"
         
         ' Insert the parameter at the current textbox location.
         SendMessage( GetDlgItem(HWND, IDC_FRMSNIPPETS_TXTCODE), _
                        EM_REPLACESEL, true, cast(LPARAM, @wszText) )
            
      case IDC_FRMSNIPPETS_LIST1
         if codeNotify = LBN_SELCHANGE THEN
            frmSnippets_SetTextboxes()
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
               frmSnippets_SwapListBoxItems( nCurSel, nCurSel - 1 )
            END IF
         end if

      case IDC_FRMSNIPPETS_CMDDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1) - 1 THEN
               frmSnippets_SwapListBoxItems( nCurSel, nCurSel + 1 )
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
            frmSnippets_LoadListBox(HWND)
            nCurSel = Min(nCurSel, ubound(gConfig.SnippetsTemp))
            ListBox_SetCurSel(hList1, nCurSel)
            frmSnippets_SetTextboxes()
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
                  frmSnippets_LoadListBox(HWND)
                  nCurSel = Min( nCurSel, ubound(gConfig.SnippetsTemp) )
                  ListBox_SetCurSel(hList1, nCurSel)
                  frmSnippets_SetTextboxes()
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
private Function frmSnippets_PositionWindows() As LRESULT

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMSNIPPETS)
   if pWindow = 0 then exit function
   
   dim as hwnd hCtrl
   dim as long nLeft, nTop, nWidth, nHeight, cx, cy
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect( HWND_FRMSNIPPETS, @rc )
  
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
   nLeft = rc.right - cx
   AfxSetWindowLocation( hCtrl, nLeft, rc.bottom - cy )
   dim as long nLeftOK = nLeft
   
   hCtrl = GetDlgItem( HWND_FRMSNIPPETS, IDC_FRMSNIPPETS_LBLPARAMETERS )
   AfxGetWindowLocation( hCtrl, nLeft, nTop )
   AfxSetWindowLocation( hCtrl, nLeft, rc.bottom - cy )
   nHeight = AfxGetWindowHeight( hCtrl )
   AfxSetWindowSize( hCtrl, nLeftOK - nLeft, nHeight )

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
private Function frmSnippets_OnSize( ByVal HWnd As HWnd, _
                                     ByVal state As UINT, _
                                     ByVal cx As Long, _
                                     ByVal cy As Long _
                                     ) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      frmSnippets_PositionWindows
   End If
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmSnippets
' ========================================================================================
private Function frmSnippets_OnClose( byval HWnd As HWnd ) As LRESULT
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND)
   if pWindow then 
      ' Save the window size and placement. Use the pWindow version in order to
      ' get the coordinates UnScaled. We do this because rcSnippets is fed to
      ' pWindow->Create and that function will rescale the values.
      pWindow->GetWindowRect( @gConfig.rcSnippets )
   end if
   
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmSnippets
' ========================================================================================
private Function frmSnippets_OnDestroy( byval HWnd As HWnd ) As LRESULT
   ' Delete the font we created and applied to the multiline textbox
   DeleteObject( AfxGetWindowFont( GetDlgItem(hwnd, IDC_FRMSNIPPETS_TXTCODE) ))
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmSnippets Window procedure
' ========================================================================================
private Function frmSnippets_WndProc( ByVal HWnd   As HWnd, _
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

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmSnippets_Show
' ========================================================================================
public Function frmSnippets_Show( ByVal hWndParent As HWnd ) as LRESULT

   DIM hBitmap AS HBITMAP 
   dim hCtrl as HWnd
   dim wszImage as wstring * 100

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   ' Resize the Form should the user have previously manually resized it.
   dim as long nLeft, nTop, nWidth, nHeight
   If IsRectEmpty( @gConfig.rcSnippets ) = false then
      nLeft   = gConfig.rcSnippets.left
      nTop    = gConfig.rcSnippets.top
      nWidth  = gConfig.rcSnippets.right - gConfig.rcSnippets.left
      nHeight = gConfig.rcSnippets.bottom - gConfig.rcSnippets.top
   else
      Dim rcWork As Rect = pWindow->GetWorkArea
      nWidth  = (rcWork.Right - rcWork.Left) * .60
      nHeight = (rcWork.Bottom - rcWork.Top) * .80
   end if

   Dim As HWnd hForm = _
   pWindow->Create(hWndParent, L(88,"User Snippets"), _
        @frmSnippets_WndProc, nLeft, nTop, nWidth, nHeight, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN or WS_THICKFRAME, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT )
   
   If IsRectEmpty( @gConfig.rcSnippets ) then
      pWindow->Center(pWindow->hWindow, hWndParent)
   end if
   
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
   pWindow->AddControl("TEXTBOX", , IDC_FRMSNIPPETS_TXTTRIGGER, "", 240, 77, 130, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMSNIPPETS_LABEL4, "(" & L(90,"Press TAB in code editor to activate") & ")", 376, 77, 350, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("LABEL", , IDC_FRMSNIPPETS_LABEL3, L(328,"Code") & ":", 240, 101, 91, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   hCtrl = _
   pWindow->AddControl("MULTILINETEXTBOX", , IDC_FRMSNIPPETS_TXTCODE, "", 240, 122, 372, 270, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or _  'WS_HSCROLL or WS_VSCROLL or _
        ES_LEFT Or ES_MULTILINE or ES_WANTRETURN or ES_AUTOHSCROLL or ES_AUTOVSCROLL or WS_HSCROLL or WS_VSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, , _
        Cast(SUBCLASSPROC, @frmSnippets_TextBox_SubclassProc), IDC_FRMSNIPPETS_TXTCODE, Cast(DWORD_PTR, @pWindow))
   dim as HFONT _hFont = pWindow->CreateFont( gConfig.EditorFontname, Val(**gConfig.EditorFontsize) )
   AfxSetWindowFont( hCtrl, _hFont )
   
   pWindow->AddControl("BUTTON", , IDC_FRMSNIPPETS_CMDINSERT, L(281, "Insert"), 8, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMSNIPPETS_CMDDELETE, L(282, "Delete"), 87, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   hCtrl = _
   pWindow->AddControl("BUTTON", , IDC_FRMSNIPPETS_CMDUP, wszTriangleUp, 166, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghMenuBar.hFontSymbolLargeBold, false )     

   hCtrl = _
   pWindow->AddControl("BUTTON", , IDC_FRMSNIPPETS_CMDDOWN, wszTriangleDown, 199, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghMenuBar.hFontSymbolLargeBold, false )     

   pWindow->AddControl("LABEL", , IDC_FRMSNIPPETS_LBLPARAMETERS, L(172,"Use right-click popup menu to insert parameters"), _
        240, 388, 212, 34, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
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
   frmSnippets_LoadListBox( hForm )
   
   ListBox_SetCurSel( GetDlgItem(hForm, IDC_FRMSNIPPETS_LIST1), 0 )
   frmSnippets_SetTextboxes()
   
   frmSnippets_PositionWindows
   
   SetFocus GetDlgItem(hForm, IDC_FRMSNIPPETS_LIST1)
   
   ' Process Windows messages(modal)
   Function = pWindow->DoEvents(SW_SHOW)

   ' Delete the frmSnippets CWindow class manually allocated memory 
   Delete pWindow

End Function

