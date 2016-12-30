'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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


' =====================================================================================
' Find forward or backwards
' =====================================================================================
Function Find_UpOrDown( ByVal flagUpDown As Long, _
                        ByVal findFlags  As Long, _
                        ByVal flagSearchCurrentOnly As BOOLEAN, _
                        ByVal hWndDialog As HWnd _
                        ) As Long

   ' flagSearchCurrentOnly will be True for all F3 / Shift F3 searches because we only
   ' want those options to search the current active document only.

   Dim As Long curPos, startPos, endPos, startLine, endLine, nCurTab, nTabCount
   Dim As Long r, i, n, idx, idxFirst, IDC_SCOPE1, IDC_SCOPE2, IDC_SCOPE3  
   Dim As String buffer      
   Dim pDoc As clsDocument Ptr 
   
   ' This routine is used by both the Find dialog and the Replace dialog. Set the
   ' hWnd's and IDC's as needed for each type of dialog.
   IDC_SCOPE1 = Iif( hWndDialog = HWND_FRMFIND, IDC_FRMFIND_OPTSCOPE1, IDC_FRMREPLACE_OPTSCOPE1)
   IDC_SCOPE2 = Iif( hWndDialog = HWND_FRMFIND, IDC_FRMFIND_OPTSCOPE2, IDC_FRMREPLACE_OPTSCOPE2)
   IDC_SCOPE3 = Iif( hWndDialog = HWND_FRMFIND, IDC_FRMFIND_OPTSCOPE3, IDC_FRMREPLACE_OPTSCOPE3)

   If Len(gFind.txtFind) = 0 Then Return -1

   Dim pDocs(Any) As clsDocument Ptr
   
   ' Create an array of Scintilla document pointers that we will need to search depending
   ' on whether searching the current window or all windows has been chosen. The current
   ' document must be added to the array twice because the initial searching may start part
   ' way through the document. The second time the document is accessed will search the 
   ' first part of the document that would have been missed the first time.
   If cBool(gFind.nScopeFind = IDC_SCOPE1) OrElse (flagSearchCurrentOnly = True) Then   
      ' Search current document
      ReDim pDocs(1) As clsDocument Ptr
      pDocs(0) = gTTabCtl.GetActiveDocumentPtr
      pDocs(1) = gTTabCtl.GetActiveDocumentPtr
      nTabCount = 2
   Elseif gFind.nScopeFind = IDC_SCOPE2 then
      ' Search all loaded documents
      nTabCount = TabCtrl_GetItemCount(gTTabCtl.hWindow)           
      nCurTab   = TabCtrl_GetCurSel(gTTabCtl.hWindow)
      ReDim pDocs(nTabCount) As clsDocument Ptr
      For idx = nCurTab To nTabCount - 1
         pDocs(n) = gTTabCtl.GetDocumentPtr(idx): n = n + 1
      Next
      For idx = 0 To nCurTab - 1 
         pDocs(n) = gTTabCtl.GetDocumentPtr(idx): n = n + 1
      Next
      pDocs(nTabCount) = gTTabCtl.GetDocumentPtr(nCurTab)
      nTabCount = nTabCount + 1
   Elseif gFind.nScopeFind = IDC_SCOPE3 then
      ' Search selected text
      ReDim pDocs(1) As clsDocument Ptr
      pDocs(0) = gTTabCtl.GetActiveDocumentPtr
      nTabCount = 1   
      flagUpDown = IDM_FINDNEXT
   End If

   
   ' cycle the search through all of the loaded pDocs
   For idx = 0 To nTabCount - 1
      pDoc = pDocs(idx)
   
      ' Determine the start and end position. If searching a Selection then
      ' set that now, otherwise deal with the group of documents.
      if gFind.nScopeFind = IDC_SCOPE3 then
         pDoc->GetSelectedLineRange(startLine, endLine, startPos, endPos )
      else
         ' Begin to search from the current position
         If flagUpDown = IDM_FINDNEXT Then
            startPos = Iif(curPos = -1, 0, SendMessageW( pDoc->hWindow, SCI_GETCURRENTPOS, 0, 0) + 1)
            endPos = SendMessageW( pDoc->hWindow, SCI_GETTEXTLENGTH, 0, 0)
         Else
            ' For backward searches the end position must be less than the start position
            endPos = 0
            startPos = Iif(curPos = -1, _
                           SendMessageW( pDoc->hWindow, SCI_GETTEXTLENGTH, 0, 0), _
                           SendMessageW( pDoc->hWindow, SCI_GETCURRENTPOS, 0, 0) - 1)
         End If
      end if

      curPos = 0               

      SendMessageW( pDoc->hWindow, SCI_SETTARGETSTART, startPos, 0)
      SendMessageW( pDoc->hWindow, SCI_SETTARGETEND, endPos, 0)
      SendMessageW( pDoc->hWindow, SCI_SETSEARCHFLAGS, findFlags, 0)

      ' Search for the text to find
      r = SendMessageW( pDoc->hWindow, SCI_SEARCHINTARGET, Len(gFind.txtFind), Cast(LPARAM, Strptr(gFind.txtFind)))
      If r = -1 Then ' SEARCH NOT FOUND
         ' Search the next loaded window or display not found messagebox if no more pDocs to search.
         If idx + 1 > nTabCount - 1 Then                              
            MessageBoxW( hWndDialog, L(228, "No matches found."), L(158, "Find"), MB_OK Or MB_ICONINFORMATION Or MB_APPLMODAL Or MB_DEFBUTTON1)
            If hWndDialog = HWND_FRMFIND Then SetFocus GetDlgItem(hWndDialog, IDC_FRMFIND_COMBOFIND)
            If hWndDialog = HWND_FRMREPLACE Then SetFocus GetDlgItem(hWndDialog, IDC_FRMREPLACE_COMBOFIND)
            Exit Function
         End If   
         curPos = -1
      Else
         ' Position the caret and select the text
         SendMessageW( pDoc->hWindow, SCI_SETCURRENTPOS, r, 0)
         SendMessageW( pDoc->hWindow, SCI_GOTOPOS, r, 0)
         if gFind.nScopeFind = IDC_SCOPE3 then
            SendMessageW( pDoc->hWindow, SCI_SETSELECTIONSTART, startPos, 0)
            SendMessageW( pDoc->hWindow, SCI_SETSELECTIONEND, endPos, 0)
         else
            SendMessageW( pDoc->hWindow, SCI_SETSELECTIONSTART, r, 0)
            SendMessageW( pDoc->hWindow, SCI_SETSELECTIONEND, r+Len(gFind.txtFind), 0)
         end if
         gTTabCtl.SetTabIndexByDocumentPtr(pDoc)
         Exit Function
      End If
   Next
      
   Function = 0
End Function



' ========================================================================================
' Add a string to the Find combobox (ensure no duplicates)
' ========================================================================================
Function frmFind_AddToFindCombo( ByRef sText As Const String ) As Long
   Dim i As Long 
   
   ' If the string already exists in the combobox then remove it
   i = Instr(gFind.txtFindCombo, sText & DELIM)
   If i Then gFind.txtFindCombo = AfxStrDelete(gFind.txtFindCombo, i, Len(sText))
  
   ' Add it to the beginning of the list
   gFind.txtFindCombo = sText & DELIM & gFind.txtFindCombo 

   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmFind
' ========================================================================================
Function frmFind_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDOK  ' Find Next
         If codeNotify = BN_CLICKED Then
            Dim findFlags As Long 
            If gFind.nWholeWord Then findFlags = (findFlags Or SCFIND_WHOLEWORD)
            If gFind.nMatchCase Then findFlags = (findFlags Or SCFIND_MATCHCASE)
         
            ' Ensure that the text in the find textbox has been added to the combobox. Need to do
            ' this to ensure that manually typed in text has been captured.
            Dim wszText As WString * MAX_PATH
            wszText = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMFIND_COMBOFIND))
            gFind.txtFind = Str(wszText)
            frmFind_AddToFindCombo(gFind.txtFind)
            
            ' Do the actual search
            Find_UpOrDown(IDM_FINDNEXT, findFlags, False, HWND_FRMFIND)
            Exit Function
         End If  
         
      Case IDCANCEL  ' Close
         If codeNotify = BN_CLICKED Then
            SendMessageW HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If

      Case IDC_FRMFIND_COMBOFIND
         ' Catch the dropdown of the Find combobox so we can populate it
         If codeNotify = CBN_DROPDOWN Then
            ' Add any previous search terms to the combobox
            Dim As Long i, nCount 
            Dim wszText As WString * MAX_PATH
            wszText = AfxGetWindowText(hwndCtl)
            gFind.txtFind = Str(wszText)
            frmFind_AddToFindCombo(gFind.txtFind)
            ComboBox_ResetContent(hwndCtl)
            nCount = AfxStrParseCount(gFind.txtFindCombo, DELIM )
            For i = 1 To nCount
               wszText = AfxStrParse(gFind.txtFindCombo, i, DELIM)
               If Len(wszText) Then ComboBox_AddString( hwndCtl, @wszText )
            Next
            wszText = gFind.txtFind
            i = ComboBox_FindStringExact( hwndCtl, -1, @wszText )
            ComboBox_SetCurSel(hwndCtl, i)
         End If

      Case IDC_FRMFIND_CHKMATCHCASE
         If codeNotify = BN_CLICKED Then
            gFind.nMatchCase = Button_GetCheck(hwndCtl)
            Exit Function
         End If

      Case IDC_FRMFIND_CHKWHOLEWORDS
         If codeNotify = BN_CLICKED Then
            gFind.nWholeWord = Button_GetCheck(hwndCtl)
            Exit Function
         End If

      Case IDC_FRMFIND_OPTSCOPE1  ' current document
         If codeNotify = BN_CLICKED Then
            If Button_GetCheck(hwndCtl) = BST_CHECKED Then gFind.nScopeFind = IDC_FRMFIND_OPTSCOPE1      
            Exit Function
         End If

      Case IDC_FRMFIND_OPTSCOPE2   ' all open documents
         If codeNotify = BN_CLICKED Then
            If Button_GetCheck(hwndCtl) = BST_CHECKED Then gFind.nScopeFind = IDC_FRMFIND_OPTSCOPE2
            Exit Function
         End If

      Case IDC_FRMFIND_OPTSCOPE3    ' selection
         If codeNotify = BN_CLICKED Then
            If Button_GetCheck(hwndCtl) = BST_CHECKED Then gFind.nScopeFind = IDC_FRMFIND_OPTSCOPE3
            Exit Function
         End If

   End Select

   Function = 0
End Function



' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmFind
' ========================================================================================
Function frmFind_OnClose(HWnd As HWnd) As LRESULT
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmFind
' ========================================================================================
Function frmFind_OnDestroy(HWnd As HWnd) As LRESULT 
   ' Save the search parameters for later instances of the find/replace dialog
   gFind.nWholeWord = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFIND_CHKWHOLEWORDS) )  
   gFind.nMatchCase = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFIND_CHKMATCHCASE) )    
   If Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFIND_OPTSCOPE1) ) Then gFind.nScopeFind = IDC_FRMFIND_OPTSCOPE1      
   If Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFIND_OPTSCOPE2) ) Then gFind.nScopeFind = IDC_FRMFIND_OPTSCOPE2

   HWND_FRMFIND = 0
   Function = 0
End Function


' ========================================================================================
' frmFind Window procedure
' ========================================================================================
Function frmFind_WndProc( ByVal HWnd   As HWnd, _
                          ByVal uMsg   As UINT, _
                          ByVal wParam As WPARAM, _
                          ByVal lParam As LPARAM _
                          ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CLOSE,    frmFind_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmFind_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmFind_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmFind_Show
' ========================================================================================
Function frmFind_Show( ByVal hWndParent As HWnd ) As Long

   ' We don't want the Replace dialog to display at the same time as the Find form.
   If IsWindowVisible(HWND_FRMREPLACE) Then DestroyWindow(HWND_FRMREPLACE)

   ' If the Find window already exists then simply set focus to it and exit.
   If IsWindow(HWND_FRMFIND) Then
      SetForegroundWindow(HWND_FRMMAIN)
      SetFocus GetDlgItem( HWND_FRMFIND, IDC_FRMFIND_COMBOFIND )
      Exit Function
   End If
 
      
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   pWindow->ClassStyle = CS_DROPSHADOW

   HWND_FRMFIND = _
   pWindow->Create(hwndParent, L(158,"Find"), @frmFind_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(510, 155)
   pWindow->Center
   
   pWindow->AddControl("LABEL", , IDC_FRMFIND_LBLFINDWHAT, L(159,"Find What:"), 7, 20, 63, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMFIND_COMBOFIND, "", 75, 17, 327, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDOK, L(160,"Find Next"), 414, 16, 90, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(161,"Close"), 414, 47, 90, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMFIND_FRAMEOPTIONS, L(167,"Options"), 16, 53, 187, 90, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFIND_CHKWHOLEWORDS, L(162,"Match Whole Words"), 26, 68, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFIND_CHKMATCHCASE, L(163,"Match Case"), 26, 88, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMFIND_FRAMESCOPE, L(164,"Scope"), 215, 53, 187, 90, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMFIND_OPTSCOPE1, L(165,"Current Document"), 225, 71, 170, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", , IDC_FRMFIND_OPTSCOPE2, L(166,"All Open Documents"), 225, 91, 170, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
' Find/Replace in Selected text is not working okay yet.
'   pWindow->AddControl("OPTION", , IDC_FRMFIND_OPTSCOPE3, L(220,"Selected Text"), 225, 111, 170, 16, _
'        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
'        WS_EX_LEFT Or WS_EX_LTRREADING)

   Dim As Long i, nCount, startPos, endPos, startLine, endLine
   Dim As String buffer      
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then 
      pDoc->GetSelectedLineRange( startLine, endLine, startPos, endPos )

      ' Fills the search box with the selected word.
      ' If there are carriage returns or/and line feeds, this mean that
      ' there is a block selected, instead of a word, so avoid it.
      buffer = pDoc->GetSelText()
      If Len(buffer) Then
         If Instr(buffer, Chr(13, 10)) = 0 Then
            gFind.txtFind = buffer
            gFind.txtLastFind = buffer
         End If
      Else
         gFind.txtFind = gFind.txtLastFind
      End If
   End If

   frmFind_AddToFindCombo(gFind.txtFind)

   If gFind.nScopeFind = 0 Then gFind.nScopeFind = IDC_FRMFIND_OPTSCOPE1
   Button_SetCheck( GetDlgItem(HWND_FRMFIND, IDC_FRMFIND_OPTSCOPE1), Iif(gFind.nScopeFind = IDC_FRMFIND_OPTSCOPE1, BST_CHECKED, BST_UNCHECKED) )
   Button_SetCheck( GetDlgItem(HWND_FRMFIND, IDC_FRMFIND_OPTSCOPE2), Iif(gFind.nScopeFind = IDC_FRMFIND_OPTSCOPE2, BST_CHECKED, BST_UNCHECKED) )
   
   SetWindowTextW( GetDlgItem(HWND_FRMFIND, IDC_FRMFIND_COMBOFIND), WStr(gFind.txtFind) )    
   EnableWindow( GetDlgItem(HWND_FRMFIND, IDOK), True )
 
   Button_SetCheck( GetDlgItem(HWND_FRMFIND, IDC_FRMFIND_CHKWHOLEWORDS), gFind.nWholeWord )
   Button_SetCheck( GetDlgItem(HWND_FRMFIND, IDC_FRMFIND_CHKMATCHCASE), gFind.nMatchCase )
   
   ShowWindow HWND_FRMFIND, SW_SHOW
   SetFocus GetDlgItem( HWND_FRMFIND, IDC_FRMFIND_COMBOFIND )

   Function = 0
End Function









