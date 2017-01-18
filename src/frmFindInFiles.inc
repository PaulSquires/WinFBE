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



' ========================================================================================
' Add a string to the Find combobox (ensure no duplicates)
' ========================================================================================
Function frmFindInFiles_AddToFindCombo( ByRef sText As Const String ) As Long
   Dim i As Long 
   
   ' If the string already exists in the combobox then remove it
   i = Instr(gFindInFiles.txtFindCombo, sText & DELIM)
   If i Then gFindInFiles.txtFindCombo = AfxStrDelete(gFindInFiles.txtFindCombo, i, Len(sText))
  
   ' Add it to the beginning of the list
   gFindInFiles.txtFindCombo = sText & DELIM & gFindInFiles.txtFindCombo 

   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDOK  ' Find Next
         If codeNotify = BN_CLICKED Then
            Dim findFlags As Long 
            If gFindInFiles.nWholeWord Then findFlags = (findFlags Or SCFIND_WHOLEWORD)
            If gFindInFiles.nMatchCase Then findFlags = (findFlags Or SCFIND_MATCHCASE)
         
            ' Ensure that the text in the find textbox has been added to the combobox. Need to do
            ' this to ensure that manually typed in text has been captured.
            Dim wszText As WString * MAX_PATH
            wszText = AfxGetWindowText( GetDlgItem(HWnd, IDC_frmFindInFiles_COMBOFIND))
            gFindInFiles.txtFind = Str(wszText)
            frmFindInFiles_AddToFindCombo(gFindInFiles.txtFind)
            
            ' Do the actual search
            'Find_UpOrDown(IDM_FINDNEXT, findFlags, False, HWND_FRMFINDINFILES)
            Exit Function
         End If  
         
      Case IDCANCEL  ' Close
         If codeNotify = BN_CLICKED Then
            SendMessageW HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_COMBOFIND
         ' Catch the dropdown of the Find combobox so we can populate it
         If codeNotify = CBN_DROPDOWN Then
            ' Add any previous search terms to the combobox
            Dim As Long i, nCount 
            Dim wszText As WString * MAX_PATH
            wszText = AfxGetWindowText(hwndCtl)
            gFindInFiles.txtFind = Str(wszText)
            frmFindInFiles_AddToFindCombo(gFindInFiles.txtFind)
            ComboBox_ResetContent(hwndCtl)
            nCount = AfxStrParseCount(gFindInFiles.txtFindCombo, DELIM )
            For i = 1 To nCount
               wszText = AfxStrParse(gFindInFiles.txtFindCombo, i, DELIM)
               If Len(wszText) Then ComboBox_AddString( hwndCtl, @wszText )
            Next
            wszText = gFindInFiles.txtFind
            i = ComboBox_FindStringExact( hwndCtl, -1, @wszText )
            ComboBox_SetCurSel(hwndCtl, i)
         End If

      Case IDC_FRMFINDINFILES_CHKMATCHCASE
         If codeNotify = BN_CLICKED Then
            gFindInFiles.nMatchCase = Button_GetCheck(hwndCtl)
            Exit Function
         End If

      Case IDC_FRMFINDINFILES_CHKWHOLEWORDS
         If codeNotify = BN_CLICKED Then
            gFindInFiles.nWholeWord = Button_GetCheck(hwndCtl)
            Exit Function
         End If

   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

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
' Process WM_CLOSE message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow GetParent(HWnd), True 
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmFindInFiles
' ========================================================================================
Function frmFindInFiles_OnDestroy(HWnd As HWnd) As LRESULT 
   ' Save the search parameters for later instances of the find/replace dialog
   gFindInFiles.nWholeWord = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKWHOLEWORDS) )  
   gFindInFiles.nMatchCase = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_CHKMATCHCASE) )    
   If Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_OPTSCOPE1) ) Then gFindInFiles.nScopeFind = IDC_FRMFINDINFILES_OPTSCOPE1      
   If Button_GetCheck( GetDlgItem(HWnd, IDC_FRMFINDINFILES_OPTSCOPE2) ) Then gFindInFiles.nScopeFind = IDC_FRMFINDINFILES_OPTSCOPE2

   HWND_FRMFINDINFILES = 0
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmFindInFiles Window procedure
' ========================================================================================
Function frmFindInFiles_WndProc( ByVal HWnd   As HWnd, _
                                 ByVal uMsg   As UINT, _
                                 ByVal wParam As WPARAM, _
                                 ByVal lParam As LPARAM _
                                 ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmFindInFiles_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmFindInFiles_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmFindInFiles_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmFindInFiles_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmFindInFiles_Show
' ========================================================================================
Function frmFindInFiles_Show( ByVal hWndParent As HWnd ) As Long

   ' We don't want the Find or Replace dialogs to display at the same time.
   If IsWindowVisible(HWND_FRMREPLACE) Then DestroyWindow(HWND_FRMREPLACE)
   If IsWindowVisible(HWND_FRMFINDINFILES) Then DestroyWindow(HWND_FRMFINDINFILES)

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   HWND_FRMFINDINFILES = _
   pWindow->Create(hwndParent, L(256,"Find In Files"), @frmFindInFiles_WndProc, 0, 0, 0, 0, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(510, 155)
   pWindow->Center
   
   pWindow->AddControl("LABEL", , IDC_FRMFINDINFILES_LBLFINDWHAT, L(159,"Find What:"), 7, 20, 63, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMFINDINFILES_COMBOFIND, "", 75, 17, 327, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("COMBOBOX", , IDC_FRMFINDINFILES_COMBOPATTERN, "", 75, 37, 295, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("COMBOBOX", , IDC_FRMFINDINFILES_COMBOFOLDER, "", 75, 57, 295, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDOK, L(160,"Find Next"), 414, 16, 90, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(161,"Close"), 414, 47, 90, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMFINDINFILES_FRAMEOPTIONS, L(167,"Options"), 16, 53, 187, 90, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKWHOLEWORDS, L(162,"Match Whole Words"), 26, 68, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKMATCHCASE, L(163,"Match Case"), 26, 88, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMFINDINFILES_CHKSUBFOLDERS, L(163,"Search Subfolders"), 26, 118, 167, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

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
            gFindInFiles.txtFind = buffer
            gFindInFiles.txtLastFind = buffer
         End If
      Else
         gFindInFiles.txtFind = gFindInFiles.txtLastFind
      End If
   End If

   frmFindInFiles_AddToFindCombo(gFindInFiles.txtFind)

   SetWindowTextW( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFINDINFILES_COMBOFIND), WStr(gFindInFiles.txtFind) )    
   EnableWindow( GetDlgItem(HWND_FRMFINDINFILES, IDOK), True )
 
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFIND_CHKWHOLEWORDS), gFindInFiles.nWholeWord )
   Button_SetCheck( GetDlgItem(HWND_FRMFINDINFILES, IDC_FRMFIND_CHKMATCHCASE), gFindInFiles.nMatchCase )
   
   ShowWindow HWND_FRMFINDINFILES, SW_SHOW
   SetFocus GetDlgItem( HWND_FRMFINDINFILES, IDC_FRMFIND_COMBOFIND )

   ' Process Windows messages (modal)
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the CWindow class manually allocated memory 
   Delete pWindow

   Function = 0
End Function



