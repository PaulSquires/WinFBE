
' ========================================================================================
' Add a string to the Find combobox (ensure no duplicates)
' ========================================================================================
Function frmReplace_AddToReplaceCombo( ByRef sText As Const String ) As Long
   Dim i As Long 
   
   ' If the string already exists in the combobox then remove it
   i = Instr(gFind.txtReplaceCombo, sText & DELIM)
   If i Then gFind.txtReplaceCombo = FF_StrDelete(gFind.txtReplaceCombo, i, Len(sText))
  
   ' Add it to the beginning of the list
   gFind.txtReplaceCombo = sText & DELIM & gFind.txtReplaceCombo 

   Function = 0
End Function


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmReplace
' ========================================================================================
Function frmReplace_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmReplace
' ========================================================================================
Function frmReplace_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDOK  ' Find Next
         If codeNotify = BN_CLICKED Then
            Dim findFlags As Long 
            If gFind.nWholeWord Then findFlags = (findFlags Or SCFIND_WHOLEWORD)
            If gFind.nMatchCase Then findFlags = (findFlags Or SCFIND_MATCHCASE)
         
            ' Ensure that the text in the find and replace textboxes have been added to the combobox. Need to do
            ' this to ensure that manually typed in text has been captured.
            Dim wszText As WString * MAX_PATH
            FF_Control_GetTextW( GetDlgItem(HWnd, IDC_FRMREPLACE_COMBOFIND), @wszText, MAX_PATH )
            gFind.txtFind = Str(wszText)
            frmFind_AddToFindCombo(gFind.txtFind)
            FF_Control_GetTextW( GetDlgItem(HWnd, IDC_FRMREPLACE_COMBOREPLACE), @wszText, MAX_PATH )
            gFind.txtReplace = Str(wszText)
            frmReplace_AddToReplaceCombo(gFind.txtReplace)
            
            ' Do the actual search
            Find_UpOrDown(IDM_FINDNEXT, findFlags, False, HWND_FRMREPLACE)
            Exit Function
         End If  
         
      Case IDCANCEL  ' Close
         If codeNotify = BN_CLICKED Then
            SendMessageW HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If

      Case IDC_FRMREPLACE_COMBOFIND
         ' Catch the dropdown of the Find combobox so we can populate it
         If codeNotify = CBN_DROPDOWN Then
            ' Add any previous search terms to the combobox
            Dim As Long i, nCount 
            Dim wszText As WString * MAX_PATH
            FF_Control_GetTextW( hwndCtl, @wszText, MAX_PATH )
            gFind.txtFind = Str(wszText)
            frmFind_AddToFindCombo(gFind.txtFind)
            ComboBox_ResetContent(hwndCtl)
            nCount = FF_ParseCount(gFind.txtFindCombo, DELIM )
            For i = 1 To nCount
               wszText = FF_Parse(gFind.txtFindCombo, DELIM, i)
               If Len(wszText) Then ComboBox_AddString( hwndCtl, @wszText )
            Next
            wszText = gFind.txtFind
            i = ComboBox_FindStringExact( hwndCtl, -1, @wszText )
            ComboBox_SetCurSel(hwndCtl, i)
         End If

      Case IDC_FRMREPLACE_COMBOREPLACE
         ' Catch the dropdown of the Find combobox so we can populate it
         If codeNotify = CBN_DROPDOWN Then
            ' Add any previous search terms to the combobox
            Dim As Long i, nCount 
            Dim wszText As WString * MAX_PATH
            FF_Control_GetTextW( hwndCtl, @wszText, MAX_PATH )
            gFind.txtReplace = Str(wszText)
            frmReplace_AddToReplaceCombo(gFind.txtReplace)
            ComboBox_ResetContent(hwndCtl)
            nCount = FF_ParseCount(gFind.txtReplaceCombo, DELIM )
            For i = 1 To nCount
               wszText = FF_Parse(gFind.txtReplaceCombo, DELIM, i)
               If Len(wszText) Then ComboBox_AddString( hwndCtl, @wszText )
            Next
            wszText = gFind.txtReplace
            i = ComboBox_FindStringExact( hwndCtl, -1, @wszText )
            ComboBox_SetCurSel(hwndCtl, i)
         End If

      Case IDC_FRMREPLACE_CHKMATCHCASE
         If codeNotify = BN_CLICKED Then
            gFind.nMatchCase = Button_GetCheck(hwndCtl)
            Exit Function
         End If

      Case IDC_FRMREPLACE_CHKWHOLEWORDS
         If codeNotify = BN_CLICKED Then
            gFind.nWholeWord = Button_GetCheck(hwndCtl)
            Exit Function
         End If

      Case IDC_FRMREPLACE_OPTSCOPE1
         If codeNotify = BN_CLICKED Then
            If Button_GetCheck(hwndCtl) = BST_CHECKED Then gFind.nScopeReplace = IDC_FRMREPLACE_OPTSCOPE1      
            Exit Function
         End If

      Case IDC_FRMREPLACE_OPTSCOPE2
         If codeNotify = BN_CLICKED Then
            If Button_GetCheck(hwndCtl) = BST_CHECKED Then gFind.nScopeReplace = IDC_FRMREPLACE_OPTSCOPE2
            Exit Function
         End If
 
      Case IDC_FRMREPLACE_CMDREPLACE
         If codeNotify = BN_CLICKED Then
            Dim startSelPos As Long     
            Dim endSelPos   As Long     
            
            Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr
            If pDoc = 0 Then Exit Function
         
            ' Ensure that the text in the replace textbox has been added to the combobox. Need to do
            ' this to ensure that manually typed in text has been captured.
            Dim wszText As WString * MAX_PATH
            FF_Control_GetTextW( GetDlgItem(HWnd, IDC_FRMREPLACE_COMBOREPLACE), @wszText, MAX_PATH )
            gFind.txtReplace = Str(wszText)
            frmReplace_AddToReplaceCombo(gFind.txtReplace)

            ' If startSelPos and endSelPos are the same there is no selection,
            startSelPos = SendMessageW( pDoc->hWindow, SCI_GETSELECTIONSTART, 0, 0)
            endSelPos   = SendMessageW( pDoc->hWindow, SCI_GETSELECTIONEND, 0, 0)
            If startSelPos <> endSelPos Then 
               If Len(wszText) = 0 Then
                  SendMessageW( pDoc->hWindow, SCI_CUT, 0, 0 )
               Else
                  SendMessageW( pDoc->hWindow, SCI_REPLACESEL, 0, Cast(LPARAM, Strptr(gFind.txtReplace)) )
               End If   
            End If
            
            ' Search for the next occurrance
            frmReplace_OnCommand(HWnd, IDOK, GetDlgItem(HWnd, IDOK), BN_CLICKED)
         End If
         
      Case IDC_FRMREPLACE_CMDREPLACEALL
         If codeNotify = BN_CLICKED Then
            Dim As Long nTabCount = TabCtrl_GetItemCount(gTTabCtl.hWindow)           
            Dim As Long idx, numChanges
            
            Dim pDocCurrent As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
            Dim pDoc As clsDocument Ptr 

            ' Ensure that the text in the replace textbox has been added to the combobox. Need to do
            ' this to ensure that manually typed in text has been captured.
            Dim wszText As WString * MAX_PATH
            FF_Control_GetTextW( GetDlgItem(HWnd, IDC_FRMREPLACE_COMBOFIND), @wszText, MAX_PATH )
            gFind.txtFind = Str(wszText)
            frmFind_AddToFindCombo(gFind.txtFind)
            FF_Control_GetTextW( GetDlgItem(HWnd, IDC_FRMREPLACE_COMBOREPLACE), @wszText, MAX_PATH )
            gFind.txtReplace = Str(wszText)
            frmReplace_AddToReplaceCombo(gFind.txtReplace)

            ' Cycle through all of the loaded documents and do the replacing.
            For idx = 0 To nTabCount - 1
               pDoc = gTTabCtl.GetDocumentPtr(idx)
               If pDoc Then 
                  ' Check to see if we are doing a replace on the current document or on all documents.
                  If gFind.nScopeReplace = IDC_FRMREPLACE_OPTSCOPE1 Then ' only current document
                     If pDoc <> pDocCurrent Then Continue For
                  End If
                      
                  Dim findFlags As Long 
                  If gFind.nWholeWord Then findFlags = (findFlags Or SCFIND_WHOLEWORD)
                  If gFind.nMatchCase Then findFlags = (findFlags Or SCFIND_MATCHCASE)
                  Dim As Long endPos = SendMessageW( pDoc->hWindow, SCI_GETTEXTLENGTH, 0, 0)
                  Dim As Long newPos = 0
               
                  ' Search the text to replace
                  Do
                     ' Set the start and end positions, and search flags, and finally do the search
                     SendMessageW( pDoc->hWindow, SCI_SETTARGETSTART, newPos, 0)
                     SendMessageW( pDoc->hWindow, SCI_SETTARGETEND, endPos, 0)
                     SendMessageW( pDoc->hWindow, SCI_SETSEARCHFLAGS, findFlags, 0)
                     newPos = SendMessageW( pDoc->hWindow, SCI_SEARCHINTARGET, Len(gFind.txtFind), Cast(LPARAM, Strptr(gFind.txtFind)) )
                     If newPos = -1 Then Exit Do
                     SendMessageW( pDoc->hWindow, SCI_REPLACETARGET, -1, Cast(LPARAM, Strptr(gFind.txtReplace)))
                     numChanges = numChanges + 1
                     newPos = newPos + 1
                  Loop
                  gTTabCtl.SetTabText(idx)
                  
               End If
            Next

            wszText = Str(numChanges) & " " & L(172,"replacements.")
            SetWindowTextW( GetDlgItem(HWnd, IDC_FRMREPLACE_LBLSTATUS), wszText)
            ShowWindow GetDlgItem(HWnd, IDC_FRMREPLACE_LBLSTATUS), SW_SHOW 
         End If
         
   End Select

   Function = 0
End Function



' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmReplace
' ========================================================================================
Function frmReplace_OnClose(HWnd As HWnd) As LRESULT
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmReplace
' ========================================================================================
Function frmReplace_OnDestroy(HWnd As HWnd) As LRESULT 
   ' Save the search parameters for later instances of the find/replace dialog
   gFind.nWholeWord = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMREPLACE_CHKWHOLEWORDS) )  
   gFind.nMatchCase = Button_GetCheck( GetDlgItem(HWnd, IDC_FRMREPLACE_CHKMATCHCASE) )    
   If Button_GetCheck( GetDlgItem(HWnd, IDC_FRMREPLACE_OPTSCOPE1) ) Then gFind.nScopeReplace = IDC_FRMREPLACE_OPTSCOPE1      
   If Button_GetCheck( GetDlgItem(HWnd, IDC_FRMREPLACE_OPTSCOPE2) ) Then gFind.nScopeReplace = IDC_FRMREPLACE_OPTSCOPE2

   HWND_FRMREPLACE = 0
   Function = 0
End Function


' ========================================================================================
' frmReplace Window procedure
' ========================================================================================
Function frmReplace_WndProc( ByVal HWnd   As HWnd, _
                             ByVal uMsg   As UINT, _
                             ByVal wParam As WPARAM, _
                             ByVal lParam As LPARAM _
                             ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmReplace_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmReplace_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmReplace_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmReplace_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmReplace_Show
' ========================================================================================
Function frmReplace_Show( ByVal hWndParent As HWnd ) As Long

   ' We don't want the Find dialog to display at the same time as the Replace form.
   If IsWindowVisible(HWND_FRMFIND) Then DestroyWindow(HWND_FRMFIND)
   
   ' If the Find window already exists then simply set focus to it and exit.
   If IsWindow(HWND_FRMREPLACE) Then
      SetForegroundWindow(HWND_FRMMAIN)
      SetFocus GetDlgItem( HWND_FRMREPLACE, IDC_FRMREPLACE_COMBOFIND )
      Exit Function
   End If
  
      
   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMREPLACE = _
   pWindow->Create(hwndParent, L(173,"Replace"), @frmReplace_WndProc, 0, 0, 440, 190, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(434, 161)
   pWindow->Center
   
   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_FRMREPLACE_LBLFINDWHAT, L(159,"Find what:"), 5, 11, 82, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", pWindow->hWindow, IDC_FRMREPLACE_COMBOFIND, "", 95, 8, 230, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("COMBOBOX", pWindow->hWindow, IDC_FRMREPLACE_COMBOREPLACE, "", 95, 37, 230, 21, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", pWindow->hWindow, IDOK, L(160,"Find Next"), 340, 7, 81, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", pWindow->hWindow, IDC_FRMREPLACE_CMDREPLACE, L(173,"Replace"), 340, 37, 81, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or  BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", pWindow->hWindow, IDC_FRMREPLACE_CMDREPLACEALL, L(174,"Replace All"), 340, 67, 81, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or  BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", pWindow->hWindow, IDCANCEL, L(161,"Close"), 340, 110, 81, 24, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", pWindow->hWindow, IDC_FRMREPLACE_CHKWHOLEWORDS, L(162,"Match Whole Words"), 28, 82, 132, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", pWindow->hWindow, IDC_FRMREPLACE_CHKMATCHCASE, L(163,"Match Case"), 28, 102, 132, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", pWindow->hWindow, IDC_FRMREPLACE_FRAMESCOPE, L(164,"Scope"), 174, 67, 150, 70, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", pWindow->hWindow, IDC_FRMREPLACE_OPTSCOPE1, L(165,"Current Document"), 183, 85, 132, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("OPTION", pWindow->hWindow, IDC_FRMREPLACE_OPTSCOPE2, L(166,"All Open Documents"), 183, 105, 132, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTORADIOBUTTON Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", pWindow->hWindow, IDC_FRMREPLACE_FRAMEOPTIONS, L(167,"Options"), 18, 67, 150, 70, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_FRMREPLACE_LBLREPLACEWITH, L(175,"Replace with:"), 5, 40, 82, 16, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_FRMREPLACE_LBLSTATUS, "Status", 19, 141, 368, 16, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
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
            gFind.txtFind = buffer
            gFind.txtLastFind = buffer
         End If
      Else
         gFind.txtFind = gFind.txtLastFind
      End If
   End If

   frmFind_AddToFindCombo(gFind.txtFind)


   If gFind.nScopeReplace = 0 Then gFind.nScopeReplace = IDC_FRMREPLACE_OPTSCOPE1
   Button_SetCheck( GetDlgItem(HWND_FRMREPLACE, IDC_FRMREPLACE_OPTSCOPE1), Iif(gFind.nScopeReplace = IDC_FRMREPLACE_OPTSCOPE1, BST_CHECKED, BST_UNCHECKED) )
   Button_SetCheck( GetDlgItem(HWND_FRMREPLACE, IDC_FRMREPLACE_OPTSCOPE2), Iif(gFind.nScopeReplace = IDC_FRMREPLACE_OPTSCOPE2, BST_CHECKED, BST_UNCHECKED) )
   
   SetWindowTextW( GetDlgItem(HWND_FRMREPLACE, IDC_FRMREPLACE_COMBOFIND), WStr(gFind.txtFind) )    
   SetWindowTextW( GetDlgItem(HWND_FRMREPLACE, IDC_FRMREPLACE_COMBOREPLACE), WStr(gFind.txtReplace) )    
   EnableWindow( GetDlgItem(HWND_FRMREPLACE, IDOK), True )
 
   Button_SetCheck( GetDlgItem(HWND_FRMREPLACE, IDC_FRMREPLACE_CHKWHOLEWORDS), gFind.nWholeWord )
   Button_SetCheck( GetDlgItem(HWND_FRMREPLACE, IDC_FRMREPLACE_CHKMATCHCASE), gFind.nMatchCase )

   ShowWindow HWND_FRMREPLACE, SW_SHOW
   SetFocus GetDlgItem( HWND_FRMREPLACE, IDC_FRMREPLACE_COMBOFIND )
   
   Function = 0
End Function
