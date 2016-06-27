

' ========================================================================================
' Enumerate the names of all the fonts. Note the difference between how to enumerate them
' (%TMPF_FIXED_PITCH has the bit cleared).
' %TMPF_FIXED_PITCH for fixed pitch fonts (like in PB edit)
' %TMPF_TRUETYPE OR %TMPF_VECTOR for True type and vector fonts
' %TMPF_DEVICE for device fonts (like printer fonts)
' Exclude what you don't want to include in list.
' ========================================================================================
Function EnumFontName(lf As LOGFONTW, tm As TEXTMETRIC, ByVal FontType As Long, HWnd As HWnd) As Long

   If (FontType And TRUETYPE_FONTTYPE) Then                      ' // True type fonts
      ComboBox_AddString( HWnd, @lf.lfFaceName )
   ElseIf (FontType And TMPF_FIXED_PITCH) = 0 Then               ' <- check if bit is cleared!
      ComboBox_AddString( HWnd, @lf.lfFaceName )                        ' // fixed pitch fonts
   ElseIf (FontType And TMPF_VECTOR) Then
      ComboBox_AddString( HWnd, @lf.lfFaceName )                        ' // vector fonts
   ElseIf (FontType And TMPF_DEVICE) Then
      ComboBox_AddString( HWnd, @lf.lfFaceName )                        ' // device fonts
   Else
      ComboBox_AddString( HWnd, @lf.lfFaceName )                        ' // system, others
   End If

   Function = True

End Function


' ========================================================================================
' Fill combobox with list of valid font names
' ========================================================================================
Sub FillFontCombo( ByVal HWnd As HWnd)
   Dim hDC As HDC
   SendMessageW HWnd, CB_RESETCONTENT, 0, 0
   hDC = GetDC(0)
   EnumFontFamilies hDC, ByVal 0, Cast(FONTENUMPROCW, @EnumFontName), Cast(LPARAM, HWnd)
   ReleaseDC 0, hDC
End Sub


' ========================================================================================
' WM_DRAWITEM procedure
' ========================================================================================
Function DrawFontCombo(ByVal HWnd As HWnd, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As Long

   Dim hFont As HFONT
   Dim lpdis As DRAWITEMSTRUCT Ptr
   Dim wText As WString * MAX_PATH

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMMAIN)
   
   lpdis = Cast(DRAWITEMSTRUCT Ptr, lParam)
   If lpdis->itemID = &HFFFFFFFF Then Exit Function   ' empty list, take a break..

   Select Case lpdis->itemAction
      Case ODA_DRAWENTIRE, ODA_SELECT
         ' If not selected or if in edit part of the combobox
         If (lpdis->itemState And ODS_SELECTED) = 0 Or _
               (lpdis->itemState And ODS_COMBOBOXEDIT) Then
            SetBkColor lpdis->hDC, GetSysColor(COLOR_WINDOW)
            SetTextColor lpdis->hDC, GetSysColor(COLOR_WINDOWTEXT)
            FillRect lpdis->hDC, @lpdis->rcItem, GetSysColorBrush(COLOR_WINDOW)
         Else
            ' Select text background
            SetBkColor lpdis->hDC, GetSysColor(COLOR_HIGHLIGHT)
            SetTextColor lpdis->hDC, GetSysColor(COLOR_HIGHLIGHTTEXT)
            FillRect lpdis->hDC, @lpdis->rcItem, GetSysColorBrush(COLOR_HIGHLIGHT)
         End If

         ' Get item's text (fontname), create font and draw text
         FF_ComboBox_GetTextW(HWnd, lpdis->itemID, wText, Sizeof(wText))
         If Len(wText) Then
            hFont = pWindow->CreateFont( wText, 10 ) 
            If hFont Then hFont = SelectObject(lpdis->hDC, hFont)
         End If
         DrawTextW lpdis->hDC, @wText, Len(wText), @lpdis->rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER
         If hFont Then DeleteObject SelectObject(lpdis->hDC, hFont)

         ' Focus rect around selected item
         If (lpdis->itemState And ODS_SELECTED) Then
            DrawFocusRect lpdis->hDC, @lpdis->rcItem
         End If
         Function = True

   End Select

End Function


' ========================================================================================
' Fill combobox with list of valid font sizes
' ========================================================================================
Sub FillFontSizeCombo( ByVal hCb As HWnd, ByVal strFontName As WString Ptr )

   Select Case Ucase(*strFontName)
      Case "FIXEDSYS"
         ComboBox_ResetContent(hCb)
         ComboBox_AddString( hCb, @WStr("9"))
      Case "TERMINAL"
         ComboBox_ResetContent(hCb)
         ComboBox_AddString( hCb, @WStr("5" ))
         ComboBox_AddString( hCb, @WStr("6" ))
         ComboBox_AddString( hCb, @WStr("12"))
         ComboBox_AddString( hCb, @WStr("14"))
      Case Else                        
         ComboBox_ResetContent(hCb)
         ComboBox_AddString( hCb, @WStr("8" ))
         ComboBox_AddString( hCb, @WStr("9" ))
         ComboBox_AddString( hCb, @WStr("10"))
         ComboBox_AddString( hCb, @WStr("11"))
         ComboBox_AddString( hCb, @WStr("12"))
         ComboBox_AddString( hCb, @WStr("14"))
         ComboBox_AddString( hCb, @WStr("16"))
         ComboBox_AddString( hCb, @WStr("18"))
         ComboBox_AddString( hCb, @WStr("20"))
         ComboBox_AddString( hCb, @WStr("22"))
         ComboBox_AddString( hCb, @WStr("24"))
         ComboBox_AddString( hCb, @WStr("26"))
         ComboBox_AddString( hCb, @WStr("28"))
         ComboBox_AddString( hCb, @WStr("32"))
         ComboBox_AddString( hCb, @WStr("36"))
   End Select

End Sub

' ========================================================================================
' Fill combobox with list of valid character sets
' ========================================================================================
Sub FillFontCharSets( ByVal hCtl As HWnd )
   ComboBox_AddString( hCtl, @WStr("Default") )
   ComboBox_AddString( hCtl, @WStr("Ansi") )
   ComboBox_AddString( hCtl, @WStr("Arabic") )
   ComboBox_AddString( hCtl, @WStr("Baltic") )
   ComboBox_AddString( hCtl, @WStr("Chinese Big 5") )
   ComboBox_AddString( hCtl, @WStr("East Europe") )
   ComboBox_AddString( hCtl, @WStr("GB 2312") )
   ComboBox_AddString( hCtl, @WStr("Greek") )
   ComboBox_AddString( hCtl, @WStr("Hangul") )
   ComboBox_AddString( hCtl, @WStr("Hebrew") )
   ComboBox_AddString( hCtl, @WStr("Johab") )
   ComboBox_AddString( hCtl, @WStr("Mac") )
   ComboBox_AddString( hCtl, @WStr("OEM") )
   ComboBox_AddString( hCtl, @WStr("Russian") )
   ComboBox_AddString( hCtl, @WStr("Shiftjis") )
   ComboBox_AddString( hCtl, @WStr("Symbol") )
   ComboBox_AddString( hCtl, @WStr("Thai") )
   ComboBox_AddString( hCtl, @WStr("Turkish") )
   ComboBox_AddString( hCtl, @WStr("Vietnamese") )
End Sub


' ========================================================================================
' SetTempColorSelection
' ========================================================================================
Function SetTempColorSelection( ByVal HWnd As HWnd, ByVal nCtrlID As Long ) As Long

   Dim As COLORREF nFg, nBg
   Dim As Long nCurSel, idx
   
   Dim As HWnd hList = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_LSTCOLORS )
   Dim As HWnd hCB1  = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_CBCOLOR1 )
   Dim As HWnd hCB2  = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_CBCOLOR2 )

   nCurSel = Listbox_GetCurSel( hList )
   If nCurSel < 0 Then Exit Function

   Select Case nCtrlID 
      Case IDC_FRMOPTIONSCOLORS_CBCOLOR1 
         idx = ComboBox_GetCurSel( hCB1 )
         nFg = ComboBox_GetItemData( hCB1, idx )
         gTempColors(nCurSel).nFg = nFg
      Case IDC_FRMOPTIONSCOLORS_CBCOLOR2 
         idx = ComboBox_GetCurSel( hCB2 )
         nBg = ComboBox_GetItemData( hCB2, idx )
         gTempColors(nCurSel).nBg = nBg
   End Select
   
   Function = 0
End Function


' ========================================================================================
' ShowComboColors
' ========================================================================================
Function ShowComboColors( ByVal HWnd As HWnd, ByVal nCurSel As Long ) As Long

   Dim fHideBack As BOOLEAN
   Dim nFg       As COLORREF
   Dim nBg       As COLORREF
   Dim NumItems  As Long
   Dim nItem     As Long
   Dim i         As Long
   
   Dim As HWnd hList = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_LSTCOLORS )
   Dim As HWnd hCB1  = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_CBCOLOR1 )
   Dim As HWnd hCB2  = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_CBCOLOR2 )

   nFg = gTempColors(nCurSel).nFg 
   nBg = gTempColors(nCurSel).nBg 

   ' Search for the color values as stored in the combobox's item area
   NumItems = ComboBox_GetCount(hCB1)
   For i = 0 To NumItems - 1
      nItem = ComboBox_GetItemData(hCB1, i)
      If nItem = nFG Then
         ComboBox_SetCurSel(hCB1, i)
         Exit For
      End If
   Next
         
   NumItems = ComboBox_GetCount(hCB2)
   For i = 0 To NumItems - 1
      nItem = ComboBox_GetItemData(hCB2, i)
      If nItem = nBG Then
         ComboBox_SetCurSel(hCB2, i)
         Exit For
      End If
   Next

   If nCurSel = 1  Then fHideBack = True   ' Caret
   If nCurSel = 3  Then fHideBack = True   ' Highlight current line
   If nCurSel = 5  Then fHideBack = True   ' Fold Margin
   If nCurSel = 15 Then fHideBack = True   ' Window

   ShowWindow hCB2, Iif( fHideBack, SW_HIDE, SW_SHOW )
   ShowWindow GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_LBLBACKGROUND), Iif( fHideBack, SW_HIDE, SW_SHOW )

   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptionsColors
' ========================================================================================
Function frmOptionsColors_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case codeNotify

      Case CBN_SELENDOK
         If (id = IDC_FRMOPTIONSCOLORS_CBCOLOR1) Or (id = IDC_FRMOPTIONSCOLORS_CBCOLOR2) Then
            SetTempColorSelection(HWnd, id)
         End If

      Case LBN_SELCHANGE
         If id = IDC_FRMOPTIONSCOLORS_LSTCOLORS Then 
            Dim As Long nCurSel = Listbox_GetCurSel( hWndCtl )
            If nCurSel < 0 Then Exit Function
            ShowComboColors( HWnd, nCurSel )
         End If
         
   End Select

   Function = 0
End Function


' ========================================================================================
' frmOptionsColors Window procedure
' ========================================================================================
Function frmOptionsColors_WndProc( ByVal HWnd   As HWnd, _
                                   ByVal uMsg   As UINT, _
                                   ByVal wParam As WPARAM, _
                                   ByVal lParam As LPARAM _
                                   ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOptionsColors_OnCommand)
      
      Case WM_DRAWITEM  ' must pass this one on to ownerdrawn combo
         If (wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR1) Or _
            (wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR2) Then
            SendMessageW GetDlgItem(HWnd, wParam), uMsg, wParam, lParam
            Return True
         End If
        If (wParam = IDC_FRMOPTIONSCOLORS_COMBOFONTNAME) Then 
           DrawFontCombo( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTNAME), wParam, lParam )
           Return True 
        End If

     Case WM_MEASUREITEM
         Scope
         If (wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR1) OrElse _
            (wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR2) OrElse _
            (wParam = IDC_FRMOPTIONSCOLORS_COMBOFONTNAME) Then
            Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
            Dim pMeasureItem As MEASUREITEMSTRUCT Ptr = Cast(MEASUREITEMSTRUCT Ptr, lParam)
            pMeasureItem->itemHeight = pMeasureItem->itemHeight * pWindow->ryRatio
            Return True
         End If
         End Scope

   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsColors_Show
' ========================================================================================
Function frmOptionsColors_Show( ByVal hWndParent As HWnd, _
                                ByVal nCmdShow   As Long = 0 _
                                ) As Long

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOPTIONSCOLORS = pWindow->Create( hWndParent, "", @frmOptionsColors_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   
   Dim As HWnd hList = _
      pWindow->AddControl("LISTBOX",, IDC_FRMOPTIONSCOLORS_LSTCOLORS, "", 11, 22, 200, 203, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("GROUPBOX",, IDC_FRMOPTIONSCOLORS_FRMCOLORS, L(133,"Colors"), 0, 0, 385, 235, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL",, IDC_FRMOPTIONSCOLORS_LBLFOREGROUND, L(134,"Color:"), 224, 21, 140, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL",, IDC_FRMOPTIONSCOLORS_LBLBACKGROUND, L(135,"Background:"), 224, 80, 140, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX",, IDC_FRMOPTIONSCOLORS_FRMFONT, L(136,"Font (applies to all styles)"), 0, 241, 385, 55, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
        
   CreateCBColorList(pWindow->hWindow, IDC_FRMOPTIONSCOLORS_CBCOLOR1, _
                 pWindow->ScaleX(224), pWindow->ScaleY(41), pWindow->ScaleX(146), pWindow->ScaleY(20) )

   CreateCBColorList(pWindow->hWindow, IDC_FRMOPTIONSCOLORS_CBCOLOR2, _
                 pWindow->ScaleX(224), pWindow->ScaleY(100), pWindow->ScaleX(146), pWindow->ScaleY(20) )
                 
   Dim As HWnd hComboFontname = _
      pWindow->AddControl("COMBOBOX", pWindow->hWindow, IDC_FRMOPTIONSCOLORS_COMBOFONTNAME, "", 11, 261, 150, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST Or CBS_OWNERDRAWFIXED Or CBS_HASSTRINGS Or CBS_SORT, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   Dim As HWnd hComboFontCharSet = _
      pWindow->AddControl("COMBOBOX", pWindow->hWindow, IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET, "", 170, 261, 150, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   Dim As HWnd hComboFontSize = _
      pWindow->AddControl("COMBOBOX", pWindow->hWindow, IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE, "", 330, 261, 40, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Load the lstList of colors
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(137,"Bookmarks")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(138,"Caret")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(139,"Comments")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(140,"Highlighted Line")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(141,"Keywords")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(142,"Fold Margin")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(143,"Fold Symbols")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(144,"Indentation Guides")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(145,"LineNumbers")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(146,"Operators")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(147,"Preprocessor")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(148,"Selection")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(149,"Strings")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(150,"Text")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(151,"WinAPI Keywords")) )
   SendMessageW( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(152,"Window")) )

   ' Save the currently selected fg and bg colors to the temporary global array
   gTempColors(0).nFg  = gConfig.clrBookmarksFG         
   gTempColors(0).nBg  = gConfig.clrBookmarksBG        
   gTempColors(1).nFg  = gConfig.clrCaretFG            
   gTempColors(1).nBg  = gConfig.clrCaretBG            
   gTempColors(2).nFg  = gConfig.clrCommentsFG         
   gTempColors(2).nBg  = gConfig.clrCommentsBG         
   gTempColors(3).nFg  = gConfig.clrHighlightedFG      
   gTempColors(3).nBg  = gConfig.clrHighlightedBG      
   gTempColors(4).nFg  = gConfig.clrKeywordFG          
   gTempColors(4).nBg  = gConfig.clrKeywordBG          
   gTempColors(5).nFg  = gConfig.clrFoldMarginFG       
   gTempColors(5).nBg  = gConfig.clrFoldMarginBG       
   gTempColors(6).nFg  = gConfig.clrFoldSymbolFG       
   gTempColors(6).nBg  = gConfig.clrFoldSymbolBG       
   gTempColors(7).nFg  = gConfig.clrIndentGuidesFG     
   gTempColors(7).nBg  = gConfig.clrIndentGuidesBG     
   gTempColors(8).nFg  = gConfig.clrLineNumbersFG      
   gTempColors(8).nBg  = gConfig.clrLineNumbersBG      
   gTempColors(9).nFg  = gConfig.clrOperatorsFG        
   gTempColors(9).nBg  = gConfig.clrOperatorsBG        
   gTempColors(10).nFg = gConfig.clrPreprocessorFG     
   gTempColors(10).nBg = gConfig.clrPreprocessorBG     
   gTempColors(11).nFg = gConfig.clrSelectionFG        
   gTempColors(11).nBg = gConfig.clrSelectionBG        
   gTempColors(12).nFg = gConfig.clrStringsFG          
   gTempColors(12).nBg = gConfig.clrStringsBG          
   gTempColors(13).nFg = gConfig.clrTextFG             
   gTempColors(13).nBg = gConfig.clrTextBG             
   gTempColors(14).nFg = gConfig.clrWinApiBG           
   gTempColors(14).nBg = gConfig.clrWinApiBG           
   gTempColors(15).nFg = gConfig.clrWindowFG           
   gTempColors(15).nBg = gConfig.clrWindowBG           


   Listbox_SetCurSel( hList, 0 )
   ShowComboColors( HWND_FRMOPTIONSCOLORS, 0 )   

   FillFontCombo( hComboFontName )
   FillFontSizeCombo( hComboFontSize, @gConfig.EditorFontname )
   FillFontCharSets( hComboFontCharSet )
   
   Dim idx As Long
   idx = SendMessageW( hComboFontName, CB_FINDSTRING, 0, Cast(LPARAM, Strptr(gConfig.EditorFontname)))
   ComboBox_SetCurSel( hComboFontName, idx )
   idx = SendMessageW( hComboFontSize, CB_FINDSTRING, 0, Cast(LPARAM, Strptr(gConfig.EditorFontSize)))
   ComboBox_SetCurSel( hComboFontSize, idx )
   idx = SendMessageW( hComboFontCharSet, CB_FINDSTRING, 0, Cast(LPARAM, Strptr(gConfig.EditorFontCharSet)))
   ComboBox_SetCurSel( hComboFontCharSet, idx )

   Function = 0
End Function
