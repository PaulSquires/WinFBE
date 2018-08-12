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
   SendMessage( HWnd, CB_RESETCONTENT, 0, 0)
   hDC = GetDC(0)
   EnumFontFamilies hDC, ByVal 0, Cast(FONTENUMPROCW, @EnumFontName), Cast(LPARAM, HWnd)
   ReleaseDC 0, hDC
End Sub


' ========================================================================================
' Fill combobox with list of Theme names
' ========================================================================================
function FillThemesCombo( byval hComboThemes as hwnd ) as long    
   SendMessage( hComboThemes, CB_RESETCONTENT, 0, 0)
   for i as long = lbound(gConfig.ThemesTemp) to ubound(gConfig.ThemesTemp)
      ComboBox_AddString( hComboThemes, gConfig.ThemesTemp(i).wszDescription )
   NEXT
   function = 0
end function      


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
         wText = AfxGetComboBoxText(HWnd, lpdis->itemID)
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

   dim as hwnd hWndCombo = GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_COMBOTHEMES)
   gConfig.idxTheme = Combobox_GetCurSel( hWndCombo )
   
   ' Must increase the nCurSel value because it is zero based and the color
   ' offset is between CLR_CARET and CLR_WINDOW
   Select Case nCtrlID 
      Case IDC_FRMOPTIONSCOLORS_CBCOLOR1 
         idx = ComboBox_GetCurSel( hCB1 )
         nFg = ComboBox_GetItemData( hCB1, idx )
         gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+CLR_CARET).nFg = nFg
      Case IDC_FRMOPTIONSCOLORS_CBCOLOR2 
         idx = ComboBox_GetCurSel( hCB2 )
         nBg = ComboBox_GetItemData( hCB2, idx )
         gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+CLR_CARET).nBg = nBg
   End Select
   
   gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+CLR_CARET).bFontBold = _
      Button_GetCheck( GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTBOLD) )
   gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+CLR_CARET).bFontItalic = _
      Button_GetCheck( GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTITALIC) )
   gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+CLR_CARET).bFontUnderline = _
      Button_GetCheck( GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTUNDERLINE) )

   Function = 0
End Function


' ========================================================================================
' ShowComboColors
' ========================================================================================
Function ShowComboColors( ByVal HWnd As HWnd ) As Long
           
   Dim as long NumItems, foundat, nItem, i 
   Dim as Boolean fHideBack, fShowStyles
   Dim as COLORREF nFg, nBg
   
   Dim As HWnd hList = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_LSTCOLORS )
   Dim As HWnd hCB1  = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_CBCOLOR1 )
   Dim As HWnd hCB2  = GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_CBCOLOR2 )

   Dim As Long nCurSel = Listbox_GetCurSel( hList )
   If nCurSel < 0 Then Exit Function

   dim as hwnd hWndCombo = GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_COMBOTHEMES)
   gConfig.idxTheme = Combobox_GetCurSel( hWndCombo )
   ' Must increase the nCurSel value because it is zero based and the color
   ' offset is between CLR_CARET and CLR_WINDOW
   nFg = gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+ CLR_CARET).nFg 
   nBg = gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+ CLR_CARET).nBg 

   ' Search for the color values as stored in the combobox's item area
   NumItems = ComboBox_GetCount(hCB1)
   foundat = MODCBCOLOR_USERSELECTED
   For i = 0 To NumItems - 1
      nItem = ComboBox_GetItemData(hCB1, i)
      If nItem = nFG Then
         foundat = i
         Exit For
      End If
   Next
   if foundat = MODCBCOLOR_USERSELECTED THEN
      SendMessage(hCB1, MSG_USER_SETCOLORCUSTOM, nFG, 0)
   END IF
   ComboBox_SetCurSel(hCB1, foundat)
         
   NumItems = ComboBox_GetCount(hCB2)
   foundat = MODCBCOLOR_USERSELECTED
   For i = 0 To NumItems - 1
      nItem = ComboBox_GetItemData(hCB2, i)
      If nItem = nBG Then
         foundat = i
         Exit For
      End If
   Next
   if foundat = MODCBCOLOR_USERSELECTED THEN
      SendMessage(hCB2, MSG_USER_SETCOLORCUSTOM, nBG, 0)
   END IF
   ComboBox_SetCurSel(hCB2, foundat)

   If nCurSel = 0  Then fHideBack = True   ' Caret
   If nCurSel = 2  Then fHideBack = True   ' Highlight current line
   If nCurSel = 4  Then fHideBack = True   ' Fold Margin
   If nCurSel = 13 Then fHideBack = True   ' Window
   ShowWindow hCB2, Iif( fHideBack, SW_HIDE, SW_SHOW )
   ShowWindow GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_LBLBACKGROUND), Iif( fHideBack, SW_HIDE, SW_SHOW )

   if nCurSel = 1  THEN fShowStyles = true    ' comments
   if nCurSel = 3  THEN fShowStyles = true    ' keywords
   if nCurSel = 6  THEN fShowStyles = true    ' line numbers
   if nCurSel = 7  THEN fShowStyles = true    ' operators
   if nCurSel = 9  THEN fShowStyles = true    ' preprocessor
   if nCurSel = 11 THEN fShowStyles = true    ' strings
   if nCurSel = 12 THEN fShowStyles = true    ' text
   Button_SetCheck( GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTBOLD), _
      gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+CLR_CARET).bFontBold)
   Button_SetCheck( GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTITALIC), _
      gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+CLR_CARET).bFontItalic)
   Button_SetCheck( GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTUNDERLINE), _
      gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+CLR_CARET).bFontUnderline)
   ShowWindow GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTBOLD), Iif( fShowStyles, SW_SHOW, SW_HIDE)
   ShowWindow GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTITALIC), Iif( fShowStyles, SW_SHOW, SW_HIDE)
   ShowWindow GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_CHKFONTUNDERLINE), Iif( fShowStyles, SW_SHOW, SW_HIDE)

   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptionsColors
' ========================================================================================
Function frmOptionsColors_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   select CASE id
      case IDM_CREATE_THEME
         dim as hwnd hWndCombo = GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_COMBOTHEMES)
         gConfig.idxTheme = Combobox_GetCurSel( hWndCombo )
         dim cwsText as CWSTR
         cwsText = AfxInputBox( hwnd, 0, 0, L(336,"New Theme"), L(337,"Theme Name"), "")
         if len(cwsText) THEN
            dim as long ub = ubound(gConfig.ThemesTemp) + 1
            redim preserve gConfig.ThemesTemp(ub) as TYPE_THEMES
            ' Copy all of the current theme attributes to the new theme.
            gConfig.ThemesTemp(ub) = gConfig.ThemesTemp(gConfig.idxTheme)
            gConfig.ThemesTemp(ub).id = AfxGuidText(AfxGuid())
            gConfig.ThemesTemp(ub).wszDescription = cwsText
            FillThemesCombo(hWndCombo)
            ub = ComboBox_GetCount(hWndCombo) - 1
            ComboBox_SetCurSel(hWndCombo, ub )
            Listbox_SetCurSel( GetDlgItem(HWND, IDC_FRMOPTIONSCOLORS_LSTCOLORS), 0 )
            ShowComboColors( HWND_FRMOPTIONSCOLORS )   
         END IF
         
      case IDM_DELETE_THEME
         dim as hwnd hWndCombo = GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_COMBOTHEMES)
         gConfig.idxTheme = Combobox_GetCurSel( hWndCombo )
         ' Don't allow the first (Default) theme to be deleted.
         If gConfig.idxTheme > 0 Then 
            if MessageBox( HWND, L(338, "Are you sure you want to delete this theme?"), L(276,"Confirm"), _
                              MB_YESNOCANCEL Or MB_ICONINFORMATION Or MB_DEFBUTTON1 ) = IDYES then
               if ubound(gConfig.ThemesTemp) = 0 THEN
                  erase gConfig.ThemesTemp
                  gConfig.idxTheme = -1
               else
                  ' remove the item from the internal array
                  for i as long = gConfig.idxTheme to ubound(gConfig.ThemesTemp) - 1
                     gConfig.ThemesTemp(i) = gConfig.ThemesTemp(i+1)
                  NEXT
                  redim preserve gConfig.ThemesTemp(ubound(gConfig.ThemesTemp)-1)
               END IF
               ' reload the combobox
               FillThemesCombo(hWndCombo)
               ComboBox_SetCurSel(hWndCombo, 0 )
               Listbox_SetCurSel( GetDlgItem(HWND, IDC_FRMOPTIONSCOLORS_LSTCOLORS), 0 )
               ShowComboColors( HWND_FRMOPTIONSCOLORS )   
            end if
         END IF

      CASE IDM_IMPORT_THEME
         dim as hwnd hWndCombo = GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_COMBOTHEMES)
         gConfig.idxTheme = Combobox_GetCurSel( hWndCombo )
         dim as CWSTR wszFilename = _
            AfxOpenFileDialog( hwnd, _
                               L(346,"Import theme"), "", curdir, _
                               L(79,"All files")+" (*.*)|*.*", "" )
         if AfxFileExists(wszFilename) THEN
            ' Attempt to open as unicode and then fallback as ansi
            dim pStream as CTextStream
            if pStream.Open(wszFilename, IOMode_ForReading, false, AfxIsFileUnicode(wszFilename)) <> S_OK then exit function
            Do Until pStream.EOS
               Dim as CWSTR wst = pStream.ReadLine
               gConfig.ImportTheme(wst, true)
            loop
            pStream.Close
            ' reload the combobox
            FillThemesCombo(hWndCombo)
            ComboBox_SetCurSel(hWndCombo, 0 )
            Listbox_SetCurSel( GetDlgItem(HWND, IDC_FRMOPTIONSCOLORS_LSTCOLORS), 0 )
            ShowComboColors( HWND_FRMOPTIONSCOLORS )   
            Exit Function
         END IF
         
   END SELECT
   
   Select Case codeNotify

      Case CBN_SELENDOK
         If (id = IDC_FRMOPTIONSCOLORS_CBCOLOR1) Or (id = IDC_FRMOPTIONSCOLORS_CBCOLOR2) Then
            dim as long nCurSel = Combobox_GetCurSel(hwndCtl)
            if nCurSel = MODCBCOLOR_USERSELECTED THEN
               PostMessage(hwnd, MSG_USER_GETCOLORCUSTOM, id, 0)
            else   
               SetTempColorSelection(HWnd, id)
            END IF
         end if
         
         If id = IDC_FRMOPTIONSCOLORS_COMBOTHEMES Then 
            gConfig.idxTheme = Combobox_GetCurSel( hwndCtl )
            if gConfig.idxTheme < 0 THEN exit function
            Listbox_SetCurSel( GetDlgItem(HWND, IDC_FRMOPTIONSCOLORS_LSTCOLORS), 0 )
            ShowComboColors( HWND_FRMOPTIONSCOLORS )
            ' Make sure the text in the edit control is not highlighted
            PostMessage( hwndCtl, CB_SETEDITSEL, 0, MAKELPARAM(-1,0) )   
         End If

         
      Case CBN_EDITCHANGE
         ' hwndCtl is the handle to the edit box
         If id = IDC_FRMOPTIONSCOLORS_COMBOTHEMES Then 
            if gConfig.idxTheme < 0 THEN exit function
            gConfig.ThemesTemp(gConfig.idxTheme).wszDescription = AfxGetWindowText(hwndCtl)
            ComboBox_ReplaceString(hwndCtl, gConfig.idxTheme, gConfig.ThemesTemp(gConfig.idxTheme).wszDescription)
            ' Make sure the text in the edit control is not highlighted
            PostMessage( hwndCtl, CB_SETEDITSEL, 0, MAKELPARAM(-1,0) )   
         End If

      Case LBN_SELCHANGE
         If id = IDC_FRMOPTIONSCOLORS_LSTCOLORS Then 
            PostMessage(HWND, MSG_USER_SHOWCOLORCOMBOBOXES, 0, 0)
         End If
      
      case BN_CLICKED
         dim as hwnd hWndCombo = GetDlgItem(hwnd, IDC_FRMOPTIONSCOLORS_COMBOTHEMES)
         gConfig.idxTheme = Combobox_GetCurSel( hWndCombo )
         
         if (id = IDC_FRMOPTIONSCOLORS_CHKFONTBOLD) orelse _
            (id = IDC_FRMOPTIONSCOLORS_CHKFONTITALIC) orelse _
            (id = IDC_FRMOPTIONSCOLORS_CHKFONTUNDERLINE) then
            SetTempColorSelection(HWnd, id)
         end if
         if id = IDC_FRMOPTIONSCOLORS_BTNACTIONS then
            ' Get screen coordinates of the button
            dim as RECT rcButton
            GetWindowRect(hwndCtl, @rcButton)
            DIM pt AS POINT = (rcButton.left, rcButton.bottom)
            'ClientToScreen(pNmh->hwndFrom, @pt)
            ' Don't allow the first (Default) theme to be deleted.
            dim as long fGrayed = iif(gConfig.idxTheme = 0, MF_GRAYED, MF_ENABLED)
            ' Create a menu and add items
            DIM hSplitMenu AS HMENU = CreatePopupMenu
            AppendMenu(hSplitMenu, MF_BYPOSITION, IDM_CREATE_THEME, L(348,"Create new theme"))
            AppendMenu(hSplitMenu, MF_BYPOSITION, IDM_IMPORT_THEME, L(346,"Import theme"))
            AppendMenu(hSplitMenu, MF_BYPOSITION or fGrayed, IDM_DELETE_THEME, L(347,"Delete theme"))
            ' Display the menu
            TrackPopupMenu(hSplitMenu, TPM_LEFTALIGN OR TPM_TOPALIGN, pt.x, pt.y, 0, hwnd, NULL)
         end if
            
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmMain
' ========================================================================================
Function frmOptionsColors_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

   ' Processs notify messages sent by the split button
   DIM pNmh AS NMHDR PTR = CAST(NMHDR PTR, pNMHDR)
   IF (id = IDC_FRMOPTIONSCOLORS_BTNACTIONS) ANDalso (pNmh->code = BCN_DROPDOWN) THEN
      PostMessage(hwnd, WM_COMMAND, MAKEWPARAM(IDC_FRMOPTIONSCOLORS_BTNACTIONS, BN_CLICKED), cast(LPARAM, pNMHDR->hwndFrom))
      FUNCTION = CTRUE
      EXIT FUNCTION
   END IF

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
      HANDLE_MSG (HWnd, WM_NOTIFY,  frmOptionsColors_OnNotify)
      
      Case WM_DRAWITEM  ' must pass this one on to ownerdrawn combo
         If (wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR1) Or _
            (wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR2) Then
            SendMessage( GetDlgItem(HWnd, wParam), uMsg, wParam, lParam)
            Return True
         End If
         If (wParam = IDC_FRMOPTIONSCOLORS_COMBOFONTNAME) Then 
            DrawFontCombo( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTNAME), wParam, lParam )
            Return True 
         End If

     Case WM_MEASUREITEM
         If (wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR1) OrElse _
            (wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR2) OrElse _
            (wParam = IDC_FRMOPTIONSCOLORS_COMBOFONTNAME) Then
            Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
            Dim pMeasureItem As MEASUREITEMSTRUCT Ptr = Cast(MEASUREITEMSTRUCT Ptr, lParam)
            pMeasureItem->itemHeight = pMeasureItem->itemHeight * pWindow->ryRatio
            Return True
         End If

      case MSG_USER_SHOWCOLORCOMBOBOXES
         ShowComboColors( HWnd )
   
      case MSG_USER_GETCOLORCUSTOM
         ' wParam contains the id of the color combobox
         dim as hwnd hCombo = GetDlgItem(hwnd, wParam)
         dim lCustomColor(15) AS LONG
         dim as COLORREF clrResult
         
         ' Set the default color to be whatever the current color is for the
         ' listbox item that is currently selected.
         dim as long nCurSel = ListBox_GetCurSel(GetDlgItem( HWnd, IDC_FRMOPTIONSCOLORS_LSTCOLORS ))
         If wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR1 then
            lCustomColor(0) = gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+ CLR_CARET).nFg 
            clrResult = AfxChooseColorDialog( hwnd, lCustomColor(0), @lCustomColor(0))
            if clrResult <> -1 then 
               gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+ CLR_CARET).nFg = clrResult
            end if   
         end if
         if wParam = IDC_FRMOPTIONSCOLORS_CBCOLOR2 Then         
            lCustomColor(0) = gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+ CLR_CARET).nBg 
            clrResult = AfxChooseColorDialog( hwnd, lCustomColor(0), @lCustomColor(0))
            if clrResult <> -1 then 
               gConfig.ThemesTemp(gConfig.idxTheme).colors(nCurSel+ CLR_CARET).nBg = clrResult
            end if
         end if
         ShowComboColors(hwnd)
         
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

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
   
   Dim As HWnd hComboThemes = _
   pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSCOLORS_COMBOTHEMES, "", 11, 22, 200, 23, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWN, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)   
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSCOLORS_BTNACTIONS, L(345,"&Actions"), 224, 21, 146, 25, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_SPLITBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hList = _
   pWindow->AddControl("LISTBOX",, IDC_FRMOPTIONSCOLORS_LSTCOLORS, "", 11, 50, 200, 225, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("GROUPBOX",, IDC_FRMOPTIONSCOLORS_FRMCOLORS, L(133,"Colors"), 0, 0, 385, 285, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL",, IDC_FRMOPTIONSCOLORS_LBLFOREGROUND, L(134,"Color:"), 224, 71, 140, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL",, IDC_FRMOPTIONSCOLORS_LBLBACKGROUND, L(135,"Background:"), 224, 130, 140, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
        
   CreateCBColorList(pWindow->hWindow, IDC_FRMOPTIONSCOLORS_CBCOLOR1, _
                 pWindow->ScaleX(224), pWindow->ScaleY(91), pWindow->ScaleX(146), pWindow->ScaleY(200) )
   
   CreateCBColorList(pWindow->hWindow, IDC_FRMOPTIONSCOLORS_CBCOLOR2, _
                 pWindow->ScaleX(224), pWindow->ScaleY(150), pWindow->ScaleX(146), pWindow->ScaleY(200) )
                 
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSCOLORS_CHKFONTBOLD, L(342,"Bold"), 224, 210, 400, 20, _
        WS_CHILD Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSCOLORS_CHKFONTITALIC, L(343,"Italic"), 224, 231, 400, 20, _
        WS_CHILD Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMOPTIONSCOLORS_CHKFONTUNDERLINE, L(344,"Underline"), 224, 252, 400, 20, _
        WS_CHILD Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   pWindow->AddControl("GROUPBOX",, IDC_FRMOPTIONSCOLORS_FRMFONT, L(136,"Font (applies to all styles)"), 0, 291, 385, 55, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hComboFontname = _
      pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSCOLORS_COMBOFONTNAME, "", 11, 311, 150, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST Or CBS_OWNERDRAWFIXED Or CBS_HASSTRINGS Or CBS_SORT, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   Dim As HWnd hComboFontCharSet = _
      pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET, "", 170, 311, 150, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   Dim As HWnd hComboFontSize = _
      pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE, "", 330, 311, 40, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   ' Load the lstList of colors
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(138,"Caret")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(139,"Comments")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(140,"Highlighted Line")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(141,"Keywords")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(142,"Fold Margin")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(143,"Fold Symbols")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(145,"LineNumbers")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(146,"Operators")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(144,"Indentation Guides")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(147,"Preprocessor")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(148,"Selection")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(149,"Strings")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(150,"Text")) )
   SendMessage( hList, LB_ADDSTRING, 0, Cast(LPARAM, @L(152,"Window")) )


   ' Copy all of the Themes to the ThemesTemp array because we will work with 
   ' temporary copies until the user hits OK.
   redim gConfig.ThemesTemp(ubound(gConfig.Themes))
   for i as long = lbound(gConfig.Themes) to ubound(gConfig.Themes)
      if gConfig.Themes(i).id = gConfig.SelectedTheme THEN gConfig.idxTheme = i
      gConfig.ThemesTemp(i) = gConfig.Themes(i)
   NEXT
   FillThemesCombo(hComboThemes)   
   ComboBox_SetCurSel( hComboThemes, gConfig.idxTheme )

   Listbox_SetCurSel( hList, 0 )
   ShowComboColors( HWND_FRMOPTIONSCOLORS )   

   FillFontCombo( hComboFontName )
   FillFontSizeCombo( hComboFontSize, gConfig.EditorFontname )
   FillFontCharSets( hComboFontCharSet )
   
   Dim idx As Long
   idx = SendMessage( hComboFontName, CB_FINDSTRING, 0, Cast(LPARAM, *gConfig.EditorFontname))
   ComboBox_SetCurSel( hComboFontName, idx )
   idx = SendMessage( hComboFontSize, CB_FINDSTRING, 0, Cast(LPARAM, *gConfig.EditorFontSize))
   ComboBox_SetCurSel( hComboFontSize, idx )
   idx = SendMessage( hComboFontCharSet, CB_FINDSTRING, 0, Cast(LPARAM, *gConfig.EditorFontCharSet))
   ComboBox_SetCurSel( hComboFontCharSet, idx )

   Function = 0
End Function
