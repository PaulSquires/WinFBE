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

#include once "frmOptionsColors.bi"
#include once "clsConfig.bi"


' ========================================================================================
' Enumerate the names of all the fonts. Note the difference between how to enumerate them
' (%TMPF_FIXED_PITCH has the bit cleared).
' %TMPF_FIXED_PITCH for fixed pitch fonts (like in PB edit)
' %TMPF_TRUETYPE OR %TMPF_VECTOR for True type and vector fonts
' %TMPF_DEVICE for device fonts (like printer fonts)
' Exclude what you don't want to include in list.
' ========================================================================================
Function frmOptionsColors_EnumFontName( _
            byref lf As LOGFONTW, _
            byref tm As TEXTMETRIC, _
            ByVal FontType As Long, _
            HWnd As HWnd _
            ) As Long

   static wszPrevText as CWSTR
   if gApp.isWineActive then
      ' Wine doesn't display the WYSIWYG version of the font. It will just be duplicate names
      ' so ensure that only one version of the font name displays.
      if lf.lfFaceName = wszPrevText then return true
      wszPrevText = lf.lfFaceName
   end if
   
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
function frmOptionsColors_FillFontCombo( ByVal HWnd As HWnd ) as long
   Dim hDC As HDC = GetDC(0)
   SendMessage( HWnd, CB_RESETCONTENT, 0, 0)
   EnumFontFamilies hDC, ByVal 0, Cast(FONTENUMPROCW, @frmOptionsColors_EnumFontName), Cast(LPARAM, HWnd)
   ReleaseDC 0, hDC
   function = 0
End function


' ========================================================================================
' WM_DRAWITEM procedure
' ========================================================================================
Function frmOptionsColors_DrawFontCombo( _
            ByVal HWnd As HWnd, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As Long

   Dim hFont As HFONT
   Dim lpdis As DRAWITEMSTRUCT Ptr
   Dim wText As WString * MAX_PATH
   Dim wPrevText As WString * MAX_PATH

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
            if gApp.isWineActive then
               SelectObject(lpdis->hDC, pWindow->Font)
            else
               hFont = pWindow->CreateFont( wText, 10 ) 
               If hFont Then hFont = SelectObject(lpdis->hDC, hFont)
            end if
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
function frmOptionsColors_FillFontSizeCombo( _
            ByVal hCb As HWnd, _
            ByVal strFontName As WString Ptr _
            ) as long

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

   function = 0
End function


' ========================================================================================
' Fill combobox with list of valid character sets
' ========================================================================================
function frmOptionsColors_FillFontCharSets( ByVal hCtl As HWnd ) as long
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
   
   function = 0
End function


' ========================================================================================
' frmOptionsColors Window procedure
' ========================================================================================
Function frmOptionsColors_WndProc( _
            ByVal HWnd   As HWnd, _
            ByVal uMsg   As UINT, _
            ByVal wParam As WPARAM, _
            ByVal lParam As LPARAM _
            ) As LRESULT

   Select Case uMsg
      
     Case WM_MEASUREITEM
         If wParam = IDC_FRMOPTIONSCOLORS_COMBOFONTNAME Then
            Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
            Dim pMeasureItem As MEASUREITEMSTRUCT Ptr = Cast(MEASUREITEMSTRUCT Ptr, lParam)
            pMeasureItem->itemHeight = pWindow->ScaleY(pMeasureItem->itemHeight)
            Return True
         End If

      Case WM_DRAWITEM  ' must pass this one on to ownerdrawn combo
         If (wParam = IDC_FRMOPTIONSCOLORS_COMBOFONTNAME) Then 
            frmOptionsColors_DrawFontCombo( GetDlgItem(HWnd, IDC_FRMOPTIONSCOLORS_COMBOFONTNAME), wParam, lParam )
            Return True 
         End If

   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' GetThemeDescription
' ========================================================================================
function GetThemeDescription( byval wszFilename as CWSTR ) as string
   ' default to the theme description being the filename
   function = AfxGetFileName(wszFilename)

   dim pStream AS CTextStream   
   if pStream.Open(wszFilename) <> S_OK then exit function  'error

   dim as CWSTR wst, keyValue, keyData
      
   Do Until pStream.EOS
      wst = pStream.ReadLine

      wst = trim(AfxStrExtract( 1, wst, "#"))     ' remove comments
      If Len(wst) = 0 Then Continue Do

      keyValue = trim(AfxStrParse(wst, 1, ":"))
      keyData = trim(AfxStrParse(wst, 2, ":"))
      
      ' is this a replaceable parameter
      if keyValue = "general.description" then
         function = keyData
         exit do
      end if

   loop

   pStream.Close

end function


' ========================================================================================
' frmOptionsColors_Show
' ========================================================================================
Function frmOptionsColors_Show( ByVal hWndParent As HWnd ) as LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOPTIONSCOLORS = pWindow->Create( hWndParent, "", @frmOptionsColors_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.

   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSCOLORS_LBLTHEMES, _
        L(422,"Themes") & ":", 0, 2, 375, 15, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   dim as HWND hListBox = _
   pWindow->AddControl("LISTBOX", , IDC_FRMOPTIONSCOLORS_LSTTHEMES, "", 0, 20, 386, 220, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   
   pWindow->AddControl("GROUPBOX",, IDC_FRMOPTIONSCOLORS_FRMFONT, L(136,"Font (applies to all styles)"), 0, 266, 385, 80, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)
   Dim As HWnd hComboFontname = _
      pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSCOLORS_COMBOFONTNAME, "", 11, 286, 310, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST Or CBS_OWNERDRAWFIXED Or CBS_HASSTRINGS Or CBS_SORT, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   Dim As HWnd hComboFontSize = _
      pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSCOLORS_COMBOFONTSIZE, "", 330, 286, 40, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   Dim As HWnd hComboFontCharSet = _
      pWindow->AddControl("COMBOBOX", , IDC_FRMOPTIONSCOLORS_COMBOFONTCHARSET, "", 11, 318, 150, 23, _
           WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
           WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL",, IDC_FRMOPTIONSCOLORS_LBLEXTRASPACE, L(421,"Extra line spacing") & ":", 170, 320, 140, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSCOLORS_TXTEXTRASPACE, "0", 330, 318, 40, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   AfxSetWindowText( GetDlgItem( HWND_FRMOPTIONSCOLORS, IDC_FRMOPTIONSCOLORS_TXTEXTRASPACE), gConfig.FontExtraSpace)

   frmOptionsColors_FillFontCombo( hComboFontName )
   frmOptionsColors_FillFontSizeCombo( hComboFontSize, gConfig.EditorFontname )
   frmOptionsColors_FillFontCharSets( hComboFontCharSet )
   
   Dim idx As Long
   idx = SendMessage( hComboFontName, CB_FINDSTRING, 0, Cast(LPARAM, *gConfig.EditorFontname))
   ComboBox_SetCurSel( hComboFontName, idx )
   idx = SendMessage( hComboFontSize, CB_FINDSTRING, 0, Cast(LPARAM, *gConfig.EditorFontSize))
   ComboBox_SetCurSel( hComboFontSize, idx )
   idx = SendMessage( hComboFontCharSet, CB_FINDSTRING, 0, Cast(LPARAM, *gConfig.EditorFontCharSet))
   ComboBox_SetCurSel( hComboFontCharSet, idx )
   
   ' Load the Themes listbox
   DIM hSearch as HANDLE
   dim WFD AS WIN32_FIND_DATAW
   DIM wszPath AS WSTRING * MAX_PATH
   dim wszCurPath AS WSTRING * MAX_PATH
   dim wszFullPath AS WSTRING * (MAX_PATH * 2)
   dim wszDisplayName AS WSTRING * (MAX_PATH * 2)
   dim as long idxMatch = -1
   
   wszPath = AfxGetExePathName + "themes\"
   wszCurPath = wszPath + "*"

   erase gThemeFilenames

   idx = 0
   hSearch = FindFirstFile( wszCurPath, @WFD )
   IF hSearch <> INVALID_HANDLE_VALUE THEN
      DO
         if (WFD.cFileName <> ".") andalso (WFD.cFileName <> "..") then
            wszFullPath = wszPath & WFD.cFileName
            wszDisplayName = GetThemeDescription(wszFullPath)
            idx = SendMessage( hListBox, LB_ADDSTRING, 0, cast(LPARAM, @wszDisplayName) )
            redim preserve gThemeFilenames(idx) as CWSTR
            gThemeFilenames(idx) = WFD.cFileName
            if ucase(gConfig.ThemeFilename) = ucase(WFD.cFileName ) then
               idxMatch = idx
            end if
         end if   
      LOOP WHILE FindNextFile( hSearch, @WFD )
      FindClose( hSearch )
   END IF
   SendMessage( hListBox, LB_SETCURSEL, idxMatch, 0 )

   Function = 0
End Function
