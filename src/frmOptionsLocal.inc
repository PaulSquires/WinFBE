'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

#include once "frmOptionsLocal.bi"


' ========================================================================================
' If a localization file has been edited then save it.
' ========================================================================================
public function frmOptionsLocal_LocalEditCheck() as Long
   if gLocalPhrasesEdit = false then exit function
               
   dim as CWSTR wszLocalName = AfxGetWindowText( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION) )
   dim as CWSTR wszFileName = AfxGetExePathName + wstr("Languages\" + wszLocalName)
   dim as CWSTR wszTemp, wszPhrase
   
   ' If the english file is selected then we do not want to allow any modifications to that file.
   if ucase(trim(AfxStrPathname( "NAMEX", wszFilename ))) = "ENGLISH.LANG" then exit function

   dim as HWND hLV = GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LVWPHRASES)
   dim as long nCount = ListView_GetItemCount(hLV)
   dim wszIndex as wstring * MAX_PATH
   
   AfxDeleteFile(wszFileName)
   
   dim pStream as CTextStream
   IF pStream.OpenForOutputW(wszFilename, true) = S_OK THEN

      wszTemp = _
      "' WinFBE localization file for " & ucase(AfxStrPathName("NAME", wszLocalName)) & " language" & vbcrlf & _
      "'" & vbcrlf & _
      "' The first element to exist must be called MAXIMUM and needs to be set to" & vbcrlf & _
      "' the value of the last key in this file." & vbcrlf & _ 
      "'" & vbcrlf & _
      "' This file should be created and saved using UTF-16 encoding (unicode)." & vbcrlf & _
      "'" & vbcrlf & _  
      "' Each line represents a key/value pair describing the position to store the" & vbcrlf & _
      "' localized word/phrase into the localization array used by WinFBE. Simple." & vbcrlf & vbcrlf & _
      "MAXIMUM:" & wstr(nCount) & vbcrlf & vbcrlf
      pStream.WriteLine(wszTemp) 

      for i as long = 0 to nCount - 1
         FF_ListView_GetItemText(hLV, i, 0, @wszIndex, MAX_PATH )
         wszTemp = wspace(1024)
         wszPhrase = AfxStrRSet(wszIndex, 5, "0") + wstr(":") + gLocalPhrases(val(wszIndex))
         ' Use ** version of CWSTR on MID statements per the documentation.
         mid(**wszTemp, 1) = wszPhrase
         if len(gLangEnglish(val(wszIndex))) then
            ' Always attempt to insert the ENGLISH comment at position 60. If the local
            ' phrase is too long then simply put the ENGLISH comment after it.
            dim as long nPos = iif( len(wszPhrase) < 60, 60, len(wszPhrase) + 5)
            mid(**wszTemp, nPos) = "; " & gLangEnglish(val(wszIndex))
         end if   
         pStream.WriteLine rtrim(wszTemp)
      next
   end if
   pStream.Close
   
   function = 0
end function


' ========================================================================================
' Do not allow the English language file to be edited or deleted
' ========================================================================================
private function frmOptionsLocal_DisableEnglish() as Long
   dim as boolean bEnable, bEnglish
   Dim as CWSTR wszFilename = ucase(AfxGetWindowText(GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION)))
   
   ' If the english file is selected then we do not want to allow any editing so disable the controls.
   bEnglish = iif( AfxStrPathname( "NAMEX", wszFilename ) = "ENGLISH.LANG", true, false)

   ' Set the default case for english localization
   if bEnglish then 
      EnableWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION), true )
      EnableWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_CMDNEW), true )
      EnableWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_CMDEDIT), false )
      EnableWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_CMDDELETE), false )
      ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LVWPHRASES), SW_HIDE )
      ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LBLENGLISH), SW_HIDE )
      ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_TXTENGLISH), SW_HIDE )
      ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LBLTRANSLATE), SW_HIDE )
      ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_TXTTRANSLATE), SW_HIDE )
      ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_FRAMEEDITAREA), SW_HIDE )
      exit function
   end if
   
   ' If a non-english file is being edited then enable the fields.   
   dim as long nShow = iif(gLocalPhrasesEdit, SW_SHOW, SW_HIDE)
   ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LVWPHRASES), nShow )
   ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LBLENGLISH), nShow )
   ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_TXTENGLISH), nShow )
   ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LBLTRANSLATE), nShow )
   ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_TXTTRANSLATE), nShow )
   ShowWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_FRAMEEDITAREA), nShow )

   EnableWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION), true )
   EnableWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_CMDNEW), true )
   EnableWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_CMDDELETE), true )
   bEnable = iif(gLocalPhrasesEdit, false, true)
   EnableWindow( GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_CMDEDIT), bEnable )

   function = 0
end function


' ========================================================================================
' Clear the global localization phrase array.
' ========================================================================================
private function frmOptionsLocal_ClearGlobalArray() as Long
   ' Clear the global array contents
   for i as long = lbound(gLocalPhrases) to ubound(gLocalPhrases) 
      gLocalPhrases(i) = ""
   next
   function = 0
end function


' ========================================================================================
' Load the selected localization file into the global array
' ========================================================================================
private function frmOptionsLocal_LoadLocalizationFile( byref wszFilename as wstring ) as Long

   If AfxFileExists( wszFileName ) = 0 Then exit function

   Dim as CBSTR wst, wKey, wData
   Dim as long nKey, nData, i, nRow, maxKey
    
   dim pStream AS CTextStream
   if pStream.OpenUnicode( wszFileName ) <> S_OK then exit function
   
   frmOptionsLocal_ClearGlobalArray()

   do until pStream.EOS
      wst = pStream.ReadLine
      

      If Len(wst) = 0 Then Continue Do
      If Left(wst, 1) = "'" Then Continue Do
      
      i = Instr(wst, ":")
      If i = 0 Then Continue Do
      
      wKey = "": wData = "": nData = 0

      wKey  = Left(wst, i-1)
      wData = Mid(**wst, i+1)    ' MID causes problems with Chinese data so ** is used.
      
      nKey  = Val(wKey)
      nData = Val(wData)

      ' MAXMIUM for a non-English does not redimension the gLocalPhrases array because
      ' that array is always sized according to the main ENGLISH array.
      If Ucase(wKey) <> "MAXIMUM" Then
         if (nKey >= lbound(gLocalPhrases)) and (nKey <= ubound(gLocalPhrases)) then
            ' Use ** to ensure that cyrillic langauge gets converted correctly. FB intrinsic
            ' functions (RTRIM) automatically convert those incorrectly when using CBSTR or CWSTR.
            gLocalPhrases(nKey) = rtrim(**AfxStrParse(wData, 1, ";"), any chr(9,32))
            ' If the English phrase is blank then blank the translation phrase
            if gLangEnglish(nKey) = "" then gLocalPhrases(nKey) = ""
            ' If the Translated phrase is blank then replace it with the English version
            if gLocalPhrases(nKey) = "" then gLocalPhrases(nKey) = gLangEnglish(nKey)
            if nKey > maxKey then maxKey = nKey
         end if
      End If   
         
   Loop
   pStream.Close
   
   ' It is possible that the number of keys in the local file is less than the number
   ' of keys in the full english file. In that case we need process the remaining 
   ' missing keys by copying the english version over to the local array.
   for i as long = maxKey + 1 to ubound(gLangEnglish)
      gLocalPhrases(i) = gLangEnglish(i)
   next
   
   function = 0
end function


' ========================================================================================
' Update the translation textboxes
' ========================================================================================
private Function frmOptionsLocal_UpdateTextBoxes( ByVal HWnd As HWnd ) As LRESULT
   dim wszIndex     as wstring * MAX_PATH
   dim wszEnglish   as wstring * MAX_PATH
   dim wszTranslate as wstring * MAX_PATH
   dim as long idx
   
   dim as HWND hLV = GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_LVWPHRASES)
   Dim As Long nCurSel = ListView_GetSelection(hLV)
   If nCurSel < 0 Then exit function
   FF_ListView_GetItemText(hLV, nCurSel, 0, @wszIndex, MAX_PATH )
   FF_ListView_GetItemText(hLV, nCurSel, 1, @wszEnglish, MAX_PATH )
   idx = val(wszIndex)
   if (idx >= lbound(gLocalPhrases)) and (idx <= ubound(gLocalPhrases)) then
      wszTranslate = gLocalPhrases(idx)
      AfxSetWindowText( GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_TXTENGLISH), wszEnglish)
      AfxSetWindowText( GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_TXTTRANSLATE), wszTranslate)
   end if
   function = 0
end function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOptionsLocal
' ========================================================================================
private Function frmOptionsLocal_OnCommand( ByVal HWnd As HWnd, _
                                            ByVal id As Long, _
                                            ByVal hwndCtl As HWnd, _
                                            ByVal codeNotify As UINT _
                                            ) As LRESULT
   Select Case id

      Case IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION
         ' select localization file
         If codeNotify = BN_CLICKED Then
            ' If a localization file is currently being edited then save it now.
            frmOptionsLocal_LocalEditCheck

            ' Display the Open File Dialog
            Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, id)
            If pwszName Then
               frmOptionsLocal_LoadLocalizationFile(*pwszName)
               gLocalPhrasesEdit = false
               ' Clear any previous listview selected line
               dim as HWND hLV = GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LVWPHRASES)
               ListView_UnselectAllItems(hLV)
               AfxSetWindowText( GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_TXTENGLISH), "")
               AfxSetWindowText( GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_TXTTRANSLATE), "")
               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION), AfxStrPathname("NAMEX", *pwszName) )
               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_LBLTRANSLATE), AfxStrPathName("NAME", *pwszName) )
               frmOptionsLocal_DisableEnglish()
               CoTaskMemFree pwszName
            End If
            Exit Function
         End If

      Case IDC_FRMOPTIONSLOCAL_CMDNEW
         ' New localization file
         If codeNotify = BN_CLICKED Then
            ' If a localization file is currently being edited then save it now.
            frmOptionsLocal_LocalEditCheck
            
            ' Display the Save File Dialog
            Dim wzFilename As WString * MAX_PATH 
            wzFilename = "*.lang"
            Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWND, @wzFilename, @wstr("lang"), id)
            If pwszName Then
               wzFilename = *pwszName
               frmOptionsLocal_ClearGlobalArray()
               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION), AfxStrPathname("NAMEX", *pwszName) )
               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_LBLTRANSLATE), AfxStrPathName("NAME", *pwszName) )
               gLocalPhrasesEdit = true
               frmOptionsLocal_DisableEnglish()
               dim as HWND hLV = GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LVWPHRASES)
               ListView_SelectItem(hLV, 0)
               frmOptionsLocal_UpdateTextBoxes(HWND)
               SetFocus( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_TXTTRANSLATE) )
               CoTaskMemFree(pwszName)
            End If
            Exit Function
         End If

      Case IDC_FRMOPTIONSLOCAL_CMDEDIT
         ' Edit the loaded localization file
         If codeNotify = BN_CLICKED Then
            dim as CWSTR wszLanguage = _
               AfxGetExePathName & wstr("Languages\") & AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION) )
            frmOptionsLocal_LoadLocalizationFile(wszLanguage)
            AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_LBLTRANSLATE), AfxStrPathName("NAME", wszLanguage) )
            gLocalPhrasesEdit = true
            frmOptionsLocal_DisableEnglish()
            dim as HWND hLV = GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LVWPHRASES)
            ListView_SelectItem(hLV, 0)
            frmOptionsLocal_UpdateTextBoxes(HWND)
            SetFocus( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_TXTTRANSLATE) )
            Exit Function
         End If

      Case IDC_FRMOPTIONSLOCAL_CMDDELETE
         ' Delete localization file
         If codeNotify = BN_CLICKED Then
            if MessageBox( HWND, L(366,"Are you sure you want to delete?"), L(276,"Confirm"), _
                  MB_YESNOCANCEL or MB_ICONQUESTION Or MB_DEFBUTTON1 ) <> IDYES then
               exit function
            end if
            dim as CWSTR wszFileName = AfxGetExePathName + wstr("Languages\") + _
                                       AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION) )
            frmOptionsLocal_ClearGlobalArray()
            AfxDeleteFile(wszFileName)
            ' Clear any previous listview selected line
            dim as HWND hLV = GetDlgItem(HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_LVWPHRASES)
            ListView_UnselectAllItems(hLV)
            AfxSetWindowText( GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_TXTENGLISH), "")
            AfxSetWindowText( GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_TXTTRANSLATE), "")
            AfxSetWindowText( GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION), wstr("english.lang"))
            AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMOPTIONSLOCAL_LBLTRANSLATE), wstr("english") )
            frmOptionsLocal_DisableEnglish()
         end if
         exit function
         
      case IDC_FRMOPTIONSLOCAL_TXTTRANSLATE
         If codeNotify = EN_CHANGE Then
            dim wszIndex     as wstring * MAX_PATH
            dim wszTranslate as wstring * MAX_PATH
            dim as HWND hLV = GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_LVWPHRASES)
            Dim As Long nCurSel = ListView_GetSelection(hLV)
            If nCurSel >= 0 Then 
               FF_ListView_GetItemText(hLV, nCurSel, 0, @wszIndex, MAX_PATH )
               dim as long idx = val(wszIndex)
               if (idx >= lbound(gLocalPhrases)) and (idx <= ubound(gLocalPhrases)) then
                  gLocalPhrases(idx) = AfxGetWindowText( GetDlgItem(HWND, IDC_FRMOPTIONSLOCAL_TXTTRANSLATE) )
               end if
            end if
         end if
         exit function

   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmOptionsLocal
' ========================================================================================
private Function frmOptionsLocal_OnNotify( ByVal HWnd As HWnd, _
                                           ByVal id As Long, _
                                           ByVal pNMHDR As NMHDR Ptr _
                                           ) As LRESULT
   SELECT CASE id
      Case IDC_FRMOPTIONSLOCAL_LVWPHRASES
         select case pNMHDR->code
            case LVN_ITEMCHANGED 
               frmOptionsLocal_UpdateTextBoxes(HWND)
         end select
   end select

   function = 0
end function


' ========================================================================================
' frmOptionsLocal Window procedure
' ========================================================================================
private Function frmOptionsLocal_WndProc( ByVal HWnd   As HWnd, _
                                          ByVal uMsg   As UINT, _
                                          ByVal wParam As WPARAM, _
                                          ByVal lParam As LPARAM _
                                          ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOptionsLocal_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,  frmOptionsLocal_OnNotify)
   End Select

   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOptionsLocal_Show
' ========================================================================================
public Function frmOptionsLocal_Show( ByVal hWndParent As HWnd ) as LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   gLocalPhrasesEdit = false

   HWND_FRMOPTIONSLOCAL = pWindow->Create( hWndParent, "", @frmOptionsLocal_WndProc, 0, 0, 0, 0, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   ' Height and width of this child form is set in frmOptions_OnNotify when the treeview option is selected.
   
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSLOCAL_CMDLOCALIZATION, L(126,"Select"), 40, 22, 75, 26, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSLOCAL_CMDNEW, L(3,"New"), 117, 22, 75, 26, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSLOCAL_CMDEDIT, L(14,"Edit"), 194, 22, 75, 26, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMOPTIONSLOCAL_CMDDELETE, L(282,"Delete"), 271, 22, 75, 26, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("GROUPBOX", , IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION, "english.lang", 0, 0, 385, 58, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)


   dim as hwnd hwndLV = _
   pWindow->AddControl("LISTVIEW", , IDC_FRMOPTIONSLOCAL_LVWPHRASES, "", 8, 78, 368, 160, _
        WS_CHILD Or WS_VISIBLE OR WS_CLIPCHILDREN OR WS_TABSTOP OR LVS_REPORT OR LVS_SHOWSELALWAYS, _
        WS_EX_CLIENTEDGE)
   ListView_MakeHeaderFlat(hwndLV)
   ' Add some extended styles
   dim as long dwExStyle = ListView_GetExtendedListViewStyle(hwndLV)
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_DOUBLEBUFFER Or LVS_EX_FLATSB
   ListView_SetExtendedListViewStyle(hwndLV, dwExStyle)

   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSLOCAL_LBLENGLISH, "English", 8, 246, 150, 16, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSLOCAL_TXTENGLISH, "", 8, 264, 368, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL or ES_READONLY, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("LABEL", , IDC_FRMOPTIONSLOCAL_LBLTRANSLATE, "English", 8, 292, 150, 16, _
        WS_CHILD Or WS_VISIBLE Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMOPTIONSLOCAL_TXTTRANSLATE, "", 8, 310, 368, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   pWindow->AddControl("GROUPBOX", , IDC_FRMOPTIONSLOCAL_FRAMEEDITAREA, "Phrases", 0, 58, 385, 288, _
        WS_CHILD Or WS_VISIBLE Or BS_TEXT Or BS_LEFT Or BS_NOTIFY Or BS_GROUPBOX, _
        WS_EX_TRANSPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING)

   AfxSetWindowText(GetDlgItem( HWND_FRMOPTIONSLOCAL, IDC_FRMOPTIONSLOCAL_FRAMELOCALIZATION), gConfig.LocalizationFile)
   frmOptionsLocal_DisableEnglish()

   ' Display the columns in the Listview and load the English phrases
   ListView_AddColumn( hwndLV, 0, @wstr("Index"), pWindow->ScaleX(50) )
   ListView_AddColumn( hwndLV, 1, @wstr("English Phrase"), pWindow->ScaleX(280))

   ' Always size the local phrase array to be the same size as the ENGLISH array.
   ReDim gLocalPhrases(Ubound(gLangEnglish)) 
   frmOptionsLocal_LoadLocalizationFile(gConfig.LocalizationFile)
   
   dim as long nRow = 0
   for i as long = LBound(gLangEnglish) to Ubound(gLangEnglish)
      ' Add to the English phrase Listview
      dim as CWSTR wKey = AfxStrRset(str(i), 5, "0")
      FF_ListView_InsertItem(hwndLV, nRow, 0, wKey, 0 )
      FF_ListView_InsertItem(hwndLV, nRow, 1, gLangEnglish(i), 0 )
      nRow = nRow + 1
   next
   
   Function = 0
   
End Function

