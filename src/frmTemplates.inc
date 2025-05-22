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

#include once "frmTemplates.bi"


' ========================================================================================
' Process WM_CREATE message for window/dialog: frmTemplates
' ========================================================================================
private Function frmTemplates_OnCreate( ByVal HWnd As HWnd, _
                                        ByVal lpCreateStructPtr As LPCREATESTRUCT _
                                        ) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   EnableWindow GetParent(HWnd), False

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmTemplates
' ========================================================================================
private Function frmTemplates_OnCommand( ByVal HWnd As HWnd, _
                                         ByVal id As Long, _
                                         ByVal hwndCtl As HWnd, _
                                         ByVal codeNotify As UINT _
                                         ) As LRESULT
   Select Case id
      Case IDOK  
         If codeNotify = BN_CLICKED Then
            ' Let the IDC_FRMTEMPLATES_LISTBOX message to process it
            SendMessage( HWnd, WM_COMMAND, MAKELONG(IDC_FRMTEMPLATES_LISTBOX, LBN_DBLCLK), 0 )
            Exit Function
         End If
      
      Case IDCANCEL ' button clicked or ESC pressed
         If codeNotify = BN_CLICKED Then
            SendMessage ( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If

      Case IDC_FRMTEMPLATES_LISTBOX
         If codeNotify = LBN_DBLCLK Then
            ' Get the handle of the Listbox
            Dim hListBox As HWnd = GetDlgItem(HWnd, IDC_FRMTEMPLATES_LISTBOX)
            ' Get the current selection
            Dim curSel As Long = SendMessage(hListBox, LB_GETCURSEL, 0, 0)
            If curSel = LB_ERR Then Exit Function
            ' Get the stored index
            Dim pwszPath As WString Ptr = Cast(WString Ptr, SendMessage(hListBox, LB_GETITEMDATA, Cast(WPARAM, curSel), 0))
            If pwszPath = 0 Then Exit Function
            ' Open the template
            If pwszPath Then
               dim pDoc as clsDocument ptr
               pDoc = frmMain_OpenFileSafely( _
                           HWND_FRMMAIN, _
                           False, _    ' bIsNewFile
                           true, _     ' bIsTemplate
                           True,  _    ' bShowInTab
                           false, _    ' bIsInclude
                           *pwszPath, _  ' wszName
                           0 )         ' pDocIn
   
               ' Ensure that the code window scrolls to the correct location
               dim as hwnd hEdit = pDoc->hWndActiveScintilla
               Dim As Long curPos = SciExec( hEdit, SCI_GETCURRENTPOS, 0, 0)
               SciExec( hEdit, SCI_GOTOPOS, curPos, 0 )

            End If
            ' Close the dialog
            SendMessage HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmTemplates
' ========================================================================================
private Function frmTemplates_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow GetParent(HWnd), True 
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmTemplates
' ========================================================================================
private Function frmTemplates_OnDestroy( byval HWnd As HWnd ) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' Templates window procedure
' ========================================================================================
private Function frmTemplates_WndProc( ByVal HWnd As HWnd, _
                                       ByVal uMsg As UINT, _
                                       ByVal wParam As WPARAM, _
                                       ByVal lParam As LPARAM _
                                       ) As LRESULT
   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmTemplates_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmTemplates_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmTemplates_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmTemplates_OnCommand)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Templates popup dialog
' ========================================================================================
public Function frmTemplates_Show( ByVal hParent As HWnd ) As LRESULT

   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hParent)->DPI
	
   pWindow->Create( hParent, L(176,"Templates"), @frmTemplates_WndProc, _
                    0, 0, 0, 0, WS_CAPTION Or WS_POPUPWINDOW, WS_EX_WINDOWEDGE )
   pWindow->SetClientSize(420,395)
   pWindow->Center(pWindow->hWindow, hParent)
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)

   ' Add a listbox
   Dim hListBox As HWnd = pWindow->AddControl("ListBox", , IDC_FRMTEMPLATES_LISTBOX, "", _
      WS_CHILD Or WS_VISIBLE Or WS_HSCROLL Or WS_VSCROLL Or WS_BORDER Or WS_TABSTOP Or _
      LBS_STANDARD Or LBS_HASSTRINGS Or LBS_SORT Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, WS_EX_CLIENTEDGE)
      
   pWindow->SetWindowPos( hListBox, Null, 8, 8, 400, 337, SWP_NOZORDER )
   SendMessage( hListBox, LB_SETHORIZONTALEXTENT, Cast(WPARAM, 600 * pWindow->rxRatio), 0 )

   ' Add the buttons
   pWindow->AddControl("Button", , IDOK, L(0,"&OK"), 245, 353, 74, 28, _
               WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_FLAT Or BS_DEFPUSHBUTTON, WS_EX_NOPARENTNOTIFY)
   pWindow->AddControl("Button", , IDCANCEL, L(1,"&Cancel"), 333, 353, 74, 28, _
               WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_FLAT, WS_EX_NOPARENTNOTIFY)

   ' Search templates
   DIM hSearch as HANDLE
   dim WFD AS WIN32_FIND_DATAW
   dim as long FileNo, nType, nCount, nItem, idx, nLin 
    
   DIM wszPath AS WSTRING * MAX_PATH
   dim wszCurPath AS WSTRING * MAX_PATH
   dim wszText AS WSTRING * MAX_PATH
   DIM wszFullPath AS WSTRING * MAX_PATH * 2
   
   wszPath = AfxGetExePathName + "Templates\"
   wszCurPath = wszPath + "*.fbtpl"

   ' Count the number of files and dimension the array
   ' REDIM PRESERVE causes problems
   hSearch = FindFirstFile(wszCurPath, @WFD)
   IF hSearch <> INVALID_HANDLE_VALUE THEN
      DO
         IF (WFD.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY THEN
            nCount = nCount + 1
         END IF
      LOOP WHILE FindNextFile(hSearch, @WFD)
      FindClose(hSearch)
   END IF
   
   ' Dimension the array
   IF nCount = 0 THEN EXIT FUNCTION
   DIM rgwszPaths(nCount - 1) AS WSTRING * MAX_PATH
   idx = 0

   ' Find the files
   dim pStream as CTextStream
   hSearch = FindFirstFile(wszCurPath, @WFD)
   IF hSearch <> INVALID_HANDLE_VALUE THEN
      DO
         IF (WFD.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY THEN
            wszFullPath = wszPath & WFD.cFileName
            ' The start of a template file contains 4 lines of meta data
            ' For example:
            ' 2       
            ' FBGUI
            ' .bas
            ' CWindow: COM: Explorer Browser in a tab page
            
            ' Get the description
            if pStream.Open(wszFullPath) = S_OK then
               nLin = 0
               DO UNTIL pStream.EOS
                  wszText = pStream.ReadLine
                  nLin = nLin + 1
                  IF nLin = 1 THEN nType = VAL(wszText)
                  IF nType < 1 OR nType > 2 THEN EXIT DO
                  IF nType = 1 AND nLin = 3 THEN EXIT DO
                  IF nType = 2 AND nLin = 4 THEN EXIT DO
               LOOP
               pStream.Close
               
               if len(wszText) then
                  ' Display the description in the listbox
                  nItem = SendMessage(hListBox, LB_ADDSTRING, 0, cast(LPARAM, @wszText))
                  ' Store the full path in the array
                  rgwszPaths(idx) = wszFullPath
                  ' Store a pointer to the element of the array in the listbox item
                  SendMessage(hListBox, LB_SETITEMDATA, nItem, cast(LPARAM, @rgwszPaths(idx) ))
                  idx = idx + 1
                  IF idx > UBOUND(rgwszPaths) THEN EXIT DO
               end if
               
            END IF
         END IF
      LOOP WHILE FindNextFile(hSearch, @WFD)
      FindClose(hSearch)
   END IF
   
   ShowWindow( pWindow->hWindow, SW_SHOW )
   
   ' Process Windows messages
   Function = pWindow->DoEvents

   ' Delete the frmTemplates CWindow class manually allocated memory 
   Delete pWindow

End Function



