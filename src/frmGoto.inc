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
' Process WM_CREATE message for window/dialog: frmGoto
' ========================================================================================
Function frmGoto_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN

   ' This is a modal popup window so disable the parent window
   EnableWindow GetParent(HWnd), False

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmGoto
' ========================================================================================
Function frmGoto_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select Case id
      Case IDOK
         If codeNotify = BN_CLICKED Then
            Dim nLine As Long
            Dim swzText As WString * MAX_PATH
            swzText = AfxGetWindowText( GetDlgItem(HWnd, IDC_FRMGOTO_TXTLINE) )
            nLine = Val(swzText) - 1
            If nLine >= 0 Then
               Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
               If pDoc Then 
                  dim as hwnd hEdit = pDoc->hWndActiveScintilla
                  SciExec( hEdit, SCI_GOTOLINE, nLine, 0 )
               END IF
            End If
            SendMessage HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage HWnd, WM_CLOSE, 0, 0
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmGoto
' ========================================================================================
Function frmGoto_OnClose(HWnd As HWnd) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableWindow GetParent(HWnd), True 
   DestroyWindow HWnd
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmGoto
' ========================================================================================
Function frmGoto_OnDestroy(HWnd As HWnd) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmGoto Window procedure
' ========================================================================================
Function frmGoto_WndProc( ByVal HWnd   As HWnd, _
                          ByVal uMsg   As UINT, _
                          ByVal wParam As WPARAM, _
                          ByVal lParam As LPARAM _
                          ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmGoto_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmGoto_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmGoto_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmGoto_OnCommand)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmGoto_Show
' ========================================================================================
Function frmGoto_Show( ByVal hWndParent As HWnd ) As Long 

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->ClassStyle = CS_DROPSHADOW
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   pWindow->Create(hWndParent, L(168,"Go to Line"), @frmGoto_WndProc, 0, 0, 287, 126, _
        WS_POPUP Or WS_DLGFRAME Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(281, 97)
   pWindow->Center
   
   pWindow->AddControl("TEXTBOX", pWindow->hWindow, IDC_FRMGOTO_TXTLINE, "", 105, 56, 67, 23, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", pWindow->hWindow, IDOK, L(0,"&OK"), 186, 14, 78, 30, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_DEFPUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", pWindow->hWindow, IDCANCEL, L(1,"&Cancel"), 186, 51, 78, 30, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_FRMGOTO_LBLCURRENTLINE, L(170,"Current line:"), 15, 15, 81, 17, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_FRMGOTO_LBLLASTLINE, L(169,"Last line:"), 15, 36, 81, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_FRMGOTO_LBLGOTOLINE, L(171,"Go to line:"), 15, 59, 81, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_FRMGOTO_LBLCURRENTVALUE, "", 107, 15, 65, 17, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", pWindow->hWindow, IDC_FRMGOTO_LBLLASTVALUE, "", 107, 36, 65, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   Dim As Long curPos, nLine, nLines
   
   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then
      dim as hwnd hEdit = pDoc->hWndActiveScintilla
      curPos  = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)
      nLine   = SciExec(hEdit, SCI_LINEFROMPOSITION, curPos, 0) + 1
      nLines  = SciExec(hEdit, SCI_GETLINECOUNT, 0, 0)
      SetWindowText( GetDlgItem(pWindow->hWindow, IDC_FRMGOTO_LBLCURRENTVALUE), WStr(nLine) )
      SetWindowText( GetDlgItem(pWindow->hWindow, IDC_FRMGOTO_LBLLASTVALUE), WStr(nLines) )
   End If
     
   ShowWindow pWindow->hWindow, SW_SHOW
   SetFocus GetDlgItem(pWindow->hWindow, IDC_FRMGOTO_TXTLINE)

   ' Process Windows messages
   Function = pWindow->DoEvents(0)
   
   ' Delete the frmGotoCompiler CWindow class manually allocated memory 
   Delete pWindow

End Function

