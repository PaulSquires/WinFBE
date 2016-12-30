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
' Show/Hide correct child controls
' ========================================================================================
Function ShowHideOutputControls( ByVal HWnd As HWnd ) As LRESULT

   Dim As HWND hTabCtl

   ' By default, hide all controls
   ShowWindow GetDlgItem(hWnd, IDC_FRMOUTPUT_LISTVIEW), SW_HIDE
   ShowWindow GetDlgItem(hWnd, IDC_FRMOUTPUT_TXTLOGFILE), SW_HIDE
   
   hTabCtl = GetDlgItem(hWnd, IDC_FRMOUTPUT_TABCONTROL)
   
   Select case TabCtrl_GetCurSel(hTabCtl)
      case 0    ' compiler results
         ShowWindow( GetDlgItem(hWnd, IDC_FRMOUTPUT_LISTVIEW), SW_SHOW )
      case 1    ' compiler log file
         ShowWindow( GetDlgItem(hWnd, IDC_FRMOUTPUT_TXTLOGFILE), SW_SHOW )
   end select
         
   Function = 0
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function PositionOutputWindows( ByVal HWnd As HWnd ) As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWnd)
   If pWindow = 0 Then Exit Function
   
   Dim As Long nTop  = pWindow->ScaleY(2)
   Dim As Long nLeft = 0 
   Dim As Long nCtlHeight, nCtlWidth
   Dim As Rect rc
   
   GetClientRect hWnd, @rc
   
   Dim As HWND hBtn    = GetDlgItem(hWnd, IDC_FRMOUTPUT_BTNCLOSE) 
   dim as hwnd hTabCtl = GetDlgItem(hWnd, IDC_FRMOUTPUT_TABCONTROL)

   nCtlWidth = AfxGetWindowWidth(hBtn) + pWindow->ScaleX(2)
   SetWindowPos hBtn, 0, rc.Right - nCtlWidth, nTop, 0, 0, SWP_NOZORDER or SWP_NOSIZE
   nCtlHeight = AfxGetWindowHeight(hTabCtl)
   SetWindowPos hTabCtl, 0, nLeft, nTop, rc.Right - nCtlWidth, nCtlHeight, SWP_NOZORDER

   ' Position the child controls
   nTop = nTop + nCtlHeight 
   SetWindowPos GetDlgItem(hWnd, IDC_FRMOUTPUT_TXTLOGFILE), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
   SetWindowPos GetDlgItem(hWnd, IDC_FRMOUTPUT_LISTVIEW), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER

   ' Determine which child controls should be shown or hidden
   ShowHideOutputControls(hWnd)
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmOutput
' ========================================================================================
Function frmOutput_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      PositionOutputWindows(HWnd)
   End If
   Function = 0
End Function
   

' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOutput
' ========================================================================================
Function frmOutput_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Select case codeNotify
      case BN_CLICKED
         If id = IDC_FRMOUTPUT_BTNCLOSE Then
            ShowWindow HWND_FRMOUTPUT, SW_HIDE
            frmMain_PositionWindows(HWND_FRMMAIN)
         end if
   end select
                       
   Function = 0
End Function
      
' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmOutput
' ========================================================================================
Function frmOutput_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
     
   Select Case pNMHDR->code 
      
      Case TCN_SELCHANGING
      Case TCN_SELCHANGE
         If id = IDC_FRMOUTPUT_TABCONTROL Then
            PositionOutputWindows(HWND)
         End If
   End Select
            
   Function = 0
End Function

            
' ========================================================================================
' Process WM_PAINT message for window/dialog: frmOutput
' ========================================================================================
Function frmOutput_OnPaint( ByVal HWnd As HWnd) As LRESULT
            
   Dim As PAINTSTRUCT ps
   Dim As HPEN hPen
   Dim As HDC hDc
   Dim As Rect rc
   
   GetClientRect HWnd, @rc
   
   hDC = BeginPaint(hWnd, @ps)

   SaveDC hDC
 
   'hPen = CreatePen( PS_SOLID, 1, BGR(63,63,63) )
   hPen = CreatePen( PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW) )
   SelectPen( hDC, hPen )
 
   ' Paint a simple line at the top of the window that will act as a 
   ' visual separator between the Output window and the Scintilla window.
   MoveToEx hDC, rc.Left, rc.Top, Null
   LineTo   hDC, rc.Right, rc.Top
   
   RestoreDC hDC, -1 
   DeletePen(hPen)
   
   EndPaint hWnd, @ps
   
   Function = 0
End Function

' ========================================================================================
' Processes messages for the subclassed frmOutput listview control.
' ========================================================================================
Function frmOutput_Listview_SubclassProc ( _
                  ByVal HWnd   As HWnd, _                 ' // Control window handle
                  ByVal uMsg   As UINT, _                 ' // Type of message
                  ByVal wParam As WPARAM, _               ' // First message parameter
                  ByVal lParam As LPARAM, _               ' // Second message parameter
                  ByVal uIdSubclass As UINT_PTR, _        ' // The subclass ID
                  ByVal dwRefData As DWORD_PTR _          ' // Pointer to reference data
                  ) As LRESULT

   Dim As HWND hLV 
   Dim As POINT pt
   Dim As HTREEITEM hItem
   Dim As BOOLEAN bIsFolder
   Dim As BOOLEAN bIsExpanded

   ' Convert our ENTER key presses into LBUTTONDBLCLK to process them similarly
   If (uMsg = WM_KEYUP) And (Loword(wParam) = VK_RETURN) Then uMsg = WM_LBUTTONDBLCLK

      
   Select Case uMsg

      Case WM_GETDLGCODE
         ' All keyboard input
         Function = DLGC_WANTALLKEYS
         Exit Function

      Case WM_LBUTTONDBLCLK
         SetDocumentErrorPosition( HWND )    
         Exit Function
         
      Case WM_KEYUP
         Select Case Loword(wParam)
            Case VK_RETURN  ' already processed in WM_LBUTTONDBLCLK 
         End Select
         Exit Function

      Case WM_CHAR   ' prevent the annoying beep!
         If wParam = VK_RETURN Then Return 0

      Case WM_DESTROY
         ' REQUIRED: Remove control subclassing
         RemoveWindowSubclass HWnd, @frmOutput_Listview_SubclassProc, uIdSubclass

   End Select

   ' Default processing of Windows messages
   Function = DefSubclassProc(HWnd, uMsg, wParam, lParam)

End Function



' ========================================================================================
' frmOutput Window procedure
' ========================================================================================
Function frmOutput_WndProc( ByVal HWnd   As HWnd, _
                            ByVal uMsg   As UINT, _
                            ByVal wParam As WPARAM, _
                            ByVal lParam As LPARAM _
                            ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND, frmOutput_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,  frmOutput_OnNotify)
      HANDLE_MSG (HWnd, WM_SIZE,    frmOutput_OnSize)
      HANDLE_MSG (HWnd, WM_PAINT,   frmOutput_OnPaint)
  
      CASE WM_DESTROY
         DeleteFont( AfxGetWindowFont(GetDlgItem(HWnd, IDC_FRMOUTPUT_TXTLOGFILE)) )
         
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProcW(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOutput_Show
' ========================================================================================
Function frmOutput_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOUTPUT= pWindow->Create( hWndParent, "", @frmOutput_WndProc, 0, 0, 0, 180, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   Dim As HWnd hCtl, hLV
   Dim As Long nWidth  = 24
   Dim As Long nHeight = 24
   
   ' Add the top tabcontrol (initially not visible)
   hCtl = pWindow->AddControl("TABCONTROL", , _
       IDC_FRMOUTPUT_TABCONTROL, "", 0, 0, 0, 24, _
       WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or _
       TCS_TABS Or TCS_FOCUSNEVER, WS_EX_LEFT Or WS_EX_LTRREADING)

   TabCtrl_AddTab( hCtl, -1, L(191, "Compiler Results") )
   TabCtrl_AddTab( hCtl, -1, L(252, "Compiler Log File") )
   TabCtrl_SetCurSel( hCtl, 0 )
   
   frmHotImgBtn( HWND_FRMOUTPUT, IDC_FRMOUTPUT_BTNCLOSE, "IMAGE_CLOSE", "IMAGE_CLOSE",  _
                  L(161, "Close"), nWidth, nHeight, GetSysColor(COLOR_BTNFACE), 0, 0, nWidth, nHeight )
                  
   hCtl = _
      pWindow->AddControl("EDITMULTILINE", , IDC_FRMOUTPUT_TXTLOGFILE, "", 0, 0, 0, 0, _
      WS_VISIBLE OR WS_TABSTOP OR WS_VSCROLL or ES_LEFT OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_NOHIDESEL OR ES_WANTRETURN, _
      WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)
   Dim As HFONT hFont = pWindow->CreateFont( "Courier New", 9 )
   AfxSetWindowFont( hCtl, hFont, True )

   hLV = _
      pWindow->AddControl("LISTVIEW", , IDC_FRMOUTPUT_LISTVIEW, "", 0, 0, 0, 0, _
      WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS or LVS_SINGLESEL, _
      WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
      Cast(WNDPROC, @frmOutput_Listview_SubclassProc), IDC_FRMOUTPUT_LISTVIEW, Cast(DWORD_PTR, @pWindow))
                                     
   ' Make the listview header flat
   ListView_MakeHeaderFlat(hLV)
          
   ' Add some extended styles
   Dim dwExStyle As DWORD
   dwExStyle = ListView_GetExtendedListViewStyle(hLV)
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_DOUBLEBUFFER Or LVS_EX_FLATSB
   ListView_SetExtendedListViewStyle(hLV, dwExStyle)

   ' Configure the ListView
   ListView_AddColumn( hLV, 0, L(253, "Line"), pWindow->ScaleX(75) )
   ListView_AddColumn( hLV, 1, L(254, "File"), pWindow->ScaleX(250) )
   ListView_AddColumn( hLV, 2, L(255, "Description"), pWindow->ScaleX(500) )
       
   PositionOutputWindows(HWND_FRMOUTPUT)
   
   Function = 0

End Function









































