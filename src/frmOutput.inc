'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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

#include once "frmOutput.bi"
#include once "clsApp.bi"


' ========================================================================================
' Clear data from all of the controls in the frmOutput windows. This is needed
' when Projects are loaded and closed.
' ========================================================================================
private function frmOutput_ResetAllControls() as long 
   ListView_DeleteAllItems( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTVIEW) )
   AfxSetWindowText( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), "" )
   ListBox_ResetContent( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTSEARCH) )
   ListView_DeleteAllItems( GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVTODO) )
   AfxSetWindowText( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), "" )
   function = 0
END FUNCTION


' ========================================================================================
' Ensure that the correct notes are shown
' ========================================================================================
public function frmOutput_ShowNotes() as long 

   dim hCtl as hwnd = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES)

   if gApp.IsProjectActive THEN
      AfxSetWindowText(hCtl, gApp.ProjectNotes)
   else   
      AfxSetWindowText(hCtl, gApp.NonProjectNotes)
   END IF

   function = 0
END FUNCTION


' ========================================================================================
' Update the TODO listview
' ========================================================================================
public function frmOutput_UpdateToDoListview() as long 
   dim as hwnd hLV = GetDlgItem( HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVTODO)
   ListView_DeleteAllItems( hLV )
   dim as long n = 0

   dim pData as DB2_DATA ptr
   
   gdb2.dbRewind()
   do 
      pData = gdb2.dbGetNext
      if pData = 0 THEN exit do
      if pData->id <> DB2_TODO THEN continue do
      FF_ListView_InsertItem( hLV, n, 0, WStr(pData->nLineNum) ) 
      FF_ListView_InsertItem( hLV, n, 1, WStr(pData->fileName) )
      FF_ListView_InsertItem( hLV, n, 2, WStr(pData->ElementValue) ) 
      n = n + 1
   loop

   function = 0
END FUNCTION


' ========================================================================================
' Process the search result that was double clicked
' ========================================================================================
private function frmOutput_DoSearchFileSelected() as LONG

   dim wszText as wstring * MAX_PATH
   dim wszFilename as wstring * MAX_PATH
   dim as long nCursel, nLineNum

   dim as HWND hList = GetDlgItem(hWnd_FRMOUTPUT, IDC_FRMOUTPUT_LISTSEARCH)
   nCurSel = ListBox_GetCurSel( hList )
   SendMessage( hList, LB_GETTEXT, nCurSel, cast(LPARAM, @wszText))
   
   'X:\FB\WinFBE - Editor\license.txt(1): WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
   
   wszFilename = AfxStrExtract( 1, wszText, "(" )
   nLineNum    = val( AfxStrExtract( wszText, "(", ")" ) ) - 1
   
   OpenSelectedDocument( wszFilename, "", nLineNum )

   function = 0
end function
  

' ========================================================================================
' Show/Hide correct child controls
' ========================================================================================
public Function frmOutput_ShowHideOutputControls( ByVal HWnd As HWnd ) As LRESULT

   Dim As HWND hTabCtl

   ' By default, hide all controls
   ShowWindow GetDlgItem(hWnd, IDC_FRMOUTPUT_LISTVIEW), SW_HIDE
   ShowWindow GetDlgItem(hWnd, IDC_FRMOUTPUT_TXTLOGFILE), SW_HIDE
   ShowWindow GetDlgItem(hWnd, IDC_FRMOUTPUT_LISTSEARCH), SW_HIDE
   ShowWindow GetDlgItem(hWnd, IDC_FRMOUTPUT_LVTODO), SW_HIDE
   ShowWindow GetDlgItem(hWnd, IDC_FRMOUTPUT_TXTNOTES), SW_HIDE
   
   hTabCtl = GetDlgItem(hWnd, IDC_FRMOUTPUT_TABCONTROL)
   
   Select case TabCtrl_GetCurSel(hTabCtl)
      case 0    ' compiler results
         ShowWindow( GetDlgItem(hWnd, IDC_FRMOUTPUT_LISTVIEW), SW_SHOW )
      case 1    ' compiler log file
         ShowWindow( GetDlgItem(hWnd, IDC_FRMOUTPUT_TXTLOGFILE), SW_SHOW )
      case 2    ' search results
         ShowWindow( GetDlgItem(hWnd, IDC_FRMOUTPUT_LISTSEARCH), SW_SHOW )
      case 3    ' ToDo list
         ShowWindow( GetDlgItem(hWnd, IDC_FRMOUTPUT_LVTODO), SW_SHOW )
      case 4    ' Notes
         ShowWindow( GetDlgItem(hWnd, IDC_FRMOUTPUT_TXTNOTES), SW_SHOW )
   end select
         
   Function = 0
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
public Function frmOutput_PositionWindows() As LRESULT
   
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(HWND_FRMOUTPUT)
   If pWindow = 0 Then Exit Function
   
   Dim As Long nTop  = pWindow->ScaleY(2)
   Dim As Long nLeft = 0 
   Dim As Long nCtlHeight, nCtlWidth
   Dim As Rect rc
   
   GetClientRect HWND_FRMOUTPUT, @rc
   
   Dim As HWND hBtn    = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_BTNCLOSE) 
   dim as hwnd hTabCtl = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABCONTROL)

   nCtlWidth = AfxGetWindowWidth(hBtn) + pWindow->ScaleX(2)
   SetWindowPos hBtn, 0, rc.Right - nCtlWidth, nTop, 0, 0, SWP_NOZORDER or SWP_NOSIZE
   nCtlHeight = AfxGetWindowHeight(hTabCtl)
   SetWindowPos hTabCtl, 0, nLeft, nTop, rc.Right - nCtlWidth, nCtlHeight, SWP_NOZORDER

   ' Position the child controls
   nTop = nTop + nCtlHeight 
   SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTLOGFILE), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
   SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTVIEW), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
   SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LISTSEARCH), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
   SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_LVTODO), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER
   SetWindowPos GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES), 0, nLeft, nTop, rc.Right, rc.Bottom - nTop, SWP_NOZORDER

   ' Determine which child controls should be shown or hidden
   frmOutput_ShowHideOutputControls(HWND_FRMOUTPUT)
   
   Function = 0
End Function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmOutput
' ========================================================================================
private Function frmOutput_OnSize( ByVal HWnd As HWnd, _
                                   ByVal state As UINT, _
                                   ByVal cx As Long, _
                                   ByVal cy As Long _
                                   ) As LRESULT
   If state <> SIZE_MINIMIZED Then
      ' Position all of the child windows
      frmOutput_PositionWindows
   End If
   Function = 0
End Function
   

' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmOutput
' ========================================================================================
private Function frmOutput_OnCommand( ByVal HWnd As HWnd, _
                                      ByVal id As Long, _
                                      ByVal hwndCtl As HWnd, _
                                      ByVal codeNotify As UINT _
                                      ) As LRESULT
   Select case codeNotify

      case LBN_DBLCLK 
         if (id = IDC_FRMOUTPUT_LISTSEARCH) orelse _
              (id = IDC_FRMOUTPUT_LVTODO) then
            frmOutput_DoSearchFileSelected()   
            Exit Function
         End If
      
      case EN_CHANGE
         ' Notes have been modified. Save them to the correct global variable
         ' to ensure that the changes are not lost when documents or projects
         ' are switched.
         if id = IDC_FRMOUTPUT_TXTNOTES THEN
            if gApp.IsProjectActive THEN
               gApp.ProjectNotes = AfxGetWindowText(hwndCtl)
            else
               gApp.NonProjectNotes = AfxGetWindowText(hwndCtl)
            end if   
         END IF   
         
   end select
                       
   Function = 0
End Function
      

' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmOutput
' ========================================================================================
private Function frmOutput_OnNotify( ByVal HWnd As HWnd, _
                                     ByVal id As Long, _
                                     ByVal pNMHDR As NMHDR Ptr _
                                     ) As LRESULT
   Select Case pNMHDR->code 
      
      Case TCN_SELCHANGING
      Case TCN_SELCHANGE
         If id = IDC_FRMOUTPUT_TABCONTROL Then
            frmOutput_PositionWindows
         End If
      case NM_CLICK
         If id = IDC_FRMOUTPUT_BTNCLOSE Then
            ShowWindow HWND_FRMOUTPUT, SW_HIDE
            frmMain_PositionWindows
         end if
   End Select
            
   Function = 0
End Function

            
' ========================================================================================
' Process WM_PAINT message for window/dialog: frmOutput
' ========================================================================================
private Function frmOutput_OnPaint( ByVal HWnd As HWnd) As LRESULT
            
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
' Processes messages for the subclassed frmOutput Compile Results and TODO listviews .
' ========================================================================================
private Function frmOutput_Listview_SubclassProc ( _
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
         SetDocumentErrorPosition( HWND, gCompile.CompileID )    
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
' Process WM_MOUSEMOVE message for window/dialog: frmOutput
' ========================================================================================
private Function frmOutput_OnMouseMove( ByVal HWnd As HWnd, _
                                        ByVal x as long, _
                                        byval y as long, _
                                        byval keyflags as UINT _
                                        ) As Long

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMMAIN)
   If pWindow = 0 Then Exit Function

   ' HITTEST (PANELS TOP/BOTTOM SPLITTER)

   ' These are relative to <hWnd>
   dim as POINT pt = (x, y)
   dim as RECT rc
   GetWindowRect HWND_FRMOUTPUT, @rc
   rc.Bottom = rc.Top + pWindow->ScaleY(3)
   GetCursorPos(@pt)
   If PtInRect( @rc, pt ) Then
      if WindowFromPoint(pt) = HWND_FRMOUTPUT then
         SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZENS), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
      end if
   End If
   If gApp.bDragActive Then
      If gApp.hWndPanel = HWND_FRMOUTPUT Then
         Dim As Long nDiff = pt.y - rc.Top
         ' Adjust the height. The positioning will be taken care of in PositionMainWindows().
         GetClientRect HWND_FRMOUTPUT, @rc
         rc.Top = rc.Top + nDiff
         SetWindowPos HWND_FRMOUTPUT, 0, 0, 0, rc.Right - rc.Left, rc.Bottom-rc.Top, SWP_NOMOVE Or SWP_NOZORDER
         frmMain_PositionWindows
         Exit Function
      End If
   End If

   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONDOWN message for window/dialog: frmOutput
' ========================================================================================
private Function frmOutput_OnLButtonDown( ByVal HWnd As HWnd, _
                                          byval fDoubleClick as Boolean, _
                                          ByVal x as long, _
                                          byval y as long, _
                                          byval keyflags as UINT _
                                          ) As Long

   DIM pWindow AS CWindow PTR = AfxCWindowPtr(HWND_FRMMAIN)
   If pWindow = 0 Then Exit Function

   ' HITTEST (PANELS TOP/BOTTOM SPLITTER)
   Dim As Rect rc
   Dim As Point pt 

   gApp.bDragActive = False 
   
   GetWindowRect HWND_FRMOUTPUT, @rc
   rc.Bottom = rc.Top + pWindow->ScaleY(3)
   GetCursorPos(@pt)
   If PtInRect( @rc, pt ) Then
      if WindowFromPoint(pt) = HWND_FRMOUTPUT then
         SetCursor LoadImage( Null, MAKEINTRESOURCEW(OCR_SIZENS), IMAGE_CURSOR, 0, 0, LR_DEFAULTSIZE Or LR_SHARED )
         gApp.bDragActive = True 
         gApp.hWndPanel   = HWND_FRMOUTPUT
         SetCapture( HWND_FRMOUTPUT )
      end if
      Exit Function
   End If

   function = 0
end function


' ========================================================================================
' Process WM_LBUTTONUP message for window/dialog: frmOutput
' ========================================================================================
private Function frmOutput_OnLButtonUp( ByVal HWnd As HWnd, _
                                        ByVal x as long, _
                                        byval y as long, _
                                        byval keyflags as UINT _
                                        ) As Long

   ' HITTEST (PANELS TOP/BOTTOM SPLITTER)
   if gApp.bDragActive then
      gApp.bDragActive = False 
      gApp.hWndPanel = 0
      ReleaseCapture()
   end if
   SetCursor( LoadCursor( null, IDC_ARROW ))
   
   function = 0
end function


' ========================================================================================
' frmOutput Window procedure
' ========================================================================================
private Function frmOutput_WndProc( ByVal HWnd   As HWnd, _
                                    ByVal uMsg   As UINT, _
                                    ByVal wParam As WPARAM, _
                                    ByVal lParam As LPARAM _
                                    ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,     frmOutput_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmOutput_OnNotify)
      HANDLE_MSG (HWnd, WM_SIZE,        frmOutput_OnSize)
      HANDLE_MSG (HWnd, WM_PAINT,       frmOutput_OnPaint)
      HANDLE_MSG (HWnd, WM_LBUTTONUP,   frmOutput_OnLButtonUp)
      HANDLE_MSG (HWnd, WM_LBUTTONDOWN, frmOutput_OnLButtonDown)
      HANDLE_MSG (HWnd, WM_MOUSEMOVE,   frmOutput_OnMouseMove)
  
   CASE WM_DESTROY
      DeleteFont( AfxGetWindowFont(GetDlgItem(HWnd, IDC_FRMOUTPUT_TXTLOGFILE)) )
      DIM pImageCtx AS CImageCtx ptr = AfxCImageCtxPtr(hwnd, IDC_FRMOUTPUT_BTNCLOSE)
      if pImageCtx then delete pImageCtx
         
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmOutput_Show
' ========================================================================================
public Function frmOutput_Show( ByVal hWndParent As HWnd ) As LRESULT

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowPtr(hwndParent)->DPI

   HWND_FRMOUTPUT = pWindow->Create( hWndParent, "", @frmOutput_WndProc, 0, 0, 0, 180, _
        WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

   Dim As HWnd hCtl, hLV, hLB
   
   ' Add the top tabcontrol (initially not visible)
   hCtl = pWindow->AddControl("TABCONTROL", , _
       IDC_FRMOUTPUT_TABCONTROL, "", 0, 0, 0, 24, _
       WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or _
       TCS_TABS Or TCS_FOCUSNEVER, WS_EX_LEFT Or WS_EX_LTRREADING)

   TabCtrl_AddTab( hCtl, -1, L(191, "Compiler Results") )
   TabCtrl_AddTab( hCtl, -1, L(252, "Compiler Log File") )
   TabCtrl_AddTab( hCtl, -1, L(262, "Search Results") )
   TabCtrl_AddTab( hCtl, -1, L(263, "TODO") )
   TabCtrl_AddTab( hCtl, -1, L(264, "Notes") )
   TabCtrl_SetCurSel( hCtl, 0 )
   
   DIM pImageCtx AS CImageCtx ptr 
   pImageCtx = new CImageCtx(pWindow, IDC_FRMOUTPUT_BTNCLOSE, "", 0, 0, 24, 24)
   pImageCtx->SetBkColor(GetSysColor(COLOR_BTNFACE))
   pImageCtx->SetBkColorHot(GetSysColor(COLOR_BTNHIGHLIGHT))
   pImageCtx->LoadImageFromResource (pWindow->InstanceHandle, wstr("IMAGE_CLOSE"))
   AfxAddTooltip(pImageCtx->hWindow, L(161, "Close"))

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
      Cast(SUBCLASSPROC, @frmOutput_Listview_SubclassProc), IDC_FRMOUTPUT_LISTVIEW, Cast(DWORD_PTR, @pWindow))
                                     
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
       
   ' Search results listbox
   hLB = _
      pWindow->AddControl("LISTBOX", , IDC_FRMOUTPUT_LISTSEARCH, "", 0, 0, 0, 0, _
      WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
      WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   AfxSetWindowFont( hLB, hFont, True )

   ' TODO listview
   hLV = _
      pWindow->AddControl("LISTVIEW", , IDC_FRMOUTPUT_LVTODO, "", 0, 0, 0, 0, _
      WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or LVS_REPORT Or LVS_SHOWSELALWAYS or LVS_SINGLESEL, _
      WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR, , _
      Cast(SUBCLASSPROC, @frmOutput_Listview_SubclassProc), IDC_FRMOUTPUT_LVTODO, Cast(DWORD_PTR, @pWindow))
                                     
   ' Make the listview header flat
   ListView_MakeHeaderFlat(hLV)
          
   ' Add some extended styles
   dwExStyle = ListView_GetExtendedListViewStyle(hLV)
   dwExStyle = dwExStyle Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_DOUBLEBUFFER Or LVS_EX_FLATSB
   ListView_SetExtendedListViewStyle(hLV, dwExStyle)

   ' Configure the ListView
   ListView_AddColumn( hLV, 0, L(253, "Line"), pWindow->ScaleX(75) )
   ListView_AddColumn( hLV, 1, L(254, "File"), pWindow->ScaleX(250) )
   ListView_AddColumn( hLV, 2, L(255, "Description"), pWindow->ScaleX(500) )

   hCtl = _
      pWindow->AddControl("EDITMULTILINE", , IDC_FRMOUTPUT_TXTNOTES, "", 0, 0, 0, 0, _
      WS_VISIBLE OR WS_TABSTOP OR WS_VSCROLL or ES_LEFT OR ES_AUTOVSCROLL OR ES_MULTILINE OR ES_NOHIDESEL OR ES_WANTRETURN, _
      WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)
   hFont = pWindow->CreateFont( "Courier New", 9 )
   AfxSetWindowFont( hCtl, hFont, True )
   AfxSetWindowText(hCtl, gApp.NonProjectNotes)
   
   frmOutput_PositionWindows
   
   Function = 0

End Function


