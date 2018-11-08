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
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
Function frmHelpViewer_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWnd, @rc)
   
   dim as hwnd hTabCtrl = GetDlgItem(HWnd, IDC_FRMHELPVIEWER_TABCONTROL)
   dim as hwnd hTree(1), hWebBrowser(1)
   hTree(0) = GetDlgItem(HWnd, IDC_FRMHELPVIEWER_TVWINFBE)
   hTree(1) = GetDlgItem(HWnd, IDC_FRMHELPVIEWER_TVWINFBX)
   hWebBrowser(0) = GetDlgItem(HWnd, IDC_FRMHELPVIEWER_WEBBROWSER_WINFBE )
   hWebBrowser(1) = GetDlgItem(HWnd, IDC_FRMHELPVIEWER_WEBBROWSER_WINFBX )
    
   dim as long nCurSel = TabCtrl_GetCurSel(hTabCtrl) 
   select case nCurSel
   case -1
      exit function   
   case 0
      ShowWindow( hTree(1), SW_HIDE)
      ShowWindow( hWebBrowser(1), SW_HIDE)
   case 1   
      ShowWindow( hTree(0), SW_HIDE)
      ShowWindow( hWebBrowser(0), SW_HIDE)
   end select
   
   dim as long nTabHeight = AfxScaleY(24)
   dim as long nTreeWidth = AfxScaleX(200)
   dim as long nTreeHeight = rc.Bottom-rc.top-nTabHeight-AfxScaleY(30)
   
   SetWindowPos( hTabCtrl, 0, 0, 0, rc.Right-rc.Left, nTabHeight, SWP_SHOWWINDOW Or SWP_NOZORDER)
   SetWindowPos( hTree(nCurSel), 0, 0, nTabHeight+AfxScaleY(30), nTreeWidth, nTreeHeight, SWP_SHOWWINDOW Or SWP_NOZORDER)
   SetWindowPos( hWebBrowser(nCurSel), 0, nTreeWidth, nTabHeight, rc.Right-rc.Left-nTreeWidth, rc.Bottom-rc.top-nTabHeight, SWP_SHOWWINDOW Or SWP_NOZORDER)

   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmHelpViewer
' ========================================================================================
Function frmHelpViewer_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT

   Function = 0
End Function


''
''  Display the HTML file held in the array index
''
function frmHelpViewer_DisplayHTML( byval idx as long ) as Long
   
   if idx >= lbound(gHTMLHelpFiles) andalso idx <= ubound(gHTMLHelpFiles) then
      if AfxFileExists(gHTMLHelpFiles(idx)) then
         dim as hwnd hTabCtrl = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_TABCONTROL)
         dim as long nCurSel = TabCtrl_GetCurSel(hTabCtrl) 

         DIM pwb AS CWebCtx PTR 
         select case nCurSel
            case 0: pwb = CAST(CWebCtx PTR, GetProp(HWND_FRMHELPVIEWER, "CWEBCTXPTR_WINFBE"))
            case 1: pwb = CAST(CWebCtx PTR, GetProp(HWND_FRMHELPVIEWER, "CWEBCTXPTR_WINFBX"))
         end select
         
         if pwb then
            ' Navigate to the path
            dim wszPath as wstring * MAX_PATH = gHTMLHelpFiles(idx)
            pwb->Navigate(wszPath)
            
            'pwb->Navigate("https://google.com")
            ' Optional: Wait for page load with a timeout of 10 seconds
            DIM lReadyState AS READYSTATE = pwb->WaitForPageLoad(10)
            ' Set the focus in the page (the page must be fully loaded)
            pwb->SetFocus
         end if

      end if
   end if   
   
   function = 0
end function


' ========================================================================================
' Process WM_NOTIFY message for window/dialog: frmHelpViewer
' ========================================================================================
Function frmHelpViewer_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT

   select case pNMHDR->code
      case TCN_SELCHANGE 
         if id = IDC_FRMHELPVIEWER_TABCONTROL then
            frmHelpViewer_PositionWindows(HWND_FRMHELPVIEWER)
         end if 
      case TVN_SELCHANGED
         Dim lpNMTV As NM_TREEVIEW Ptr = Cast(NM_TREEVIEW Ptr, pNMHDR)
         dim as long idx = FF_TreeView_GetlParam( pNMHDR->hWndFrom, lpNMTV->itemNew.hItem)
         frmHelpViewer_DisplayHTML(idx)
   end select
   
   function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmHelpViewer
' ========================================================================================
Function frmHelpViewer_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT

   If state <> SIZE_MINIMIZED Then 
      frmHelpViewer_PositionWindows(HWnd)
   End If   

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmHelpViewer
' ========================================================================================
Function frmHelpViewer_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ' Never close the window; simply hide it.
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmHelpViewer
' ========================================================================================
Function frmHelpViewer_OnDestroy(HWnd As HWnd) As LRESULT
   HWND_FRMHELPVIEWER = 0
   ' Delete the CWewbCtx class
   DIM pwb1 AS CWebCtx PTR = CAST(CWebCtx PTR, GetProp(hwnd, "CWEBCTXPTR_WINFBE"))
   IF pwb1 THEN Delete pwb1
   DIM pwb2 AS CWebCtx PTR = CAST(CWebCtx PTR, GetProp(hwnd, "CWEBCTXPTR_WINFBX"))
   IF pwb2 THEN Delete pwb2
   
   ' Delete the popup CWindow class
   DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
   IF pWindow THEN Delete pWindow
   Function = 0
End Function


' ========================================================================================
' frmHelpViewer Window procedure
' ========================================================================================
Function frmHelpViewer_WndProc( ByVal HWnd   As HWnd, _
                                ByVal uMsg   As UINT, _
                                ByVal wParam As WPARAM, _
                                ByVal lParam As LPARAM _
                                ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_COMMAND,     frmHelpViewer_OnCommand)
      HANDLE_MSG (HWnd, WM_NOTIFY,      frmHelpViewer_OnNotify)
      HANDLE_MSG (HWnd, WM_SIZE,        frmHelpViewer_OnSize)
      HANDLE_MSG (HWnd, WM_CLOSE,       frmHelpViewer_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,     frmHelpViewer_OnDestroy)
   End Select

   ' for messages that we don't deal with
   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


''
''
function LoadHelpTreeview( byval hTreeview as hwnd, byval sPath as CWSTR, byval hParent as HTREEITEM ) as long

   DIM pFinder AS CFindFile

   if pFinder.FindFile(sPath & "*.*") <> S_OK then exit function

   dim as HTREEITEM addMode = iif( hParent = null, TVI_ROOT, TVI_FIRST )
   dim SubFolders(any) as CWSTR
   dim as CWSTR wszRoot 

   do
      
      if pFinder.IsDots then 
      
      elseif pFinder.IsFolder then
         dim as long ub = ubound(SubFolders) 
         redim preserve SubFolders(ub + 1)
         SubFolders(ub + 1) = pFinder.FilePath
         wszRoot = pFinder.Root
      else
         ' Find any *.html files 
         if ucase(pFinder.FileExt) = "HTML" then
            dim as long ub = ubound(gHTMLHelpFiles) + 1
            redim preserve gHTMLHelpFiles(ub)
            gHTMLHelpFiles(ub) = pFinder.FilePath
            TreeView_AddItem(hTreeView, hParent, addMode, pFinder.Filename, ub, 0, 0)
         end if
         
      end if

      if pFinder.FindNext = 0 then exit do
   loop

   pFinder.Close()

   ' Process all of the folders that were found
   for i as long = lbound(SubFolders) to ubound(SubFolders)       
      dim as CWSTR wszFolder = mid(SubFolders(i), len(wszRoot) + 1)
      dim as HTREEITEM hNode = TreeView_AddItem(hTreeView, hParent, addMode, wszFolder, -1, 0, 0)
      LoadHelpTreeview(hTreeview, SubFolders(i) & "\", hNode)
   next   

   function = 0
end function

' ========================================================================================
' frmHelpViewer_Show
' ========================================================================================
Function frmHelpViewer_Show( ByVal hWndParent As HWnd, _
                             ByVal nCmdShow   As Long = -1 _
                             ) As Long

   ' If function list already exists then toggle its visibility
   ' unless nCmdShow is explicitly set to show it.
   If IsWindow(HWND_FRMHELPVIEWER) Then
      If nCmdShow <> SW_SHOW Then nCmdShow = Iif(IsWindowVisible(HWND_FRMHELPVIEWER), SW_HIDE, SW_SHOW)
      dim as hwnd hTabCtrl = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_TABCONTROL)
      select case gHelpViewerIndex 
         case IDM_HELPWINFBE: TabCtrl_SetCurSel(hTabCtrl, 0) 
         case IDM_HELPWINFBX: TabCtrl_SetCurSel(hTabCtrl, 1) 
      end select
      frmHelpViewer_PositionWindows(HWND_FRMHELPVIEWER)
      ShowWindow HWND_FRMHELPVIEWER, nCmdShow
      Exit Function 
   Else 
      if nCmdShow = -1 then nCmdShow = SW_HIDE
   End If
      

   '  Create the main window and child controls
   Dim pWindow As CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim rcWork  As Rect = pWindow->GetWorkArea
   Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .80
   Dim nWidth  As Long = (rcWork.Right - rcWork.Left) * .80

   HWND_FRMHELPVIEWER = _
   pWindow->Create( hWndParent, "Help Viewer", @frmHelpViewer_WndProc, 0, 0, nWidth, nHeight, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->Center(pWindow->hWindow, hWndParent)
   pWindow->brush = GetSysColorBrush(COLOR_WINDOW)
   
   ' Set the small and large icon for the main window (must be set after main window is created)
   pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
   pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)


   dim as hwnd hTabCtrl = pWindow->AddControl("TABCONTROL", , _
      IDC_FRMHELPVIEWER_TABCONTROL, "", 0, 0, 0, 24, _
      WS_CHILD Or WS_TABSTOP Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or _
      TCS_TABS Or TCS_FOCUSNEVER Or TCS_FORCEICONLEFT, WS_EX_LEFT Or WS_EX_LTRREADING)
      
      TabCtrl_AddTab(hTabCtrl, 0, "WinFBE Editor", 0)
      TabCtrl_AddTab(hTabCtrl, 0, "WinFBX Library", 0)
      select case gHelpViewerIndex 
         case IDM_HELPWINFBE: TabCtrl_SetCurSel(hTabCtrl, 0) 
         case IDM_HELPWINFBX: TabCtrl_SetCurSel(hTabCtrl, 1) 
      end select
      
   Dim As HWnd hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMHELPVIEWER_TVWINFBX, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP or _
        TVS_INFOTIP or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT or _
        TVS_HASBUTTONS or TVS_HASLINES or TVS_LINESATROOT, _   
        WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

      ' Use the new style Explorer Treeview (triangles instead of boxes).
      ' Set an undocumented extended style that enables the treeview glyphs to resize
      ' according to the high dpi setting.
      ' https://stackoverflow.com/questions/38772670/ctreectrl-with-explorer-theme-not-dpi-aware
      SendMessage(hTree, TVM_SETEXTENDEDSTYLE, &H1000, &H1000)
      SetWindowTheme(hTree, @wstr("EXPLORER"), 0)
      SendMessage( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   

'   Dim As HFONT hFont = pWindow->CreateFont("Segoe UI", 9) 
'   SendMessage( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
'   SendMessage( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, GetSysColor(COLOR_BTNFACE)) )
'   SendMessage( hTree, TVM_SETITEMHEIGHT, pWindow->ScaleY(20), 0 )

   ' Load the Treeview 
   LoadHelpTreeview(hTree, AfxGetExePath & "\Help\WinFBX\", null)
   
   
   hTree = _
        pWindow->AddControl("TREEVIEW", , IDC_FRMHELPVIEWER_TVWINFBE, "", 0, 0, 0, 0, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP or _
        TVS_INFOTIP or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT or _
        TVS_HASBUTTONS or TVS_HASLINES or TVS_LINESATROOT, _   
        WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

      ' Use the new style Explorer Treeview (triangles instead of boxes).
      ' Set an undocumented extended style that enables the treeview glyphs to resize
      ' according to the high dpi setting.
      ' https://stackoverflow.com/questions/38772670/ctreectrl-with-explorer-theme-not-dpi-aware
      SendMessage(hTree, TVM_SETEXTENDEDSTYLE, &H1000, &H1000)
      SetWindowTheme(hTree, @wstr("EXPLORER"), 0)
      SendMessage( hTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   

'   Dim As HFONT hFont = pWindow->CreateFont("Segoe UI", 9) 
'   SendMessage( hTree, WM_SETFONT, Cast(WPARAM, hFont), CTRUE )
'   SendMessage( hTree, TVM_SETBKCOLOR, 0, Cast(LPARAM, GetSysColor(COLOR_BTNFACE)) )
'   SendMessage( hTree, TVM_SETITEMHEIGHT, pWindow->ScaleY(20), 0 )

   ' Load the Treeview 
   LoadHelpTreeview(hTree, AfxGetExePath & "\Help\WinFBE\", null)

   
   ' Add a WebBrowser controls
   DIM pwb1 AS CWebCtx PTR = NEW CWebCtx(pWindow, IDC_FRMHELPVIEWER_WEBBROWSER_WINFBE, 0, 0, pWindow->ClientWidth, pWindow->ClientHeight)
   SetProp(pWindow->hWindow, "CWEBCTXPTR_WINFBE", CAST(HANDLE, pwb1))
   DIM pwb2 AS CWebCtx PTR = NEW CWebCtx(pWindow, IDC_FRMHELPVIEWER_WEBBROWSER_WINFBX, 0, 0, pWindow->ClientWidth, pWindow->ClientHeight)
   SetProp(pWindow->hWindow, "CWEBCTXPTR_WINFBX", CAST(HANDLE, pwb2))
   
   frmHelpViewer_PositionWindows(HWND_FRMHELPVIEWER)
   
   ' Display the window
   ShowWindow HWND_FRMHELPVIEWER, nCmdShow
   UpdateWindow(HWND_FRMHELPVIEWER)

   Function = 0
End Function





