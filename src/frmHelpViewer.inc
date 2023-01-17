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


#include once "frmHelpViewer.bi"


' ========================================================================================
' Highlight the treeview node when moving between help documents
' ========================================================================================
private function frmHelpViewer_HighlightTreeviewNode( byref wszLocationURL as WSTRING ) as Long
   for i as long = lbound(gHTMLHelp) to Ubound(gHTMLHelp)
      if wszLocationURL = gHTMLHelp(i).wszLocationURL then
         TreeView_SelectItem( gHTMLHelp(i).hTreeview, gHTMLHelp(i).TreeviewNode )
         exit for
      end if
   next   
   function = 0
end function


' ========================================================================================
' Show context help or general contents if word does not exist
' ========================================================================================
public Function ShowContextHelp( byval id as long ) As Long

   Dim HtmlHelpW As Function ( BYVAL hwndCaller AS HWnd, _
                               BYVAL pswzFile AS WSTRING Ptr, _
                               BYVAL uCommand AS UNIT, _
                               BYVAL dwData AS DWORD_PTR _
                               ) AS HWND

   dim as CWSTR wszHelpFilename, wszLabel
   
   ' Ensure that the CurDrive parameter is converted if applicable
   gConfig.CompilerHelpFile = ProcessFromCurdriveApp( gConfig.CompilerHelpFile ) 
   gConfig.WinFBXHelpfile   = ProcessFromCurdriveApp( gConfig.WinFBXHelpfile ) 

   select Case id
      case IDM_HELP
         wszLabel = "FreeBASIC "
         ' Convert relative path to absolute path if needed.
         if AfxPathIsRelative(gConfig.CompilerHelpFile) then
            wszHelpFilename = AfxPathCombine(AfxGetExePathName, gConfig.CompilerHelpFile)
         else
            wszHelpFilename = gConfig.CompilerHelpFile
         END IF

      case IDM_HELPSHORTCUTS, IDM_HELPWINFBE, IDM_HELPWINFBX
         frmHelpViewer_Show( HWND_FRMMAIN, id )
         exit function

   END SELECT


   If AfxFileExists(wszHelpFilename) = 0 Then
      MessageBoxW( HWND_FRMMAIN, wszLabel & L(244,"Help file not found."), L(201,"Error"), _
                    MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 Or MB_APPLMODAL )
      Exit Function
   End If


   HtmlHelpW = DyLibSymbol( gpHelpLib, "HtmlHelpW" )

   If (gpHelpLib = 0) OrElse (HtmlHelpW = 0) Then
      MessageBox( HWND_FRMMAIN, L(243,"Error loading HtmlHelp."), L(201,"Error"), _
                    MB_OK Or MB_ICONWARNING Or MB_DEFBUTTON1 Or MB_APPLMODAL )
      Exit Function
   End If

   ' If we are currently in an active document then attempt to lookup the 
   ' word immediately under the caret.
   Dim pDoc as clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   
   If pDoc = 0 Then
      HtmlHelpW( 0, wszHelpFilename, HH_DISPLAY_TOC, NULL )
      Exit Function
   ENd If

   Dim wszKeyword as WString * MAX_PATH
   
   wszKeyword = WStr( pDoc->GetWord )
   
   Dim li As HH_AKLINK
   With li
      .cbStruct     = SizeOf(HH_AKLINK) 
      .fReserved    = FALSE 
      .pszKeywords  = @wszKeyword
      .pszUrl       = Null
      .pszMsgText   = Null
      .pszMsgTitle  = Null
      .pszWindow    = Null
      .fIndexOnFail = FALSE 
   End With
   
   ' Open the help and show the topic
   HtmlHelpW( 0, wszHelpFilename, HH_DISPLAY_TOC, Null )   '<-- needed?
   
   If HtmlHelpW( 0, wszHelpFilename, HH_KEYWORD_LOOKUP, Cast(DWORD_PTR, @li) ) = 0 Then
      ' Normal case search failed, try a ucase search
      wszKeyword     = UCase(wszKeyword)
      li.pszKeywords = @wszKeyword
      HtmlHelpW( 0, wszHelpFilename, HH_KEYWORD_LOOKUP, Cast(DWORD_PTR, @li) )
   End If

   Function = 0
   
End Function


' ========================================================================================
' Position all child windows. Called manually and/or by WM_SIZE
' ========================================================================================
private Function frmHelpViewer_PositionWindows() As LRESULT
   
   ' Get the entire client area
   Dim As Rect rc
   GetClientRect(HWND_FRMHELPVIEWER, @rc)
   
   dim as hwnd hTabCtrl = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_TABCONTROL)
   dim as hwnd hTree(1), hWebBrowser(1), hBtn(3)
   hTree(0) = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_TVWINFBE)
   hTree(1) = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_TVWINFBX)
   hWebBrowser(0) = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_WEBBROWSER_WINFBE )
   hWebBrowser(1) = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_WEBBROWSER_WINFBX )
   hBtn(0) = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_BACK ) 
   hBtn(1) = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_FORWARD ) 
   hBtn(2) = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_PRINT ) 
   hBtn(3) = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_FIND )
    
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
   
   dim as long nSpace     = AfxScaleY(4)
   dim as long nBtnHeight = AfxScaleY(28)
   dim as long nBtnWidth  = AfxScaleY(74)
   dim as long nTabHeight = AfxScaleY(24)
   dim as long nTreeWidth = AfxScaleX(200)
   dim as long nTreeHeight = rc.Bottom-rc.top-AfxScaleY(70)
   
   SetWindowPos( hTabCtrl, 0, 0, 0, rc.Right-rc.Left, nTabHeight, SWP_SHOWWINDOW Or SWP_NOZORDER)
   SetWindowPos( hTree(nCurSel), 0, 0, AfxScaleY(70), nTreeWidth, nTreeHeight, SWP_SHOWWINDOW Or SWP_NOZORDER)
   
   dim nLeft as long = 0
   for i as long = 0 to 3
      SetWindowPos( hBtn(i), 0, nTreeWidth + nLeft, nTabHeight + (nSpace * 2), nBtnWidth, nBtnHeight, SWP_SHOWWINDOW Or SWP_NOZORDER)
      nLeft = nLeft + nBtnWidth + nSpace               
   next
   SetWindowPos( hWebBrowser(nCurSel), 0, nTreeWidth, AfxScaleY(70), _
                  rc.Right-rc.Left-nTreeWidth, nTreeHeight, SWP_SHOWWINDOW Or SWP_NOZORDER)

   Function = 0
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmHelpViewer
' ========================================================================================
private Function frmHelpViewer_OnCommand( ByVal HWnd As HWnd, _
                                          ByVal id As Long, _
                                          ByVal hwndCtl As HWnd, _
                                          ByVal codeNotify As UINT _
                                          ) As LRESULT
   DIM pwb AS CWebCtx PTR 

   select case TabCtrl_GetCurSel( GetDlgItem(HWnd, IDC_FRMHELPVIEWER_TABCONTROL) ) 
      case -1:  exit function   
      case  0:  pwb = CAST(CWebCtx PTR, GetProp(hwnd, "CWEBCTXPTR_WINFBE"))
      case  1:  pwb = CAST(CWebCtx PTR, GetProp(hwnd, "CWEBCTXPTR_WINFBX")) 
   end select
   if pwb = 0 then exit function
   
   Select Case id
      Case IDC_FRMHELPVIEWER_BACK
         If codeNotify = BN_CLICKED Then
            pwb->GoBack
            pwb->WaitForPageLoad(10)
            frmHelpViewer_HighlightTreeviewNode(pwb->LocationURL)
         End If
      Case IDC_FRMHELPVIEWER_FORWARD
         If codeNotify = BN_CLICKED Then
            pwb->GoForward
            pwb->WaitForPageLoad(10)
            frmHelpViewer_HighlightTreeviewNode(pwb->LocationURL)
         End If
      Case IDC_FRMHELPVIEWER_PRINT
         If codeNotify = BN_CLICKED Then
            pwb->PrintPreview
         End If
      Case IDC_FRMHELPVIEWER_FIND
         If codeNotify = BN_CLICKED Then
            pwb->Find
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Display the HTML file held in the array index
' ========================================================================================
private function frmHelpViewer_DisplayHTML( byval idx as long ) as Long
   
   if idx >= lbound(gHTMLHelp) andalso idx <= ubound(gHTMLHelp) then
      if AfxFileExists(gHTMLHelp(idx).wszFilename) then
         dim as hwnd hTabCtrl = GetDlgItem(HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_TABCONTROL)
         dim as long nCurSel = TabCtrl_GetCurSel(hTabCtrl) 

         DIM pwb AS CWebCtx PTR 
         select case nCurSel
            case 0: pwb = CAST(CWebCtx PTR, GetProp(HWND_FRMHELPVIEWER, "CWEBCTXPTR_WINFBE"))
            case 1: pwb = CAST(CWebCtx PTR, GetProp(HWND_FRMHELPVIEWER, "CWEBCTXPTR_WINFBX"))
         end select
         
         if pwb then
            ' Navigate to the path
            dim wszPath as wstring * MAX_PATH = gHTMLHelp(idx).wszFilename
            pwb->Navigate(wszPath)
            ' Optional: Wait for page load with a timeout of 10 seconds
            DIM lReadyState AS READYSTATE = pwb->WaitForPageLoad(10)
            gHTMLHelp(idx).wszLocationURL = pwb->LocationURL
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
private Function frmHelpViewer_OnNotify( ByVal HWnd As HWnd, _
                                         ByVal id As Long, _
                                         ByVal pNMHDR As NMHDR Ptr _
                                         ) As LRESULT
   select case pNMHDR->code
      case TCN_SELCHANGE 
         if id = IDC_FRMHELPVIEWER_TABCONTROL then
            frmHelpViewer_PositionWindows
         end if 
      case TVN_SELCHANGED
         Dim lpNMTV As NM_TREEVIEW Ptr = Cast(NM_TREEVIEW Ptr, pNMHDR)
         dim as long idx = TreeView_GetlParam( pNMHDR->hWndFrom, lpNMTV->itemNew.hItem)
         frmHelpViewer_DisplayHTML(idx)
   end select
   
   function = 0
end function


' ========================================================================================
' Process WM_SIZE message for window/dialog: frmHelpViewer
' ========================================================================================
private Function frmHelpViewer_OnSize( ByVal HWnd As HWnd, _
                                       ByVal state As UINT, _
                                       ByVal cx As Long, _
                                       ByVal cy As Long _
                                       ) As LRESULT
   If state <> SIZE_MINIMIZED Then 
      frmHelpViewer_PositionWindows
   End If   
   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmHelpViewer
' ========================================================================================
private Function frmHelpViewer_OnClose( ByVal HWnd As HWnd ) As LRESULT
   ' Only hide the window rather than destroy it
   ShowWindow( HWnd, SW_HIDE )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmHelpViewer
' ========================================================================================
private Function frmHelpViewer_OnDestroy( byval HWnd As HWnd ) As LRESULT
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
private Function frmHelpViewer_WndProc( ByVal HWnd   As HWnd, _
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

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' Load the Treeview showing the hlep topics (recursive). Store array index in lParam
' ========================================================================================
private function LoadHelpTreeview( byval hTreeview as hwnd, _
                                   byval sPath as CWSTR, _
                                   byval hParent as HTREEITEM _
                                   ) as long

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
            dim as long ub = ubound(gHTMLHelp) + 1
            redim preserve gHTMLHelp(ub)
            gHTMLHelp(ub).wszFilename = pFinder.FilePath
            gHTMLHelp(ub).TreeviewNode = TreeView_AddItem(hTreeView, hParent, addMode, pFinder.Filename, ub, 0, 0)
            gHTMLHelp(ub).hTreeView = hTreeView
            Treeview_SortChildren( hTreeView, hParent, 0 )
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
public Function frmHelpViewer_Show( ByVal hWndParent As HWnd, _
                                    byval idmHelpMessage as long _
                                    ) As LRESULT

   static as Long idxNodeShortcuts = -1
   dim as hwnd hTree1, hTree2, hTabCtrl
   
   if IsWindow( HWND_FRMHELPVIEWER ) = 0 then

      '  Create the main window and child controls
      Dim pWindow As CWindow Ptr = New CWindow
      pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

      Dim rcWork  As Rect = pWindow->GetWorkArea
      Dim nHeight As Long = (rcWork.Bottom - rcWork.Top) * .80
      Dim nWidth  As Long = (rcWork.Right - rcWork.Left) * .80

      HWND_FRMHELPVIEWER = _
      pWindow->Create( 0, "Help Viewer", @frmHelpViewer_WndProc, 0, 0, nWidth, nHeight, _
           WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_MAXIMIZEBOX or WS_MINIMIZEBOX or WS_THICKFRAME or _
           WS_CLIPCHILDREN OR WS_CLIPSIBLINGS, _
           WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
      pWindow->Center(pWindow->hWindow, hWndParent)
      pWindow->brush = GetSysColorBrush(COLOR_WINDOW)
      
      ' Set the small and large icon for the main window (must be set after main window is created)
      pWindow->BigIcon   =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 32, 32, LR_SHARED)
      pWindow->SmallIcon =  LoadImage( pWindow->InstanceHandle, "IMAGE_AAA_MAINICON", IMAGE_ICON, 16, 16, LR_SHARED)


      hTree1 = _
           pWindow->AddControl("TREEVIEW", , IDC_FRMHELPVIEWER_TVWINFBE, "", 0, 0, 0, 0, _
           WS_CHILD Or WS_VISIBLE Or WS_TABSTOP or _
           TVS_INFOTIP or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT or _
           TVS_HASBUTTONS or TVS_HASLINES or TVS_LINESATROOT, _   
           WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

         ' Use the new style Explorer Treeview (triangles instead of boxes).
         ' Set an undocumented extended style that enables the treeview glyphs to resize
         ' according to the high dpi setting.
         ' https://stackoverflow.com/questions/38772670/ctreectrl-with-explorer-theme-not-dpi-aware
         SendMessage(hTree1, TVM_SETEXTENDEDSTYLE, &H1000, &H1000)
         SetWindowTheme(hTree1, @wstr("EXPLORER"), 0)
         SendMessage( hTree1, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   

         ' Load the Treeview 
         LoadHelpTreeview(hTree1, AfxGetExePath & "\Help\WinFBE\", null)

      
      hTree2 = _
           pWindow->AddControl("TREEVIEW", , IDC_FRMHELPVIEWER_TVWINFBX, "", 0, 0, 0, 0, _
           WS_CHILD Or WS_VISIBLE Or WS_TABSTOP or _
           TVS_INFOTIP or TVS_SHOWSELALWAYS Or TVS_FULLROWSELECT Or TVS_TRACKSELECT or _
           TVS_HASBUTTONS or TVS_HASLINES or TVS_LINESATROOT, _   
           WS_EX_LEFT Or WS_EX_RIGHTSCROLLBAR)

         ' Use the new style Explorer Treeview (triangles instead of boxes).
         ' Set an undocumented extended style that enables the treeview glyphs to resize
         ' according to the high dpi setting.
         ' https://stackoverflow.com/questions/38772670/ctreectrl-with-explorer-theme-not-dpi-aware
         SendMessage(hTree2, TVM_SETEXTENDEDSTYLE, &H1000, &H1000)
         SetWindowTheme(hTree2, @wstr("EXPLORER"), 0)
         SendMessage( hTree2, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER)   

         ' Load the Treeview 
         LoadHelpTreeview(hTree2, AfxGetExePath & "\Help\WinFBX\", null)
      
      
      ' Add the naviagation and print buttons
      pWindow->AddControl("BUTTON", , IDC_FRMHELPVIEWER_BACK, "Back", 0, 0, 0, 0, _
           WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING)
      
      pWindow->AddControl("BUTTON", , IDC_FRMHELPVIEWER_FORWARD, "Forward", 0, 0, 0, 0, _
           WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING)

      pWindow->AddControl("BUTTON", , IDC_FRMHELPVIEWER_PRINT, "Print", 0, 0, 0, 0, _
           WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING)

      pWindow->AddControl("BUTTON", , IDC_FRMHELPVIEWER_FIND, "Find", 0, 0, 0, 0, _
           WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
           WS_EX_LEFT Or WS_EX_LTRREADING)

      ' Add a WebBrowser controls
      DIM pwb1 AS CWebCtx PTR = NEW CWebCtx(pWindow, IDC_FRMHELPVIEWER_WEBBROWSER_WINFBE, 0, 0, pWindow->ClientWidth, pWindow->ClientHeight)
      SetProp(pWindow->hWindow, "CWEBCTXPTR_WINFBE", CAST(HANDLE, pwb1))
      DIM pwb2 AS CWebCtx PTR = NEW CWebCtx(pWindow, IDC_FRMHELPVIEWER_WEBBROWSER_WINFBX, 0, 0, pWindow->ClientWidth, pWindow->ClientHeight)
      SetProp(pWindow->hWindow, "CWEBCTXPTR_WINFBX", CAST(HANDLE, pwb2))
   
      hTabCtrl = pWindow->AddControl("TABCONTROL", , _
         IDC_FRMHELPVIEWER_TABCONTROL, "", 0, 0, 0, 24, _
         WS_CHILD Or WS_TABSTOP Or TCS_SINGLELINE Or TCS_RAGGEDRIGHT Or TCS_HOTTRACK Or _
         TCS_TABS Or TCS_FOCUSNEVER Or TCS_FORCEICONLEFT, WS_EX_LEFT Or WS_EX_LTRREADING)
      
      TabCtrl_AddTab(hTabCtrl, 0, "WinFBE Editor", 0)
      TabCtrl_AddTab(hTabCtrl, 0, "WinFBX Library", 0)

   
      ' Set the treeviews to their default topic (the first topic found in the array for each treeview).
      dim as long idxNodeTree1 = -1
      dim as long idxNodeTree2 = -1
      
      for i as long = lbound(gHTMLHelp) to ubound(gHTMLHelp) 
         if gHTMLHelp(i).hTreeview = hTree1 then
            if idxNodeTree1 = -1 then idxNodeTree1 = i
         elseif gHTMLHelp(i).hTreeview = hTree2 then
            if idxNodeTree2 = -1 then idxNodeTree2 = i
         end if
         if AfxStrPathName( "NAMEX", ucase( gHTMLHelp(i).wszFilename )) = "KEYBOARD SHORTCUTS.HTML" then
            if idxNodeShortcuts = -1 then idxNodeShortcuts = i
         end if
      next 

      ' Set the default first page to show for each treeview
      TabCtrl_SetCurSel( hTabCtrl, 0 ) 
      if idxNodeTree1 > -1 then TreeView_SelectItem( hTree1, gHTMLHelp(idxNodeTree1).TreeviewNode )
      frmHelpViewer_DisplayHTML( idxNodeTree1 )

      TabCtrl_SetCurSel( hTabCtrl, 1 ) 
      if idxNodeTree2 > -1 then TreeView_SelectItem( hTree2, gHTMLHelp(idxNodeTree2).TreeviewNode )
      frmHelpViewer_DisplayHTML( idxNodeTree2 )
   end if


   ' Must set the active tab page first because frmHelpViewer_DisplayHTML and frmHelpViewer_PositionWindows
   ' both test what the active tab page in order to perform their work.
   hTabCtrl = GetDlgItem( HWND_FRMHELPVIEWER, IDC_FRMHELPVIEWER_TABCONTROL )
   select case idmHelpMessage
      case IDM_HELPSHORTCUTS
         TabCtrl_SetCurSel( hTabCtrl, 0 )
         frmHelpViewer_DisplayHTML( idxNodeShortcuts )
      
      case IDM_HELPWINFBE
         TabCtrl_SetCurSel( hTabCtrl, 0 ) 
      
      case IDM_HELPWINFBX
         TabCtrl_SetCurSel( hTabCtrl, 1 ) 
   end select
      
   frmHelpViewer_PositionWindows

         
   ' Display the window
   ShowWindow HWND_FRMHELPVIEWER, SW_SHOW
   UpdateWindow(HWND_FRMHELPVIEWER)

   Function = 0
End Function

