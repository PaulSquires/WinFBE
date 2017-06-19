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
' Create the frmMain rebar/toolbar
' ========================================================================================
Function frmMain_CreateToolbar( ByVal pWindow As CWindow Ptr ) As BOOLEAN

   Dim As HINSTANCE hInst = pWindow->InstanceHandle

   ' Create rebar and toolbar controls
   Dim As HWnd hRebar = pWindow->AddControl("Rebar", pWindow->hWindow, IDC_FRMMAIN_REBAR)
   Dim As HWnd hToolBar = pWindow->AddControl("Toolbar", pWindow->hWindow, IDC_FRMMAIN_TOOLBAR, "", 0, 0, 0, 0, _
      WS_CHILD Or WS_VISIBLE Or TBSTYLE_TOOLTIPS Or TBSTYLE_FLAT Or CCS_NODIVIDER Or CCS_NOPARENTALIGN)

   HWND_FRMMAIN_TOOLBAR = hToolBar
   
   ' Allow drop down arrows
   SendMessageW hToolBar, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS

   ' Create the ImageLists that will hold the images for the toolbar
   Dim As HIMAGELIST hImageListNormal, hImageListHot, hImageListDisabled
   Dim cx As Long = 24 * (pWindow->DPI \ 96)

   hImageListNormal   = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 20, 1)
   hImageListDisabled = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 20, 1)
   hImageListHot      = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 20, 1)

   ' Add buttons to the toolbar
   Dim As Long i

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_NEW_FILE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_NEW_FILE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_NEW_FILE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FILENEW, 0, TBSTYLE_DROPDOWN
  
   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_OPEN_FILE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_OPEN_FILE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_OPEN_FILE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FILEOPEN, 0, TBSTYLE_DROPDOWN

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_SAVE_FILE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_SAVE_FILE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_SAVE_FILE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FILESAVE

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_CUT")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_CUT")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_CUT", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_CUT

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_COPY")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_COPY")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_COPY", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_COPY

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_PASTE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_PASTE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_PASTE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_PASTE

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_FIND")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_FIND")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_FIND", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FIND

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_REPLACE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_REPLACE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_REPLACE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_REPLACE

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_UNDO")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_UNDO")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_UNDO", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_UNDO

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_REDO")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_REDO")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_REDO", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_REDO

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_INDENT")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_INDENT")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_INDENT", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_INDENTBLOCK

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_INDENT_REMOVE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_INDENT_REMOVE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_INDENT_REMOVE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_UNINDENTBLOCK

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_COMMENT")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_COMMENT")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_COMMENT", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_COMMENTBLOCK

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_COMMENT_REMOVE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_COMMENT_REMOVE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_COMMENT_REMOVE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_UNCOMMENTBLOCK

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_FUNCTIONLIST")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_FUNCTIONLIST")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_FUNCTIONLIST", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FUNCTIONLIST

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_COMPILE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_COMPILE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_COMPILE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_COMPILE

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_COMPILE_AND_EXECUTE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_COMPILE_AND_EXECUTE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_COMPILE_AND_EXECUTE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_BUILDEXECUTE

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_EXECUTE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_EXECUTE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_EXECUTE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_RUNEXE

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_HELP")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_HELP")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_HELP", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_HELP

   SendMessage hToolBar, TB_SETIMAGELIST, 0, Cast(LPARAM, hImageListNormal)
   SendMessage hToolBar, TB_SETDISABLEDIMAGELIST, 0, Cast(LPARAM, hImageListDisabled)
   SendMessage hToolBar, TB_SETHOTIMAGELIST, 0, Cast(LPARAM, hImageListHot)

   ' Size the toolbar
   SendMessage hToolBar, TB_AUTOSIZE, 0, 0

   ' Add the band containing the toolbar control to the rebar
   ' The size of the REBARBANDINFOW is different in Vista/Windows 7
   Dim rc As Rect
   Dim rbbi As REBARBANDINFOW
   If (AfxWindowsVersion >= 600) AndAlso (AfxComCtlVersion >= 600) Then
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   Else
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   End If

   ' Insert the toolbar in the rebar control
   rbbi.fMask      = RBBIM_STYLE Or RBBIM_CHILD Or RBBIM_CHILDSIZE Or _
                     RBBIM_SIZE Or RBBIM_ID Or RBBIM_IDEALSIZE Or RBBIM_COLORS
   rbbi.fStyle     = RBBS_CHILDEDGE Or RBBS_GRIPPERALWAYS
   rbbi.hwndChild  = hToolbar
   rbbi.cxMinChild = 270 * pWindow->rxRatio
   rbbi.cyMinChild = Hiword(SendMessageW(hToolBar, TB_GETBUTTONSIZE, 0, 0))
   rbbi.cx         = 270 * pWindow->rxRatio
   rbbi.cxIdeal    = 270 * pWindow->rxRatio
   rbbi.clrFore    = GetSysColor(COLOR_BTNTEXT)
   rbbi.clrBack    = GetSysColor(COLOR_BTNFACE)

   ' Insert band into rebar
   SendMessage hRebar, RB_INSERTBANDW, -1, Cast(LPARAM, @rbbi)

   Function = True

End Function


' ========================================================================================
' Disable toolbar buttons
' ========================================================================================
Function frmMain_DisableToolbarButtons() As Long

   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_FILESAVE)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_CUT)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_COPY)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_PASTE)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_FIND)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_REPLACE)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNDO)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_REDO)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_INDENTBLOCK)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNINDENTBLOCK)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_COMMENTBLOCK)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNCOMMENTBLOCK)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_FUNCTIONLIST)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_COMPILE)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_BUILDEXECUTE)
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_RUNEXE)

   Function = 0
End Function


' ========================================================================================
' Change toolbar button's state
' ========================================================================================
Function frmMain_ChangeToolbarButtonsState() As Long

   ' Disables toolbar buttons if there is not any file being edited
   Dim hEdit       As HWnd
   Dim startSelPos As Long
   Dim endSelPos   As Long
   
   ' Disable all toolbar icons by default
   frmMain_DisableToolbarButtons

   Dim pDoc As clsDocument Ptr = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then hEdit = pDoc->hWindow

   ' Enables or disables toolbar buttons
   If IsWindow(hEdit) Then
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_FILESAVE)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_CUT)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_COPY)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_PASTE)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_FIND)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_REPLACE)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNDO)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_REDO)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_INDENTBLOCK)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNINDENTBLOCK)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_COMMENTBLOCK)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNCOMMENTBLOCK)
   
      If cbool(SciExec(hEdit, SCI_GETMODIFY, 0, 0)) = false and (pDoc->UserModified = false) Then 
         FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_FILESAVE) 
      Else 
         FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_FILESAVE)
      End If
   
      startSelPos = SciExec(hEdit, SCI_GETSELECTIONSTART, 0, 0)
      endSelPos   = SciExec(hEdit, SCI_GETSELECTIONEND, 0, 0)
      If startSelPos = endSelPos Then
         FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_CUT)
         FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_COPY)
      Else
         FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_CUT)
         FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_COPY)
      End If
   
      If SciExec(hEdit, SCI_CANPASTE, 0, 0) = 0 Then 
         FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_PASTE) 
      Else 
         FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_PASTE)
      End If
      
      If SciExec(hEdit, SCI_CANUNDO, 0, 0) = 0 Then 
         FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNDO) 
      Else 
         FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNDO)
      End If
   
      If SciExec(hEdit, SCI_CANREDO, 0, 0) = 0 Then 
         FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_REDO) 
      Else 
         FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_REDO)
      End If

   End If
   
   If gApp.IsProjectActive OrElse cBool(IsWindow(hEdit)) Then
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_PROJECTMANAGER)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_BUILDEXECUTE)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_COMPILE)
   Else
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_PROJECTMANAGER)
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_BUILDEXECUTE)
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_COMPILE)
   End If
   
   If AfxFileExists(gConfig.LastRunFilename) Then
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_RUNEXE)
   Else
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_RUNEXE)
   End If
   
   ' If the Explorer treeview window is visible then at least one document or project is open
   if IsWindowVisible( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE) ) THEN
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_FUNCTIONLIST)
   else   
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_FUNCTIONLIST)
   END IF
   
   Function = 0
End Function

