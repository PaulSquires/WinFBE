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
   Dim cx As Long = 24 * pWindow->DPI \ 96

   hImageListNormal   = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 20, 1)
   hImageListDisabled = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 20, 1)
   hImageListHot      = ImageList_Create( cx, cx, ILC_COLOR32 Or ILC_MASK, 20, 1)

   ' Add buttons to the toolbar
   Dim As Long i

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_NEW_FILE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_NEW_FILE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_NEW_FILE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FILENEW, 0, TBSTYLE_DROPDOWN
  
   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_OPEN_FILE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_OPEN_FILE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_OPEN_FILE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FILEOPEN, 0, TBSTYLE_DROPDOWN

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_SAVE_FILE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_SAVE_FILE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_SAVE_FILE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FILESAVE

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_CUT")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_CUT")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_CUT", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_CUT

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_COPY")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_COPY")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_COPY", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_COPY

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_PASTE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_PASTE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_PASTE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_PASTE

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_FIND")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_FIND")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_FIND", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FIND

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_REPLACE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_REPLACE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_REPLACE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_REPLACE

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_UNDO")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_UNDO")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_UNDO", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_UNDO

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_REDO")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_REDO")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_REDO", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_REDO

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_INDENT")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_INDENT")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_INDENT", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_INDENTBLOCK

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_INDENT_REMOVE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_INDENT_REMOVE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_INDENT_REMOVE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_UNINDENTBLOCK

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_COMMENT")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_COMMENT")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_COMMENT", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_COMMENTBLOCK

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_COMMENT_REMOVE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_COMMENT_REMOVE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_COMMENT_REMOVE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_UNCOMMENTBLOCK

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_FUNCTIONLIST")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_FUNCTIONLIST")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_FUNCTIONLIST", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_FUNCTIONLIST

   Toolbar_AddSeparator hToolBar

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_COMPILE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_COMPILE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_COMPILE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_COMPILE

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_COMPILE_AND_EXECUTE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_COMPILE_AND_EXECUTE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_COMPILE_AND_EXECUTE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_BUILDEXECUTE

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_QUICKRUN")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_QUICKRUN")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_QUICKRUN", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_QUICKRUN

   i = AfxGdipAddIconFromRes(hImageListNormal, hInst, "IMAGE_TB_EXECUTE")
   i = AfxGdipAddIconFromRes(hImageListHot, hInst, "IMAGE_TB_EXECUTE")
   i = AfxGdipAddIconFromRes(hImageListDisabled, hInst, "IMAGE_TB_EXECUTE", 10, True)
   Toolbar_AddButton hToolBar, i, IDM_RUNEXE

   Toolbar_AddSeparator hToolBar

   SendMessage hToolBar, TB_SETIMAGELIST, 0, Cast(LPARAM, hImageListNormal)
   SendMessage hToolBar, TB_SETDISABLEDIMAGELIST, 0, Cast(LPARAM, hImageListDisabled)
   SendMessage hToolBar, TB_SETHOTIMAGELIST, 0, Cast(LPARAM, hImageListHot)

   ' Size the toolbar
   SendMessage hToolBar, TB_AUTOSIZE, 0, 0

   ' Add the band containing the toolbar control to the rebar
   ' The size of the REBARBANDINFOW is different in Vista/Windows 7
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

   ' Get the last button's rect and position the combobox immediately the right
   dim rc as RECT
   Toolbar_GetItemRect(hToolBar, 26, @rc)
   
   ' Combobox to switch amongst the different build options
   HWND_FRMMAIN_COMBOBUILDS = _
   pWindow->AddControl("COMBOBOX", hToolBar, IDC_FRMMAIN_COMBOBUILDS, "", _
        pWindow->UnScaleX(rc.right + 10), pWindow->UnScaleY(8), 200, 20, _
        WS_CHILD OR WS_VISIBLE OR WS_VSCROLL OR WS_BORDER OR WS_TABSTOP OR CBS_DROPDOWNLIST OR CBS_HASSTRINGS, _
        WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

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
   FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_QUICKRUN)
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
   If pDoc Then hEdit = pDoc->hWndActiveScintilla

   ' Enable/Disable depending on if Visual Designer or Code Editor
   if (cbool(pDoc <> 0)) andalso (pDoc->IsDesigner = true) andalso (IsDesignerView(pDoc) = true) THEN
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_FILESAVE)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_CUT)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_COPY)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_PASTE)
      If pDoc->UserModified = false Then 
         FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_FILESAVE) 
      End If

   else
      ' Code editor is active
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
         End If
      
         startSelPos = SciExec(hEdit, SCI_GETSELECTIONSTART, 0, 0)
         endSelPos   = SciExec(hEdit, SCI_GETSELECTIONEND, 0, 0)
         If startSelPos = endSelPos Then
            FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_CUT)
            FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_COPY)
         End If
      
         If SciExec(hEdit, SCI_CANPASTE, 0, 0) = 0 Then 
            FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_PASTE) 
         End If
         
         If SciExec(hEdit, SCI_CANUNDO, 0, 0) = 0 Then 
            FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_UNDO) 
         End If
      
         If SciExec(hEdit, SCI_CANREDO, 0, 0) = 0 Then 
            FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_REDO) 
         End If
   
         ' If the Explorer treeview window is visible then at least one document or project is open
         if IsWindowVisible( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE) ) THEN
            FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_FUNCTIONLIST)
         else   
            FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_FUNCTIONLIST)
         END IF
   
      end if
   End If
   
   If gApp.IsProjectActive OrElse cBool(IsWindow(hEdit)) Then
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_PROJECTMANAGER)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_BUILDEXECUTE)
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_COMPILE)
      ' QuickRun is not available for projects.
      If gApp.IsProjectActive = false then 
         FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_QUICKRUN)
      end if
   Else
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_PROJECTMANAGER)
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_BUILDEXECUTE)
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_COMPILE)
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_QUICKRUN)
   End If
   
   ' Generate the gCompile.OutputFilename
   If AfxFileExists(GetRunExecutableFilename) Then
      FF_Toolbar_EnableButton(HWND_FRMMAIN_TOOLBAR, IDM_RUNEXE)
   Else
      FF_Toolbar_DisableButton(HWND_FRMMAIN_TOOLBAR, IDM_RUNEXE)
   End If
      
   Function = 0
End Function

