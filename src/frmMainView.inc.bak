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

#include once "frmMain.bi"
#include once "clsDocument.bi"


' ========================================================================================
' View fold actions
' ========================================================================================
function OnCommand_ViewFold( _
            ByVal id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT
   
   select case id      
      Case IDM_FOLDTOGGLE: pDoc->FoldToggle( pDoc->GetCurrentLineNumber() )
      Case IDM_FOLDBELOW:  pDoc->FoldToggleOnwards( pDoc->GetCurrentLineNumber() )
      Case IDM_FOLDALL:    pDoc->FoldAll()
      Case IDM_UNFOLDALL:  pDoc->UnFoldAll()
   end select
   
   function = 0
end function

' ========================================================================================
' View zoom actions
' ========================================================================================
function OnCommand_ViewZoom( _
            ByVal id as long, _
            byval pDoc as clsDocument ptr _
            ) as LRESULT
   
   dim as HWND hEdit = pDoc->hWndActiveScintilla
   select case id      
      Case IDM_ZOOMIN:  SciExec(hEdit, SCI_ZOOMIN, 0, 0)
      Case IDM_ZOOMOUT: SciExec(hEdit, SCI_ZOOMOUT, 0, 0)
   end select
   
   function = 0
end function


' ========================================================================================
' View FunctionList, Explorer, Output, Bookmarks windows
' ========================================================================================
function OnCommand_ViewFunctionList() as LRESULT
   ShowWindow( HWND_FRMPANEL, SW_SHOW )
   gPanel.hActiveChild = HWND_FRMFUNCTIONS
   frmPanel_PositionWindows
   frmMain_PositionWindows
   function = 0
end function

function OnCommand_ViewExplorer() as LRESULT
   ShowWindow( HWND_FRMPANEL, SW_SHOW )
   gPanel.hActiveChild = HWND_FRMEXPLORER
   frmPanel_PositionWindows
   frmMain_PositionWindows
   function = 0
end function

function OnCommand_ViewBookmarksList() as LRESULT
   ShowWindow( HWND_FRMPANEL, SW_SHOW )
   gPanel.hActiveChild = HWND_FRMBOOKMARKS
   frmPanel_PositionWindows
   frmMain_PositionWindows
   function = 0
end function

function OnCommand_ViewOutput() as LRESULT
   ShowWindow( HWND_FRMOUTPUT, Iif(IsWindowVisible(HWND_FRMOUTPUT), SW_HIDE, SW_SHOW) )
   frmMain_PositionWindows
   function = 0
end function

' ========================================================================================
' View Restore Main window size
' ========================================================================================
function OnCommand_ViewRestoreMain() as LRESULT
   Dim rc As Rect
   SystemParametersInfo( SPI_GETWORKAREA, 0, @rc, 0 )
   MoveWindow( HWND_FRMMAIN, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, CTRUE )
   function = 0
end function
               
' ========================================================================================
' View Todo, Notes
' ========================================================================================
function OnCommand_ViewToDo() as LRESULT
   ' toggle close the Output window if it is open and already at ToDo 
   if IsWindowVisible(HWND_FRMOUTPUT) then
      if gOutputTabsCurSel = 3 then
         ShowWindow( HWND_FRMOUTPUT, SW_HIDE )
         frmMain_PositionWindows
         exit function
      end if
   else      
      ShowWindow( HWND_FRMOUTPUT, SW_SHOW )
   end if
   frmMain_PositionWindows
   gOutputTabsCurSel = 3
   frmOutput_PositionWindows
   AfxRedrawWindow( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABS) )
   function = 0
end function

function OnCommand_ViewNotes() as LRESULT
   ' toggle close the Output window if it is open and already at Notes
   if IsWindowVisible(HWND_FRMOUTPUT) then
      if gOutputTabsCurSel = 4 then
         ShowWindow( HWND_FRMOUTPUT, SW_HIDE )
         frmMain_PositionWindows
         exit function
      end if
   else      
      ShowWindow( HWND_FRMOUTPUT, SW_SHOW )
   end if
   frmMain_PositionWindows
   gOutputTabsCurSel = 4
   frmOutput_PositionWindows
   AfxRedrawWindow( GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TABS) )
   function = 0
end function
