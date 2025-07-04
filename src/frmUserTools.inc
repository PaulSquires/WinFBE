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

#include once "frmUserTools.bi"
#include once "clsConfig.bi"
#include once "modMRU.bi"


' ========================================================================================
' Update the UserTools items in the topmenu array
' ========================================================================================
function updateUserToolsMenuItems() as long
   ' clear MRU items already existing in the gTopMenu array. We overwrite and extend
   ' the gTopMenu array rather than erase it because existing menus depend on the 
   ' array index that aready exist. The function also returns the width to use for
   ' the resulting popup menu (based on text width metrics of each filename).
   clearMRUFilesItems( IDM_USERTOOLS )
   
   dim wszText as WSTRING * 256
   dim as boolean hasTools = false
   dim as long txtWidth = 0
   dim as long nMenuWidth = -1

   setTopMenuMRUItem( getNextFreeMRUindex(), IDM_USERTOOLS, IDM_USERTOOLSDIALOG, 0, false, false )
   setTopMenuMRUItem( getNextFreeMRUindex(), IDM_USERTOOLS, 0, 0, false, true )
   
   for i as long = lbound(gConfig.Tools) to ubound(gConfig.Tools)
      if gConfig.Tools(i).IsDisplayMenu then
         hasTools = true
         setTopMenuMRUItem( getNextFreeMRUindex(), IDM_USERTOOLS, IDM_USERTOOLSBASE + i, 0, false, false )
         wszText = gConfig.Tools(i).wszDescription
         if gConfig.Tools(i).wszKey then
            wszText = wszText & chr(9) & gConfig.Tools(i).wszKey
         end if
         txtWidth = getTextWidth( HWND_FRMMAIN_MENUBAR, wszText, ghMenuBar.hFontMenuBar, 30 )
         if txtWidth > nMenuWidth then nMenuWidth = txtWidth
      end if   
   next

   if hasTools = false then
      setTopMenuMRUItem( getNextFreeMRUindex(), IDM_USERTOOLS, IDM_USERTOOLSLIST, 0, false, false )  ' (empty)
   end if   
   
   function = nMenuWidth
end function


' ========================================================================================
' Prepare and Execute a user tool
' ========================================================================================
function frmUserTools_ExecuteUserTool( ByVal nToolNum As Long ) As Long            
            
   Dim ShExecInfo  As SHELLEXECUTEINFOW  
   Dim pDocMain    As clsDocument Ptr 
   Dim pDocCurrent As clsDocument Ptr 
   
   dim wszCommand       As CWSTR
   dim wszParameters    As CWSTR
   dim wszProjectName   As CWSTR
   dim wszMainSource    As CWSTR
   dim wszExeName       As CWSTR
   dim wszWord          As CWSTR
   dim wszWorkingFolder as CWSTR 
   dim sCurrentDir      As String
   
   dim nWndStyle      As Long = SW_SHOWNORMAL

    
   ' Make sure that the incoming Tool Number is valid
   If (nToolNum < LBound(gConfig.Tools)) Orelse (nToolNum > UBound(gConfig.Tools)) Then exit function
       
   If gConfig.Tools(nToolNum).IsPromptRun Then
      If MessageBox( HWND_FRMMAIN, L(304,"Please confirm that you wish to run the following User Tool:") + vbcrlf + _
                     gConfig.Tools(nToolNum).wszDescription, L(276,"Confirm"), _
                     MB_ICONQUESTION Or MB_YESNOCANCEL ) <> IDYES Then 
         exit function
      End If               
   End If

   if gConfig.Tools(nToolNum).IsMinimized then nWndStyle = SW_SHOWMINNOACTIVE

   ' Prepare the command line based on any substitute parameters
   If Len(gConfig.Tools(nToolNum).wszParameters) Then
        
      pDocCurrent = gTTabCtl.GetActiveDocumentPtr()
      
      ' Do the substitutions for the command line parameters
      If gApp.IsProjectActive Then 
         wszProjectName = AfxStrPathname("NAMEX", gApp.ProjectFilename)
         pDocMain = gApp.GetMainDocumentPtr()
      else
         pDocMain = pDocCurrent
      END IF 

      if pDocMain THEN wszMainSource = chr(34) + pDocMain->DiskFilename + chr(34)
          
      wszParameters = gConfig.Tools(nToolNum).wszParameters
      wszParameters = AfxStrReplace(wszParameters, "<P>", wszProjectName)
      wszParameters = AfxStrReplace(wszParameters, "<p>", wszProjectName)
      wszParameters = AfxStrReplace(wszParameters, "<S>", wszMainSource)
      wszParameters = AfxStrReplace(wszParameters, "<s>", wszMainSource)

      dim as long idxBuild = frmBuildConfig_getActiveBuildIndex()
      if (idxBuild > -1) andalso (pDocMain <> 0) THEN
         wszExeName = gCompile.OutputFilename
         if len(wszExeName) = 0 then
            wszExeName = AfxStrPathname("PATH", pDocMain->DiskFilename) + _
                         AfxStrPathname("NAME", pDocMain->DiskFilename) + _
                         ".exe"
            If Instr(Ucase(gConfig.Builds(idxBuild).wszOptions), " -DLL") orelse _
               Instr(Ucase(gConfig.Builds(idxBuild).wszOptions), " -DYLIB") Then 
               wszExeName = AfxStrPathname("PATH", pDocMain->DiskFilename) + _
                            AfxStrPathname("NAME", pDocMain->DiskFilename) + _
                            ".dll"
            end if
            If Instr(Ucase(gConfig.Builds(idxBuild).wszOptions), " -LIB") then 
               wszExeName = AfxStrPathname("PATH", pDocMain->DiskFilename) + _
                            AfxStrPathname("NAME", pDocMain->DiskFilename) + _
                            ".a"
            end if              
         end if
      END IF
      wszParameters = AfxStrReplace(wszParameters, "<E>", wszExeName)
      wszParameters = AfxStrReplace(wszParameters, "<e>", wszExeName)
       
      if pDocCurrent Then 
          ' Determine the word underneath the cursor
         wszWord = Trim(pDocCurrent->GetWord, Any " ()*&^%$#@!~`:;'><?/.,-+=\|}{[]")
      End If
      wszParameters = AfxStrReplace(wszParameters, "<W>", wszWord)
      wszParameters = AfxStrReplace(wszParameters, "<w>", wszWord)
        
   end if
       
   ' Change the working folder
   sCurrentDir = CurDir
   ' Is the specified working folder valid? If it is not, then maybe try same path 
   ' but use the current drive.
   wszWorkingFolder = gConfig.Tools(nToolNum).wszWorkingFolder


   If Len(wszWorkingFolder) Then
      ' Convert relative path to absolute path if needed.
      if AfxPathIsRelative(wszWorkingFolder) then
         wszWorkingFolder = AfxPathCombine(AfxGetExePathName, wszWorkingFolder)
      END IF
      if AfxPathIsDirectory(wszWorkingFolder) = false then
         wszWorkingFolder = left(AfxGetExePathName, 1) + mid(wszWorkingFolder, 2)
      end if
      ChDir gConfig.Tools(nToolNum).wszWorkingFolder
   End If
       
   ' Is the specified exe command to execute valid? If it is not, then maybe try same path 
   ' but use the current drive.
   wszCommand = gConfig.Tools(nToolNum).wszCommand 
   ' Convert relative path to absolute path if needed.
   if AfxPathIsRelative(wszCommand) then
      wszCommand = AfxPathCombine(AfxGetExePathName, wszCommand)
   END IF
   if AfxPathFileExists(wszCommand) = false THEN
      wszCommand = left(AfxGetExePathName, 1) + mid(wszCommand, 2)
   END IF
   wszCommand = wchr(34) + wszCommand + wchr(34)

   With ShExecInfo
      .cbSize       = Len(SHELLEXECUTEINFOW)
      .fMask        = SEE_MASK_NOCLOSEPROCESS
      .HWnd         = 0
      .lpVerb       = Null
      .lpFile       = wszCommand
      .lpParameters = wszParameters   
      .lpDirectory  = 0
      .nShow        = nWndStyle
      .hInstApp     = 0 
   End With
   ShellExecuteEx(@ShExecInfo)
       
   if gConfig.Tools(nToolNum).IsWaitFinish then
      ' Give the process 100 ms 
      while WaitForSingleObject(ShExecInfo.hProcess, 100) = WAIT_TIMEOUT
      wend
   end if   

   ' Restore the current drive/directory
   ChDir sCurrentDir

   function = 0    
End Function

  
' ========================================================================================
' Create the keyboard shortcut text description.
' ========================================================================================
function createToolsMenuShortcut( byval nCtrlID as long ) as CWSTR
   dim wszShortcut as CWSTR
      
   If gConfig.Tools(nCtrlID).IsCtrl Then wszShortcut = "Ctrl+"
   If gConfig.Tools(nCtrlID).IsShift Then wszShortcut = wszShortcut + "Shift+"
   If gConfig.Tools(nCtrlID).IsAlt Then wszShortcut = wszShortcut + "Alt+"
   wszShortcut = trim(wszShortcut + UCase(Trim(gConfig.Tools(nCtrlID).wszKey)), any " +")
   
   return wszShortcut
end function



' ========================================================================================
' Create an accelerator table based on selections for the user tools.
' Also adds the Tools to the top menu.
' ========================================================================================
function frmUserTools_CreateAcceleratorTable() As Long

   dim NumKeys  As Long
   dim nKey     As Long
   dim fVirt    As Long
     
   If ghAccelUserTools Then DestroyAcceleratorTable(ghAccelUserTools)
   Dim a(any) As ACCEL
      
   NumKeys = 0
   For y as long = LBound(gConfig.Tools) To UBound(gConfig.Tools)
          
      fVirt = FNOINVERT Or FVIRTKEY
      
      If gConfig.Tools(y).IsCtrl Then fVirt = fVirt Or FCONTROL
      If gConfig.Tools(y).IsShift Then fVirt = fVirt Or FSHIFT
      If gConfig.Tools(y).IsAlt Then fVirt = fVirt Or FALT
          
      nKey = 0   ' important to reset
      Select Case UCase(Trim(gConfig.Tools(y).wszKey))
         Case "F1":  nKey = VK_F1
         Case "F2":  nKey = VK_F2
         Case "F3":  nKey = VK_F3
         Case "F4":  nKey = VK_F4
         Case "F5":  nKey = VK_F5
         Case "F6":  nKey = VK_F6
         Case "F7":  nKey = VK_F7
         Case "F8":  nKey = VK_F8
         Case "F9":  nKey = VK_F9
         Case "F10": nKey = VK_F10
         Case "F11": nKey = VK_F11
         Case "F12": nKey = VK_F12
         Case "0":   nKey = VK_0
         Case "1":   nKey = VK_1
         Case "2":   nKey = VK_2
         Case "3":   nKey = VK_3
         Case "4":   nKey = VK_4
         Case "5":   nKey = VK_5
         Case "6":   nKey = VK_6
         Case "7":   nKey = VK_7
         Case "8":   nKey = VK_8
         Case "9":   nKey = VK_9
         Case "A":   nKey = VK_A
         Case "B":   nKey = VK_B
         Case "C":   nKey = VK_C
         Case "D":   nKey = VK_D
         Case "E":   nKey = VK_E
         Case "F":   nKey = VK_F
         Case "G":   nKey = VK_G
         Case "H":   nKey = VK_H
         Case "I":   nKey = VK_I
         Case "J":   nKey = VK_J
         Case "K":   nKey = VK_K
         Case "L":   nKey = VK_L
         Case "M":   nKey = VK_M
         Case "N":   nKey = VK_N
         Case "O":   nKey = VK_O
         Case "P":   nKey = VK_P  
         Case "Q":   nKey = VK_Q  
         Case "R":   nKey = VK_R  
         Case "S":   nKey = VK_S  
         Case "T":   nKey = VK_T  
         Case "U":   nKey = VK_U  
         Case "V":   nKey = VK_V  
         Case "W":   nKey = VK_W
         Case "X":   nKey = VK_X
         Case "Y":   nKey = VK_Y
         Case "Z":   nKey = VK_Z  
      End Select
         
      If nKey > 0 Then
         NumKeys = NumKeys + 1
         ReDim Preserve a(1 To NumKeys) As ACCEL
         a(NumKeys).fVirt = fVirt
         a(NumKeys).key   = nKey
         a(NumKeys).cmd   = IDM_USERTOOLSBASE + y
      End If

   Next                                    
      
   If NumKeys Then
      ghAccelUserTools = CreateAcceleratorTable(CAST(LPACCEL, @a(1)), NumKeys)
   End If

   function = 0
End Function


' ========================================================================================
' Load all of the Tools descriptions into the listbox
' ========================================================================================
function frmUserTools_LoadListBox( byval hParent as hwnd ) as Long
   dim hList1 as hwnd = GetDlgItem(hParent, IDC_FRMUSERTOOLS_LSTTOOLS)

   ListBox_ResetContent(hList1)
   for i as long = lbound(gConfig.ToolsTemp) to ubound(gConfig.ToolsTemp)
      ListBox_AddString(hList1, gConfig.ToolsTemp(i).wszDescription.sptr)
   NEXT

   function = 0
end function


' ========================================================================================
' Swap two entries in the Tools Listbox
' ========================================================================================
function frmUserTools_SwapListBoxItems( _
   byval Item1 as long, _
   byval Item2 as long _
   ) as Long
   dim as hwnd hList1 = GetDlgItem( HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_LSTTOOLS)
   
   ' Swap the array values
   swap gConfig.ToolsTemp(Item1), gConfig.ToolsTemp(Item2)
   
   ListBox_ReplaceString(hList1, Item1, gConfig.ToolsTemp(Item1).wszDescription)
   ListBox_ReplaceString(hList1, Item2, gConfig.ToolsTemp(Item2).wszDescription)

   function = 0
end function


' ========================================================================================
' Set the UserTools textboxes and options depending on what listbox entry is selected
' ========================================================================================
function frmUserTools_SetTextboxes() as long
   dim as long nCurSel = ListBox_GetCurSel(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_LSTTOOLS))
   dim as Boolean fEnabled = iif( nCurSel < 0, false, true)

   AfxSetWindowText(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTTOOLNAME), iif(fEnabled, gConfig.ToolsTemp(nCurSel).wszDescription, wstr("")))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTTOOLNAME), fEnabled)
   AfxSetWindowText(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTCOMMAND), iif(fEnabled, gConfig.ToolsTemp(nCurSel).wszCommand, wstr("")))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTCOMMAND), fEnabled)
   AfxSetWindowText(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTPARAMETERS), iif(fEnabled, gConfig.ToolsTemp(nCurSel).wszParameters, wstr("")))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTPARAMETERS), fEnabled)
   AfxSetWindowText(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTKEY), iif(fEnabled, gConfig.ToolsTemp(nCurSel).wszKey, wstr("")))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTKEY), fEnabled)
   AfxSetWindowText(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTWORKINGFOLDER), iif(fEnabled, gConfig.ToolsTemp(nCurSel).wszWorkingFolder, wstr("")))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_TXTWORKINGFOLDER), fEnabled)
   Button_SetCheck(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKCTRL), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsCtrl, 0))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKCTRL), fEnabled)
   Button_SetCheck(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKALT), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsAlt, 0))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKALT), fEnabled)
   Button_SetCheck(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKSHIFT), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsShift, 0))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKSHIFT), fEnabled)
   Button_SetCheck(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKPROMPT), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsPromptRun, 0))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKPROMPT), fEnabled)
   Button_SetCheck(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKMINIMIZED), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsMinimized, 0))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKMINIMIZED), fEnabled)
   Button_SetCheck(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKWAIT), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsWaitFinish, 0))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKWAIT), fEnabled)
   Button_SetCheck(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKDISPLAYMENU), iif(fEnabled, gConfig.ToolsTemp(nCurSel).IsDisplayMenu, 0))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_CHKDISPLAYMENU), fEnabled)
   ComboBox_SetCurSel(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_COMBOACTION), iif(fEnabled, gConfig.ToolsTemp(nCurSel).Action, 0))
   EnableWindow(GetDlgItem(HWND_FRMUSERTOOLS, IDC_FRMUSERTOOLS_COMBOACTION), fEnabled)

   function = 0
end function

' ========================================================================================
' Process WM_CREATE message for window/dialog: frmUserTools
' ========================================================================================
function frmUserTools_OnCreate( _
   byval HWnd As HWnd, _
   byval lpCreateStructPtr As LPCREATESTRUCT _
   ) as boolean

   ' This is a modal popup window so disable the parent window
   DisableAllModeless()

   '  Message cracker macro expects a True to be returned for a successful
   '  OnCreate handler even though returning -1 from a standard WM_CREATE
   '  call would stop creating the window. This is just one of those Windows
   '  inconsistencies.
   Return True
End Function


' ========================================================================================
' Process WM_COMMAND message for window/dialog: frmUserTools
' ========================================================================================
function frmUserTools_OnCommand( _
   ByVal HWnd As HWnd, _
   ByVal id As Long, _
   ByVal hwndCtl As HWnd, _
   ByVal codeNotify As UINT _
   ) As LRESULT

   dim as hwnd hList1 = GetDlgItem( HWND, IDC_FRMUSERTOOLS_LSTTOOLS)
   dim as long nCurSel = ListBox_GetCurSel(hList1)

   Select Case id
      case IDC_FRMUSERTOOLS_LSTTOOLS
         if codeNotify = LBN_SELCHANGE THEN
            frmUserTools_SetTextboxes()
         END IF
           
      case IDC_FRMUSERTOOLS_TXTTOOLNAME
         if codeNotify = EN_CHANGE THEN
            ' Update the temp array and the Listbox
            if nCurSel > -1 THEN
               gConfig.ToolsTemp(nCurSel).wszDescription = AfxGetWindowText(hwndCtl)
               ListBox_ReplaceString(hList1, nCurSel, gConfig.ToolsTemp(nCurSel).wszDescription)
            END IF
         end if
          
      CASE IDC_FRMUSERTOOLS_TXTCOMMAND
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).wszCommand = AfxGetWindowText(hwndCtl)
         end if
          
      case IDC_FRMUSERTOOLS_CMDBROWSEEXE
         if codeNotify = BN_CLICKED THEN
            Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWnd, IDC_FRMUSERTOOLS_CMDBROWSEEXE)
            If pwszName Then
               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMUSERTOOLS_TXTCOMMAND), pwszName )
               CoTaskMemFree pwszName
            End If
         end if
         
      CASE IDC_FRMUSERTOOLS_TXTPARAMETERS
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).wszParameters = AfxGetWindowText(hwndCtl)
         end if

      CASE IDC_FRMUSERTOOLS_TXTKEY
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).wszKey = AfxGetWindowText(hwndCtl)
         end if

      CASE IDC_FRMUSERTOOLS_TXTWORKINGFOLDER
         if codeNotify = EN_CHANGE THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).wszWorkingFolder = AfxGetWindowText(hwndCtl)
         end if
                   
      case IDC_FRMUSERTOOLS_CMDBROWSEFOLDER
         if codeNotify = BN_CLICKED THEN
            dim as CWSTR cwsFolder = AfxBrowseForFolder(HWND, L(293,"Working Folder:"), curdir)
            if len(cwsFolder) THEN
               AfxSetWindowText( GetDlgItem(HWnd, IDC_FRMUSERTOOLS_TXTWORKINGFOLDER), cwsFolder)
            END IF
         end if
         
      case IDC_FRMUSERTOOLS_CHKCTRL
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsCtrl = Button_GetCheck(hwndCtl) 
         end if

      case IDC_FRMUSERTOOLS_CHKALT
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsAlt = Button_GetCheck(hwndCtl) 
         end if

      case IDC_FRMUSERTOOLS_CHKSHIFT
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsShift = Button_GetCheck(hwndCtl) 
         end if

      case IDC_FRMUSERTOOLS_CHKPROMPT
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsPromptRun = Button_GetCheck(hwndCtl) 
         end if

      case IDC_FRMUSERTOOLS_CHKMINIMIZED
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsMinimized = Button_GetCheck(hwndCtl) 
         end if

      case IDC_FRMUSERTOOLS_CHKWAIT
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsWaitFinish = Button_GetCheck(hwndCtl) 
         end if

      case IDC_FRMUSERTOOLS_CHKDISPLAYMENU
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1 THEN gConfig.ToolsTemp(nCurSel).IsDisplayMenu = Button_GetCheck(hwndCtl) 
         end if

      case IDC_FRMUSERTOOLS_COMBOACTION
         if codeNotify = CBN_SELCHANGE THEN
            if nCurSel > -1 then gConfig.ToolsTemp(nCurSel).Action = ComboBox_GetCurSel(hwndCtl) 
         END IF

      case IDC_FRMUSERTOOLS_CMDUP
         if codeNotify = BN_CLICKED THEN
            if nCurSel > 0 THEN
               frmUserTools_SwapListboxItems(nCurSel, nCurSel-1)
            END IF
         end if

      case IDC_FRMUSERTOOLS_CMDDOWN
         if codeNotify = BN_CLICKED THEN
            if nCurSel < ListBox_GetCount(hList1)-1 THEN
               frmUserTools_SwapListboxItems(nCurSel, nCurSel+1)
            END IF
         end if

      case IDC_FRMUSERTOOLS_CMDINSERT
         if codeNotify = BN_CLICKED THEN
            if ubound(gConfig.ToolsTemp) = -1 THEN
               redim gConfig.ToolsTemp(0)
               nCurSel = 0
            Else
               redim preserve gConfig.ToolsTemp(ubound(gConfig.ToolsTemp)+1)
               if nCurSel = -1 THEN nCurSel = 0
               ' insert the item above current entry in the internal array
               for i as long = ubound(gConfig.ToolsTemp) to nCurSel + 1 step -1
                  gConfig.ToolsTemp(i) = gConfig.ToolsTemp(i-1)
               NEXT
            END IF
            gConfig.ToolsTemp(nCurSel).wszDescription   = ""
            gConfig.ToolsTemp(nCurSel).wszCommand       = ""
            gConfig.ToolsTemp(nCurSel).wszParameters    = ""
            gConfig.ToolsTemp(nCurSel).wszKey           = ""
            gConfig.ToolsTemp(nCurSel).wszWorkingFolder = ""
            gConfig.ToolsTemp(nCurSel).IsCtrl           = 0
            gConfig.ToolsTemp(nCurSel).IsAlt            = 0
            gConfig.ToolsTemp(nCurSel).IsShift          = 0
            gConfig.ToolsTemp(nCurSel).IsPromptRun      = 0
            gConfig.ToolsTemp(nCurSel).IsMinimized      = 0
            gConfig.ToolsTemp(nCurSel).IsWaitFinish     = 0
            gConfig.ToolsTemp(nCurSel).IsDisplayMenu    = 0
            ' reload the listbox
            frmUserTools_LoadListBox(HWND)
            nCurSel = Min(nCurSel, ubound(gConfig.ToolsTemp))
            ListBox_SetCurSel(hList1, nCurSel)
            frmUserTools_SetTextboxes()
            SetFocus GetDlgItem(HWND, IDC_FRMUSERTOOLS_TXTTOOLNAME)
         end if

      case IDC_FRMUSERTOOLS_CMDDELETE
         if codeNotify = BN_CLICKED THEN
            if nCurSel > -1  THEN
               if MessageBox( HWND, L(300, "Are you sure you want to delete this user tool?"), L(276,"Confirm"), _
                                 MB_YESNOCANCEL Or MB_ICONINFORMATION Or MB_DEFBUTTON1 ) = IDYES then
                  if ubound(gConfig.ToolsTemp) = 0 THEN
                     erase gConfig.ToolsTemp
                     nCurSel = -1
                  else
                     ' remove the item from the internal array
                     for i as long = nCurSel to ubound(gConfig.ToolsTemp) - 1
                        gConfig.ToolsTemp(i) = gConfig.ToolsTemp(i+1)
                     NEXT
                     redim preserve gConfig.ToolsTemp(ubound(gConfig.ToolsTemp)-1)
                  END IF
                  ' reload the listbox
                  frmUserTools_LoadListBox(HWND)
                  nCurSel = Min(nCurSel, ubound(gConfig.ToolsTemp))
                  ListBox_SetCurSel(hList1, nCurSel)
                  frmUserTools_SetTextboxes()
                  SetFocus hList1
               end if
            END IF
         end if

      Case IDC_FRMUSERTOOLS_CMDOK    
         If codeNotify = BN_CLICKED Then
            ' Copy the temporary items to the main array  
            redim gConfig.Tools(ubound(gConfig.ToolsTemp))
            for i as long = lbound(gConfig.ToolsTemp) to ubound(gConfig.ToolsTemp)
               gConfig.Tools(i) = gConfig.ToolsTemp(i)   
            next
            erase gConfig.ToolsTemp
            frmUserTools_CreateAcceleratorTable            
            gConfig.SaveConfigFile
            updateUserToolsMenuItems()
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         end if
             
      Case IDCANCEL
         If codeNotify = BN_CLICKED Then
            SendMessage( HWnd, WM_CLOSE, 0, 0 )
            Exit Function
         End If
   End Select

   Function = 0
End Function


' ========================================================================================
' Process WM_CLOSE message for window/dialog: frmUserTools
' ========================================================================================
private Function frmUserTools_OnClose( byval HWnd As HWnd ) As LRESULT
   ' Enables parent window keeping parent's zorder
   EnableAllModeless()
   DestroyWindow( HWnd )
   Function = 0
End Function


' ========================================================================================
' Process WM_DESTROY message for window/dialog: frmUserTools
' ========================================================================================
private Function frmUserTools_OnDestroy( byval HWnd As HWnd ) As LRESULT
   PostQuitMessage(0)
   Function = 0
End Function


' ========================================================================================
' frmUserTools Window procedure
' ========================================================================================
private Function frmUserTools_WndProc( ByVal HWnd   As HWnd, _
                                       ByVal uMsg   As UINT, _
                                       ByVal wParam As WPARAM, _
                                       ByVal lParam As LPARAM _
                                       ) As LRESULT

   Select Case uMsg
      HANDLE_MSG (HWnd, WM_CREATE,   frmUserTools_OnCreate)
      HANDLE_MSG (HWnd, WM_CLOSE,    frmUserTools_OnClose)
      HANDLE_MSG (HWnd, WM_DESTROY,  frmUserTools_OnDestroy)
      HANDLE_MSG (HWnd, WM_COMMAND,  frmUserTools_OnCommand)
   End Select

   Function = DefWindowProc(HWnd, uMsg, wParam, lParam)

End Function


' ========================================================================================
' frmUserTools_Show
' ========================================================================================
public Function frmUserTools_Show( ByVal hWndParent As HWnd ) as LRESULT

   DIM hBitmap AS HBITMAP
   dim hCtrl as HWnd
   dim wszImage as wstring * 100
   
   '  Create the main window and child controls
   Dim pWindow as CWindow Ptr = New CWindow
   pWindow->DPI = AfxCWindowOwnerPtr(hwndParent)->DPI

   Dim As HWnd hForm = _
   pWindow->Create(hWndParent, L(289,"User Tools"), _
        @frmUserTools_WndProc, 0, 0, 683, 465, _
        WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, _
        WS_EX_DLGMODALFRAME Or WS_EX_CONTROLPARENT Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->SetClientSize(677, 436)
   pWindow->Center

   pWindow->AddControl("LISTBOX", , IDC_FRMUSERTOOLS_LSTTOOLS, "", 9, 10, 218, 362, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or LBS_NOTIFY Or LBS_NOINTEGRALHEIGHT, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMUSERTOOLS_CMDINSERT, L(281, "Insert"), 8, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMUSERTOOLS_CMDDELETE, L(282, "Delete"), 87, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   
   hCtrl = _
   pWindow->AddControl("BUTTON", , IDC_FRMUSERTOOLS_CMDUP, wszTriangleUp, 166, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghMenuBar.hFontSymbolLargeBold, false )     

   hCtrl = _
   pWindow->AddControl("BUTTON", , IDC_FRMUSERTOOLS_CMDDOWN, wszTriangleDown, 199, 388, 28, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   AfxSetWindowFont( hCtrl, ghMenuBar.hFontSymbolLargeBold, false )     

   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL1, L(290,"Tool Name") & ":", 230, 25, 95, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMUSERTOOLS_TXTTOOLNAME, "", 334, 22, 305, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL2, L(291,"Command") & ":", 230, 50, 95, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMUSERTOOLS_TXTCOMMAND, "", 334, 47, 267, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMUSERTOOLS_CMDBROWSEEXE, "...", 610, 46, 30, 23, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL3, L(292,"Parameters") & ":", 230, 75, 95, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMUSERTOOLS_TXTPARAMETERS, "", 334, 72, 305, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL4, "<P> Project Name  <S> Main Source File  <W> Current Word  <E> Exe/DLL Name", 334, 101, 302, 36, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL5, L(293,"Working Folder") & ":", 230, 177, 95, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMUSERTOOLS_TXTWORKINGFOLDER, "", 334, 174, 267, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("BUTTON", , IDC_FRMUSERTOOLS_CMDBROWSEFOLDER, "...", 610, 173, 30, 23, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL6, L(294,"Accelerator") & ":", 230, 205, 95, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMUSERTOOLS_CHKCTRL, "Ctrl", 335, 202, 48, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMUSERTOOLS_CHKALT, "Alt", 391, 202, 42, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMUSERTOOLS_CHKSHIFT, "Shift", 441, 202, 48, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL7, "Key", 559, 204, 38, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("TEXTBOX", , IDC_FRMUSERTOOLS_TXTKEY, "", 498, 201, 52, 20, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or ES_LEFT Or ES_AUTOHSCROLL, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("CHECKBOX", , IDC_FRMUSERTOOLS_CHKPROMPT, L(295,"Prompt for confirmation when Tool is invoked"), 334, 259, 347, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMUSERTOOLS_CHKMINIMIZED, L(296,"Run minimized"), 334, 281, 214, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMUSERTOOLS_CHKWAIT, L(297,"Wait for Tool to complete before continuing"), 334, 303, 344, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("CHECKBOX", , IDC_FRMUSERTOOLS_CHKDISPLAYMENU, L(299,"Display this item in the editor menu"), 334, 325, 344, 19, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_NOTIFY Or BS_AUTOCHECKBOX Or BS_LEFT Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("COMBOBOX", , IDC_FRMUSERTOOLS_COMBOACTION, "", 334, 228, 308, 22, _
        WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or WS_TABSTOP Or CBS_DROPDOWNLIST, _
        WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)
   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL8, L(298,"Action:"), 230, 232, 95, 18, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_RIGHT Or SS_NOTIFY, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("LABEL", , IDC_FRMUSERTOOLS_LABEL9, "", 238, 370, 425, 2, _
        WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY Or SS_SUNKEN, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDC_FRMUSERTOOLS_CMDOK, L(0,"OK"), 509, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)
   pWindow->AddControl("BUTTON", , IDCANCEL, L(1,"Cancel"), 591, 388, 74, 28, _
        WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_TEXT Or BS_PUSHBUTTON Or BS_NOTIFY Or BS_CENTER Or BS_VCENTER, _
        WS_EX_LEFT Or WS_EX_LTRREADING)

   HWND_FRMUSERTOOLS = hForm

   ' Copy all of the Tools to the ToolsTemp array because we will work with 
   ' temporary copies until the user hits OK.
   redim gConfig.ToolsTemp(ubound(gConfig.Tools))
   for i as long = lbound(gConfig.Tools) to ubound(gConfig.Tools)
      gConfig.ToolsTemp(i) = gConfig.Tools(i)   
   NEXT
   frmUserTools_LoadListBox(hForm)

   ' Load the Actions combobox
   dim hCombo as hwnd = GetDlgItem(hForm, IDC_FRMUSERTOOLS_COMBOACTION)
   ComboBox_AddString( hCombo, @L(303,"Invoke only when selected by user") ) 
   ComboBox_AddString( hCombo, @L(302,"Invoke during pre-compile") ) 
   ComboBox_AddString( hCombo, @L(301,"Invoke during post-compile") ) 
   ComboBox_AddString( hCombo, @L(402,"Invoke immediately after WinFBE starts") ) 
   
   dim hList as hwnd = GetDlgItem(hForm, IDC_FRMUSERTOOLS_LSTTOOLS)
   ListBox_SetCurSel(hList, 0)
   frmUserTools_SetTextboxes()
   SetFocus GetDlgItem(hForm, IDC_FRMUSERTOOLS_LSTTOOLS)

   ' Process Windows messages
   Function = pWindow->DoEvents(SW_SHOW)
   
   ' Delete the frmUserTools CWindow class manually allocated memory 
   Delete pWindow

   function = 0
end function

