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


''  CONFIGURATION MODULE



''
''  CONSTRUCTOR
''
Constructor clsConfig
   _ConfigFilename = Exepath & "\Settings\WinFBE.ini" 
End Constructor


''
''  SAVE CONFIGURATION TO DISK FILE
''
Function clsConfig.SaveToFile() as Long

   ' Determine the current editor positioning 
   Dim WinPla As WINDOWPLACEMENT
   WinPla.Length = Sizeof(WinPla)
   GetWindowPlacement(HWND_FRMMAIN, @WinPla)
   With this
      .StartupLeft   = WinPla.rcNormalPosition.Left 
      .StartupTop    = WinPla.rcNormalPosition.Top
      .StartupRight  = WinPla.rcNormalPosition.Right
      .StartupBottom = WinPla.rcNormalPosition.Bottom
      .StartupMaximized = Iif( WinPla.showCmd = SW_MAXIMIZE, True, False )
   End With

   Dim f As Long = Freefile
   If Open( _ConfigFilename For Output Encoding "utf16" As #f ) <> 0 Then Exit Function

   Print #f, "'  WINFBE CONFIGURATION" 
   print #f, ""
   print #f, "[Editor]"
   print #f, "MultipleInstances="     ; this.MultipleInstances    
   print #f, "CompileAutosave="       ; this.CompileAutosave      
   print #f, "CloseFuncList="         ; this.CloseFuncList        
   print #f, "HideCompile="           ; this.HideCompile          
   print #f, "SyntaxHighlighting="    ; this.SyntaxHighlighting   
   print #f, "Codetips="              ; this.Codetips             
   print #f, "AutoComplete="          ; this.AutoComplete
   print #f, "LeftMargin="            ; this.LeftMargin           
   print #f, "FoldMargin="            ; this.FoldMargin           
   print #f, "AutoIndentation="       ; this.AutoIndentation      
   print #f, "ConfineCaret="          ; this.ConfineCaret         
   print #f, "LineNumbering="         ; this.LineNumbering        
   print #f, "HighlightCurrentLine="  ; this.HighlightCurrentLine 
   print #f, "IndentGuides="          ; this.IndentGuides         
   print #f, "TabIndentSpaces="       ; this.TabIndentSpaces      
   print #f, "UnicodeEncoding="       ; this.UnicodeEncoding      
   print #f, "KeywordCase="           ; this.KeywordCase          
   print #f, "LocalizationFile="      ; this.LocalizationFile     
   print #f, "TabSize="               ; this.TabSize              
   print #f, "EditorFontname="        ; this.EditorFontname       
   print #f, "EditorFontsize="        ; this.EditorFontsize       
   print #f, "EditorFontCharSet="     ; this.EditorFontCharSet    
   print #f, ""
   print #f, "[Startup]" 
   print #f, "StartupLeft="           ; this.StartupLeft          
   print #f, "StartupTop="            ; this.StartupTop           
   print #f, "StartupRight="          ; this.StartupRight         
   print #f, "StartupBottom="         ; this.StartupBottom        
   print #f, "StartupMaximized="      ; this.StartupMaximized     
   print #f, "ShowExplorer="          ; this.ShowExplorer         
   print #f, "ShowExplorerWidth="     ; this.ShowExplorerWidth    
   print #f, ""

   ' For each folder location determine if it resides on the same drive as
   ' the WinFBE application. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   this.FBWINCompiler32  = ProcessToCurdrive(this.FBWINCompiler32)
   this.FBWINCompiler64  = ProcessToCurdrive(this.FBWINCompiler64)
   this.CompilerHelpfile = ProcessToCurdrive(this.CompilerHelpfile)
   this.Win32APIHelpfile = ProcessToCurdrive(this.Win32APIHelpfile)

   Print #f, "[Compiler]" 
   print #f, "FBWINCompiler32="       ; this.FBWINCompiler32      
   print #f, "FBWINCompiler64="       ; this.FBWINCompiler64      
   print #f, "CompilerSwitches="      ; this.CompilerSwitches     
   print #f, "CompilerHelpfile="      ; this.CompilerHelpfile     
   print #f, "Win32APIHelpfile="      ; this.Win32APIHelpfile     
   print #f, "DefaultCompiler="       ; this.DefaultCompiler      
   print #f, "DefaultCompileMode="    ; this.DefaultCompileMode   
   
   print #f, ""
   Print #f, "[MRU]"  
   For i As Long = 0 To 9
      this.MRU(i) = ProcessToCurdrive(this.MRU(i))
      Print #f, "MRU_"; Right("00" & Str(i), 2); "="; this.MRU(i) 
   Next         
   
   Print #f, ""
   Print #f, "[MRUPROJECTS]"  
   For i As Long = 0 To 9
      this.MRUProject(i) = ProcessToCurdrive(this.MRUProject(i))
      Print #f, "MRUPROJECT_"; Right("00" & Str(i), 2); "="; this.MRUProject(i) 
   Next         

   Print #f, ""
   print #f, "[Colors]" 
   print #f, "clrCaretFG="            ; this.clrCaretFG           
   print #f, "clrCaretBG="            ; this.clrCaretBG           
   print #f, "clrCommentsFG="         ; this.clrCommentsFG        
   print #f, "clrCommentsBG="         ; this.clrCommentsBG        
   print #f, "clrHighlightedFG="      ; this.clrHighlightedFG     
   print #f, "clrHighlightedBG="      ; this.clrHighlightedBG     
   print #f, "clrKeywordFG="          ; this.clrKeywordFG         
   print #f, "clrKeywordBG="          ; this.clrKeywordBG         
   print #f, "clrFoldMarginFG="       ; this.clrFoldMarginFG      
   print #f, "clrFoldMarginBG="       ; this.clrFoldMarginBG      
   print #f, "clrFoldSymbolFG="       ; this.clrFoldSymbolFG      
   print #f, "clrFoldSymbolBG="       ; this.clrFoldSymbolBG      
   print #f, "clrLineNumbersFG="      ; this.clrLineNumbersFG     
   print #f, "clrLineNumbersBG="      ; this.clrLineNumbersBG     
   print #f, "clrBookmarksFG="        ; this.clrBookmarksFG       
   print #f, "clrBookmarksBG="        ; this.clrBookmarksBG       
   print #f, "clrOperatorsFG="        ; this.clrOperatorsFG       
   print #f, "clrOperatorsBG="        ; this.clrOperatorsBG       
   print #f, "clrIndentGuidesFG="     ; this.clrIndentGuidesFG    
   print #f, "clrIndentGuidesBG="     ; this.clrIndentGuidesBG    
   print #f, "clrPreprocessorFG="     ; this.clrPreprocessorFG    
   print #f, "clrPreprocessorBG="     ; this.clrPreprocessorBG    
   print #f, "clrSelectionFG="        ; this.clrSelectionFG       
   print #f, "clrSelectionBG="        ; this.clrSelectionBG       
   print #f, "clrStringsFG="          ; this.clrStringsFG         
   print #f, "clrStringsBG="          ; this.clrStringsBG         
   print #f, "clrTextFG="             ; this.clrTextFG            
   print #f, "clrTextBG="             ; this.clrTextBG            
   print #f, "clrWinAPIFG="           ; this.clrWinAPIFG          
   print #f, "clrWinAPIBG="           ; this.clrWinAPIBG          
   print #f, "clrWindowFG="           ; this.clrWindowFG          
   print #f, "clrWindowBG="           ; this.clrWindowBG          
   
   Close #f
   
   Function = 0
End Function



''
''  LOAD CONFIGURATION FROM DISK FILE
''
Function clsConfig.LoadFromFile() as Long

   Dim st    As WString * MAX_PATH
   Dim wKey  As WString * MAX_PATH
   Dim wData As WString * MAX_PATH
   Dim nData As Long  
   Dim i     as Long

   If Dir(_ConfigFilename) = "" Then Exit Function
    
   Dim f as Long = Freefile
   If Open( _ConfigFilename For Input Encoding "utf16" As #f ) <> 0 Then Exit Function

   Do Until EOF(f)
      Line Input #f, st
      If Len(st) = 0 Then Continue Do
      If Left(st, 1) = "'" Then Continue Do
      If Left(st, 1) = "[" Then Continue Do
      
      i = InStr(st, "=")
      If i = 0 Then Continue Do
      
      wKey  = Left(st, i-1)
      wData = Mid(st, i+1)
      nData = Val(wData)
      
      wData = ProcessFromCurdrive(wData)

      If Left(wKey, 4) = "MRU_" Then  ' is this an MRU entry
         i = Val(Right(wKey,2))
         If (i >= 0) And (i <= 9) Then 
            this.MRU(i) = wData
            Continue Do
         End If
      End If
      
      If Left(wKey, 11) = "MRUPROJECT_" Then  ' is this an MRU Project entry
         i = Val(Right(wKey,2))
         If (i >= 0) And (i <= 9) Then 
            this.MRUProject(i) = wData
            Continue Do
         End If
      End If

      Select Case wKey
         Case "MultipleInstances":    this.MultipleInstances    = nData
         Case "CompileAutosave":      this.CompileAutosave      = nData
         Case "CloseFuncList":        this.CloseFuncList        = nData
         Case "HideCompile":          this.HideCompile          = nData
         Case "SyntaxHighlighting":   this.SyntaxHighlighting   = nData
         Case "Codetips":             this.Codetips             = nData
         Case "AutoComplete":         this.AutoComplete         = nData
         Case "LeftMargin":           this.LeftMargin           = nData
         Case "FoldMargin":           this.FoldMargin           = nData
         Case "AutoIndentation":      this.AutoIndentation      = nData
         Case "ConfineCaret":         this.ConfineCaret         = nData
         Case "LineNumbering":        this.LineNumbering        = nData
         Case "HighlightCurrentLine": this.HighlightCurrentLine = nData
         Case "IndentGuides":         this.IndentGuides         = nData
         Case "TabIndentSpaces":      this.TabIndentSpaces      = nData
         Case "LocalizationFile":     this.LocalizationFile     = wData
         Case "TabSize":              this.TabSize              = wData
         Case "UnicodeEncoding":      this.UnicodeEncoding      = nData
         Case "EditorFontname":       this.EditorFontname       = wData
         Case "EditorFontsize":       this.EditorFontsize       = wData
         Case "EditorFontCharSet":    this.EditorFontCharSet    = wData
         Case "KeywordCase":          this.KeywordCase          = nData
         Case "StartupLeft":          this.StartupLeft          = nData
         Case "StartupTop":           this.StartupTop           = nData
         Case "StartupRight":         this.StartupRight         = nData
         Case "StartupBottom":        this.StartupBottom        = nData
         Case "StartupMaximized":     this.StartupMaximized     = nData
         Case "ShowExplorer":         this.ShowExplorer         = nData
         Case "ShowExplorerWidth":    this.ShowExplorerWidth    = nData
         Case "FBWINCompiler32":      this.FBWINCompiler32      = wData
         Case "FBWINCompiler64":      this.FBWINCompiler64      = wData
         Case "CompilerSwitches":     this.CompilerSwitches     = wData
         Case "CompilerHelpfile":     this.CompilerHelpfile     = wData
         Case "Win32APIHelpfile":     this.Win32APIHelpfile     = wData
         Case "DefaultCompiler":      this.DefaultCompiler      = wData
         Case "DefaultCompileMode":   this.DefaultCompileMode   = wData
         Case "clrCaretFG":           this.clrCaretFG           = nData
         Case "clrCaretBG":           this.clrCaretBG           = nData
         Case "clrCommentsFG":        this.clrCommentsFG        = nData
         Case "clrCommentsBG":        this.clrCommentsBG        = nData
         Case "clrHighlightedFG":     this.clrHighlightedFG     = nData
         Case "clrHighlightedBG":     this.clrHighlightedBG     = nData
         Case "clrKeywordFG":         this.clrKeywordFG         = nData
         Case "clrKeywordBG":         this.clrKeywordBG         = nData
         Case "clrFoldMarginFG":      this.clrFoldMarginFG      = nData
         Case "clrFoldMarginBG":      this.clrFoldMarginBG      = nData
         Case "clrFoldSymbolFG":      this.clrFoldSymbolFG      = nData
         Case "clrFoldSymbolBG":      this.clrFoldSymbolBG      = nData
         Case "clrLineNumbersFG":     this.clrLineNumbersFG     = nData
         Case "clrLineNumbersBG":     this.clrLineNumbersBG     = nData
         Case "clrBookmarksFG":       this.clrBookmarksFG       = nData
         Case "clrBookmarksBG":       this.clrBookmarksBG       = nData
         Case "clrOperatorsFG":       this.clrOperatorsFG       = nData
         Case "clrOperatorsBG":       this.clrOperatorsBG       = nData
         Case "clrIndentGuidesFG":    this.clrIndentGuidesFG    = nData
         Case "clrIndentGuidesBG":    this.clrIndentGuidesBG    = nData
         Case "clrPreprocessorFG":    this.clrPreprocessorFG    = nData
         Case "clrPreprocessorBG":    this.clrPreprocessorBG    = nData
         Case "clrSelectionFG":       this.clrSelectionFG       = nData
         Case "clrSelectionBG":       this.clrSelectionBG       = nData
         Case "clrStringsFG":         this.clrStringsFG         = nData
         Case "clrStringsBG":         this.clrStringsBG         = nData
         Case "clrTextFG":            this.clrTextFG            = nData
         Case "clrTextBG":            this.clrTextBG            = nData
         Case "clrWinAPIFG":          this.clrWinAPIFG          = nData
         Case "clrWinAPIBG":          this.clrWinAPIBG          = nData
         Case "clrWindowFG":          this.clrWindowFG          = nData
         Case "clrWindowBG":          this.clrWindowBG          = nData 
      End Select
   
   Loop
   Close #f

   Function = 0
End Function


' ========================================================================================
' Save all options for the Project to a diskfile
' ========================================================================================
Function clsConfig.ProjectSaveToFile() As BOOLEAN    

   Dim wszText As WString * MAX_PATH
   Dim As Long nCount 
   Dim pDoc As clsDocument Ptr
   Dim As Long i, nActiveTab

   Dim f As Long = Freefile
   If Open( gApp.ProjectFilename For Output Encoding "utf16" As #f ) <> 0 Then Return False

   Print #f, "' WINFBE PROJECT FILE"
   Print #f, "ProjectType=";            gApp.ProjectType
   Print #f, "ProjectErrorOption=";     gApp.ProjectErrorOption
   Print #f, "ProjectDebug=";           gApp.ProjectDebug
   Print #f, "ProjectThread=";          gApp.ProjectThread
   Print #f, "ProjectShowConsole=";     gApp.ProjectShowConsole
   Print #f, "ProjectOther32=";         gApp.ProjectOther32
   Print #f, "ProjectOther64=";         gApp.ProjectOther64
   Print #f, "DefaultCompiler=";        gConfig.DefaultCompiler       
   Print #f, "DefaultCompileMode=";     gConfig.DefaultCompileMode    
   Print #f, "CommandLine=";            gConfig.CommandLine

   ' Save all of the loaded tabs first
   nCount = TabCtrl_GetItemCount(gTTabCtl.hWindow) 
   nActiveTab = TabCtrl_GetCurSel( gTTabCtl.hWindow )

   Print #f, "ActiveTab="; nActiveTab

   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM

   dim cwzRelative as CWSTR
   
   For i = 0 To nCount - 1
      TabCtrl_GetItem( gTTabCtl.hWindow, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If pDoc->IsProjectFile Then
         ' Only deal with files that are no longer "new"
         if pDoc->IsNewFlag = false then 
            wszText = pDoc->DiskFilename
            cwzRelative = AfxPathRelativePathTo( gApp.ProjectFilename, FILE_ATTRIBUTE_NORMAL, wszText, FILE_ATTRIBUTE_NORMAL)
            if AfxPathIsRelative(cwzRelative) then wszText = cwzRelative
            wszText = ProcessToCurdrive(wszText)
            Print #f, "File="; wszText
            Print #f, "FileType="; pDoc->ProjectFileType
            Print #f, "TabIndex="; -1     ' use -1 rather than True
            Print #f, "Bookmarks="; pDoc->GetBookmarks()
            Print #f, "FirstLine="; SendMessageW( pDoc->hWindow, SCI_GETFIRSTVISIBLELINE, 0, 0) 
            Print #f, "Position="; SendMessageW( pDoc->hWindow, SCI_GETCURRENTPOS, 0, 0)  
            Print #f, "FileEnd=[-]" 
         end if
      End If   
   Next
         
   ' Save all other non-displayed documents for the project
   nCount = gApp.GetDocumentCount 
   For i = 0 To nCount - 1
      pDoc = gApp.GetDocumentPtr(i)
      If pDoc Then 
         If pDoc->IsProjectFile Then
            ' Only deal with files that are no longer "new"
            if pDoc->IsNewFlag = false then 
               If gTTabCtl.GetTabIndexByDocumentPtr(pDoc) = -1 Then
                  wszText = pDoc->DiskFilename
                  cwzRelative = AfxPathRelativePathTo( gApp.ProjectFilename, FILE_ATTRIBUTE_NORMAL, wszText, FILE_ATTRIBUTE_NORMAL)
                  if AfxPathIsRelative(cwzRelative) then wszText = cwzRelative
                  wszText = ProcessToCurdrive(wszText)
                  Print #f, "File="; wszText
                  Print #f, "FileType="; pDoc->ProjectFileType
                  Print #f, "TabIndex="; 0    
                  Print #f, "FileEnd=[-]"
               End If
            end if
         End If   
      End If
   Next

   Close #f
   
   Function = True   ' successul save
End Function


' ========================================================================================
' Load all options for the Project from a diskfile
' ========================================================================================
Function clsConfig.ProjectLoadFromFile() As BOOLEAN    

   Dim st    As WString * MAX_PATH
   Dim wKey  As WString * MAX_PATH
   Dim wData As WString * 1024
   
   Dim sBookmarks As String
   Dim As Long nData, i, idx, nFirstLine, nPosition, nActiveTab

   ' Info saved for each File and used to create the file once FileEnd is found
   Dim pDoc As clsDocument Ptr
   Dim wszFilename As WString * MAX_PATH
   Dim As Long nFileType, nTabIndex
   Dim As BOOLEAN bLoadInTab
   
   If AfxFileExists(gApp.ProjectFilename) = 0 Then Exit Function
   gApp.ProjectName = AfxStrPathname( "NAMEX", gApp.ProjectFilename )
    
   Dim f As Long = Freefile
   If Open( gApp.ProjectFilename For Input Encoding "utf16" As #f ) <> 0 Then Exit Function

   gApp.IsProjectActive = True

   ' Variable length array to hold sequence of TabCtrl tabs to open. We save the sequence
   ' here and open them only after all over documents have been loaded. This will look
   ' visually more appealing.
   Dim nTabs(any) As clsDocument Ptr
   
   ' Prevent the constant updating of the app caption bar with filename, (pDoc->DisplayStats)
   gProjectLoading = true   
   
   ' Prevent Treeview flickering during adding
   dim as hwnd hTV = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   LockWindowUpdate(hTV)

   
   Do Until Eof(f)
      Line Input #f, st
      If Len(st) = 0 Then Continue Do
      If Left(st, 1) = "'" Then Continue Do
      If Left(st, 1) = "[" Then Continue Do
      
      i = Instr(st, "=")
      If i = 0 Then Continue Do
      
      wKey  = Left(st, i-1)
      wData = Mid(st, i+1)
      
      Select Case UCase(wData)
         Case "TRUE":   nData = True
         Case "FALSE":  nData = False
         Case Else:     nData = Val(wData)
      End Select

      wData = ProcessFromCurdrive(wData)

      Select Case wKey
         Case "ProjectType":            gApp.ProjectType           = nData
         Case "ProjectErrorOption":     gApp.ProjectErrorOption    = nData
         Case "ProjectDebug":           gApp.ProjectDebug          = nData
         Case "ProjectThread":          gApp.ProjectThread         = nData
         Case "ProjectShowConsole":     gApp.ProjectShowConsole    = nData
         Case "ProjectOther32":         gApp.ProjectOther32        = wData
         Case "ProjectOther64":         gApp.ProjectOther64        = wData
         Case "DefaultCompiler":        gConfig.DefaultCompiler    = wData
         Case "DefaultCompileMode":     gConfig.DefaultCompileMode = wData
         Case "CommandLine":            gConfig.CommandLine        = wData
         Case "ActiveTab":              nActiveTab = nData
         
         Case "File":                wszFilename = wData
         Case "FileType":            nFileType = nData
         
         Case "TabIndex":            bLoadInTab = nData
         Case "Bookmarks":           sBookmarks = Str(wData)
         Case "FirstLine":           nFirstLine = nData
         Case "Position":            nPosition = nData
         Case "FileEnd":
            ' If this is a relative filename then convert it back.
            if AfxPathIsRelative(wszFilename) then 
               wszFilename = AfxPathCombine( AfxStrPathName("PATH", gApp.ProjectFilename), wszFilename)
            end if
            If AfxFileExists(wszFilename) Then
               pDoc = frmMain_OpenFileSafely(HWND_FRMMAIN, _
                                       False, _    ' bIsNewFile
                                       False, _    ' bIsTemplate
                                       FALSE, _    ' bShowInTab
                                       false, _    ' bIsInclude
                                       @wszFilename, _  ' pwszName
                                       0 )  ' pDocIn
               pDoc->ProjectFileType = nFileType
 
               If bLoadInTab Then
                  If nTabIndex > Ubound(nTabs) Then
                     Redim Preserve nTabs(nTabIndex) As clsDocument Ptr
                  End If
                  nTabs(nTabIndex) = pDoc
                  nTabIndex += 1
               End If
                   
               ' Set the saved position
               SendMessageW( pDoc->hWindow, SCI_SETFIRSTVISIBLELINE, nFirstLine, 0) 
               SendMessageW( pDoc->hWindow, SCI_GOTOPOS, nPosition, 0) 
               nFirstLine = 0: nPosition = 0
               
               ' Set any bookmarks
               pDoc->SetBookmarks(sBookmarks)
               sBookmarks = ""
            
            End If

      End Select
   
   Loop                        
   
   LockWindowUpdate(0)
   
   ' Display all of the tabs
   For i = LBound(nTabs) To UBound(nTabs)
      idx = gTTabCtl.GetTabIndexFromFilename(nTabs(i)->DiskFilename)
      If idx = -1 Then idx = gTTabCtl.AddTab( nTabs(i) )  ' Add the new document to the top tabcontrol
   Next   
   gTTabCtl.SetFocusTab( nActiveTab )
   pDoc = gTTAbCtl.GetActiveDocumentPtr()
   If pDoc Then
      Treeview_SelectItem( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), _
                           pDoc->hNodeExplorer )
   end if
   
   Close #f

   gProjectLoading = false

   Function = True   ' successful open
End Function


