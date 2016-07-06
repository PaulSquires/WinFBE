''  CONFIGURATION MODULE



''
''  CONSTRUCTOR
''
Constructor clsConfig
   _ConfigFilename = Exepath & "\WinFBE.ini" 
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

   Print #f, _
   "'  WINFBE CONFIGURATION" & vbcrlf & vbcrlf & _
   "[Editor]" & vbcrlf & _
   "SyntaxHighlighting="    & this.SyntaxHighlighting    & vbcrlf & _
   "Codetips="              & this.Codetips              & vbcrlf & _
   "LeftMargin="            & this.LeftMargin            & vbcrlf & _
   "FoldMargin="            & this.FoldMargin            & vbcrlf & _
   "AutoIndentation="       & this.AutoIndentation       & vbcrlf & _
   "ConfineCaret="          & this.ConfineCaret          & vbcrlf & _
   "LineNumbering="         & this.LineNumbering         & vbcrlf & _
   "HighlightCurrentLine="  & this.HighlightCurrentLine  & vbcrlf & _
   "IndentGuides="          & this.IndentGuides          & vbcrlf & _
   "TabIndentSpaces="       & this.TabIndentSpaces       & vbcrlf & _
   "HideCompile="           & this.HideCompile           & vbcrlf & _
   "MultipleInstances="     & this.MultipleInstances     & vbcrlf & _
   "CompileAutosave="       & this.CompileAutosave       & vbcrlf & _
   "LocalizationFile="      & this.LocalizationFile      & vbcrlf & _
   "TabSize="               & this.TabSize               & vbcrlf & _
   "KeywordCase="           & this.KeywordCase           & vbcrlf & _
   "EditorFontname="        & this.EditorFontname        & vbcrlf & _
   "EditorFontsize="        & this.EditorFontsize        & vbcrlf & _
   "EditorFontCharSet="     & this.EditorFontCharSet     & vbcrlf & vbcrlf & _
   "[Startup]" & vbcrlf & _
   "StartupLeft="           & this.StartupLeft           & vbcrlf & _
   "StartupTop="            & this.StartupTop            & vbcrlf & _
   "StartupRight="          & this.StartupRight          & vbcrlf & _
   "StartupBottom="         & this.StartupBottom         & vbcrlf & _
   "StartupMaximized="      & this.StartupMaximized      & vbcrlf 
   
   Print #f, "[Compiler]" 
   ' For each folder location determine if it resides on the same drive as
   ' the WinFBE application. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   ProcessToCurdrive(@this.FBWINCompiler32)
   ProcessToCurdrive(@this.FBWINCompiler64)
   ProcessToCurdrive(@this.CompilerHelpfile)
   ProcessToCurdrive(@this.Win32APIHelpfile)

   Print #f, _
   "FBWINCompiler32="       & this.FBWINCompiler32       & vbcrlf & _
   "FBWINCompiler64="       & this.FBWINCompiler64       & vbcrlf & _
   "CompilerSwitches="      & this.CompilerSwitches      & vbcrlf & _
   "CompilerHelpfile="      & this.CompilerHelpfile      & vbcrlf & _
   "Win32APIHelpfile="      & this.Win32APIHelpfile      & vbcrlf & _
   "DefaultCompiler="       & this.DefaultCompiler       & vbcrlf & _
   "DefaultCompileMode="    & this.DefaultCompileMode    & vbcrlf 
   
   Print #f, "[MRU]"  
   For i As Long = 0 To 9
      ProcessToCurdrive(@this.MRU(i))
      Print #f, "MRU_" & Right("00" & Str(i), 2) & "="   & this.MRU(i) 
   Next         
   
   Print #f, ""
   
   Print #f, "[MRUPROJECTS]"  
   For i As Long = 0 To 9
      ProcessToCurdrive(@this.MRUProject(i))
      Print #f, "MRUPROJECT_" & Right("00" & Str(i), 2) & "="   & this.MRUProject(i) 
   Next         

   Print #f, vbcrlf & _
   "[Colors]" & vbcrlf & _
   "clrCaretFG="            & this.clrCaretFG           & vbcrlf & _
   "clrCaretBG="            & this.clrCaretBG           & vbcrlf & _
   "clrCommentsFG="         & this.clrCommentsFG        & vbcrlf & _
   "clrCommentsBG="         & this.clrCommentsBG        & vbcrlf & _
   "clrHighlightedFG="      & this.clrHighlightedFG     & vbcrlf & _
   "clrHighlightedBG="      & this.clrHighlightedBG     & vbcrlf & _
   "clrKeywordFG="          & this.clrKeywordFG         & vbcrlf & _
   "clrKeywordBG="          & this.clrKeywordBG         & vbcrlf & _
   "clrFoldMarginFG="       & this.clrFoldMarginFG      & vbcrlf & _
   "clrFoldMarginBG="       & this.clrFoldMarginBG      & vbcrlf & _
   "clrFoldSymbolFG="       & this.clrFoldSymbolFG      & vbcrlf & _
   "clrFoldSymbolBG="       & this.clrFoldSymbolBG      & vbcrlf & _
   "clrLineNumbersFG="      & this.clrLineNumbersFG     & vbcrlf & _
   "clrLineNumbersBG="      & this.clrLineNumbersBG     & vbcrlf & _
   "clrBookmarksFG="        & this.clrBookmarksFG       & vbcrlf & _
   "clrBookmarksBG="        & this.clrBookmarksBG       & vbcrlf & _
   "clrOperatorsFG="        & this.clrOperatorsFG       & vbcrlf & _
   "clrOperatorsBG="        & this.clrOperatorsBG       & vbcrlf & _
   "clrIndentGuidesFG="     & this.clrIndentGuidesFG    & vbcrlf & _
   "clrIndentGuidesBG="     & this.clrIndentGuidesBG    & vbcrlf & _
   "clrPreprocessorFG="     & this.clrPreprocessorFG    & vbcrlf & _
   "clrPreprocessorBG="     & this.clrPreprocessorBG    & vbcrlf & _
   "clrSelectionFG="        & this.clrSelectionFG       & vbcrlf & _
   "clrSelectionBG="        & this.clrSelectionBG       & vbcrlf & _
   "clrStringsFG="          & this.clrStringsFG         & vbcrlf & _
   "clrStringsBG="          & this.clrStringsBG         & vbcrlf & _
   "clrTextFG="             & this.clrTextFG            & vbcrlf & _
   "clrTextBG="             & this.clrTextBG            & vbcrlf & _
   "clrWinAPIFG="           & this.clrWinAPIFG          & vbcrlf & _
   "clrWinAPIBG="           & this.clrWinAPIBG          & vbcrlf & _
   "clrWindowFG="           & this.clrWindowFG          & vbcrlf & _
   "clrWindowBG="           & this.clrWindowBG          & vbcrlf 
   
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
      
      ProcessFromCurdrive(@wData)

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
         Case "SyntaxHighlighting":   this.SyntaxHighlighting   = nData
         Case "Codetips":             this.Codetips             = nData
         Case "LeftMargin":           this.LeftMargin           = nData
         Case "FoldMargin":           this.FoldMargin           = nData
         Case "AutoIndentation":      this.AutoIndentation      = nData
         Case "ConfineCaret":         this.ConfineCaret         = nData
         Case "LineNumbering":        this.LineNumbering        = nData
         Case "HighlightCurrentLine": this.HighlightCurrentLine = nData
         Case "IndentGuides":         this.IndentGuides         = nData
         Case "TabIndentSpaces":      this.TabIndentSpaces      = nData
         Case "HideCompile":          this.HideCompile          = nData
         Case "MultipleInstances":    this.MultipleInstances    = nData
         Case "CompileAutosave":      this.CompileAutosave      = nData
         Case "LocalizationFile":     this.LocalizationFile     = wData
         Case "TabSize":              this.TabSize              = wData
         Case "EditorFontname":       this.EditorFontname       = wData
         Case "EditorFontsize":       this.EditorFontsize       = wData
         Case "EditorFontCharSet":    this.EditorFontCharSet    = wData
         Case "KeywordCase":          this.KeywordCase          = nData
         Case "StartupLeft":          this.StartupLeft          = nData
         Case "StartupTop":           this.StartupTop           = nData
         Case "StartupRight":         this.StartupRight         = nData
         Case "StartupBottom":        this.StartupBottom        = nData
         Case "StartupMaximized":     this.StartupMaximized     = nData
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





