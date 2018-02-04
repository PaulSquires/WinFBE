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


''  CONFIGURATION MODULE



''
''  CONSTRUCTOR
''
Constructor clsConfig
   _ConfigFilename = AfxGetExePathName & "Settings\WinFBE.ini" 
   this.LoadKeywords
   ' Set default path to WinFBX should it not yet be specified
   if len(this.AfxPath) = 0 THEN
      this.AfxPath = AfxGetExePathName & "src"
   END IF
   
End Constructor

''
''  DESTRUCTOR
''
DESTRUCTOR clsConfig
'
end DESTRUCTOR


''
''
Function clsConfig.LoadKeywords() as Long
   Dim st as String 
   Dim f  as Long = Freefile
   
   If Open( AfxGetExePathName & "Settings\freebasic_keywords.txt" For Input As #f) = 0 Then
      this.FBKeywords  = ""
      Do Until EOF(f)
         Line Input #f, st
         If Len(st) Then this.FBKeywords = this.FBKeywords & st & " "
      Loop    
      this.FBKeywords = Trim(this.FBKeywords)
      Close #f
   End If
   
   Function = 0
End Function

''
''
Function clsConfig.SaveKeywords() As Long
   Dim f As Long = Freefile
   
   If Open( AfxGetExePathName & "Settings\freebasic_keywords.txt" For Output As #f) = 0 Then
      Print #f, this.FBKeywords
      Close #f
   End If
   
   Function = 0
End Function

  
''
''
Function clsConfig.LoadCodetips( ByRef sFilename As String ) as boolean
   Dim as string st, sFuncName, sFuncDeclare
   Dim as long f, i
   
   dim parser as clsParser

   parser.fileName = ""     ' no need to store this sFilename
   parser.idxProject = -1   ' do not associate it with any project
   
   f = Freefile
   If Open( sFilename For Input As #f ) = 0 Then
      Do Until Eof(f)
         Line Input #f, st
         
         st = Trim(st)
         If Len(st) = 0 Then Continue Do
         
         i = Instr(st, "=")
         If i Then
            parser.funcName   = Left(st, i-1)
            parser.funcParams = Mid(st, i+1)
            gdb2.dbAdd(parser, DB2_FUNCTION)
         End If   
      Loop
      Close #f
      return false
   End If

   Function = true   ' error
End Function


''
''
Function clsConfig.LoadCodetipsWinAPI( ByRef sFilename As String ) as boolean
   Dim as string st, sID, sFuncName, sFuncDeclare
   dim as string sTypeName, sTypeElements, sParts, sPart
   Dim as long f, i, nCount
   
   dim parser as clsParser

   parser.fileName = ""     ' no need to store this sFilename
   parser.idxProject = -1   ' do not associate it with any project

   ''
   ''  Codetips definition file for WinAPI
   ''
   ''  Format:
   ''  1st character:  F=FUNCTION, T=TYPE, S=STANDARD_DATATYPE
   ''  2nd character:  colon.
   ''  TYPE elements are separated by | pipe symbol. Each element name/datatype is separated by comma.
   ''  STANDARD datatypes are separated by | pipe symobol (datatype followed by description of the datatype.
   ''

   f = Freefile
   If Open( sFilename For Input As #f ) = 0 Then
      Do Until Eof(f)
         Line Input #f, st
         
         st = Trim(st)
         If Len(st) = 0 Then Continue Do
         
         i = Instr(st, ":")
         If i = 0 Then continue do
         
         sID = left(st, i - 1)
         st = mid(st, i + 1)
                   
         select CASE sID
            case "S"    ' standard datatype
               i = instr(st, "|")
               parser.typeName = Left(st, i-1)
               parser.Description = mid(st, i+1)
               parser.IsWinApi = true
               gdb2.dbAdd(parser, DB2_STANDARDDATATYPE)  
               
            CASE "F"    ' function
               i = instr(st, "(")
               parser.funcName   = Left(st, i-1)
               parser.funcParams = st
               parser.IsWinApi = true
               gdb2.dbAdd(parser, DB2_FUNCTION)

            case "T"     ' type
               i = instr(st, "|")
               parser.typeName = Left(st, i-1)
               parser.typeAlias = parser.typeName
               parser.IsWinApi = true
               sParts = mid(st, i + 1)
               gdb2.dbAdd(parser, DB2_TYPE)

               ' Each element of the TYPE structure is added to the database. The
               ' element is in the format elementname,elementtype
               parser.funcName = ""
               nCount = AfxStrParseCount(sParts, "|")
               for i = 1 to nCount
                  sPart = AfxStrParse(sParts, i, "|")
                  parser.varName = AfxStrParse(sPart, 1, ",") 
                  parser.varType = AfxStrParse(sPart, 2, ",") 
                  parser.bInTypePublic = true
                  gdb2.dbAdd(parser, DB2_VARIABLE)
               NEXT

         END SELECT
  
      Loop
      Close #f
      return false
   End If

   Function = true   ' error
End Function
      

''
''  INITIALIZE THE TOOLBOX CONTROLS
''
Function clsConfig.InitializeToolBox() as Long
   
   redim gToolBox(25) as TOOLBOX_TYPE
   
   gToolBox(0).nToolType = CTRL_POINTER
   gToolBox(0).wszName   = "Pointer"
   gToolBox(0).wszImage  = "IMAGE_POINTER"
   gToolBox(0).wszCursor = ""
   gToolBox(0).wszClassName = ""
   
   gToolBox(1).nToolType = CTRL_LABEL
   gToolBox(1).wszName   = "Label"
   gToolBox(1).wszImage  = "IMAGE_LABEL"
   gToolBox(1).wszCursor = "IMAGE_CURSOR_LABEL"
   gToolBox(1).wszClassName = "STATIC"

   gToolBox(2).nToolType = CTRL_BUTTON
   gToolBox(2).wszName   = "Button"
   gToolBox(2).wszImage  = "IMAGE_BUTTON"
   gToolBox(2).wszCursor = "IMAGE_CURSOR_BUTTON"
   gToolBox(2).wszClassName = "BUTTON"

   gToolBox(3).nToolType = CTRL_TEXTBOX
   gToolBox(3).wszName   = "TextBox"
   gToolBox(3).wszImage  = "IMAGE_TEXTBOX"
   gToolBox(3).wszCursor = "IMAGE_CURSOR_TEXTBOX"
   gToolBox(3).wszClassName = "TEXTBOX"

   gToolBox(4).nToolType = CTRL_CHECKBOX
   gToolBox(4).wszName   = "CheckBox"
   gToolBox(4).wszImage  = "IMAGE_CHECKBOX"
   gToolBox(4).wszCursor = "IMAGE_CURSOR_CHECKBOX"
   gToolBox(4).wszClassName = "CHECKBOX"

   gToolBox(5).nToolType = CTRL_OPTION
   gToolBox(5).wszName   = "Option"
   gToolBox(5).wszImage  = "IMAGE_OPTION"
   gToolBox(5).wszCursor = "IMAGE_CURSOR_OPTION"
   gToolBox(5).wszClassName = "RADIOBUTTON"

   gToolBox(6).nToolType = CTRL_FRAME
   gToolBox(6).wszName   = "Frame"
   gToolBox(6).wszImage  = "IMAGE_FRAME"
   gToolBox(6).wszCursor = "IMAGE_CURSOR_FRAME"
   gToolBox(6).wszClassName = "FRAME"

   gToolBox(7).nToolType = CTRL_PICTURE
   gToolBox(7).wszName   = "Picture"
   gToolBox(7).wszImage  = "IMAGE_PICTURE"
   gToolBox(7).wszCursor = "IMAGE_CURSOR_PICTURE"
   gToolBox(7).wszClassName = "STATIC"

   gToolBox(8).nToolType = CTRL_COMBOBOX
   gToolBox(8).wszName   = "ComboBox"
   gToolBox(8).wszImage  = "IMAGE_COMBOBOX"
   gToolBox(8).wszCursor = "IMAGE_CURSOR_COMBOBOX"
   gToolBox(8).wszClassName = "COMBOBOX"

   gToolBox(9).nToolType = CTRL_LISTBOX
   gToolBox(9).wszName   = "ListBox"
   gToolBox(9).wszImage  = "IMAGE_LISTBOX"
   gToolBox(9).wszCursor = "IMAGE_CURSOR_LISTBOX"
   gToolBox(9).wszClassName = "LISTBOX"

   gToolBox(10).nToolType = CTRL_IMAGEBUTTON
   gToolBox(10).wszName   = "ImageButton"
   gToolBox(10).wszImage  = "IMAGE_IMGBUTTON"
   gToolBox(10).wszCursor = "IMAGE_CURSOR_IMGBUTTON"
   gToolBox(10).wszClassName = "BUTTON"

   gToolBox(11).nToolType = CTRL_HSCROLL
   gToolBox(11).wszName   = "HScroll"
   gToolBox(11).wszImage  = "IMAGE_HSCROLL"
   gToolBox(11).wszCursor = "IMAGE_CURSOR_HSCROLL"
   gToolBox(11).wszClassName = "HSCROLLBAR"

   gToolBox(12).nToolType = CTRL_VSCROLL
   gToolBox(12).wszName   = "VScroll"
   gToolBox(12).wszImage  = "IMAGE_VSCROLL"
   gToolBox(12).wszCursor = "IMAGE_CURSOR_VSCROLL"
   gToolBox(12).wszClassName = "VSCROLLBAR"

   gToolBox(13).nToolType = CTRL_TIMER
   gToolBox(13).wszName   = "Timer"
   gToolBox(13).wszImage  = "IMAGE_TIMER"
   gToolBox(13).wszCursor = "IMAGE_CURSOR_TIMER"
   gToolBox(13).wszClassName = ""

   gToolBox(14).nToolType = CTRL_TABCONTROL
   gToolBox(14).wszName   = "TabControl"
   gToolBox(14).wszImage  = "IMAGE_TABCONTROL"
   gToolBox(14).wszCursor = "IMAGE_CURSOR_TABCONTROL"
   gToolBox(14).wszClassName = "TABCONTROL"

   gToolBox(15).nToolType = CTRL_RICHEDIT
   gToolBox(15).wszName   = "RichEdit"
   gToolBox(15).wszImage  = "IMAGE_RICHEDIT"
   gToolBox(15).wszCursor = "IMAGE_CURSOR_RICHEDIT"
   gToolBox(15).wszClassName = "RICHEDIT"

   gToolBox(16).nToolType = CTRL_PROGRESSBAR
   gToolBox(16).wszName   = "ProgressBar"
   gToolBox(16).wszImage  = "IMAGE_PROGRESSBAR"
   gToolBox(16).wszCursor = "IMAGE_CURSOR_PROGRESSBAR"
   gToolBox(16).wszClassName = "PROGRESSBAR"

   gToolBox(17).nToolType = CTRL_UPDOWN
   gToolBox(17).wszName   = "UpDown"
   gToolBox(17).wszImage  = "IMAGE_UPDOWN"
   gToolBox(17).wszCursor = "IMAGE_CURSOR_UPDOWN"
   gToolBox(17).wszClassName = "UPDOWN"

   gToolBox(18).nToolType = CTRL_LISTVIEW
   gToolBox(18).wszName   = "ListView"
   gToolBox(18).wszImage  = "IMAGE_LISTVIEW"
   gToolBox(18).wszCursor = "IMAGE_CURSOR_LISTVIEW"
   gToolBox(18).wszClassName = "LISTVIEW"

   gToolBox(19).nToolType = CTRL_TREEVIEW
   gToolBox(19).wszName   = "TreeView"
   gToolBox(19).wszImage  = "IMAGE_TREEVIEW"
   gToolBox(19).wszCursor = "IMAGE_CURSOR_TREEVIEW"
   gToolBox(19).wszClassName = "TREEVIEW"

   gToolBox(20).nToolType = CTRL_SLIDER
   gToolBox(20).wszName   = "Slider"
   gToolBox(20).wszImage  = "IMAGE_SLIDER"
   gToolBox(20).wszCursor = "IMAGE_CURSOR_SLIDER"
   gToolBox(20).wszClassName = "SLIDER"

   gToolBox(21).nToolType = CTRL_DATETIMEPICKER
   gToolBox(21).wszName   = "DateTimePicker"
   gToolBox(21).wszImage  = "IMAGE_DATETIMEPICKER"
   gToolBox(21).wszCursor = "IMAGE_CURSOR_DATETIMEPICKER"
   gToolBox(21).wszClassName = "DATETIMEPICKER"

   gToolBox(22).nToolType = CTRL_MONTHCALENDAR
   gToolBox(22).wszName   = "MonthCalendar"
   gToolBox(22).wszImage  = "IMAGE_MONTHCALENDAR"
   gToolBox(22).wszCursor = "IMAGE_CURSOR_MONTHCALENDAR"
   gToolBox(22).wszClassName = "MONTHCALENDAR"

   gToolBox(23).nToolType = CTRL_WEBBROWSER
   gToolBox(23).wszName   = "WebBrowser"
   gToolBox(23).wszImage  = "IMAGE_WEBBROWSER"
   gToolBox(23).wszCursor = "IMAGE_CURSOR_WEBBROWSER"
   gToolBox(23).wszClassName = ""

   gToolBox(24).nToolType = CTRL_CUSTOM
   gToolBox(24).wszName   = "Custom"
   gToolBox(24).wszImage  = "IMAGE_CUSTOM"
   gToolBox(24).wszCursor = "IMAGE_CURSOR_CUSTOM"
   gToolBox(24).wszClassName = ""

   gToolBox(25).nToolType = CTRL_OCX
   gToolBox(25).wszName   = "OCX"
   gToolBox(25).wszImage  = "IMAGE_OCX"
   gToolBox(25).wszCursor = "IMAGE_CURSOR_OCX"
   gToolBox(25).wszClassName = ""

   function = 0
end function


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
   print #f, "AskExit="               ; this.AskExit
   print #f, "HideToolbar="           ; this.HideToolbar
   print #f, "HideStatusbar="         ; this.HideStatusbar
   print #f, "MultipleInstances="     ; this.MultipleInstances    
   print #f, "CompileAutosave="       ; this.CompileAutosave      
   print #f, "CloseFuncList="         ; this.CloseFuncList        
   print #f, "HideCompile="           ; this.HideCompile          
   print #f, "SyntaxHighlighting="    ; this.SyntaxHighlighting   
   print #f, "Codetips="              ; this.Codetips             
   print #f, "AutoComplete="          ; this.AutoComplete
   print #f, "RightEdge="             ; this.RightEdge
   print #f, "RightEdgePosition="     ; this.RightEdgePosition           
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
   this.AfxPath          = ProcessToCurdrive(this.AfxPath)
   this.CompilerHelpfile = ProcessToCurdrive(this.CompilerHelpfile)
   this.Win32APIHelpfile = ProcessToCurdrive(this.Win32APIHelpfile)

   Print #f, "[Compiler]" 
   print #f, "FBWINCompiler32="       ; this.FBWINCompiler32      
   print #f, "FBWINCompiler64="       ; this.FBWINCompiler64      
   print #f, "WinFBXPath="            ; this.AfxPath     
   print #f, "CompilerSwitches="      ; this.CompilerSwitches     
   print #f, "CompilerHelpfile="      ; this.CompilerHelpfile     
   print #f, "Win32APIHelpfile="      ; this.Win32APIHelpfile     
   print #f, "RunViaCommandWindow="   ; this.RunViaCommandWindow
   

   print #f, ""
   Print #f, "[UserTools]"  
   for i as long = lbound(this.Tools) to ubound(this.Tools)
      print #f, "USERTOOL_"; Right("00" & Str(i), 2); "=";
      print #f, this.Tools(i).wszDescription; "|-|";
      print #f, this.Tools(i).wszCommand; "|-|"; 
      print #f, this.Tools(i).wszParameters; "|-|"; 
      print #f, this.Tools(i).wszKey; "|-|"; 
      print #f, this.Tools(i).wszWorkingFolder; "|-|"; 
      print #f, this.Tools(i).IsCtrl; "|-|";  
      print #f, this.Tools(i).IsAlt; "|-|";  
      print #f, this.Tools(i).IsShift; "|-|";  
      print #f, this.Tools(i).IsPromptRun; "|-|";  
      print #f, this.Tools(i).IsMinimized; "|-|";  
      print #f, this.Tools(i).IsWaitFinish; "|-|";  
      print #f, this.Tools(i).IsDisplayMenu; "|-|";    
      print #f, this.Tools(i).Action   
   NEXT

   print #f, ""
   Print #f, "[Builds]"  
   for i as long = lbound(this.Builds) to ubound(this.Builds)
      print #f, "BUILD_"; Right("00" & Str(i), 2); "=";
      print #f, this.Builds(i).id; "|-|";
      print #f, this.Builds(i).wszDescription; "|-|";
      print #f, this.Builds(i).wszOptions; "|-|"; 
      print #f, this.Builds(i).IsDefault; "|-|";  
      print #f, this.Builds(i).Is32bit; "|-|";  
      print #f, this.Builds(i).Is64bit  
   NEXT

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
   print #f, "[Themes]" 
   print #f, "SelectedTheme="; this.SelectedTheme
   for i as long = lbound(this.Themes) to ubound(this.Themes)
      print #f, "THEME_"; Right("00" & Str(i), 2); "=";
      print #f, this.Themes(i).id; ",";
      print #f, AfxStrReplaceAny(this.Themes(i).wszDescription, ",|", "_"); 
      for ii as long = CLR_CARET to CLR_WINDOW
         print #f, "|"; this.Themes(i).colors(ii).nFg; ","; _
                        this.Themes(i).colors(ii).nBg; ","; _
                        this.Themes(i).colors(ii).bUseDefaultBg; ","; _
                        this.Themes(i).colors(ii).bFontBold; ","; _
                        this.Themes(i).colors(ii).bFontItalic; ","; _
                        this.Themes(i).colors(ii).bFontUnderline;
      NEXT
      print #f, ""
   NEXT

   dim hCtl as hwnd = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES)
   dim wszText as CWSTR = wstr("NOTES-START") + vbcrlf + _
                          gApp.NonProjectNotes + vbcrlf + _
                          wstr("NOTES-END") + vbcrlf
   print #f, ""
   print #f, "[Notes]"
   print #f, wszText

   Close #f
   
   Function = 0
End Function


''
''  IMPORT THEME CONTAINED IN INCOMING STRING (THEME_00=)
''
Function clsConfig.ImportTheme( byref st as wstring, _
                                byval bImportExternal as Boolean = false _
                                ) as Long

   Dim wKey   As WString * 1024
   Dim wData  As WString * 1024
   Dim wData2 As WString * 1024
   dim as long i
   dim theme as TYPE_THEMES
   
   i = InStr(st, "=")
   If i = 0 Then exit function
   
   wKey  = Left(st, i-1)
   wData = Mid(st, i+1)

   If Left(wKey, 6) = "THEME_" Then  ' is this a THEME entry
      ' Theme data is saved in parse blocks separated by a "|" pipe symbol. Within
      ' each of of those parsed pipe blocks the data is further divided and separated
      ' by "," comma blocks. The first pipe block is always the theme id and description.
      ' From pipe block 2 to parsecount, the pipe blocks are the actual colors. Each
      ' color block is comma separated and composed of foreground color, background color,
      ' and 0/1 indicating if to use default background color.
      dim as long nNumPipes = AfxStrParseCount(wData, "|")
      if nNumPipes > ubound(theme.colors) THEN
         nNumPipes = ubound(theme.colors)
      END IF
      for xx as long = 1 to nNumPipes
         wData2 = AfxStrParse(wData, xx, "|")
         if xx = 1 THEN
            ' We set the id later depending on whether existing or import
            theme.id = AfxStrParse(wData2, 1, ",")
            theme.wszDescription = AfxStrParse(wData2, 2, ",")
         else
            theme.colors(xx).nFg = val(**AfxStrParse(wData2, 1, ","))
            theme.colors(xx).nBg = val(**AfxStrParse(wData2, 2, ","))   
            theme.colors(xx).bUseDefaultBg = val(**AfxStrParse(wData2, 3, ","))   
            theme.colors(xx).bFontBold = val(**AfxStrParse(wData2, 4, ","))   
            theme.colors(xx).bFontItalic = val(**AfxStrParse(wData2, 5, ","))   
            theme.colors(xx).bFontUnderline = val(**AfxStrParse(wData2, 6, ","))   
         END IF
      NEXT
   
      if bImportExternal THEN
         theme.id = AfxGuidText(AfxGuid())
         i = ubound(this.ThemesTemp) + 1
         redim preserve this.ThemesTemp(i)
         this.ThemesTemp(i) = theme
      else
         i = ubound(this.Themes) + 1
         redim preserve this.Themes(i)
         this.Themes(i) = theme
      end if   
   end if
   
   function = 0
end function


''
''  LOAD CONFIGURATION FROM DISK FILE
''
Function clsConfig.LoadFromFile() as Long

   Dim st     As WString * 1024
   Dim wKey   As WString * 1024
   Dim wData  As WString * 1024
   Dim wData2 As WString * 1024
   Dim nData  As Long  
   Dim i      as Long
   dim bReadingNote as Boolean

   If AfxFileExists(_ConfigFilename) = 0 Then Exit Function
    
   Dim f as Long = Freefile
   If Open( _ConfigFilename For Input Encoding "utf16" As #f ) <> 0 Then Exit Function

   gApp.NonProjectNotes = ""
   
   Do Until EOF(f)
      Line Input #f, st
      If Len(st) = 0 Then Continue Do
      If Left(st, 1) = "'" Then Continue Do
      If Left(st, 1) = "[" Then Continue Do
      
      If Left(st, 11) = "NOTES-START" Then  
         bReadingNote = true
         continue do
      end if   
      If Left(st, 9) = "NOTES-END" Then  
         bReadingNote = false
         continue do
      end if
      if bReadingNote THEN
         gApp.NonProjectNotes = gApp.NonProjectNotes + st + vbcrlf
         continue do
      END IF

      i = InStr(st, "=")
      If i = 0 Then Continue Do
      
      wKey  = Left(st, i-1)
      wData = Mid(st, i+1)
      nData = Val(wData)
      
      wData = ProcessFromCurdrive(wData)


      If Left(wKey, 9) = "USERTOOL_" Then  ' is this a User Tool entry
         i = Val(Right(wKey,2))
         If i > ubound(this.Tools) Then
            redim preserve this.Tools(i) 
            this.Tools(i).wszDescription   = AfxStrParse(wData, 1, "|-|")
            this.Tools(i).wszCommand       = AfxStrParse(wData, 2, "|-|")
            this.Tools(i).wszParameters    = AfxStrParse(wData, 3, "|-|")
            this.Tools(i).wszKey           = AfxStrParse(wData, 4, "|-|")
            this.Tools(i).wszWorkingFolder = AfxStrParse(wData, 5, "|-|")
            this.Tools(i).IsCtrl           = val(**AfxStrParse(wData, 6, "|-|"))
            this.Tools(i).IsAlt            = val(**AfxStrParse(wData, 7, "|-|"))
            this.Tools(i).IsShift          = val(**AfxStrParse(wData, 8, "|-|"))
            this.Tools(i).IsPromptRun      = val(**AfxStrParse(wData, 9, "|-|"))
            this.Tools(i).IsMinimized      = val(**AfxStrParse(wData, 10, "|-|"))
            this.Tools(i).IsWaitFinish     = val(**AfxStrParse(wData, 11, "|-|"))
            this.Tools(i).IsDisplayMenu    = val(**AfxStrParse(wData, 12, "|-|"))
            this.Tools(i).Action           = val(**AfxStrParse(wData, 13, "|-|"))
            Continue Do
         End If
      End If

      If Left(wKey, 6) = "BUILD_" Then  ' is this a BUILD entry
         i = Val(Right(wKey,2))
         If i > ubound(this.Builds) Then
            redim preserve this.Builds(i) 
            this.Builds(i).id             = AfxStrParse(wData, 1, "|-|")
            this.Builds(i).wszDescription = AfxStrParse(wData, 2, "|-|")
            this.Builds(i).wszOptions     = AfxStrParse(wData, 3, "|-|")
            this.Builds(i).IsDefault      = val(**AfxStrParse(wData, 4, "|-|"))
            this.Builds(i).Is32bit        = val(**AfxStrParse(wData, 5, "|-|"))
            this.Builds(i).Is64bit        = val(**AfxStrParse(wData, 6, "|-|"))
            Continue Do
         End If
      End If

      If Left(wKey, 6) = "THEME_" Then  ' is this a THEME entry
         this.ImportTheme(st, false)
         continue do
      End If

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
         Case "SelectedTheme":        this.SelectedTheme        = wData
         Case "AskExit":              this.AskExit              = nData
         Case "HideToolbar":          this.HideToolbar          = nData
         Case "HideStatusbar":        this.HideStatusbar        = nData
         Case "MultipleInstances":    this.MultipleInstances    = nData
         Case "CompileAutosave":      this.CompileAutosave      = nData
         Case "CloseFuncList":        this.CloseFuncList        = nData
         Case "HideCompile":          this.HideCompile          = nData
         Case "SyntaxHighlighting":   this.SyntaxHighlighting   = nData
         Case "Codetips":             this.Codetips             = nData
         Case "AutoComplete":         this.AutoComplete         = nData
         Case "RightEdge":            this.RightEdge            = nData
         Case "RightEdgePosition":    this.RightEdgePosition    = wData
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
         Case "WinFBXPath":           this.AfxPath              = wData
         Case "CompilerSwitches":     this.CompilerSwitches     = wData
         Case "CompilerHelpfile":     this.CompilerHelpfile     = wData
         Case "Win32APIHelpfile":     this.Win32APIHelpfile     = wData
         case "RunViaCommandWindow":  this.RunViaCommandWindow  = nData
      End Select
   
   Loop
   Close #f
   
   ' If no builds exist then create some default ones...
   if ubound(this.Builds) = -1 THEN
      redim this.Builds(11) 
      this.Builds(0).id = AfxGuidText(AfxGuid())
      this.Builds(0).wszDescription = "Win32 GUI (Release)"
      this.Builds(0).wszOptions = "-s gui"
      this.Builds(0).Is32bit = 1
      this.Builds(0).Is64bit = 0
      
      this.Builds(1).id = AfxGuidText(AfxGuid())
      this.Builds(1).wszDescription = "Win32 GUI (Debug)"
      this.Builds(1).wszOptions = "-g -exx -s gui"
      this.Builds(1).Is32bit = 1
      this.Builds(1).Is64bit = 0
      
      this.Builds(2).id = AfxGuidText(AfxGuid())
      this.Builds(2).wszDescription = "Win32 Console (Release)"
      this.Builds(2).wszOptions = "-s console"
      this.Builds(2).IsDefault = 1
      this.Builds(2).Is32bit = 1
      this.Builds(2).Is64bit = 0
      
      this.Builds(3).id = AfxGuidText(AfxGuid())
      this.Builds(3).wszDescription = "Win32 Console (Debug)"
      this.Builds(3).wszOptions = "-g -exx -s console"
      this.Builds(3).Is32bit = 1
      this.Builds(3).Is64bit = 0
      
      this.Builds(4).id = AfxGuidText(AfxGuid())
      this.Builds(4).wszDescription = "Win32 Windows DLL"
      this.Builds(4).wszOptions = "-s gui -dll -export"
      this.Builds(4).Is32bit = 1
      this.Builds(4).Is64bit = 0

      this.Builds(5).id = AfxGuidText(AfxGuid())
      this.Builds(5).wszDescription = "Win32 Static Library"
      this.Builds(5).wszOptions = "-lib"
      this.Builds(5).Is32bit = 1
      this.Builds(5).Is64bit = 0

      this.Builds(6).id = AfxGuidText(AfxGuid())
      this.Builds(6).wszDescription = "Win64 GUI (Release)"
      this.Builds(6).wszOptions = "-s gui"
      this.Builds(6).Is32bit = 0
      this.Builds(6).Is64bit = 1
      
      this.Builds(7).id = AfxGuidText(AfxGuid())
      this.Builds(7).wszDescription = "Win64 GUI (Debug)"
      this.Builds(7).wszOptions = "-g -exx -s gui"
      this.Builds(7).Is32bit = 0
      this.Builds(7).Is64bit = 1
      
      this.Builds(8).id = AfxGuidText(AfxGuid())
      this.Builds(8).wszDescription = "Win64 Console (Release)"
      this.Builds(8).wszOptions = "-s console"
      this.Builds(8).Is32bit = 0
      this.Builds(8).Is64bit = 1
      
      this.Builds(9).id = AfxGuidText(AfxGuid())
      this.Builds(9).wszDescription = "Win64 Console (Debug)"
      this.Builds(9).wszOptions = "-g -exx -s console"
      this.Builds(9).Is32bit = 0
      this.Builds(9).Is64bit = 1

      this.Builds(10).id = AfxGuidText(AfxGuid())
      this.Builds(10).wszDescription = "Win64 Windows DLL"
      this.Builds(10).wszOptions = "-s gui -dll -export"
      this.Builds(10).Is32bit = 0
      this.Builds(10).Is64bit = 1

      this.Builds(11).id = AfxGuidText(AfxGuid())
      this.Builds(11).wszDescription = "Win64 Static Library"
      this.Builds(11).wszOptions = "-lib"
      this.Builds(11).Is32bit = 0
      this.Builds(11).Is64bit = 1

   END IF

   ' Create a default Theme if none exist
   if ubound(this.Themes) = -1 THEN
      redim this.Themes(1) 
   
      with this.Themes(0)
         .id = AfxGuidText(AfxGuid())
         .wszDescription = "Default Theme"
         .colors(CLR_CARET).nFg        = BGR(0,0,0)          ' black          
         .colors(CLR_CARET).nBg        = -1
         .colors(CLR_COMMENTS).nFg     = BGR(0,128,0)        ' green          
         .colors(CLR_COMMENTS).nBg     = BGR(255,255,255)    ' white
         .colors(CLR_HIGHLIGHTED).nFg  = BGR(255,255,0)      ' yellow
         .colors(CLR_HIGHLIGHTED).nBg  = -1
         .colors(CLR_KEYWORD).nFg      = BGR(0,0,255)        ' blue
         .colors(CLR_KEYWORD).nBg      = BGR(255,255,255)    ' white
         .colors(CLR_FOLDMARGIN).nFg   = BGR(237,236,235)    ' pale gray
         .colors(CLR_FOLDMARGIN).nBg   = -1
         .colors(CLR_FOLDSYMBOL).nFg   = BGR(255,255,255)    ' white
         .colors(CLR_FOLDSYMBOL).nBg   = BGR(0,0,0)          ' black
         .colors(CLR_LINENUMBERS).nFg  = BGR(0,0,0)          ' black
         .colors(CLR_LINENUMBERS).nBg  = BGR(204,204,204)    ' light gray
         .colors(CLR_BOOKMARKS).nFg    = BGR(0,0,0)          ' black
         .colors(CLR_BOOKMARKS).nBg    = BGR(0,0,255)        ' blue
         .colors(CLR_OPERATORS).nFg    = BGR(196,0,0)        ' red
         .colors(CLR_OPERATORS).nBg    = BGR(255,255,255)    ' white
         .colors(CLR_INDENTGUIDES).nFg = BGR(255,255,255)    ' white
         .colors(CLR_INDENTGUIDES).nBg = BGR(0,0,0)          ' black
         .colors(CLR_PREPROCESSOR).nFg = BGR(196,0,0)        ' red
         .colors(CLR_PREPROCESSOR).nBg = BGR(255,255,255)    ' white
         .colors(CLR_SELECTION).nFg    = BGR(255,255,255)    ' white
         .colors(CLR_SELECTION).nBg    = BGR(038,079,120)    ' light blue
         .colors(CLR_STRINGS).nFg      = BGR(173,0,173)      ' Purple (little darker than Magenta)
         .colors(CLR_STRINGS).nBg      = BGR(255,255,255)    ' white
         .colors(CLR_TEXT).nFg         = BGR(0,0,0)          ' black
         .colors(CLR_TEXT).nBg         = BGR(255,255,255)    ' white
         .colors(CLR_WINAPI).nFg       = BGR(0,0,255)        ' blue
         .colors(CLR_WINAPI).nBg       = BGR(255,255,255)    ' white
         .colors(CLR_WINDOW).nFg       = BGR(255,255,255)    ' white
         .colors(CLR_WINDOW).nBg       = -1
      end with
      this.SelectedTheme = this.Themes(0).id

      with this.Themes(1)
         .id = AfxGuidText(AfxGuid())
         .wszDescription = "Visual Studio Code (Dark)"
         .colors(CLR_CARET).nFg        = BGR(204,204,204)    ' light gray          
         .colors(CLR_CARET).nBg        = -1
         .colors(CLR_COMMENTS).nFg     = BGR(095,125,049)    ' moss green          
         .colors(CLR_COMMENTS).nBg     = BGR(030,030,030)    ' light black
         .colors(CLR_HIGHLIGHTED).nFg  = BGR(050,050,050)    ' charcoal black
         .colors(CLR_HIGHLIGHTED).nBg  = -1
         .colors(CLR_KEYWORD).nFg      = BGR(148,202,240)    ' powder blue
         .colors(CLR_KEYWORD).nBg      = BGR(030,030,030)    ' light black
         .colors(CLR_FOLDMARGIN).nFg   = BGR(030,030,030)    ' light black
         .colors(CLR_FOLDMARGIN).nBg   = -1
         .colors(CLR_FOLDSYMBOL).nFg   = BGR(030,030,030)    ' light black
         .colors(CLR_FOLDSYMBOL).nBg   = BGR(090,090,090)    ' evening black
         .colors(CLR_LINENUMBERS).nFg  = BGR(090,090,090)    ' evening black
         .colors(CLR_LINENUMBERS).nBg  = BGR(030,030,030)    ' light black
         .colors(CLR_BOOKMARKS).nFg    = BGR(090,090,090)    ' evening black
         .colors(CLR_BOOKMARKS).nBg    = BGR(030,030,030)    ' light black
         .colors(CLR_OPERATORS).nFg    = BGR(204,204,204)    ' light gray
         .colors(CLR_OPERATORS).nBg    = BGR(030,030,030)    ' light black
         .colors(CLR_INDENTGUIDES).nFg = BGR(030,030,030)    ' light black
         .colors(CLR_INDENTGUIDES).nBg = BGR(204,204,204)    ' light gray
         .colors(CLR_PREPROCESSOR).nFg = BGR(165,134,181)    ' light purple
         .colors(CLR_PREPROCESSOR).nBg = BGR(030,030,030)    ' light black
         .colors(CLR_SELECTION).nFg    = BGR(204,204,204)    ' light gray
         .colors(CLR_SELECTION).nBg    = BGR(038,079,120)    ' light blue
         .colors(CLR_STRINGS).nFg      = BGR(195,137,080)    ' orange brown
         .colors(CLR_STRINGS).nBg      = BGR(030,030,030)    ' light black
         .colors(CLR_TEXT).nFg         = BGR(204,204,204)    ' light gray
         .colors(CLR_TEXT).nBg         = BGR(030,030,030)    ' light black
         .colors(CLR_WINAPI).nFg       = BGR(204,204,204)    ' light gray
         .colors(CLR_WINAPI).nBg       = BGR(030,030,030)    ' light black
         .colors(CLR_WINDOW).nFg       = BGR(030,030,030)    ' light black
         .colors(CLR_WINDOW).nBg       = -1
      end with
   end if
   for i as long = lbound(this.Themes) to ubound(this.Themes)
      if this.Themes(i).id = this.SelectedTheme THEN 
         this.idxTheme = i: exit for
      end if   
   NEXT
   
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

   dim as long idx = gApp.GetActiveProjectIndex()

   Dim f As Long = Freefile
   If Open( gApp.Projects(idx).ProjectFilename For Output Encoding "utf16" As #f ) <> 0 Then Return False

   Print #f, "' WINFBE PROJECT FILE"
   Print #f, "ProjectBuild=";       gApp.Projects(idx).ProjectBuild
   Print #f, "ProjectOther32=";     gApp.Projects(idx).ProjectOther32
   Print #f, "ProjectOther64=";     gApp.Projects(idx).ProjectOther64
   Print #f, "ProjectCommandLine="; gApp.Projects(idx).ProjectCommandLine

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
      ' Only deal with files that are no longer "new"
      if pDoc->IsNewFlag = false then 
         if pDoc->ProjectIndex = idx THEN
            wszText = pDoc->DiskFilename
            cwzRelative = AfxPathRelativePathTo( gApp.Projects(idx).ProjectFilename, FILE_ATTRIBUTE_NORMAL, wszText, FILE_ATTRIBUTE_NORMAL)
            if AfxPathIsRelative(cwzRelative) then wszText = cwzRelative
            wszText = ProcessToCurdrive(wszText)
            Print #f, "File="; wszText
            Print #f, "FileType="; pDoc->ProjectFileType
            Print #f, "TabIndex="; -1     ' use -1 rather than True
            Print #f, "Bookmarks="; pDoc->GetBookmarks()
            Print #f, "FirstLine="; SendMessage( pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0) 
            Print #f, "Position="; SendMessage( pDoc->hWindow(0), SCI_GETCURRENTPOS, 0, 0)  
            Print #f, "FirstLine1="; SendMessage( pDoc->hWindow(1), SCI_GETFIRSTVISIBLELINE, 0, 0) 
            Print #f, "Position1="; SendMessage( pDoc->hWindow(1), SCI_GETCURRENTPOS, 0, 0)  
            PRINT #f, "SplitPosition="; pDoc->SplitY
            print #f, "FocusEdit="; iif(pDoc->hWndActiveScintilla = pDoc->hWindow(0), 0, 1)
            Print #f, "FileEnd=[-]" 
         END IF
      end if
   Next
         
   ' Save all other non-displayed documents for the project
   nCount = gApp.Projects(idx).GetDocumentCount 
   For i = 0 To nCount - 1
      pDoc = gApp.Projects(idx).GetDocumentPtr(i)
      If pDoc Then 
         ' Only deal with files that are no longer "new"
         if pDoc->IsNewFlag = false then 
            If gTTabCtl.GetTabIndexByDocumentPtr(pDoc) = -1 Then
               wszText = pDoc->DiskFilename
               cwzRelative = AfxPathRelativePathTo( gApp.Projects(idx).ProjectFilename, FILE_ATTRIBUTE_NORMAL, wszText, FILE_ATTRIBUTE_NORMAL)
               if AfxPathIsRelative(cwzRelative) then wszText = cwzRelative
               wszText = ProcessToCurdrive(wszText)
               Print #f, "File="; wszText
               Print #f, "FileType="; pDoc->ProjectFileType
               Print #f, "TabIndex="; 0    
               Print #f, "FileEnd=[-]"
            End If
         end if
      End If
   Next

   ' If no project is active then save NOTES to config file.
   if gApp.IsProjectActive THEN
      dim wszText as CWSTR = wstr("NOTES-START") + vbcrlf + _
                             gApp.Projects(idx).ProjectNotes + vbcrlf + _
                             wstr("NOTES-END") + vbcrlf
      print #f, ""
      print #f, "[Notes]"
      print #f, wszText                       
   END IF

   Close #f

   Function = True   ' successful save
End Function


' ========================================================================================
' Load all options for the Project from a diskfile
' ========================================================================================
Function clsConfig.ProjectLoadFromFile( byref wzFile as WSTRING) As BOOLEAN    

   Dim st    As WString * MAX_PATH
   Dim wKey  As WString * MAX_PATH
   Dim wData As WString * 1024
      
   Dim sBookmarks As String
   Dim As Long nData, i, iTab, nActiveTab, nFocusEdit 
   dim as long nFirstLine, nPosition, nFirstLine1, nPosition1, nSplitPosition

   ' Info saved for each File and used to create the file once FileEnd is found
   Dim pDoc As clsDocument Ptr
   Dim wszFilename As WString * MAX_PATH
   Dim As Long nFileType, nTabIndex
   Dim As BOOLEAN bLoadInTab
   dim as Boolean bReadingNote
   dim as long Idx    ' project index

   dim as hwnd hTV = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   
   If AfxFileExists(wzFile) = 0 Then Exit Function

   SetCompileStatusBarMessage( "", 0 )
   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )
   AfxDoEvents()

   ' Determine the ProjectIndex for where to store this project
   idx = gApp.GetNewProjectIndex()

   with gApp.Projects(idx)
      .ProjectFilename = wzFile
      .ProjectName = AfxStrPathname( "NAMEX", gApp.Projects(idx).ProjectFilename )
      .hExplorerRootNode = FF_TreeView_InsertItem (hTV, TVI_ROOT, UCase(gApp.Projects(idx).ProjectName) )
      .ProjectNotes = ""
   end with
   Treeview_SetBold(hTV, gApp.Projects(idx).hExplorerRootNode, TRUE)
   
    
   Dim f As Long = Freefile
   If Open( gApp.Projects(idx).ProjectFilename For Input Encoding "utf16" As #f ) <> 0 Then Exit Function

   ' Variable length array to hold sequence of TabCtrl tabs to open. We save the sequence
   ' here and open them only after all over documents have been loaded. This will look
   ' visually more appealing.
   Dim nTabs(any) As clsDocument Ptr

  
   ' Prevent the constant updating of the app caption bar with filename, (pDoc->DisplayStats)
   gProjectLoading = true 
   gApp.ProjectOverrideIndex = idx
   gApp.IsWindowIncludes = false
   gdb2.dbDeleteWinAPI
         
   ' Set InUse here prior to parsing the files of the project
   gApp.Projects(idx).InUse = true

   ' Prevent Treeview flickering during adding
   LockWindowUpdate(hTV)

   Do Until Eof(f)
      AfxDoEvents()
      Line Input #f, st
      If Len(st) = 0 Then Continue Do
      If Left(st, 1) = "'" Then Continue Do
      If Left(st, 1) = "[" Then Continue Do
      
      If Left(st, 11) = "NOTES-START" Then  
         bReadingNote = true
         continue do
      end if   
      If Left(st, 9) = "NOTES-END" Then  
         bReadingNote = false
         continue do
      end if
      if bReadingNote THEN
         gApp.Projects(idx).ProjectNotes = gApp.Projects(idx).ProjectNotes + st + vbcrlf
         continue do
      END IF

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
         Case "ProjectBuild":        gApp.Projects(idx).ProjectBuild       = wData
         Case "ProjectOther32":      gApp.Projects(idx).ProjectOther32     = wData
         Case "ProjectOther64":      gApp.Projects(idx).ProjectOther64     = wData
         Case "ProjectCommandLine":  gApp.Projects(idx).ProjectCommandLine = wData
         Case "ActiveTab":           nActiveTab = nData
         
         Case "File":                wszFilename = wData
         Case "FileType":            nFileType   = nData
         
         Case "TabIndex":            bLoadInTab  = nData
         Case "Bookmarks":           sBookmarks  = Str(wData)
         Case "FirstLine":           nFirstLine  = nData
         Case "Position":            nPosition   = nData
         Case "FirstLine1":          nFirstLine1 = nData
         Case "Position1":           nPosition1  = nData
         Case "SplitPosition":       nSplitPosition = nData
         Case "FocusEdit":           nFocusEdit  = nData
            
         Case "FileEnd":
            ' If this is a relative filename then convert it back.
            if AfxPathIsRelative(wszFilename) then 
               wszFilename = AfxPathCombine( AfxStrPathName("PATH", gApp.Projects(idx).ProjectFilename), wszFilename)
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
               pDoc->ProjectIndex = idx
               pDoc->SplitY = nSplitPosition
               
               If bLoadInTab Then
                  If nTabIndex > Ubound(nTabs) Then
                     Redim Preserve nTabs(nTabIndex) As clsDocument Ptr
                  End If
                  nTabs(nTabIndex) = pDoc
                  nTabIndex += 1
               End If
                   
               ' Set the saved position
               SciExec( pDoc->hWindow(0), SCI_SETFIRSTVISIBLELINE, nFirstLine, 0) 
               SciExec( pDoc->hWindow(0), SCI_GOTOPOS, nPosition, 0) 
               SciExec( pDoc->hWindow(1), SCI_SETFIRSTVISIBLELINE, nFirstLine1, 0) 
               SciExec( pDoc->hWindow(1), SCI_GOTOPOS, nPosition1, 0) 
               pDoc->hWndActiveScintilla = pDoc->hWindow(nFocusEdit)
               nFirstLine  = 0: nPosition  = 0
               nFirstLine1 = 0: nPosition1 = 0
               nFocusEdit = 0
               
               ' Set any bookmarks
               pDoc->SetBookmarks(sBookmarks)
               sBookmarks = ""
            
            End If

      End Select
   
   Loop                        

   LockWindowUpdate(0)
   
   ' Display all of the tabs
   For i = LBound(nTabs) To UBound(nTabs)
      iTab = gTTabCtl.GetTabIndexFromFilename(nTabs(i)->DiskFilename)
      If iTab = -1 Then iTab = gTTabCtl.AddTab( nTabs(i) )  ' Add the new document to the top tabcontrol
   Next   
   gTTabCtl.SetFocusTab( nActiveTab )
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc Then
      Treeview_SelectItem( hTV, pDoc->hNodeExplorer )
   end if
   
   Close #f

   gProjectLoading = false
   gApp.ProjectOverrideIndex = 0

   ' If no active build configuration then assign the currently active selection
   if gApp.Projects(idx).ProjectBuild = "" then 
      gApp.Projects(idx).ProjectBuild = GetSelectedBuildGUID()
   end if   
   
   ' If no other Projects have been designated the as the active project then make
   ' this project the active one.
   if gApp.GetActiveProjectIndex = 0 THEN
      gApp.SetActiveProject(gApp.Projects(idx).hExplorerRootNode)
   END IF

   Treeview_EnsureVisible(hTV, gApp.Projects(idx).hExplorerRootNode)
   
   ' Show the correct notes for this project
   frmOutput_ShowNotes

   SetCursor( hCurSave )

   Function = True   ' successful open
End Function


