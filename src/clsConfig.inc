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


''  CONFIGURATION MODULE

#include once "clsConfig.bi"
#include once "modRoutines.bi"
#include once "frmFunctions.bi"
#include once "frmBookmarks.bi"
#include once "frmExplorer.bi"
#include once "frmOutput.bi"
#include once "frmBuildConfig.bi"
#include once "frmMain.bi"
#include once "modVDColors.bi"
#include once "modVDDesignForm.bi"
#include once "modVDApplyProperties.bi"
#include once "modParser.bi"
#include once "frmTopTabs.bi"
#include once "frmMainOnCommand.bi"
               
''
''  CONSTRUCTOR
''
constructor clsConfig
   _ConfigFilename            = AfxGetExePathName & "Settings\WinFBE.ini" 
   _SnippetsFilename          = AfxGetExePathName & "Settings\user_snippets.ini" 
   _SnippetsDefaultFilename   = AfxGetExePathName & "Settings\user_snippets_default.ini" 
   _FBKeywordsFilename        = AfxGetExePathName & "Settings\freebasic_keywords.txt" 
   _WinApiKeywordsFilename    = AfxGetExePathName & "Settings\winapi_keywords.txt" 
   _FBKeywordsDefaultFilename = AfxGetExePathName & "Settings\freebasic_keywords_default.txt" 
   _FBCodetipsFilename        = AfxGetExePathName & "Settings\codetips.ini" 
   _WinAPICodetipsFilename    = AfxGetExePathName & "Settings\codetips_winapi.ini" 
   _WinFormsXCodetipsFilename = AfxGetExePathName & "Settings\codetips_winformsx.ini" 
   _WinFBXCodetipsFilename    = AfxGetExePathName & "Settings\codetips_winfbx.ini" 
end constructor

''
''
function clsConfig.LoadKeywords() as long
   if AfxFileExists(_FBKeywordsFilename) = false then 
      ' The FB keywords file does not exist. Try copying the "default" keywords over
      ' to the main file that the user can modify.
      if AfxFileExists(_FBKeywordsDefaultFilename) then 
         if AfxCopyFile( _FBKeywordsDefaultFilename.sptr, _FBKeywordsFilename.sptr, true ) = false then
            exit function
         end if
      end if
   end if
   if AfxFileExists(_FBKeywordsFilename) = false then exit function

   dim pStream as CTextStream
   if pStream.Open(_FBKeywordsFilename) <> S_OK then exit function
   this.FBKeywords = pStream.ReadAll
   pStream.Close

   if pStream.Open(_WinApiKeywordsFilename) <> S_OK then exit function
   this.WinApiKeywords = pStream.ReadAll
   pStream.Close
   function = 0
end function

''
''
function clsConfig.SaveKeywords() as long
   dim pStream as CTextStream
   if pStream.Create(_FBKeywordsFilename) <> S_OK then exit function
   pStream.WriteLine this.FBKeywords
   pStream.Close
   if pStream.Create(_WinApiKeywordsFilename) <> S_OK then exit function
   pStream.WriteLine this.WinApiKeywords
   pStream.Close
   function = 0
end function

  
  
''
''
function clsConfig.LoadCodetipsFB() as boolean

   dim as CWSTR wst
   dim as long i
   
   dim parser as ctxParser

   ' First, delete all exisiting codetips of this type
   gdb2.dbDeleteByFileType( DB2_FILETYPE_FB )
   
   parser.nFileType = DB2_FILETYPE_FB
   
   dim pStream as CTextStream
   if pStream.Open(_FBCodetipsFilename) <> S_OK then return true  ' error
   
   do until pStream.EOS
      wst = pStream.ReadLine
      wst = trim(wst)
      if len(wst) = 0 then continue do
      
      i = instr(wst, "=")
      if i then
         parser.functionName   = left(wst, i-1)
         parser.functionParams = mid(wst, i+1)
         gdb2.dbAdd( @parser, DB2_FUNCTION )
      end if   
   loop
   pStream.Close

   return false  ' no error
end function


''
''
function clsConfig.LoadCodetipsGeneric( _
            byval wszFilename as CWSTR, _
            byval nFileType as long _
            ) as boolean
                                        
   dim as CWSTR wst, wsID, wsTypeName, wsTypeElements, wsParts, wsPart, wsBaseType
   dim as long i, nCount
   
   dim parser as ctxParser

   parser.nFileType = nFileType

   ''
   ''  Codetips definition file for WinAPI/WinFBX
   ''
   ''  Format:
   ''  1st character:  F=function, T=TYPE, S=STANDARD_DATATYPE
   ''  2nd character:  colon.
   ''  TYPE elements are separated by | pipe symbol. Each element name/datatype is separated by comma.
   ''  STANDARD datatypes are separated by | pipe symobol (datatype followed by description of the datatype.
   ''
   
   dim pStream as CTextStream
   if pStream.Open(wszFilename) <> S_OK then return true   ' error
   
   do until pStream.EOS
      wst = pStream.ReadLine

      wst = trim(wst)
      if len(wst) = 0 then continue do
      
      i = instr(wst, ":")
      if i = 0 then continue do
      
      wsID = left(wst, i - 1)
      wst  = mid(wst, i + 1)
                
      select case wsID
         case "S"    ' standard datatype
            i = instr(wst, "|")
            parser.functionName = left(wst, i-1)
            'parser.Description = mid(wst, i+1)
            gdb2.dbAdd( @parser, DB2_STANDARDDATATYPE )  
            
         case "F"    ' function
            i = instr(wst, "(")
            parser.functionName   = trim(left(wst, i-1))
            parser.functionParams = wst
            gdb2.dbAdd( @parser, DB2_FUNCTION )

         case "T"     ' type
            i = instr(wst, "|")
            parser.typeName = left(wst, i-1)
            parser.typeAlias = parser.typeName
            wsParts = mid(wst, i + 1)
            gdb2.dbAdd( @parser, DB2_TYPE )

            ' Each element of the TYPE structure is added to the database. The
            ' element is in the format elementname,elementtype
            parser.functionName = ""
            nCount = AfxStrParseCount(wsParts, "|")
            for i = 1 to nCount
               wsPart = AfxStrParse(wsParts, i, "|")
               parser.functionName = parser.typeName
               parser.varName = AfxStrParse(wsPart, 1, ",") 
               parser.varType = AfxStrParse(wsPart, 2, ",") 
               parser.varScope = DIMSCOPE.SCOPETYPE
               gdb2.dbAdd( @parser, DB2_VARIABLE )
            next

      end select
         
   loop
   pStream.Close
   
   return false  ' no error
end function
    

''
''
function clsConfig.LoadCodetipsWinForms( byval wszFilename as CWSTR ) as boolean
                                        
   dim as CWSTR wst
   
   dim parser as ctxParser

   parser.nFileType = DB2_FILETYPE_WINFORMSX

   ''
   ''  Codetips definition file for WinFormsX
   ''
   ''  Format:
   ''
   ''  TYPE:<typename>
   ''  <type_element>|FORWARDTYPE|(CallTip param1 as type, ...) as long
   ''  <type_element>|FORWARDTYPE|(CallTip param1 as type, ...) as long
   ''  <type_element>|FORWARDTYPE|(CallTip param1 as type, ...) as long
   ''  ...etc
   ''  ENDTYPE
   ''
   
   dim arrBase( any ) as CWSTR
   
   dim pStream as CTextStream
   if pStream.Open(wszFilename) <> S_OK then return true   ' error
   
   dim as boolean bReadingType = false
   dim as boolean bIsBase = false
   
   do until pStream.EOS
      wst = pStream.ReadLine

      wst = trim(wst)
      if len(wst) = 0 then continue do
      
      parser.objectStartLine = 0
      parser.objectEndLine = 0

      if left(wst, 5) = "TYPE:" then
         bReadingType = true
         parser.typeName = mid(wst, 6) 
         parser.typeAlias = parser.typeName   
         parser.typeExtends = ""
         if parser.typeName = "BASE" then
            bIsBase = true
         else   
            bIsBase = false
            gdb2.dbAdd( @parser, DB2_TYPE )
         end if
         continue do
      end if
      
      if left(wst, 7) = "ENDTYPE" then
         bReadingType = false
         bIsBase = false
      end if
      
      if bReadingType then
         ' Each element of the TYPE structure is added to the database. 
         parser.functionName = parser.typeName
         parser.varName  = ""
         parser.varType  = ""
         parser.functionParams = ""
         parser.varScope = DIMSCOPE.SCOPEGLOBAL
         if bIsBase then
            dim as long ub = ubound(arrBase) + 1
            redim preserve arrBase(ub)
            arrBase(ub) = wst
         elseif wst = "+BASE" then
            ' add all of the BASE properties to this type
            for i as long = lbound(arrBase) to ubound(arrBase) 
               ' If functionParams exist then this is a function within a TYPE
               dim as CWSTR functionParams = trim(AfxStrParse(arrBase(i), 3, "|"))
               if len(functionParams) then
                  parser.functionName = parser.typeName & "." & trim(AfxStrParse(arrBase(i), 1, "|"))
                  parser.functionParams = functionParams
                  gdb2.dbAdd( @parser, DB2_FUNCTION )
               else
                  parser.varName    = trim(AfxStrParse(arrBase(i), 1, "|"))
                  parser.varType    = trim(AfxStrParse(arrBase(i), 2, "|"))
                  gdb2.dbAdd( @parser, DB2_VARIABLE )
               end if   
            next 
         else
            ' If functionParams exist then this is a function within a TYPE
            dim as CWSTR functionParams = trim(AfxStrParse(wst, 3, "|"))
            if len(functionParams) then
               parser.functionName = parser.typeName & "." & trim(AfxStrParse(wst, 1, "|"))
               parser.functionParams = functionParams
               gdb2.dbAdd( @parser, DB2_FUNCTION )
            else
               parser.varName    = trim(AfxStrParse(wst, 1, "|"))
               parser.varType    = trim(AfxStrParse(wst, 2, "|"))
               gdb2.dbAdd( @parser, DB2_VARIABLE )
            end if   
         end if
      end if
     
   loop
   pStream.Close

   return false  ' no error
end function


''
''
function clsConfig.LoadCodetipsWinFormsX() as boolean
   gdb2.dbDeleteByFileType( DB2_FILETYPE_WINFORMSX )
   function = this.LoadCodetipsWinForms(_WinFormsXCodetipsFilename)

   dim parser as ctxParser
   parser.nFileType = DB2_FILETYPE_WINFORMSX
   
   parser.varName = "Application"
   parser.varType = "wfxApplication"
   parser.varScope = DIMSCOPE.SCOPEGLOBAL
   parser.functionName = ""
   gdb2.dbAdd( @parser, DB2_VARIABLE )
   
   parser.varName = "Colors"
   parser.varType = "wfxColors"
   parser.varScope = DIMSCOPE.SCOPEGLOBAL
   parser.functionName = ""
   gdb2.dbAdd( @parser, DB2_VARIABLE )
   
   function = 0
end function

''
''
function clsConfig.LoadCodetipsWinFBX() as boolean
   gdb2.dbDeleteByFileType( DB2_FILETYPE_WINFBX )
   function = this.LoadCodetipsGeneric(_WinFBXCodetipsFilename, DB2_FILETYPE_WINFBX)
end function

''
''
function clsConfig.LoadCodetipsWinAPI() as boolean
   gdb2.dbDeleteByFileType( DB2_FILETYPE_WINAPI )
   function = this.LoadCodetipsGeneric(_WinAPICodetipsFilename, DB2_FILETYPE_WINAPI)
end function


''
''
function clsConfig.LoadCodetips() as long
   this.LoadCodetipsFB()
   this.LoadCodetipsWinAPI()
   this.LoadCodetipsWinFormsX()
   this.LoadCodetipsWinFBX()
   function = 0
end function


''
''  INITIALIZE THE TOOLBOX CONTROLS
''
function clsConfig.InitializeToolBox() as long
   
   redim gToolBox(50) as TOOLBOX_TYPE
   
   dim as long n = 0
   gToolBox(n).nToolType = CTRL_POINTER
   gToolBox(n).wszToolBoxName   = "Pointer"
   gToolBox(n).wszControlName   = "Pointer"
   gToolBox(n).wszImage  = "IMAGE_POINTER"
   gToolBox(n).wszCursor = ""
   gToolBox(n).wszClassName = ""
   
   n = n + 1
   gToolBox(n).nToolType = CTRL_BUTTON
   gToolBox(n).wszToolBoxName   = "Button"
   gToolBox(n).wszControlName   = "Button"
   gToolBox(n).wszImage  = "IMAGE_BUTTON"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_BUTTON"
   gToolBox(n).wszClassName = "BUTTON"

   n = n + 1
   gToolBox(n).nToolType = CTRL_CHECKBOX
   gToolBox(n).wszToolBoxName   = "CheckBox"
   gToolBox(n).wszControlName   = "Check"
   gToolBox(n).wszImage  = "IMAGE_CHECKBOX"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_CHECKBOX"
   gToolBox(n).wszClassName = "CHECKBOX"

   n = n + 1
   gToolBox(n).nToolType = CTRL_COMBOBOX
   gToolBox(n).wszToolBoxName = "ComboBox"
   gToolBox(n).wszControlName = "Combo"
   gToolBox(n).wszImage  = "IMAGE_COMBOBOX"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_COMBOBOX"
   gToolBox(n).wszClassName = "COMBOBOX"

   n = n + 1
   gToolBox(n).nToolType = CTRL_DATETIMEPICKER
   gToolBox(n).wszToolBoxName   = "DateTimePicker"
   gToolBox(n).wszControlName   = "DateTimePicker"
   gToolBox(n).wszImage  = "IMAGE_DATETIMEPICKER"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_DATETIMEPICKER"
   gToolBox(n).wszClassName = "DATETIMEPICKER"

   n = n + 1
   gToolBox(n).nToolType = CTRL_FRAME
   gToolBox(n).wszToolBoxName   = "Frame"
   gToolBox(n).wszControlName   = "Frame"
   gToolBox(n).wszImage  = "IMAGE_FRAME"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_FRAME"
   gToolBox(n).wszClassName = "GROUPBOX"
            
   n = n + 1
   gToolBox(n).nToolType = CTRL_LABEL
   gToolBox(n).wszToolBoxName   = "Label"
   gToolBox(n).wszControlName   = "Label"
   gToolBox(n).wszImage  = "IMAGE_LABEL"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_LABEL"
   gToolBox(n).wszClassName = "LABEL"

   n = n + 1
   gToolBox(n).nToolType = CTRL_LISTBOX
   gToolBox(n).wszToolBoxName   = "ListBox"
   gToolBox(n).wszControlName   = "List"
   gToolBox(n).wszImage  = "IMAGE_LISTBOX"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_LISTBOX"
   gToolBox(n).wszClassName = "LISTBOX"

   n = n + 1
   gToolBox(n).nToolType = CTRL_LISTVIEW
   gToolBox(n).wszToolBoxName   = "ListView"
   gToolBox(n).wszControlName   = "ListView"
   gToolBox(n).wszImage  = "IMAGE_LISTVIEW"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_LISTVIEW"
   gToolBox(n).wszClassName = "LISTVIEW"

   n = n + 1
   gToolBox(n).nToolType = CTRL_MASKEDEDIT
   gToolBox(n).wszToolBoxName   = "MaskedEdit"
   gToolBox(n).wszControlName   = "Masked"
   gToolBox(n).wszImage  = "IMAGE_MASKED"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_MASKED"
   gToolBox(n).wszClassName = "MASKEDEDIT"

   n = n + 1
   gToolBox(n).nToolType = CTRL_MONTHCALENDAR
   gToolBox(n).wszToolBoxName   = "MonthCalendar"
   gToolBox(n).wszControlName   = "MonthCal"
   gToolBox(n).wszImage  = "IMAGE_MONTHCALENDAR"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_MONTHCALENDAR"
   gToolBox(n).wszClassName = "MONTHCALENDAR"

   n = n + 1
   gToolBox(n).nToolType = CTRL_OPTION
   gToolBox(n).wszToolBoxName   = "OptionButton"
   gToolBox(n).wszControlName   = "Option"
   gToolBox(n).wszImage  = "IMAGE_OPTION"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_OPTION"
   gToolBox(n).wszClassName = "RADIOBUTTON"

   n = n + 1
   gToolBox(n).nToolType = CTRL_PICTUREBOX
   gToolBox(n).wszToolBoxName   = "PictureBox"
   gToolBox(n).wszControlName   = "Picture"
   gToolBox(n).wszImage  = "IMAGE_PICTUREBOX"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_PICTUREBOX"
   gToolBox(n).wszClassName = "PICTUREBOX"

   n = n + 1
   gToolBox(n).nToolType = CTRL_PROGRESSBAR
   gToolBox(n).wszToolBoxName = "ProgressBar"
   gToolBox(n).wszControlName = "Progress"
   gToolBox(n).wszImage = "IMAGE_PROGRESSBAR"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_PROGRESSBAR"
   gToolBox(n).wszClassName = "PROGRESSBAR"

   n = n + 1
   gToolBox(n).nToolType = CTRL_RICHEDIT
   gToolBox(n).wszToolBoxName = "RichEdit"
   gToolBox(n).wszControlName = "RichEdit"
   gToolBox(n).wszImage = "IMAGE_RICHEDIT"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_RICHEDIT"
   gToolBox(n).wszClassName = "RICHEDIT"

   n = n + 1
   gToolBox(n).nToolType = CTRL_TABCONTROL
   gToolBox(n).wszToolBoxName   = "TabControl"
   gToolBox(n).wszControlName   = "TabControl"
   gToolBox(n).wszImage  = "IMAGE_TABCONTROL"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_TABCONTROL"
   gToolBox(n).wszClassName = "TABCONTROL"

   n = n + 1
   gToolBox(n).nToolType = CTRL_TEXTBOX
   gToolBox(n).wszToolBoxName   = "TextBox"
   gToolBox(n).wszControlName   = "Text"
   gToolBox(n).wszImage  = "IMAGE_TEXTBOX"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_TEXTBOX"
   gToolBox(n).wszClassName = "TEXTBOX"

   n = n + 1
   gToolBox(n).nToolType = CTRL_TIMER
   gToolBox(n).wszToolBoxName   = "Timer"
   gToolBox(n).wszControlName   = "Timer"
   gToolBox(n).wszImage  = "IMAGE_TIMER"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_TIMER"
   gToolBox(n).wszClassName = ""

   n = n + 1
   gToolBox(n).nToolType = CTRL_TREEVIEW
   gToolBox(n).wszToolBoxName   = "TreeView"
   gToolBox(n).wszControlName   = "TreeView"
   gToolBox(n).wszImage  = "IMAGE_TREEVIEW"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_TREEVIEW"
   gToolBox(n).wszClassName = "TREEVIEW"

   n = n + 1
   gToolBox(n).nToolType = CTRL_UPDOWN
   gToolBox(n).wszToolBoxName   = "UpDown"
   gToolBox(n).wszControlName   = "UpDown"
   gToolBox(n).wszImage  = "IMAGE_UPDOWN"
   gToolBox(n).wszCursor = "IMAGE_CURSOR_UPDOWN"
   gToolBox(n).wszClassName = "UPDOWN"

   redim preserve gToolBox(n) as TOOLBOX_TYPE


   ' System Colors
   redim gColors(172)
   gColors(0).SetColor("ActiveBorder", COLOR_SYSTEM, GetSysColor(COLOR_ACTIVEBORDER))
   gColors(1).SetColor("ActiveCaption", COLOR_SYSTEM, GetSysColor(COLOR_ACTIVECAPTION))
   gColors(2).SetColor("ActiveCaptionText", COLOR_SYSTEM, GetSysColor(COLOR_CAPTIONTEXT))
   gColors(3).SetColor("AppWorkspace", COLOR_SYSTEM, GetSysColor(COLOR_APPWORKSPACE))
   gColors(4).SetColor("ButtonFace", COLOR_SYSTEM, GetSysColor(COLOR_BTNFACE))
   gColors(5).SetColor("ButtonHighlight", COLOR_SYSTEM, GetSysColor(COLOR_BTNHILIGHT))
   gColors(5).SetColor("ButtonShadow", COLOR_SYSTEM,GetSysColor(COLOR_BTNSHADOW) )
   gColors(6).SetColor("Control", COLOR_SYSTEM, GetSysColor(COLOR_3DFACE))
   gColors(7).SetColor("ControlDark", COLOR_SYSTEM, GetSysColor(COLOR_3DSHADOW))
   gColors(8).SetColor("ControlDarkDark", COLOR_SYSTEM, GetSysColor(COLOR_3DDKSHADOW))
   gColors(9).SetColor("ControlLight", COLOR_SYSTEM, GetSysColor(COLOR_3DLIGHT))
   gColors(10).SetColor("ControlLightLight", COLOR_SYSTEM, GetSysColor(COLOR_3DHILIGHT))
   gColors(11).SetColor("ControlText", COLOR_SYSTEM, GetSysColor(COLOR_BTNTEXT))
   gColors(12).SetColor("Desktop", COLOR_SYSTEM, GetSysColor(COLOR_DESKTOP))
   gColors(13).SetColor("GradientActiveCaption", COLOR_SYSTEM, GetSysColor(COLOR_GRADIENTACTIVECAPTION))
   gColors(14).SetColor("GradientInactiveCaption", COLOR_SYSTEM, GetSysColor(COLOR_GRADIENTINACTIVECAPTION))
   gColors(15).SetColor("GrayText", COLOR_SYSTEM, GetSysColor(COLOR_GRAYTEXT))
   gColors(16).SetColor("Highlight", COLOR_SYSTEM, GetSysColor(COLOR_HIGHLIGHT))
   gColors(17).SetColor("HighlightText", COLOR_SYSTEM, GetSysColor(COLOR_HIGHLIGHTTEXT))
   gColors(18).SetColor("HotTrack", COLOR_SYSTEM, GetSysColor(COLOR_HOTLIGHT))
   gColors(19).SetColor("InactiveBorder", COLOR_SYSTEM, GetSysColor(COLOR_INACTIVEBORDER))
   gColors(20).SetColor("InactiveCaption", COLOR_SYSTEM, GetSysColor(COLOR_INACTIVECAPTION))
   gColors(21).SetColor("InactiveCaptionText", COLOR_SYSTEM, GetSysColor(COLOR_INACTIVECAPTIONTEXT))
   gColors(22).SetColor("Info", COLOR_SYSTEM, GetSysColor(COLOR_INFOBK))
   gColors(23).SetColor("InfoText", COLOR_SYSTEM, GetSysColor(COLOR_INFOTEXT))
   gColors(24).SetColor("Menu", COLOR_SYSTEM, GetSysColor(COLOR_MENU))
   gColors(25).SetColor("MenuBar", COLOR_SYSTEM, GetSysColor(COLOR_MENUBAR))
   gColors(26).SetColor("MenuHighlight", COLOR_SYSTEM, GetSysColor(COLOR_MENUHILIGHT))
   gColors(27).SetColor("MenuText", COLOR_SYSTEM, GetSysColor(COLOR_MENUTEXT))
   gColors(28).SetColor("ScrollBar", COLOR_SYSTEM, GetSysColor(COLOR_SCROLLBAR))
   gColors(29).SetColor("Window", COLOR_SYSTEM, GetSysColor(COLOR_WINDOW))
   gColors(30).SetColor("WindowFrame", COLOR_SYSTEM, GetSysColor(COLOR_WINDOWFRAME))
   gColors(31).SetColor("WindowText", COLOR_SYSTEM, GetSysColor(COLOR_WINDOWTEXT))

   ' QUICK COLORS
   gColors(32).SetColor("AliceBlue"            , COLOR_COLORS, BGR(240,248,255))
   gColors(33).SetColor("AntiqueWhite"         , COLOR_COLORS, BGR(250,235,215))
   gColors(34).SetColor("Aqua"                 , COLOR_COLORS, BGR(  0,255,255))
   gColors(35).SetColor("Aquamarine"           , COLOR_COLORS, BGR(127,255,212)) 
   gColors(36).SetColor("Azure"                , COLOR_COLORS, BGR(240,255,255))
   gColors(37).SetColor("Beige"                , COLOR_COLORS, BGR(245,245,220))
   gColors(38).SetColor("Bisque"               , COLOR_COLORS, BGR(255,228,196))
   gColors(39).SetColor("Black"                , COLOR_COLORS, BGR(  0,  0,  0))
   gColors(40).SetColor("BlanchedAlmond"       , COLOR_COLORS, BGR(255,255,205))
   gColors(41).SetColor("Blue"                 , COLOR_COLORS, BGR(  0,  0,255))
   gColors(42).SetColor("BlueViolet"           , COLOR_COLORS, BGR(138, 43,226))
   gColors(43).SetColor("Brown"                , COLOR_COLORS, BGR(165, 42, 42))
   gColors(44).SetColor("Burlywood"            , COLOR_COLORS, BGR(222,184,135))
   gColors(45).SetColor("CadetBlue"            , COLOR_COLORS, BGR( 95,158,160))
   gColors(46).SetColor("Chartreuse"           , COLOR_COLORS, BGR(127,255,  0))
   gColors(47).SetColor("Chocolate"            , COLOR_COLORS, BGR(210,105, 30))
   gColors(48).SetColor("Coral"                , COLOR_COLORS, BGR(255,127, 80))
   gColors(49).SetColor("CornflowerBlue"       , COLOR_COLORS, BGR(100,149,237))
   gColors(50).SetColor("Cornsilk"             , COLOR_COLORS, BGR(255,248,220))
   gColors(51).SetColor("Crimson"              , COLOR_COLORS, BGR(220, 20, 60))
   gColors(52).SetColor("Cyan"                 , COLOR_COLORS, BGR(  0,255,255))
   gColors(53).SetColor("DarkBlue"             , COLOR_COLORS, BGR(  0,  0,139))
   gColors(54).SetColor("DarkCyan"             , COLOR_COLORS, BGR(  0,139,139))
   gColors(55).SetColor("DarkGoldenRod"        , COLOR_COLORS, BGR(184,134, 11)) 
   gColors(56).SetColor("DarkGray"             , COLOR_COLORS, BGR(169,169,169))
   gColors(57).SetColor("DarkGreen"            , COLOR_COLORS, BGR(  0,100,  0))
   gColors(58).SetColor("DarkKhaki"            , COLOR_COLORS, BGR(189,183,107))
   gColors(59).SetColor("DarkMagenta"          , COLOR_COLORS, BGR(139,  0,139))
   gColors(60).SetColor("DarkOliveGreen"       , COLOR_COLORS, BGR( 85,107, 47)) 
   gColors(61).SetColor("DarkOrange"           , COLOR_COLORS, BGR(255,140,  0)) 
   gColors(62).SetColor("DarkOrchid"           , COLOR_COLORS, BGR(153, 50,204))
   gColors(63).SetColor("DarkRed"              , COLOR_COLORS, BGR(139,  0,  0))
   gColors(64).SetColor("DarkSalmon"           , COLOR_COLORS, BGR(233,150,122))
   gColors(65).SetColor("DarkSeaGreen"         , COLOR_COLORS, BGR(143,188,143))
   gColors(66).SetColor("DarkSlateBlue"        , COLOR_COLORS, BGR( 72, 61,139))
   gColors(67).SetColor("DarkSlateGray"        , COLOR_COLORS, BGR( 47, 79, 79))
   gColors(68).SetColor("DarkTurquoise"        , COLOR_COLORS, BGR(  0,206,209))
   gColors(69).SetColor("DarkViolet"           , COLOR_COLORS, BGR(148,  0,211))
   gColors(70).SetColor("DeepPink"             , COLOR_COLORS, BGR(255, 20,147))
   gColors(71).SetColor("DeepSkyBlue"          , COLOR_COLORS, BGR(  0,191,255))
   gColors(72).SetColor("DimGray"              , COLOR_COLORS, BGR(105,105,105))
   gColors(73).SetColor("DodgerBlue"           , COLOR_COLORS, BGR( 30,144,255))
   gColors(74).SetColor("FireBrick"            , COLOR_COLORS, BGR(178, 34, 34))
   gColors(75).SetColor("FloralWhite"          , COLOR_COLORS, BGR(255,250,240))
   gColors(76).SetColor("ForestGreen"          , COLOR_COLORS, BGR( 34,139, 34))
   gColors(77).SetColor("Fuchsia"              , COLOR_COLORS, BGR(255,  0,255))
   gColors(78).SetColor("Gainsboro"            , COLOR_COLORS, BGR(220,220,220))
   gColors(79).SetColor("GhostWhite"           , COLOR_COLORS, BGR(248,248,255))
   gColors(80).SetColor("Gold"                 , COLOR_COLORS, BGR(255,215,  0))
   gColors(81).SetColor("GoldenRod"            , COLOR_COLORS, BGR(218,165, 32))
   gColors(82).SetColor("Gray"                 , COLOR_COLORS, BGR(127,127,127))
   gColors(83).SetColor("Green"                , COLOR_COLORS, BGR(  0,128,  0))
   gColors(84).SetColor("GreenYellow"          , COLOR_COLORS, BGR(173,255, 47))
   gColors(85).SetColor("HoneyDew"             , COLOR_COLORS, BGR(240,255,240))
   gColors(86).SetColor("HotPink"              , COLOR_COLORS, BGR(255,105,180))
   gColors(87).SetColor("IndianRed"            , COLOR_COLORS, BGR(205, 92, 92))
   gColors(88).SetColor("Indigo"               , COLOR_COLORS, BGR( 75,  0,130))
   gColors(89).SetColor("Ivory"                , COLOR_COLORS, BGR(255,255,240))
   gColors(90).SetColor("Khaki"                , COLOR_COLORS, BGR(240,230,140))
   gColors(91).SetColor("Lavender"             , COLOR_COLORS, BGR(230,230,250))
   gColors(92).SetColor("LavenderBlush"        , COLOR_COLORS, BGR(255,240,245))
   gColors(93).SetColor("Lawngreen"            , COLOR_COLORS, BGR(124,252,  0))
   gColors(94).SetColor("LemonChiffon"         , COLOR_COLORS, BGR(255,250,205))
   gColors(95).SetColor("LightBlue"            , COLOR_COLORS, BGR(173,216,230))
   gColors(96).SetColor("LightCoral"           , COLOR_COLORS, BGR(240,128,128))
   gColors(97).SetColor("LightCyan"            , COLOR_COLORS, BGR(224,255,255))
   gColors(98).SetColor("LightGoldenRodYellow" , COLOR_COLORS, BGR(250,250,210))
   gColors(99).SetColor("LightGreen"           , COLOR_COLORS, BGR(144,238,144))
   gColors(100).SetColor("LightGrey"           , COLOR_COLORS, BGR(211,211,211))
   gColors(101).SetColor("LightPink"           , COLOR_COLORS, BGR(255,182,193))
   gColors(102).SetColor("LightSalmon"         , COLOR_COLORS, BGR(255,160,122))
   gColors(103).SetColor("LightSeaGreen"       , COLOR_COLORS, BGR( 32,178,170))
   gColors(104).SetColor("LightSkyBlue"        , COLOR_COLORS, BGR(135,206,250))
   gColors(105).SetColor("LightSlateGray"      , COLOR_COLORS, BGR(119,136,153))
   gColors(106).SetColor("LightSteelBlue"      , COLOR_COLORS, BGR(176,196,222))
   gColors(107).SetColor("LightYellow"         , COLOR_COLORS, BGR(255,255,224))
   gColors(108).SetColor("Lime"                , COLOR_COLORS, BGR(  0,255,  0))
   gColors(109).SetColor("LimeGreen"           , COLOR_COLORS, BGR( 50,205, 50))
   gColors(110).SetColor("Linen"               , COLOR_COLORS, BGR(250,240,230))
   gColors(111).SetColor("Magenta"             , COLOR_COLORS, BGR(255,  0,255))
   gColors(112).SetColor("Maroon"              , COLOR_COLORS, BGR(128,  0,  0))
   gColors(113).SetColor("MediumAquamarine"    , COLOR_COLORS, BGR(102,205,170))
   gColors(114).SetColor("MediumBlue"          , COLOR_COLORS, BGR(  0,  0,205))
   gColors(115).SetColor("MediumOrchid"        , COLOR_COLORS, BGR(186, 85,211))
   gColors(116).SetColor("MediumPurple"        , COLOR_COLORS, BGR(147,112,219))
   gColors(117).SetColor("MediumSeaGreen"      , COLOR_COLORS, BGR( 60,179,113))
   gColors(118).SetColor("MediumSlateBlue"     , COLOR_COLORS, BGR(123,104,238))
   gColors(119).SetColor("MediumSpringGreen"   , COLOR_COLORS, BGR(  0,250,154))
   gColors(120).SetColor("MediumTurquoise"     , COLOR_COLORS, BGR( 72,209,204))
   gColors(121).SetColor("MediumVioletRed"     , COLOR_COLORS, BGR(199, 21,133))
   gColors(122).SetColor("MidnightBlue"        , COLOR_COLORS, BGR( 25, 25,112))
   gColors(123).SetColor("MintCream"           , COLOR_COLORS, BGR(245,255,250))
   gColors(124).SetColor("MistyRose"           , COLOR_COLORS, BGR(255,228,225))
   gColors(125).SetColor("Moccasin"            , COLOR_COLORS, BGR(255,228,181))
   gColors(126).SetColor("NavajoWhite"         , COLOR_COLORS, BGR(255,222,173))
   gColors(127).SetColor("Navy"                , COLOR_COLORS, BGR(  0,  0,128))
   gColors(128).SetColor("Navyblue"            , COLOR_COLORS, BGR(159,175,223))
   gColors(129).SetColor("OldLace"             , COLOR_COLORS, BGR(253,245,230))
   gColors(130).SetColor("Olive"               , COLOR_COLORS, BGR(128,128,  0))
   gColors(131).SetColor("OliveDrab"           , COLOR_COLORS, BGR(107,142, 35))
   gColors(132).SetColor("Orange"              , COLOR_COLORS, BGR(255,165,  0))
   gColors(133).SetColor("OrangeRed"           , COLOR_COLORS, BGR(255, 69,  0))
   gColors(134).SetColor("Orchid"              , COLOR_COLORS, BGR(218,112,214))
   gColors(135).SetColor("PaleGoldenRod"       , COLOR_COLORS, BGR(238,232,170))
   gColors(136).SetColor("PaleGreen"           , COLOR_COLORS, BGR(152,251,152))
   gColors(137).SetColor("PaleTurquoise"       , COLOR_COLORS, BGR(175,238,238))
   gColors(138).SetColor("PaleVioletRed"       , COLOR_COLORS, BGR(219,112,147))
   gColors(139).SetColor("PapayaWhip"          , COLOR_COLORS, BGR(255,239,213))
   gColors(140).SetColor("PeachPuff"           , COLOR_COLORS, BGR(255,218,185))
   gColors(141).SetColor("Peru"                , COLOR_COLORS, BGR(205,133, 63))
   gColors(142).SetColor("Pink"                , COLOR_COLORS, BGR(255,192,203))
   gColors(143).SetColor("Plum"                , COLOR_COLORS, BGR(221,160,221))
   gColors(144).SetColor("PowderBlue"          , COLOR_COLORS, BGR(176,224,230))
   gColors(145).SetColor("Purple"              , COLOR_COLORS, BGR(128,  0,128))
   gColors(146).SetColor("Red"                 , COLOR_COLORS, BGR(255,  0,  0))
   gColors(147).SetColor("RosyBrown"           , COLOR_COLORS, BGR(188,143,143))
   gColors(148).SetColor("RoyalBlue"           , COLOR_COLORS, BGR( 65,105,225))
   gColors(149).SetColor("SaddleBrown"         , COLOR_COLORS, BGR(139, 69, 19))
   gColors(150).SetColor("Salmon"              , COLOR_COLORS, BGR(250,128,114))
   gColors(151).SetColor("SandyBrown"          , COLOR_COLORS, BGR(244,164, 96))
   gColors(152).SetColor("SeaGreen"            , COLOR_COLORS, BGR( 46,139, 87))
   gColors(153).SetColor("SeaShell"            , COLOR_COLORS, BGR(255,245,238))
   gColors(154).SetColor("Sienna"              , COLOR_COLORS, BGR(160, 82, 45))
   gColors(155).SetColor("Silver"              , COLOR_COLORS, BGR(192,192,192))
   gColors(156).SetColor("SkyBlue"             , COLOR_COLORS, BGR(135,206,235))
   gColors(157).SetColor("SlateBlue"           , COLOR_COLORS, BGR(106, 90,205))
   gColors(158).SetColor("SlateGray"           , COLOR_COLORS, BGR(112,128,144))
   gColors(159).SetColor("Snow"                , COLOR_COLORS, BGR(255,250,250))
   gColors(160).SetColor("SpringGreen"         , COLOR_COLORS, BGR(  0,255,127))
   gColors(161).SetColor("SteelBlue"           , COLOR_COLORS, BGR( 70,130,180))
   gColors(162).SetColor("Tan"                 , COLOR_COLORS, BGR(210,180,140))
   gColors(163).SetColor("Teal"                , COLOR_COLORS, BGR(  0,128,128))
   gColors(164).SetColor("Thistle"             , COLOR_COLORS, BGR(216,191,216))
   gColors(165).SetColor("Tomato"              , COLOR_COLORS, BGR(255, 99, 71))
   gColors(166).SetColor("Turquoise"           , COLOR_COLORS, BGR( 64,224,208))
   gColors(167).SetColor("Violet"              , COLOR_COLORS, BGR(238,130,238))
   gColors(168).SetColor("Wheat"               , COLOR_COLORS, BGR(245,222,179))
   gColors(169).SetColor("White"               , COLOR_COLORS, BGR(255,255,255))
   gColors(170).SetColor("WhiteSmoke"          , COLOR_COLORS, BGR(245,245,245))
   gColors(171).SetColor("Yellow"              , COLOR_COLORS, BGR(255,255,  0))
   gColors(172).SetColor("YellowGreen"         , COLOR_COLORS, BGR(139,205, 50))
   
   function = 0
end function


' ========================================================================================
' Enumerate the names of all the fonts. Note the difference between how to enumerate them
' (%TMPF_FIXED_PITCH has the bit cleared).
' %TMPF_FIXED_PITCH for fixed pitch fonts (like in PB edit)
' %TMPF_TRUETYPE OR %TMPF_VECTOR for true type and vector fonts
' %TMPF_DEVICE for device fonts (like printer fonts)
' Exclude what you don't want to include in list.
' ========================================================================================
function fonts_EnumFontName( _
            byref lf as LOGFONTW, _
            byref tm as TEXTMETRIC, _
            ByVal FontType as long, _
            HWnd as HWnd _
            ) as long
   dim as long ub = ubound(gFontNames)
   redim preserve gFontNames(ub + 1) as CWSTR
   gFontNames(ub + 1) = lf.lfFaceName 
   function = true
end function

' ========================================================================================
' Search the gFontNames() array to see if the font name exists
' ========================================================================================
function isFontNameExist( byval wszFontName as CWSTR ) as boolean
   wszFontName = ucase(wszFontName)
   for i as long = lbound(gFontnames) to ubound(gFontNames)
      if wszFontName = ucase(gFontNames(i)) then return true   
   next
   return false
end function

' ========================================================================================
' Get a list of all font names installed on this system. We check this list to ensure
' that the currently user selected font is actually available on the system that they
' are using WinFBE under.
' function returns true of the defined font name was valid on the system so no changes
' were made to the config file fontname value.
' ========================================================================================
function doFontSanityCheck() as long
   erase gFontNames
   dim hDC as HDC = GetDC(0)
   EnumFontFamilies( hDC, ByVal 0, Cast(FONTENUMPROCW, @fonts_EnumFontName), 0 )
   ReleaseDC( 0, hDC )
   
   ' Is the font defined in the config file okay?
   if isFontNameExist( gConfig.EditorFontname ) then return true

   'print "Defined font not found. Finding an alternative..."
   
   ' Test if we are running under wine
   if gApp.isWineActive then
      'print "Linux Wine is active"
      ' Test valid Linux fonts   
      if isFontNameExist( "DejaVu Sans Mono" ) then gConfig.EditorFontname = "DejaVu Sans Mono": return false
      if isFontNameExist( "Ubuntu Mono" ) then gConfig.EditorFontname = "Ubuntu Mono": return false
      if isFontNameExist( "Victor Mono" ) then gConfig.EditorFontname = "Victor Mono": return false
      if isFontNameExist( "Inconsolata" ) then gConfig.EditorFontname = "Inconsolata": return false
      if isFontNameExist( "Fira Code" ) then gConfig.EditorFontname = "Fira Code": return false
   else
      ' Test other valid Windows fonts   
      if isFontNameExist( "Consolas" ) then gConfig.EditorFontname = "Consolas": return false
      if isFontNameExist( "MonoLisa" ) then gConfig.EditorFontname = "MonoLisa": return false
      if isFontNameExist( "Apercu Mono" ) then gConfig.EditorFontname = "Apercu Mono": return false
      if isFontNameExist( "Fira Code" ) then gConfig.EditorFontname = "Fira Code": return false
      if isFontNameExist( "Dank Mono" ) then gConfig.EditorFontname = "Dank Mono": return false
      if isFontNameExist( "Hack" ) then gConfig.EditorFontname = "Hack": return false
      if isFontNameExist( "Courier New" ) then gConfig.EditorFontname = "Courier New": return false
   end if

   function = false
end function


''
''  SAVE CONFIGURATION TO DISK FILE
''
function clsConfig.SaveConfigFile() as long

   ' Determine the current editor positioning 
   ' Do not save if editor is minimized
   if isiconic( HWND_FRMMAIN ) = 0 then
      dim WinPla as WINDOWPLACEMENT
      WinPla.Length = sizeof(WinPla)
      GetWindowPlacement( HWND_FRMMAIN, @WinPla )
      With this
         .StartupLeft   = WinPla.rcNormalPosition.Left 
         .StartupTop    = WinPla.rcNormalPosition.Top
         .StartupRight  = WinPla.rcNormalPosition.Right
         .StartupBottom = WinPla.rcNormalPosition.Bottom
         .StartupMaximized = iif( WinPla.showCmd = SW_MAXIMIZE, true, false )
      end With
      
      dim as RECT rc
      GetWindowRect( HWND_FRMMAIN, @rc )
      if this.StartupLeft <> rc.Left then this.StartupLeft = rc.Left
      if this.StartupTop <> rc.Top then this.StartupTop = rc.Top
      if this.StartupRight <> rc.Right then this.StartupRight = rc.Right
      if this.StartupBottom <> rc.Bottom then this.StartupBottom = rc.Bottom
   end if
   
   this.CompilerBuild = frmBuildConfig_GetSelectedBuildGUID()

   dim pStream as CTextStream   '(utf16)
   if pStream.Create(_ConfigFilename, true, true) <> S_OK then return true   ' error
   
   pStream.WriteLine "'  WINFBE CONFIGURATION" 
   pStream.WriteLine ""
   pStream.WriteLine "WinFBEversion="         & APPVERSION
   pStream.WriteLine ""
   pStream.WriteLine "[Editor]"
   pStream.WriteLine "AskExit="               & this.AskExit
   pStream.WriteLine "CheckForUpdates="       & this.CheckForUpdates
   pStream.WriteLine "EnableProjectCache="    & this.EnableProjectCache
   pStream.WriteLine "LastUpdateCheck="       & this.LastUpdateCheck
   pStream.WriteLine "AutoSaveFiles="         & this.AutoSaveFiles
   pStream.WriteLine "AutoSaveInterval="      & this.AutoSaveInterval
   pStream.WriteLine "RestoreSession="        & this.RestoreSession
   pStream.WriteLine "LastActiveSessionFile=" & ProcessToCurdriveApp(this.wszLastActiveSession)
   pStream.WriteLine "MultipleInstances="     & this.MultipleInstances    
   pStream.WriteLine "CompileAutosave="       & this.CompileAutosave      
   pStream.WriteLine "CloseFuncList="         & this.CloseFuncList        
   pStream.WriteLine "SyntaxHighlighting="    & this.SyntaxHighlighting   
   pStream.WriteLine "Codetips="              & this.Codetips             
   pStream.WriteLine "AutoComplete="          & this.AutoComplete
   pStream.WriteLine "CharacterAutoComplete=" & this.CharacterAutoComplete
   pStream.WriteLine "RightEdge="             & this.RightEdge
   pStream.WriteLine "RightEdgePosition="     & this.RightEdgePosition           
   pStream.WriteLine "LeftMargin="            & this.LeftMargin           
   pStream.WriteLine "FoldMargin="            & this.FoldMargin           
   pStream.WriteLine "AutoIndentation="       & this.AutoIndentation      
   pStream.WriteLine "FornextVariable="       & this.FornextVariable
   pStream.WriteLine "ConfineCaret="          & this.ConfineCaret         
   pStream.WriteLine "LineNumbering="         & this.LineNumbering        
   pStream.WriteLine "HighlightCurrentLine="  & this.HighlightCurrentLine 
   pStream.WriteLine "IndentGuides="          & this.IndentGuides         
   pStream.WriteLine "TabIndentSpaces="       & this.TabIndentSpaces      
   pStream.WriteLine "UnicodeEncoding="       & this.UnicodeEncoding      
   pStream.WriteLine "KeywordCase="           & this.KeywordCase          
   pStream.WriteLine "LocalizationFile="      & this.LocalizationFile     
   pStream.WriteLine "TabSize="               & this.TabSize              
   pStream.WriteLine "PositionMiddle="        & this.PositionMiddle
   pStream.WriteLine "BraceHighlight="        & this.BraceHighlight
   pStream.WriteLine "OccurrenceHighlight="   & this.OccurrenceHighlight
   pStream.WriteLine "EditorFontname="        & this.EditorFontname       
   pStream.WriteLine "EditorFontsize="        & this.EditorFontsize       
   pStream.WriteLine "EditorFontCharSet="     & this.EditorFontCharSet    
   pStream.WriteLine "FontExtraSpace="        & this.FontExtraSpace
   pStream.WriteLine "Theme="                 & this.ThemeFilename
   pStream.WriteLine ""
   pStream.WriteLine "[Startup]" 
   pStream.WriteLine "StartupLeft="           & this.StartupLeft          
   pStream.WriteLine "StartupTop="            & this.StartupTop           
   pStream.WriteLine "StartupRight="          & this.StartupRight         
   pStream.WriteLine "StartupBottom="         & this.StartupBottom        
   pStream.WriteLine "StartupMaximized="      & this.StartupMaximized     
   pStream.WriteLine "ShowPanel="             & this.ShowPanel
   pStream.WriteLine "ShowPanelWidth="        & this.ShowPanelWidth    
   pStream.WriteLine "ToolBoxLeft="           & this.ToolBoxLeft          
   pStream.WriteLine "ToolBoxTop="            & this.ToolBoxTop           
   pStream.WriteLine "ToolBoxRight="          & this.ToolBoxRight         
   pStream.WriteLine "ToolBoxBottom="         & this.ToolBoxBottom        
   pStream.WriteLine ""

   ' for each folder location determine if it resides on the same drive as
   ' the WinFBE application. if it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   this.FBWINCompiler32  = ProcessToCurdriveApp(this.FBWINCompiler32)
   this.FBWINCompiler64  = ProcessToCurdriveApp(this.FBWINCompiler64)
   this.CompilerHelpfile = ProcessToCurdriveApp(this.CompilerHelpfile)
   this.WinFBXHelpfile   = ProcessToCurdriveApp(this.WinFBXHelpfile)

   pStream.WriteLine "[Compiler]" 
   pStream.WriteLine "FBWINCompiler32="       & this.FBWINCompiler32      
   pStream.WriteLine "FBWINCompiler64="       & this.FBWINCompiler64      
   pStream.WriteLine "CompilerBuild="         & this.CompilerBuild     
   pStream.WriteLine "CompilerSwitches="      & this.CompilerSwitches     
   pStream.WriteLine "CompilerHelpfile="      & this.CompilerHelpfile     
   pStream.WriteLine "WinFBXHelpfile="        & this.WinFBXHelpfile     
   pStream.WriteLine "RunViaCommandWindow="   & this.RunViaCommandWindow
   pStream.WriteLine "DisableCompileBeep="    & this.DisableCompileBeep
   pStream.WriteLine "WinFBXPath="            & this.WinFBXPath
   

   pStream.WriteLine ""
   pStream.WriteLine "[Categories]"  
   dim as long nNext = 0
   for i as long = lbound(this.Cat) to ubound(this.Cat)
      if left(this.Cat(i).wszDescription, 2) <> "%%" then
         if trim(this.Cat(i).wszDescription) <> "" then
            pStream.Write "CATEGORY_" & right("00" & str(nNext), 2) & "="
            pStream.Write this.Cat(i).idFileType & "|-|"
            pStream.WriteLine this.Cat(i).wszDescription
            nNext = nNext + 1
         end if
      end if   
   next
   pStream.WriteLine ""
   
   pStream.WriteLine ""
   pStream.WriteLine "[UserTools]"  
   for i as long = lbound(this.Tools) to ubound(this.Tools)
      pStream.Write "USERTOOL_" & right("00" & str(i), 2) & "="
      pStream.Write this.Tools(i).wszDescription & "|-|"
      pStream.Write ProcessToCurdriveApp(this.Tools(i).wszCommand) & "|-|" 
      pStream.Write this.Tools(i).wszParameters & "|-|"
      pStream.Write this.Tools(i).wszKey & "|-|" 
      pStream.Write ProcessToCurdriveApp(this.Tools(i).wszWorkingFolder) & "|-|"
      pStream.Write this.Tools(i).IsCtrl & "|-|"
      pStream.Write this.Tools(i).IsAlt & "|-|"  
      pStream.Write this.Tools(i).IsShift & "|-|"
      pStream.Write this.Tools(i).IsPromptRun & "|-|"
      pStream.Write this.Tools(i).IsMinimized & "|-|"  
      pStream.Write this.Tools(i).IsWaitFinish & "|-|"
      pStream.Write this.Tools(i).IsDisplayMenu & "|-|"    
      pStream.WriteLine this.Tools(i).Action   
   next

   pStream.WriteLine ""
   pStream.WriteLine "[Builds]"  
   for i as long = lbound(this.Builds) to ubound(this.Builds)
      pStream.Write "BUILD_" & right("00" & str(i), 2) & "="
      pStream.Write this.Builds(i).id & "|-|"
      pStream.Write this.Builds(i).wszDescription & "|-|"
      pStream.Write this.Builds(i).wszOptions & "|-|" 
      pStream.Write this.Builds(i).IsDefault & "|-|"  
      pStream.Write this.Builds(i).Is32bit & "|-|"  
      pStream.WriteLine this.Builds(i).Is64bit
   next

   
   pStream.WriteLine ""
   pStream.WriteLine "[MRU]"  
   for i as long = 0 To 9
      this.MRU(i) = ProcessToCurdriveApp(this.MRU(i))
      pStream.WriteLine "MRU_" & right("00" & str(i), 2) & "=" & this.MRU(i) 
   next         
   
   pStream.WriteLine ""
   pStream.WriteLine "[MRUPROJECTS]"  
   for i as long = 0 To 9
      this.MRUProject(i) = ProcessToCurdriveApp(this.MRUProject(i))
      pStream.WriteLine "MRUPROJECT_" & right("00" & str(i), 2) & "=" & this.MRUProject(i) 
   next         
   
   dim hCtl as hwnd = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES)
   dim wszText as CWSTR = wstr("NOTES-START") + vbcrlf + _
                          gApp.NonProjectNotes + vbcrlf + _
                          wstr("NOTES-END") + vbcrlf
   pStream.WriteLine ""
   pStream.WriteLine "[Notes]"
   pStream.WriteLine wszText

   pStream.Close 
   
   function = 0
end function


' Save the MRU to the configuration file. Only write the MRU items
' the ini file rather than overwriting the whole file. This enables
' WinFBE to work better with external tools that manually modify the
' WinFBE.ini file.
function clsConfig.WriteMRU() as long 
   dim as CWSTR wszKeyName, wszValue
   
   ' Delete the section name
   dim as CWSTR wszSectionName = "MRU"
   WritePrivateProfileString( wszSectionName, null, null, _ConfigFilename )
   
   ' Write the new values.
   for i as long = 0 To 9
      this.MRU(i) = ProcessToCurdriveApp(this.MRU(i))
      wszKeyName = "MRU_" & right("00" & str(i), 2)
      wszValue   = this.MRU(i) 
      WritePrivateProfileString(wszSectionName, wszKeyName, wszValue, _ConfigFilename)
   next         
    
   function = 0
end function


' Save the MRU Projects to the configuration file. Only write the MRU items
' the ini file rather than overwriting the whole file. This enables
' WinFBE to work better with external tools that manually modify the
' WinFBE.ini file.
function clsConfig.WriteMRUProjects() as long 
   dim as CWSTR wszKeyName, wszValue
   
   ' Delete the section name
   dim as CWSTR wszSectionName = "MRUPROJECTS"
   WritePrivateProfileString( wszSectionName, null, null, _ConfigFilename )
   
   ' Write the new values.
   for i as long = 0 To 9
      this.MRUProject(i) = ProcessToCurdriveApp(this.MRUProject(i))
      wszKeyName = "MRUPROJECT_" & right("00" & str(i), 2) 
      wszValue   = this.MRUProject(i) 
      WritePrivateProfileString( wszSectionName, wszKeyName, wszValue, _ConfigFilename )
   next         
    
   function = 0
end function
   

''
''  LOAD SNIPPETS FROM DISK FILE
''
function clsConfig.LoadSnippets() as long

   dim as CWSTR wst
   dim as long ub     

   if AfxFileExists(_SnippetsFilename) = false then 
      ' The snippets file does not exist. Try copying the "default" snippets over
      ' to the main file that the user will modify.
      if AfxFileExists(_SnippetsDefaultFilename) then 
         if AfxCopyFile( _SnippetsDefaultFilename.sptr, _SnippetsFilename.sptr, true ) = false then
            exit function
         end if
      end if
   end if
   if AfxFileExists(_SnippetsFilename) = false then exit function
   
    
   dim pStream as CTextStream   '(utf16)
   if pStream.OpenUnicode(_SnippetsFilename) <> S_OK then return true   ' error

   
   do until pStream.EOS
      wst = pStream.ReadLine

      if len(wst) = 0 then continue do
      
      if left(wst, 15) = "' SNIPPET START" then  
         ub = ubound(this.Snippets) 
         redim preserve this.Snippets(ub + 1) 

      elseif left(wst, 14) = "' DESCRIPTION=" then  
         ub = ubound(this.Snippets) 
         this.Snippets(ub).wszDescription = mid(wst, 15)
         
      elseif left(wst, 10) = "' TRIGGER=" then  
         ub = ubound(this.Snippets) 
         this.Snippets(ub).wszTrigger = mid(wst, 11)
      
      elseif left(wst, 7) = "' CODE=" then  
         ub = ubound(this.Snippets) 
         this.Snippets(ub).wszCode = AfxStrReplace( mid(wst, 8), "\LF", vbcrlf )
         
      elseif left(wst, 13) = "' SNIPPET END" then  
      end if

   loop
   pStream.Close

   function = 0
end function


''
''  SAVE SNIPPETS TO DISK FILE
''
function clsConfig.SaveSnippets() as long

   dim pStream as CTextStream   '(utf16)
   if pStream.Create(_SnippetsFilename, true, true) <> S_OK then return true   ' error
   
   for i as long = lbound(this.Snippets) to ubound(this.Snippets)
      pStream.WriteLine "' SNIPPET START" 
      pStream.WriteLine "' DESCRIPTION=" & this.Snippets(i).wszDescription 
      pStream.WriteLine "' TRIGGER="     & this.Snippets(i).wszTrigger
      pStream.WriteLine "' CODE=" & AfxStrReplace( this.Snippets(i).wszCode, vbcrlf, "\LF" )
      pStream.WriteLine "' SNIPPET END" 
   next

   pStream.Close 

   function = 0
end function

   
''
''  Set default values for Explorer Categories should none exist
''
function clsConfig.SetCategoryDefaults() as long
   if ubound(this.Cat) = -1 then
      redim this.Cat(CATINDEX_NORMAL)

      this.Cat(CATINDEX_FILES).idFileType = FILETYPE_UNDEFINED
      this.Cat(CATINDEX_FILES).wszDescription = L(2, "File")
      
      this.Cat(CATINDEX_MAIN).idFileType = FILETYPE_MAIN
      this.Cat(CATINDEX_MAIN).wszDescription = L(212, "Main")
      
      this.Cat(CATINDEX_RESOURCE).idFileType = FILETYPE_RESOURCE
      this.Cat(CATINDEX_RESOURCE).wszDescription = L(213, "Resource")
      
      this.Cat(CATINDEX_HEADER).idFileType = FILETYPE_HEADER
      this.Cat(CATINDEX_HEADER).wszDescription = L(175, "Header")
      
      this.Cat(CATINDEX_MODULE).idFileType = FILETYPE_MODULE
      this.Cat(CATINDEX_MODULE).wszDescription = L(211, "Module")
      
      this.Cat(CATINDEX_NORMAL).idFileType = FILETYPE_NORMAL
      this.Cat(CATINDEX_NORMAL).wszDescription = L(210, "Normal")
   end if

   function = 0
end function


''
''  LOAD CONFIGURATION FROM DISK FILE
''
function clsConfig.LoadConfigFile() as long

   dim as CWSTR wst, wKey, wData, wData2
   dim nData  as long  
   dim i      as long
   dim bReadingNote as boolean
   dim bNewConfigFile as boolean

   ' Remove any existing Builds
   ' We can not recreate the Categories each time the config reloads because the
   ' structure contains references to existing hNodeExplorer nodes.
   erase this.Builds
   
   if AfxFileExists(_ConfigFilename) = 0 then
      bNewConfigFile = true
   else
      dim pStream as CTextStream   '(utf16)
      if pStream.OpenUnicode(_ConfigFilename) <> S_OK then return true   ' error

      gApp.NonProjectNotes = ""
      
      do until pStream.EOS
         wst = pStream.ReadLine

         if len(wst) = 0 then continue do
         if left(wst, 1) = "'" then continue do
         if left(wst, 1) = "[" then continue do
         
         if left(wst, 11) = "NOTES-START" then  
            bReadingNote = true
            continue do
         end if   
         if left(wst, 9) = "NOTES-END" then  
            bReadingNote = false
            continue do
         end if
         if bReadingNote then
            gApp.NonProjectNotes = gApp.NonProjectNotes + wst + vbcrlf
            continue do
         end IF

         i = instr(wst, "=")
         if i = 0 then continue do
         
         wKey = "": wData = "": nData = 0
         
         wKey  = left(wst, i-1)
         wData = mid(wst.wstr, i+1)
         nData = val(wData)
         

         if left(wKey, 9) = "CATEGORY_" then  ' is this a CATEGORY entry
            i = val(right(wKey,2))
            if i > ubound(this.Cat) then
               redim preserve this.Cat(i) 
               this.Cat(i).idFileType     = AfxStrParse(wData, 1, "|-|")
               this.Cat(i).wszDescription = AfxStrParse(wData, 2, "|-|")
               continue do
            end if
         end if
         
            
         if left(wKey, 9) = "USERTOOL_" then  ' is this a User Tool entry
            i = val(right(wKey,2))
            if i > ubound(this.Tools) then
               redim preserve this.Tools(i)
               this.Tools(i).wszDescription   = AfxStrParse(wData, 1, "|-|")
               this.Tools(i).wszCommand       = ProcessFromCurdriveApp( AfxStrParse(wData, 2, "|-|") ) 
               this.Tools(i).wszParameters    = AfxStrParse(wData, 3, "|-|")
               this.Tools(i).wszKey           = AfxStrParse(wData, 4, "|-|")
               this.Tools(i).wszWorkingFolder = ProcessFromCurdriveApp( AfxStrParse(wData, 5, "|-|") ) 
               this.Tools(i).IsCtrl           = val(AfxStrParse(wData, 6, "|-|"))
               this.Tools(i).IsAlt            = val(AfxStrParse(wData, 7, "|-|"))
               this.Tools(i).IsShift          = val(AfxStrParse(wData, 8, "|-|"))
               this.Tools(i).IsPromptRun      = val(AfxStrParse(wData, 9, "|-|"))
               this.Tools(i).IsMinimized      = val(AfxStrParse(wData, 10, "|-|"))
               this.Tools(i).IsWaitFinish     = val(AfxStrParse(wData, 11, "|-|"))
               this.Tools(i).IsDisplayMenu    = val(AfxStrParse(wData, 12, "|-|"))
               this.Tools(i).Action           = val(AfxStrParse(wData, 13, "|-|"))
               continue do
            end If
         end If

         if left(wKey, 6) = "BUILD_" then  ' is this a BUILD entry
            i = val(right(wKey,2))
            if i > ubound(this.Builds) then
               redim preserve this.Builds(i) 
               this.Builds(i).id             = AfxStrParse(wData, 1, "|-|")
               this.Builds(i).wszDescription = AfxStrParse(wData, 2, "|-|")
               this.Builds(i).wszOptions     = AfxStrParse(wData, 3, "|-|")
               this.Builds(i).IsDefault      = val(AfxStrParse(wData, 4, "|-|"))
               this.Builds(i).Is32bit        = val(AfxStrParse(wData, 5, "|-|"))
               this.Builds(i).Is64bit        = val(AfxStrParse(wData, 6, "|-|"))
               continue do
            end If
         end If

         if left(wKey, 4) = "MRU_" then  ' is this an MRU entry
            i = val(right(wKey,2))
            if (i >= 0) And (i <= 9) then 
               this.MRU(i) = ProcessFromCurdriveApp(wData)
               continue do
            end If
         end If
         
         if left(wKey, 11) = "MRUPROJECT_" then  ' is this an MRU Project entry
            i = val(right(wKey,2))
            if (i >= 0) And (i <= 9) then 
               this.MRUProject(i) = ProcessFromCurdriveApp(wData)
               continue do
            end If
         end If

         
         select case wKey
            case "WinFBEversion":          this.WinFBEversion         = wData
            case "AskExit":                this.AskExit               = nData
            case "EnableProjectCache":     this.EnableProjectCache    = nData
            case "CheckForUpdates":        this.CheckForUpdates       = nData
            case "LastUpdateCheck":        this.LastUpdateCheck       = nData
            case "AutoSaveFiles":          this.AutoSaveFiles         = nData
            case "AutoSaveInterval":       this.AutoSaveInterval      = nData
            case "RestoreSession":         this.RestoreSession        = nData
            case "LastActiveSessionFile":  this.wszLastActiveSession  = ProcessFromCurdriveApp(wData)
            case "MultipleInstances":      this.MultipleInstances     = nData
            case "CompileAutosave":        this.CompileAutosave       = nData
            case "CloseFuncList":          this.CloseFuncList         = nData
            case "SyntaxHighlighting":     this.SyntaxHighlighting    = nData
            case "Codetips":               this.Codetips              = nData
            case "AutoComplete":           this.AutoComplete          = nData
            case "CharacterAutoComplete":  this.CharacterAutoComplete = nData
            case "RightEdge":              this.RightEdge             = nData
            case "RightEdgePosition":      this.RightEdgePosition     = wData
            case "LeftMargin":             this.LeftMargin            = nData
            case "FoldMargin":             this.FoldMargin            = nData
            case "AutoIndentation":        this.AutoIndentation       = nData
            case "FornextVariable":        this.FornextVariable       = nData
            case "ConfineCaret":           this.ConfineCaret          = nData
            case "LineNumbering":          this.LineNumbering         = nData
            case "HighlightCurrentLine":   this.HighlightCurrentLine  = nData
            case "IndentGuides":           this.IndentGuides          = nData
            case "TabIndentSpaces":        this.TabIndentSpaces       = nData
            case "PositionMiddle":         this.PositionMiddle        = nData
            case "BraceHighlight":         this.BraceHighlight        = nData
            case "OccurrenceHighlight":    this.OccurrenceHighlight   = nData
            case "LocalizationFile":       this.LocalizationFile      = ProcessFromCurdriveApp(wData)
            case "TabSize":                this.TabSize               = wData
            case "UnicodeEncoding":        this.UnicodeEncoding       = nData
            case "EditorFontname":         this.EditorFontname        = wData
            case "EditorFontsize":         this.EditorFontsize        = wData
            case "EditorFontCharSet":      this.EditorFontCharSet     = wData
            case "FontExtraSpace":         this.FontExtraSpace        = wData
            case "Theme":                  this.ThemeFilename         = wData
            case "KeywordCase":            this.KeywordCase           = nData
            case "StartupLeft":            this.StartupLeft           = nData
            case "StartupTop":             this.StartupTop            = nData
            case "StartupRight":           this.StartupRight          = nData
            case "StartupBottom":          this.StartupBottom         = nData
            case "StartupMaximized":       this.StartupMaximized      = nData
            case "ShowPanel":              this.ShowPanel             = nData
            case "ShowPanelWidth":         this.ShowPanelWidth        = nData
            case "ToolBoxLeft":            this.ToolBoxLeft           = nData
            case "ToolBoxTop":             this.ToolBoxTop            = nData
            case "ToolBoxRight":           this.ToolBoxRight          = nData
            case "ToolBoxBottom":          this.ToolBoxBottom         = nData
            case "FBWINCompiler32":        this.FBWINCompiler32       = ProcessFromCurdriveApp(wData)
            case "FBWINCompiler64":        this.FBWINCompiler64       = ProcessFromCurdriveApp(wData)
            case "WinFBXPath":             this.WinFBXPath            = ProcessFromCurdriveApp(wData)
            case "CompilerBuild":          this.CompilerBuild         = wData
            case "CompilerSwitches":       this.CompilerSwitches      = wData
            case "CompilerHelpfile":       this.CompilerHelpfile      = ProcessFromCurdriveApp(wData)
            case "WinFBXHelpfile":         this.WinFBXHelpfile        = ProcessFromCurdriveApp(wData)
            case "RunViaCommandWindow":    this.RunViaCommandWindow   = nData
            case "DisableCompileBeep":     this.DisableCompileBeep    = nData
         end select
      
      loop
      pStream.Close
   end if
   
   
   ' Set some defaults if the config file was missing.
   ' if no Tools exist then create some default ones...
   if (ubound(this.Tools) = -1) and (bNewConfigFile = true) then
      dim wszTools(4) as CWSTR
      wszTools(0) = "ASCII Chart|-|.\Tools\asciichart32.exe|-||-|1|-||-|1|-|0|-|0|-|0|-|0|-|0|-|1|-|0"
      wszTools(1) = "GUID Generator|-|.\Tools\GUIDgen32.exe|-||-|2|-||-|1|-|0|-|0|-|0|-|0|-|0|-|1|-|0"
      wszTools(2) = "Registers|-|.\Tools\Registers.chm|-||-|3|-||-|1|-|0|-|0|-|0|-|0|-|0|-|1|-|0"
      wszTools(3) = "Arch settings|-|.\Tools\ArchSettings.exe|-||-|4|-||-|1|-|0|-|0|-|0|-|0|-|0|-|1|-|0"
      wszTools(4) = "Set compiler switches|-|.\Tools\SetCompilerSwitchesII.exe|-||-|5|-||-|1|-|0|-|0|-|0|-|0|-|0|-|1|-|0"
      for i = lbound(wszTools) to ubound(wszTools)
         if i > ubound(this.Tools) then
            redim preserve this.Tools(i)
            wData = wszTools(i)
            this.Tools(i).wszDescription   = AfxStrParse(wData, 1, "|-|")
            this.Tools(i).wszCommand       = ProcessFromCurdriveApp( AfxStrParse(wData, 2, "|-|") ) 
            this.Tools(i).wszParameters    = AfxStrParse(wData, 3, "|-|")
            this.Tools(i).wszKey           = AfxStrParse(wData, 4, "|-|")
            this.Tools(i).wszWorkingFolder = ProcessFromCurdriveApp( AfxStrParse(wData, 5, "|-|") ) 
            this.Tools(i).IsCtrl           = val(AfxStrParse(wData, 6, "|-|"))
            this.Tools(i).IsAlt            = val(AfxStrParse(wData, 7, "|-|"))
            this.Tools(i).IsShift          = val(AfxStrParse(wData, 8, "|-|"))
            this.Tools(i).IsPromptRun      = val(AfxStrParse(wData, 9, "|-|"))
            this.Tools(i).IsMinimized      = val(AfxStrParse(wData, 10, "|-|"))
            this.Tools(i).IsWaitFinish     = val(AfxStrParse(wData, 11, "|-|"))
            this.Tools(i).IsDisplayMenu    = val(AfxStrParse(wData, 12, "|-|"))
            this.Tools(i).Action           = val(AfxStrParse(wData, 13, "|-|"))
         end If
      next
   end if
      

   ' Set some defaults if the config file was missing or corrupt
   ' if no Categories exist then create some default ones... 
   ' NOTE: The default for Categories are not set here because some
   ' category names depend on the translation from the Localization
   ' file and that file is not loaded until the Config file is read.
   ' We do the setting of default Categories in gConfig.SetCategoryDefaults
   ' in WinMain after the localization is correctly loaded.
   
         
   ' Set some defaults if the config file was missing or corrupt
   ' if no builds exist then create some default ones... We do this
   ' no matter if it is a bNewConfigFile or not.
   if ubound(this.Builds) = -1 then
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

   end IF

   ' v3.0.0 if we are loading a config file that is pre-version 3.0.0 then we will
   ' blank out some values in order to allow WinFBE to use the new version 3.0.0 
   ' values. The new values will eventually then be saved in a new v3 config file.
   dim as CWSTR wszConfigVersion = AfxStrReplace(this.WinFBEversion, ".", "")
   if val(wszConfigVersion) < 300 then
      this.FBWINCompiler32 = ""
      this.FBWINCompiler64 = ""
      this.EditorFontname = ""
      this.ThemeFilename = "winfbe_default_dark.theme"
      this.EditorFontsize = 11
      this.FontExtraSpace = 10
      this.RightEdge = 0
      this.CharacterAutoComplete = 0
      this.OccurrenceHighlight = 0
   end if
   
   

   ' Attempt to fill in any missing compiler paths and help files
   dim as CWSTR wszText
   if len(this.FBWINCompiler32) = 0 then
      wszText = AfxGetExePathName & "toolchains\" & gwszDefaultToolchain & "\fbc32.exe"      
      if AfxFileExists(wszText) then this.FBWINCompiler32 = wszText
   end if
   if len(this.FBWINCompiler64) = 0 then
      wszText = AfxGetExePathName & "toolchains\" & gwszDefaultToolchain & "\fbc64.exe"      
      if AfxFileExists(wszText) then this.FBWINCompiler64 = wszText
   end if
   if len(this.CompilerHelpfile) = 0 then
      wszText = AfxGetExePathName & "Help\freebasic_manual.chm"
      if AfxFileExists(wszText) then this.CompilerHelpfile = wszText
   end if

   ' do some checks to see if the editor font exists. It is possible that the
   ' user could be accessing WinFBE from a different computer of using it under
   ' Wine where the previously selected font no longer exists.
   if doFontSanityCheck() = false then
'      print "Defined font was not found. Now using this font: ", gConfig.EditorFontname
   end if

   ' Save the last write time so that it can be checked in the message loop
   ' in order to detect external changes to it.
   _DateFileTime = AfxGetFileLastWriteTime( _ConfigFilename )

   function = 0
end function


' ========================================================================================
' Save current session (open files) to a diskfile
' ========================================================================================
function clsConfig.SaveSessionFile( byref wszSessionFile as wstring ) as boolean    
   dim as CWSTR wszText 
   dim as long nCount 

   dim pDoc as clsDocument ptr

   dim pStream as CTextStream  ' (utf16)
   if pStream.Create( wszSessionFile, true, true ) <> S_OK then return false
   
   pStream.WriteLine "' WINFBE SESSION FILE"
   
   if gApp.IsProjectActive then
      wszText = ProcessToCurdriveApp( gApp.ProjectFilename )
      pStream.WriteLine "ProjectName=" & wszText
   else
      ' Save all of the loaded tabs
      pStream.WriteLine "BuildConfig=" & this.CompilerBuild
      pStream.WriteLine "ActiveTab=" & gTTabCtl.CurSel
      pStream.WriteLine "CommandLine=" & gApp.wszCommandLine

      nCount = gTTabCtl.GetItemCount

      for i as long = 0 To nCount - 1
         if gTTabCtl.IsSafeIndex(i) = false then continue for
         pDoc = gTTabCtl.tabs(i).pDoc
         ' Only deal with files that are no longer "new"
         if pDoc->IsNewFlag = false then 
            wszText = ProcessToCurdriveApp(pDoc->DiskFilename)
            pStream.WriteLine "File=" & wszText
            pStream.WriteLine "FileType=" & pDoc->ProjectFileType
            pStream.WriteLine "TabIndex=" & -1     ' use -1 rather than true
            pStream.WriteLine "Bookmarks=" & pDoc->GetBookmarks()
            pStream.WriteLine "FoldPoints=" & pDoc->GetFoldPoints()
            pStream.WriteLine "IsDesigner=" & pDoc->IsDesigner
            pStream.WriteLine "FirstLine=" & SendMessage( pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0) 
            pStream.WriteLine "Position=" & SendMessage( pDoc->hWindow(0), SCI_GETCURRENTPOS, 0, 0)  
            pStream.WriteLine "FirstLine1=" & SendMessage( pDoc->hWindow(1), SCI_GETFIRSTVISIBLELINE, 0, 0) 
            pStream.WriteLine "Position1=" & SendMessage( pDoc->hWindow(1), SCI_GETCURRENTPOS, 0, 0)  
            pStream.WriteLine "SplitPosition=" & pDoc->SplitY
            pStream.WriteLine "FocusEdit=" & iif(pDoc->hWndActiveScintilla = pDoc->hWindow(0), 0, 1)
            pStream.WriteLine "FileEnd=[-]" 
         end if
      next
         
   end if
   pStream.Close

   this.wszLastActiveSession = wszSessionFile
   this.SaveConfigFile

   function = true   ' successful save
end function


' ========================================================================================
' Load previously saved session file from a diskfile
' ========================================================================================
function clsConfig.LoadSessionFile( byref wszSessionFile as wstring ) as boolean    

   dim pDoc as clsDocument ptr
   dim as CWSTR wst, wKey, wData, wszFilename, wszProjectName
      
   dim sBookmarks as string
   dim sFoldPoints as string
   dim as long nData, i, iTab, nActiveTab, nFocusEdit
   dim as long nFileIndex = -1
   dim as long nFirstLine, nPosition, nFirstLine1, nPosition1, nSplitPosition

   ' Info saved for each File and used to create the file once Fileend is found
   dim as long nTabIndex
   dim as CWSTR wszFileType
   dim as boolean bLoadInTab
   dim as boolean bIsDesigner
   dim as boolean bDesignerView
   
   dim pStream as CTextStream  ' (utf16)
   if AfxFileExists(wszSessionFile) = 0 then exit function
   if pStream.OpenUnicode( wszSessionFile ) <> S_OK then return false

   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )
   
   do until pStream.EOS
      wst = pStream.ReadLine
      
      if len(wst) = 0 then continue do
      if left(wst, 1) = "'" then continue do
      if left(wst, 1) = "[" then continue do
      
      i = instr(wst, "=")
      if i = 0 then continue do
      
      wKey  = left(wst, i-1)
      wData = mid(wst, i+1)
      
      select case ucase(wData)
         case "TRUE":   nData = true
         case "FALSE":  nData = false
         case Else:     nData = val(wData)
      end select

      wData = ProcessFromCurdriveApp(wData)

      select case wKey
         case "ProjectName"
            wszProjectName = wData
            if AfxFileExists(wszProjectName) then exit do
            
         case "BuildConfig":         this.CompilerBuild = wData
         case "CommandLine":         gApp.wszCommandLine = wData
         case "ActiveTab":           nActiveTab = nData
         
         case "File":                wszFilename = wData
         case "FileType":            wszFileType = wData
         
         case "IsDesigner":          bIsDesigner = nData
         case "TabIndex":            bLoadInTab  = nData
         case "Bookmarks":           sBookmarks  = str(wData)
         case "FoldPoints":          sFoldPoints = str(wData)
         case "FirstLine":           nFirstLine  = nData
         case "Position":            nPosition   = nData
         case "FirstLine1":          nFirstLine1 = nData
         case "Position1":           nPosition1  = nData
         case "SplitPosition":       nSplitPosition = nData
         case "FocusEdit":           nFocusEdit  = nData
            
         case "FileEnd":
            if AfxFileExists(wszFilename) then
               pDoc = frmMain_OpenFileSafely( _
                        HWND_FRMMAIN, _
                        false, _    ' bIsNewFile
                        false, _    ' bIsTemplate
                        false, _    ' bShowInTab (we'll manually add new tab below)
                        false, _    ' bIsInclude
                        wszFilename, _  ' wszName
                        0, _            ' pDocIn
                        bIsDesigner, _  ' bIsDesigner
                        wszFileType )
    
               iTab = gTTabCtl.AddTab( pDoc )  ' Add the new document to the top tabcontrol
               
               ' Set the saved positions
               pDoc->SplitY = nSplitPosition
               if pDoc->SplitY then pDoc->bEditorIsSplit = true
               SciExec( pDoc->hWindow(0), SCI_SETFIRSTVISIBLELINE, nFirstLine, 0) 
               SciExec( pDoc->hWindow(0), SCI_GOTOPOS, nPosition, 0) 
               SciExec( pDoc->hWindow(1), SCI_SETFIRSTVISIBLELINE, nFirstLine1, 0) 
               SciExec( pDoc->hWindow(1), SCI_GOTOPOS, nPosition1, 0) 
               pDoc->hWndActiveScintilla = pDoc->hWindow(nFocusEdit)
               pDoc->SetBookmarks(sBookmarks)
               pDoc->SetFoldPoints(sFoldPoints)
            end if
      
      end select
   
   loop                        
   pStream.Close
                     
   if AfxFileExists(wszProjectName) then 
      frmMain_OpenProjectSafely( HWND_FRMMAIN, wszProjectName ) 
      return true
   end if

   ' Load all of the filenames into the Explorer listbox. Configure the node
   ' array to allow all types of files to be shown in the listbox
   for i as long = lbound(gConfig.Cat) to ubound(gConfig.Cat)
      gConfig.Cat(i).bShow = true
   next
   LoadExplorerFiles()
   
   ' Load all of the bookmarks into the Bookmarks listbox. 
   LoadBookmarksFiles()

   ' Load all of the functions into the Function List listbox. 
   LoadFunctionsFiles()

   ' Display the active editing window
   gTTabCtl.CurSel = nActiveTab
   gTTabCtl.DisplayScintilla( nActiveTab, true )
   frmTopTabs_PositionWindows()
   
   ' Highlight the selected tab file in the Explorer listbox
   if gTTabCtl.IsSafeIndex(nActiveTab) then
      frmExplorer_SelectItemData( gTTabCtl.tabs(nActiveTab).pDoc )
   end if
   
   frmMain_SetStatusbar
   SetCursor( hCurSave )

   ' Need to iterate the loaded forms that are currently visually active and
   ' re-apply properties because loading of Image properties will fail until
   ' all the forms are loaded in the editor.
   dim as long nCount = gTTabCtl.GetItemCount 
   for i as long = 0 to nCount - 1
      ' Get the document pointer and then save file to disk
      if gTTabCtl.IsSafeIndex(i) = false then continue for
      pDoc = gTTabCtl.tabs(i).pDoc
      if pDoc then
         if pDoc->IsDesigner then
            ' Apply all control properties 
            for ii as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
               dim pCtrl as clsControl ptr = pDoc->Controls.ItemAt(ii)
               if pCtrl then ApplyControlProperties(pDoc, pCtrl)
            next         
         end if
      end if
   next

   this.wszLastActiveSession = wszSessionFile
   this.SaveConfigFile
   
   frmMain_PositionWindows
   frmMain_SetFocusToCurrentCodeWindow
   
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   if pDoc then 
      if pDoc->bEditorIsSplit then AfxRedrawWindow( HWND_FRMMAIN )
   end if
   
   function = true   ' successful open
end function


' ========================================================================================
' Save all options for the Project to a diskfile
' ========================================================================================
function clsConfig.ProjectSaveToFile() as boolean    

   dim as CWSTR cwzRelative, wszText 
   dim as long nCount 
   dim as long i

   dim pDoc as clsDocument ptr

   dim pStream as CTextStream  ' (utf16)
   if pStream.Create(gApp.ProjectFilename, true, true) <> S_OK then return false
   
   pStream.WriteLine "' WINFBE PROJECT FILE"
   pStream.WriteLine "ProjectBuild="       & gApp.ProjectBuild
   pStream.WriteLine "ProjectOther32="     & gApp.ProjectOther32
   pStream.WriteLine "ProjectOther64="     & gApp.ProjectOther64
   pStream.WriteLine "ProjectCommandLine=" & gApp.ProjectCommandLine
   pStream.WriteLine "ProjectDefaultFont=" & gApp.ProjectDefaultFont

   ' Save all of the loaded tabs first
   pStream.WriteLine "ActiveTab=" & gTTabCtl.CurSel

   nCount = gTTabCtl.GetItemCount

   for i = 0 To nCount - 1
      if gTTabCtl.IsSafeIndex(i) = false then continue for
      pDoc = gTTabCtl.tabs(i).pDoc
      ' Only deal with files that are no longer "new"
      if pDoc->IsNewFlag = false then 
         wszText = pDoc->DiskFilename
         cwzRelative = AfxPathRelativePathTo( gApp.ProjectFilename, FILE_ATTRIBUTE_NORMAL, wszText, FILE_ATTRIBUTE_NORMAL)
         if len(cwzRelative) then
            if AfxPathIsRelative(cwzRelative) then wszText = cwzRelative
         end if   
         wszText = ProcessToCurdriveProject(wszText)
         pStream.WriteLine "File=" & wszText
         pStream.WriteLine "FileType=" & pDoc->ProjectFileType
         pStream.WriteLine "TabIndex=" & -1     ' use -1 rather than true
         pStream.WriteLine "Bookmarks=" & pDoc->GetBookmarks()
         pStream.WriteLine "FoldPoints=" & pDoc->GetFoldPoints()
         pStream.WriteLine "IsDesigner=" & pDoc->IsDesigner
         pStream.WriteLine "FirstLine=" & SendMessage( pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0) 
         pStream.WriteLine "Position=" & SendMessage( pDoc->hWindow(0), SCI_GETCURRENTPOS, 0, 0)  
         pStream.WriteLine "FirstLine1=" & SendMessage( pDoc->hWindow(1), SCI_GETFIRSTVISIBLELINE, 0, 0) 
         pStream.WriteLine "Position1=" & SendMessage( pDoc->hWindow(1), SCI_GETCURRENTPOS, 0, 0)  
         pStream.WriteLine "SplitPosition=" & pDoc->SplitY
         pStream.WriteLine "FocusEdit=" & iif(pDoc->hWndActiveScintilla = pDoc->hWindow(0), 0, 1)
         pStream.WriteLine "FileEnd=[-]" 
      end if
   next
         
   ' Save all other non-displayed documents for the project

   pDoc = gApp.pDocList
   do until pDoc = 0
      ' Only deal with files that are no longer "new"
      if pDoc->IsNewFlag = false then 
         if gTTabCtl.GetTabIndexByDocumentPtr(pDoc) = -1 then
            wszText = pDoc->DiskFilename
            cwzRelative = AfxPathRelativePathTo( gApp.ProjectFilename, FILE_ATTRIBUTE_NORMAL, wszText, FILE_ATTRIBUTE_NORMAL)
            if AfxPathIsRelative(cwzRelative) then wszText = cwzRelative
            wszText = ProcessToCurdriveProject(wszText)
            pStream.WriteLine "File=" & wszText
            pStream.WriteLine "FileType=" & pDoc->ProjectFileType
            pStream.WriteLine "TabIndex=" & 0    
            pStream.WriteLine "Bookmarks=" & pDoc->GetBookmarks()
            pStream.WriteLine "FoldPoints=" & pDoc->GetFoldPoints()
            pStream.WriteLine "IsDesigner=" & pDoc->IsDesigner
            pStream.WriteLine "FirstLine=" & SendMessage( pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0) 
            pStream.WriteLine "Position=" & SendMessage( pDoc->hWindow(0), SCI_GETCURRENTPOS, 0, 0)  
            pStream.WriteLine "FirstLine1=" & SendMessage( pDoc->hWindow(1), SCI_GETFIRSTVISIBLELINE, 0, 0) 
            pStream.WriteLine "Position1=" & SendMessage( pDoc->hWindow(1), SCI_GETCURRENTPOS, 0, 0)  
            pStream.WriteLine "SplitPosition=" & pDoc->SplitY
            pStream.WriteLine "FocusEdit=" & iif(pDoc->hWndActiveScintilla = pDoc->hWindow(0), 0, 1)
            pStream.WriteLine "FileEnd=[-]"
         end If
      end if
      pDoc = pDoc->pDocnext
   loop

   ' if project is active then save NOTES to config file.
   if gApp.IsProjectActive then
      dim wszText as CWSTR = wstr("NOTES-START") + vbcrlf + _
                             gApp.ProjectNotes + vbcrlf + _
                             wstr("NOTES-END") + vbcrlf
      pStream.WriteLine ""
      pStream.WriteLine "[Notes]"
      pStream.WriteLine wszText                       
   end IF

   pStream.Close

   function = true   ' successful save
end function


' ========================================================================================
' Load all options for the Project from a diskfile
' ========================================================================================
function clsConfig.ProjectLoadFromFile( byval wszFile as CWSTR ) as boolean    

   dim pDoc as clsDocument ptr
   dim as CWSTR wst, wKey, wData, wszFilename
      
   dim sBookmarks as string
   dim sFoldPoints as string
   dim as long nData, i, iTab, nActiveTab, nFocusEdit
   dim as long nFileIndex = -1
   dim as long nFirstLine, nPosition, nFirstLine1, nPosition1, nSplitPosition

   ' Info saved for each File and used to create the file once Fileend is found
   dim as long nTabIndex
   dim as CWSTR wszFileType
   dim as boolean bLoadInTab
   dim as boolean bReadingNote
   dim as boolean bIsDesigner
   dim as boolean bDesignerView
   
   if AfxFileExists(wszFile) = 0 then exit function

   dim as HCURSOR hCurSave = GetCursor()
   SetCursor( LoadCursor(0, IDC_WAIT) )

   gApp.ProjectFilename = wszFile
   gApp.ProjectName = AfxStrPathname( "NAMEX", gApp.ProjectFilename )
   gApp.ProjectNotes = ""
   
   ' As of version 3.02 we no longer use the Project Fast Cache feature so
   ' delete any old cache disk file that will just now be taking up space.
   dim as CWSTR ProjectCacheFilename = _
                  AfxStrPathname("PATH", gApp.ProjectFilename) & _
                  AfxStrPathname("NAME", gApp.ProjectFilename) & _
                  "_db_data.ini"
   AfxDeleteFile( ProjectCacheFilename )
   

   dim pStream as CTextStream  ' (utf16)
   if pStream.OpenUnicode(gApp.ProjectFilename) <> S_OK then return false

   ' Variable length array to hold sequence of all project files
   dim docData(any) as PROJECT_FILELOAD_DATA
   
   gApp.IsProjectLoading = true 
         
   gApp.FileLoadingCount = 0
   
   do until pStream.EOS
      wst = pStream.ReadLine
      
      if len(wst) = 0 then continue do
      if left(wst, 1) = "'" then continue do
      if left(wst, 1) = "[" then continue do
      
      if left(wst, 11) = "NOTES-START" then  
         bReadingNote = true
         continue do
      end if   
      if left(wst, 9) = "NOTES-END" then  
         bReadingNote = false
         continue do
      end if
      if bReadingNote then
         gApp.ProjectNotes = gApp.ProjectNotes + wst + vbcrlf
         continue do
      end IF

      i = instr(wst, "=")
      if i = 0 then continue do
      
      wKey  = left(wst, i-1)
      wData = mid(wst, i+1)
      
      select case ucase(wData)
         case "TRUE":   nData = true
         case "FALSE":  nData = false
         case Else:     nData = val(wData)
      end select

      wData = ProcessFromCurdriveProject(wData)

      select case wKey
         case "ProjectBuild":        gApp.ProjectBuild       = wData
         case "ProjectOther32":      gApp.ProjectOther32     = wData
         case "ProjectOther64":      gApp.ProjectOther64     = wData
         case "ProjectCommandLine":  gApp.ProjectCommandLine = wData
         case "ProjectDefaultFont":  gApp.ProjectDefaultFont = wData
         case "ActiveTab":           nActiveTab = nData
         case "File":                wszFilename = wData
         case "FileType":            wszFileType  = wData
         case "Bookmarks":           sBookmarks  = str(wData)
         case "FoldPoints":          sFoldPoints = str(wData)
         case "IsDesigner":          bIsDesigner = nData
         case "TabIndex":            bLoadInTab  = nData
         case "FirstLine":           nFirstLine  = nData
         case "Position":            nPosition   = nData
         case "FirstLine1":          nFirstLine1 = nData
         case "Position1":           nPosition1  = nData
         case "SplitPosition":       nSplitPosition = nData
         case "FocusEdit":           nFocusEdit  = nData
            
         case "FileEnd":
            ' if this is a relative filename then convert it back.
            if AfxPathIsRelative(wszFilename) then 
               wszFilename = AfxPathCombine( AfxStrPathName("PATH", gApp.ProjectFilename), wszFilename)
            end if
            if AfxFileExists(wszFilename) then

               wszFilename = OnCommand_FileAutoSaveFileCheck( wszFilename )

               nFileIndex = nFileIndex + 1
               if nFileIndex > ubound(docData) then
                  redim preserve docData( ubound(docData) + 100 ) 
               end If

               with docData(nFileIndex)
                  .wszFilename    = wszFilename
                  .wszFileType    = wszFileType
                  .bLoadInTab     = bLoadInTab
                  .wszBookmarks   = sBookmarks
                  .wszFoldPoints  = sFoldPoints
                  .IsDesigner     = bIsDesigner
                  .nFirstLine     = nFirstLine
                  .nPosition      = nPosition
                  .nFirstLine1    = nFirstLine1
                  .nPosition1     = nPosition1
                  .nSplitPosition = nSplitPosition
                  .nFocusEdit     = nFocusEdit
               end with

               nFirstLine  = 0: nPosition  = 0
               nFirstLine1 = 0: nPosition1 = 0
               nFocusEdit  = 0: nSplitPosition = 0
               sBookmarks  = ""
               sFoldPoints = ""
               bLoadInTab    = false
               bIsDesigner   = false
               bDesignerView = false
            end if
            
      end select
   
   loop                        
   pStream.Close
   

   ' Load/Display all of the project files
   for i as long = lbound(docData) to ubound(docData)
   
      if AfxFileExists( docData(i).wszFilename ) = false then continue for
      
      gApp.FileLoadingCount = gApp.FileLoadingCount + 1
      gApp.wszPanelText = docData(i).wszFilename
      frmMain_SetStatusbar

      pDoc = frmMain_OpenFileSafely( _
                  HWND_FRMMAIN, _
                  false, _    ' bIsNewFile
                  false, _    ' bIsTemplate
                  false, _    ' bShowInTab (we'll manually add new tab below)
                  false, _    ' bIsInclude
                  docData(i).wszFilename, _  ' wszName
                  0, _        ' pDocIn
                  docData(i).IsDesigner, _    ' bIsDesigner
                  docData(i).wszFileType )

      if docData(i).bLoadInTab then
         iTab = gTTabCtl.AddTab( pDoc )  ' Add the new document to the top tabcontrol
      end if

      ' Set the saved positions
      pDoc->SplitY = docData(i).nSplitPosition
      if pDoc->SplitY then pDoc->bEditorIsSplit = true
      SciExec( pDoc->hWindow(0), SCI_SETFIRSTVISIBLELINE, docData(i).nFirstLine, 0) 
      SciExec( pDoc->hWindow(0), SCI_GOTOPOS, docData(i).nPosition, 0) 
      SciExec( pDoc->hWindow(1), SCI_SETFIRSTVISIBLELINE, docData(i).nFirstLine1, 0) 
      SciExec( pDoc->hWindow(1), SCI_GOTOPOS, docData(i).nPosition1, 0) 
      pDoc->hWndActiveScintilla = pDoc->hWindow(docData(i).nFocusEdit)
      pDoc->SetBookmarks(docData(i).wszBookmarks)
      pDoc->SetFoldPoints(docData(i).wszFoldPoints)
   next
                  
   gApp.IsProjectLoading = false
   gApp.IsProjectActive = true
   gApp.wszPanelText = ""   ' reset filename parsing text that displays in StatusBar panel
   gApp.hIconPanel = 0
   gApp.wszLastOpenFolder = ""
   
   ' Load all of the filenames into the Explorer listbox. Configure the node
   ' array to allow all types of files to be shown in the listbox
   for i as long = lbound(gConfig.Cat) to ubound(gConfig.Cat)
      gConfig.Cat(i).bShow = true
   next

   LoadExplorerFiles()

   ' Load all of the bookmarks into the Bookmarks listbox. 
   LoadBookmarksFiles()

   ' Load all of the functions into the Function List listbox. 
   LoadFunctionsFiles()
   
   ' Display the active editing window
   if gTTabCtl.IsSafeIndex(nActiveTab) = false then 
      ' active tab as stored in the config file is not valid. Could be that the
      ' filename no longer exists of trouble loading that particular project file.
      ' Try to use an alternate tab number
      nActiveTab = gTTabCtl.GetItemCount - 1
   end if
   gTTabCtl.SetFocusTab(nActiveTab)      

   ' if no active build configuration then assign the currently active selection
   if gApp.ProjectBuild = "" then 
      gApp.ProjectBuild = frmBuildConfig_GetSelectedBuildGUID()
   end if   

   frmMain_SetStatusbar
   SetCursor( hCurSave )

   ' Show the correct notes and TODO data for this project
   frmOutput_ShowNotes
   frmOutput_UpdateToDoListview

   ' Need to iterate the loaded forms that are currently visually active and
   ' re-apply properties because loading of Image properties will fail until
   ' all the forms are loaded in the editor.
   dim as long nCount = gTTabCtl.GetItemCount 
   for i as long = 0 to nCount - 1
      ' Get the document pointer and then save file to disk
      if gTTabCtl.IsSafeIndex(i) = false then continue for
      pDoc = gTTabCtl.tabs(i).pDoc
      if pDoc then
         if pDoc->IsDesigner then
            ' Apply all control properties 
            for ii as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
               dim pCtrl as clsControl ptr = pDoc->Controls.ItemAt(ii)
               if pCtrl then 
                  ApplyControlProperties(pDoc, pCtrl)
               end if   
            next         
         end if
      end if
   next

   frmMain_PositionWindows
   frmMain_SetFocusToCurrentCodeWindow

   function = true   ' successful open
end function


' ========================================================================================
' Reload the config file should it have been changed by an external program.
' ========================================================================================
function clsConfig.ReloadConfigFileTest() as boolean    

   if gApp.IsShutdown = true then return false
   
   dim as FILETIME ft
   
   ' Compare the disk file date time to the value currently stored in static variable.
   ft = AfxGetFileLastWriteTime(_ConfigFilename)
   if AfxFileTimeToVariantTime(ft) <> AfxFileTimeToVariantTime(_DateFileTime) then
      gConfig.LoadConfigFile()  ' this function also updates the _DateFileTime value
      ' Apply the newly saved options to all open Scintilla windows
      dim pDoc as clsDocument ptr = gApp.pDocList
      do until pDoc = 0
         pDoc->ApplyProperties
         pDoc = pDoc->pDocnext
      loop   
   end if
   
   function = 0
end function
