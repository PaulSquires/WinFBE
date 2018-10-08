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
   dim  As WString * MAX_PATH MainAppPath	
   GetModuleFileNameW(NULL,@MainAppPath,MAX_PATH) ' Get MainAppPath
   with this
      .MainAppPath = AfxStrPathName ("PATH", MainAppPath)
      .MainHelpPath = .MainAppPath & "Help\"   
   end with
   _ConfigFilename     = AfxGetExePathName & "Settings\WinFBE.ini" 
   _FBKeywordsFilename = AfxGetExePathName & "Settings\freebasic_keywords.txt" 
   _FBCodetipsFilename = AfxGetExePathName & "Settings\codetips.ini" 
   _WinAPICodetipsFilename = AfxGetExePathName & "Settings\codetips_winapi.ini" 
   _WinFormsXCodetipsFilename = AfxGetExePathName & "Settings\codetips_winformsx.ini" 
   _WinFBXCodetipsFilename = AfxGetExePathName & "Settings\codetips_winfbx.ini" 
End Constructor

''
''
Function clsConfig.LoadKeywords() as Long
   dim pStream AS CTextStream
   if pStream.Open(_FBKeywordsFilename) <> S_OK then exit function
   dim as CWSTR wst = pStream.ReadAll
   this.FBKeywords = AfxStrShrink(wst, " " & wchr(13,10))
   pStream.Close
   Function = 0
End Function

''
''
Function clsConfig.SaveKeywords() As Long
   dim pStream AS CTextStream
   if pStream.Create(_FBKeywordsFilename) <> S_OK then exit function
   pStream.WriteLine this.FBKeywords
   pStream.Close
   Function = 0
End Function

  
''
''
Function clsConfig.LoadCodetipsFB() as boolean

   Dim as CWSTR wst
   Dim as long i
   
   dim parser as clsParser

   parser.fileName = ""     ' no need to store this sFilename
   parser.idxProject = -1   ' do not associate it with any project
   
   dim pStream AS CTextStream
   if pStream.Open(_FBCodetipsFilename) <> S_OK then return true  ' error
   gdb2.dbRewind4PublicProject()
   do until pStream.EOS
      wst = pStream.ReadLine
      wst = Trim(wst)
      If Len(wst) = 0 Then Continue Do
      
      i = Instr(wst, "=")
      If i Then
         parser.funcName   = Left(wst, i-1)
         parser.funcParams = Mid(**wst, i+1)
         gdb2.dbAdd4DB(parser, DB2_FUNCTION)
      End If   
   Loop
   pStream.Close

   return false  ' no error
End Function


''
''
Function clsConfig.LoadCodetipsGeneric( byref wszFilename as wstring, byval IsWinAPI as boolean ) as boolean
   Dim as CWSTR wst, wsID, wsTypeName, wsTypeElements, wsParts, wsPart, wsBaseType
   Dim as long i, nCount
   
   dim parser as clsParser

   parser.fileName = ""     ' no need to store this sFilename
   parser.idxProject = -1   ' do not associate it with any project
   parser.IsWinApi = IsWinAPI

   ''
   ''  Codetips definition file for WinAPI/WinFBX/WinFormsX
   ''
   ''  Format:
   ''  1st character:  F=FUNCTION, T=TYPE, S=STANDARD_DATATYPE
   ''  2nd character:  colon.
   ''  TYPE elements are separated by | pipe symbol. Each element name/datatype is separated by comma.
   ''  STANDARD datatypes are separated by | pipe symobol (datatype followed by description of the datatype.
   ''
   
   dim pStream AS CTextStream
   if pStream.Open(wszFilename) <> S_OK then return true   ' error
   gdb2.dbRewind4PublicProject()
   
   do until pStream.EOS
      wst = pStream.ReadLine

      wst = Trim(wst)
      If Len(wst) = 0 Then Continue Do
      
      i = Instr(wst, ":")
      If i = 0 Then continue do
      
      wsID = left(wst, i - 1)
      wst = mid(**wst, i + 1)
                
      select CASE **wsID
         case "S"    ' standard datatype
            i = instr(wst, "|")
            parser.typeName = Left(wst, i-1)
            parser.Description = mid(**wst, i+1)
            gdb2.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
            
         CASE "F"    ' function
            i = instr(wst, "(")
            parser.funcName   = trim(Left(wst, i-1))
            parser.funcParams = wst
            gdb2.dbAdd4DB(parser, DB2_FUNCTION)

         case "T"     ' type
            i = instr(wst, "|")
            parser.typeName = Left(wst, i-1)
            parser.typeAlias = parser.typeName
            wsParts = mid(**wst, i + 1)
            
            if IsWinAPI = false then
               ' Check for the WFXBASE type. We save that string and add it to subsequent
               ' types for the visual designer.
               select case ucase(parser.typeName)
                  CASE "WFXBASE" 
                     wsBaseType = wsParts
                     continue do
                  case "WFXFORM", "WFXBUTTON", "WFXLABEL", "WFXTEXTBOX", "WFXLISTBOX", _
                       "WFXCHECKBOX"
                     if len(wsBaseType) then wsParts = wsParts & "|" & wsBaseType
               END SELECT
            END IF
            gdb2.dbAdd4DB(parser, DB2_TYPE)

            ' Each element of the TYPE structure is added to the database. The
            ' element is in the format elementname,elementtype
            parser.funcName = ""
            nCount = AfxStrParseCount(wsParts, "|")
            for i = 1 to nCount
               wsPart = AfxStrParse(wsParts, i, "|")
               parser.varName = AfxStrParse(wsPart, 1, ",") 
               parser.varType = AfxStrParse(wsPart, 2, ",") 
               parser.bInTypePublic = true
               gdb2.dbAdd4DB(parser, DB2_VARIABLE)
            NEXT

      END SELECT
         
   Loop
   pStream.Close
   
   return false  ' no error
End Function
    
''
''
Function clsConfig.LoadCodetipsWinFormsX() as boolean
   function = this.LoadCodetipsGeneric(_WinFormsXCodetipsFilename, false)
   dim parser as clsParser
   gdb2.dbRewind4PublicProject()
   parser.fileName = ""     ' no need to store this sFilename
   parser.idxProject = -1   ' do not associate it with any project
   parser.IsWinApi = false
   parser.typeName = "Application"
   parser.typeAlias = "wfxApplication"
   gdb2.dbAdd4DB(parser, DB2_TYPE)
   parser.typeName = "Colors"
   parser.typeAlias = "wfxColors"
   gdb2.dbAdd4DB(parser, DB2_TYPE)
end function

''
''
Function clsConfig.LoadCodetipsWinFBX() as boolean
   function = this.LoadCodetipsGeneric(_WinFBXCodetipsFilename, false)
end function

''
''
Function clsConfig.LoadCodetipsWinAPI() as boolean
   function = this.LoadCodetipsGeneric(_WinAPICodetipsFilename, true)
end function


''
''  INITIALIZE THE TOOLBOX CONTROLS
''
Function clsConfig.InitializeToolBox() as Long
   
   redim gToolBox(7) as TOOLBOX_TYPE
   
   gToolBox(0).nToolType = CTRL_POINTER
   gToolBox(0).wszToolBoxName   = "Pointer"
   gToolBox(0).wszControlName   = "Pointer"
   gToolBox(0).wszImage  = "IMAGE_POINTER"
   gToolBox(0).wszCursor = ""
   gToolBox(0).wszClassName = ""
   
   gToolBox(1).nToolType = CTRL_BUTTON
   gToolBox(1).wszToolBoxName   = "Button"
   gToolBox(1).wszControlName   = "Button"
   gToolBox(1).wszImage  = "IMAGE_BUTTON"
   gToolBox(1).wszCursor = "IMAGE_CURSOR_BUTTON"
   gToolBox(1).wszClassName = "BUTTON"

   gToolBox(2).nToolType = CTRL_CHECKBOX
   gToolBox(2).wszToolBoxName   = "CheckBox"
   gToolBox(2).wszControlName   = "Check"
   gToolBox(2).wszImage  = "IMAGE_CHECKBOX"
   gToolBox(2).wszCursor = "IMAGE_CURSOR_CHECKBOX"
   gToolBox(2).wszClassName = "CHECKBOX"

   gToolBox(3).nToolType = CTRL_FRAME
   gToolBox(3).wszToolBoxName   = "Frame"
   gToolBox(3).wszControlName   = "Frame"
   gToolBox(3).wszImage  = "IMAGE_FRAME"
   gToolBox(3).wszCursor = "IMAGE_CURSOR_FRAME"
   gToolBox(3).wszClassName = "GROUPBOX"

   gToolBox(4).nToolType = CTRL_LABEL
   gToolBox(4).wszToolBoxName   = "Label"
   gToolBox(4).wszControlName   = "Label"
   gToolBox(4).wszImage  = "IMAGE_LABEL"
   gToolBox(4).wszCursor = "IMAGE_CURSOR_LABEL"
   gToolBox(4).wszClassName = "LABEL"

   gToolBox(5).nToolType = CTRL_LISTBOX
   gToolBox(5).wszToolBoxName   = "ListBox"
   gToolBox(5).wszControlName   = "List"
   gToolBox(5).wszImage  = "IMAGE_LISTBOX"
   gToolBox(5).wszCursor = "IMAGE_CURSOR_LISTBOX"
   gToolBox(5).wszClassName = "LISTBOX"

   gToolBox(6).nToolType = CTRL_OPTION
   gToolBox(6).wszToolBoxName   = "OptionButton"
   gToolBox(6).wszControlName   = "Option"
   gToolBox(6).wszImage  = "IMAGE_OPTION"
   gToolBox(6).wszCursor = "IMAGE_CURSOR_OPTION"
   gToolBox(6).wszClassName = "RADIOBUTTON"

   gToolBox(7).nToolType = CTRL_TEXTBOX
   gToolBox(7).wszToolBoxName   = "TextBox"
   gToolBox(7).wszControlName   = "Text"
   gToolBox(7).wszImage  = "IMAGE_TEXTBOX"
   gToolBox(7).wszCursor = "IMAGE_CURSOR_TEXTBOX"
   gToolBox(7).wszClassName = "TEXTBOX"


'   gToolBox(8).nToolType = CTRL_PICTURE
'   gToolBox(8).wszName   = "Picture"
'   gToolBox(8).wszImage  = "IMAGE_PICTURE"
'   gToolBox(8).wszCursor = "IMAGE_CURSOR_PICTURE"
'   gToolBox(8).wszClassName = "STATIC"

'   gToolBox(9).nToolType = CTRL_COMBOBOX
'   gToolBox(9).wszName   = "ComboBox"
'   gToolBox(9).wszImage  = "IMAGE_COMBOBOX"
'   gToolBox(9).wszCursor = "IMAGE_CURSOR_COMBOBOX"
'   gToolBox(9).wszClassName = "COMBOBOX"

'   gToolBox(10).nToolType = CTRL_IMAGEBUTTON
'   gToolBox(10).wszName   = "ImageButton"
'   gToolBox(10).wszImage  = "IMAGE_IMGBUTTON"
'   gToolBox(10).wszCursor = "IMAGE_CURSOR_IMGBUTTON"
'   gToolBox(10).wszClassName = "BUTTON"

'   gToolBox(11).nToolType = CTRL_HSCROLL
'   gToolBox(11).wszName   = "HScroll"
'   gToolBox(11).wszImage  = "IMAGE_HSCROLL"
'   gToolBox(11).wszCursor = "IMAGE_CURSOR_HSCROLL"
'   gToolBox(11).wszClassName = "HSCROLLBAR"

'   gToolBox(12).nToolType = CTRL_VSCROLL
'   gToolBox(12).wszName   = "VScroll"
'   gToolBox(12).wszImage  = "IMAGE_VSCROLL"
'   gToolBox(12).wszCursor = "IMAGE_CURSOR_VSCROLL"
'   gToolBox(12).wszClassName = "VSCROLLBAR"

'   gToolBox(13).nToolType = CTRL_TIMER
'   gToolBox(13).wszName   = "Timer"
'   gToolBox(13).wszImage  = "IMAGE_TIMER"
'   gToolBox(13).wszCursor = "IMAGE_CURSOR_TIMER"
'   gToolBox(13).wszClassName = ""

'   gToolBox(14).nToolType = CTRL_TABCONTROL
'   gToolBox(14).wszName   = "TabControl"
'   gToolBox(14).wszImage  = "IMAGE_TABCONTROL"
'   gToolBox(14).wszCursor = "IMAGE_CURSOR_TABCONTROL"
'   gToolBox(14).wszClassName = "TABCONTROL"

'   gToolBox(15).nToolType = CTRL_RICHEDIT
'   gToolBox(15).wszName   = "RichEdit"
'   gToolBox(15).wszImage  = "IMAGE_RICHEDIT"
'   gToolBox(15).wszCursor = "IMAGE_CURSOR_RICHEDIT"
'   gToolBox(15).wszClassName = "RICHEDIT"

'   gToolBox(16).nToolType = CTRL_PROGRESSBAR
'   gToolBox(16).wszName   = "ProgressBar"
'   gToolBox(16).wszImage  = "IMAGE_PROGRESSBAR"
'   gToolBox(16).wszCursor = "IMAGE_CURSOR_PROGRESSBAR"
'   gToolBox(16).wszClassName = "PROGRESSBAR"

'   gToolBox(17).nToolType = CTRL_UPDOWN
'   gToolBox(17).wszName   = "UpDown"
'   gToolBox(17).wszImage  = "IMAGE_UPDOWN"
'   gToolBox(17).wszCursor = "IMAGE_CURSOR_UPDOWN"
'   gToolBox(17).wszClassName = "UPDOWN"

'   gToolBox(18).nToolType = CTRL_LISTVIEW
'   gToolBox(18).wszName   = "ListView"
'   gToolBox(18).wszImage  = "IMAGE_LISTVIEW"
'   gToolBox(18).wszCursor = "IMAGE_CURSOR_LISTVIEW"
'   gToolBox(18).wszClassName = "LISTVIEW"

'   gToolBox(19).nToolType = CTRL_TREEVIEW
'   gToolBox(19).wszName   = "TreeView"
'   gToolBox(19).wszImage  = "IMAGE_TREEVIEW"
'   gToolBox(19).wszCursor = "IMAGE_CURSOR_TREEVIEW"
'   gToolBox(19).wszClassName = "TREEVIEW"

'   gToolBox(20).nToolType = CTRL_SLIDER
'   gToolBox(20).wszName   = "Slider"
'   gToolBox(20).wszImage  = "IMAGE_SLIDER"
'   gToolBox(20).wszCursor = "IMAGE_CURSOR_SLIDER"
'   gToolBox(20).wszClassName = "SLIDER"

'   gToolBox(21).nToolType = CTRL_DATETIMEPICKER
'   gToolBox(21).wszName   = "DateTimePicker"
'   gToolBox(21).wszImage  = "IMAGE_DATETIMEPICKER"
'   gToolBox(21).wszCursor = "IMAGE_CURSOR_DATETIMEPICKER"
'   gToolBox(21).wszClassName = "DATETIMEPICKER"

'   gToolBox(22).nToolType = CTRL_MONTHCALENDAR
'   gToolBox(22).wszName   = "MonthCalendar"
'   gToolBox(22).wszImage  = "IMAGE_MONTHCALENDAR"
'   gToolBox(22).wszCursor = "IMAGE_CURSOR_MONTHCALENDAR"
'   gToolBox(22).wszClassName = "MONTHCALENDAR"

'   gToolBox(23).nToolType = CTRL_WEBBROWSER
'   gToolBox(23).wszName   = "WebBrowser"
'   gToolBox(23).wszImage  = "IMAGE_WEBBROWSER"
'   gToolBox(23).wszCursor = "IMAGE_CURSOR_WEBBROWSER"
'   gToolBox(23).wszClassName = ""

'   gToolBox(24).nToolType = CTRL_CUSTOM
'   gToolBox(24).wszName   = "Custom"
'   gToolBox(24).wszImage  = "IMAGE_CUSTOM"
'   gToolBox(24).wszCursor = "IMAGE_CURSOR_CUSTOM"
'   gToolBox(24).wszClassName = ""

'   gToolBox(25).nToolType = CTRL_OCX
'   gToolBox(25).wszName   = "OCX"
'   gToolBox(25).wszImage  = "IMAGE_OCX"
'   gToolBox(25).wszCursor = "IMAGE_CURSOR_OCX"
'   gToolBox(25).wszClassName = ""

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


''
''  SAVE CONFIGURATION TO DISK FILE
''
Function clsConfig.SaveConfigFile() as Long

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

   dim pStream AS CTextStream   '(utf16)
   if pStream.Create(_ConfigFilename, true, true) <> S_OK then return true   ' error
   
   pStream.WriteLine "'  WINFBE CONFIGURATION" 
   pStream.WriteLine ""
   pStream.WriteLine "WinFBEversion="         & APPVERSION
   pStream.WriteLine ""
   pStream.WriteLine "[Editor]"
   pStream.WriteLine "AskExit="               & this.AskExit
   pStream.WriteLine "HideToolbar="           & this.HideToolbar
   pStream.WriteLine "HideStatusbar="         & this.HideStatusbar
   pStream.WriteLine "MultipleInstances="     & this.MultipleInstances    
   pStream.WriteLine "CompileAutosave="       & this.CompileAutosave      
   pStream.WriteLine "CloseFuncList="         & this.CloseFuncList        
   pStream.WriteLine "HideCompile="           & this.HideCompile          
   pStream.WriteLine "SyntaxHighlighting="    & this.SyntaxHighlighting   
   pStream.WriteLine "Codetips="              & this.Codetips             
   pStream.WriteLine "AutoComplete="          & this.AutoComplete
   pStream.WriteLine "RightEdge="             & this.RightEdge
   pStream.WriteLine "RightEdgePosition="     & this.RightEdgePosition           
   pStream.WriteLine "LeftMargin="            & this.LeftMargin           
   pStream.WriteLine "FoldMargin="            & this.FoldMargin           
   pStream.WriteLine "AutoIndentation="       & this.AutoIndentation      
   pStream.WriteLine "ConfineCaret="          & this.ConfineCaret         
   pStream.WriteLine "LineNumbering="         & this.LineNumbering        
   pStream.WriteLine "HighlightCurrentLine="  & this.HighlightCurrentLine 
   pStream.WriteLine "IndentGuides="          & this.IndentGuides         
   pStream.WriteLine "TabIndentSpaces="       & this.TabIndentSpaces      
   pStream.WriteLine "UnicodeEncoding="       & this.UnicodeEncoding      
   pStream.WriteLine "KeywordCase="           & this.KeywordCase          
   pStream.WriteLine "LocalizationFile="      & this.LocalizationFile     
   pStream.WriteLine "TabSize="               & this.TabSize              
   pStream.WriteLine "EditorFontname="        & this.EditorFontname       
   pStream.WriteLine "EditorFontsize="        & this.EditorFontsize       
   pStream.WriteLine "EditorFontCharSet="     & this.EditorFontCharSet    
   pStream.WriteLine ""
   pStream.WriteLine "[Startup]" 
   pStream.WriteLine "StartupLeft="           & this.StartupLeft          
   pStream.WriteLine "StartupTop="            & this.StartupTop           
   pStream.WriteLine "StartupRight="          & this.StartupRight         
   pStream.WriteLine "StartupBottom="         & this.StartupBottom        
   pStream.WriteLine "StartupMaximized="      & this.StartupMaximized     
   pStream.WriteLine "ShowExplorer="          & this.ShowExplorer         
   pStream.WriteLine "ShowExplorerWidth="     & this.ShowExplorerWidth    
   pStream.WriteLine ""

   ' For each folder location determine if it resides on the same drive as
   ' the WinFBE application. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.
   this.FBWINCompiler32  = ProcessToCurdrive(this.FBWINCompiler32)
   this.FBWINCompiler64  = ProcessToCurdrive(this.FBWINCompiler64)
   this.CompilerHelpfile = ProcessToCurdrive(this.CompilerHelpfile)
   this.Win32APIHelpfile = ProcessToCurdrive(this.Win32APIHelpfile)
   this.WinFBXHelpfile   = ProcessToCurdrive(this.WinFBXHelpfile)

   pStream.WriteLine "[Compiler]" 
   pStream.WriteLine "FBWINCompiler32="       & this.FBWINCompiler32      
   pStream.WriteLine "FBWINCompiler64="       & this.FBWINCompiler64      
   pStream.WriteLine "CompilerSwitches="      & this.CompilerSwitches     
   pStream.WriteLine "CompilerHelpfile="      & this.CompilerHelpfile     
   pStream.WriteLine "Win32APIHelpfile="      & this.Win32APIHelpfile     
   pStream.WriteLine "WinFBXHelpfile="        & this.WinFBXHelpfile     
   pStream.WriteLine "RunViaCommandWindow="   & this.RunViaCommandWindow
   pStream.WriteLine "WinFBXPath="            & this.WinFBXPath
   

   pStream.WriteLine ""
   pStream.WriteLine "[UserTools]"  
   for i as long = lbound(this.Tools) to ubound(this.Tools)
      pStream.Write "USERTOOL_" & Right("00" & Str(i), 2) & "="
      pStream.Write this.Tools(i).wszDescription & "|-|"
      pStream.Write this.Tools(i).wszCommand & "|-|" 
      pStream.Write this.Tools(i).wszParameters & "|-|"
      pStream.Write this.Tools(i).wszKey & "|-|" 
      pStream.Write this.Tools(i).wszWorkingFolder & "|-|"
      pStream.Write this.Tools(i).IsCtrl & "|-|"
      pStream.Write this.Tools(i).IsAlt & "|-|"  
      pStream.Write this.Tools(i).IsShift & "|-|"
      pStream.Write this.Tools(i).IsPromptRun & "|-|"
      pStream.Write this.Tools(i).IsMinimized & "|-|"  
      pStream.Write this.Tools(i).IsWaitFinish & "|-|"
      pStream.Write this.Tools(i).IsDisplayMenu & "|-|"    
      pStream.WriteLine this.Tools(i).Action   
   NEXT

   pStream.WriteLine ""
   pStream.WriteLine "[Builds]"  
   for i as long = lbound(this.Builds) to ubound(this.Builds)
      pStream.Write "BUILD_" & Right("00" & Str(i), 2) & "="
      pStream.Write this.Builds(i).id & "|-|"
      pStream.Write this.Builds(i).wszDescription & "|-|"
      pStream.Write this.Builds(i).wszOptions & "|-|" 
      pStream.Write this.Builds(i).IsDefault & "|-|"  
      pStream.Write this.Builds(i).Is32bit & "|-|"  
      pStream.WriteLine this.Builds(i).Is64bit  
   NEXT

   pStream.WriteLine ""
   pStream.WriteLine "[MRU]"  
   For i As Long = 0 To 9
      this.MRU(i) = ProcessToCurdrive(this.MRU(i))
      pStream.WriteLine "MRU_" & Right("00" & Str(i), 2) & "=" & this.MRU(i) 
   Next         
   
   pStream.WriteLine ""
   pStream.WriteLine "[MRUPROJECTS]"  
   For i As Long = 0 To 9
      this.MRUProject(i) = ProcessToCurdrive(this.MRUProject(i))
      pStream.WriteLine "MRUPROJECT_" & Right("00" & Str(i), 2) & "=" & this.MRUProject(i) 
   Next         

   pStream.WriteLine ""
   pStream.WriteLine "[Themes]" 
   pStream.WriteLine "SelectedTheme=" & this.SelectedTheme
   for i as long = lbound(this.Themes) to ubound(this.Themes)
      pStream.Write "THEME_" & Right("00" & Str(i), 2) & "=" & _
                    this.Themes(i).id & "," & _
                    AfxStrReplaceAny(this.Themes(i).wszDescription, ",|", "_") 
      for ii as long = CLR_CARET to CLR_WINDOW
         pStream.Write "|" & this.Themes(i).colors(ii).nFg & "," & _
                        this.Themes(i).colors(ii).nBg & "," & _
                        this.Themes(i).colors(ii).bUseDefaultBg & "," & _
                        this.Themes(i).colors(ii).bFontBold & "," & _
                        this.Themes(i).colors(ii).bFontItalic & "," & _
                        this.Themes(i).colors(ii).bFontUnderline
      NEXT
      pStream.WriteLine ""
   NEXT

   dim hCtl as hwnd = GetDlgItem(HWND_FRMOUTPUT, IDC_FRMOUTPUT_TXTNOTES)
   dim wszText as CWSTR = wstr("NOTES-START") + vbcrlf + _
                          gApp.NonProjectNotes + vbcrlf + _
                          wstr("NOTES-END") + vbcrlf
   pStream.WriteLine ""
   pStream.WriteLine "[Notes]"
   pStream.WriteLine wszText

   pStream.Close 
   
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
Function clsConfig.LoadConfigFile() as Long

   Dim as CWSTR wst, wKey, wData, wData2
   Dim nData  As Long  
   Dim i      as Long
   dim bReadingNote as Boolean

   If AfxFileExists(_ConfigFilename) = 0 Then Exit Function
    
   dim pStream AS CTextStream   '(utf16)
   if pStream.OpenUnicode(_ConfigFilename) <> S_OK then return true   ' error

   gApp.NonProjectNotes = ""
   
   Do Until pStream.EOS
      wst = pStream.ReadLine

      If Len(wst) = 0 Then Continue Do
      If Left(wst, 1) = "'" Then Continue Do
      If Left(wst, 1) = "[" Then Continue Do
      
      If WStrNIsEqual(wst, "NOTES-START", 11) = true Then 
         bReadingNote = true
         continue do
      end if   
      If WStrNIsEqual(wst, "NOTES-END", 9) = true Then  
         bReadingNote = false
         continue do
      end if
      if bReadingNote THEN
         gApp.NonProjectNotes = gApp.NonProjectNotes + wst + vbcrlf
         continue do
      END IF

      i = InStr(wst, "=")
      If i = 0 Then Continue Do
      
      wKey = "": wData = "": nData = 0
      
      wKey  = Left(wst, i-1)
      wData = Mid(wst.wstr, i+1)
      nData = Val(wData)
      
      wData = ProcessFromCurdrive(wData)


      If WStrNIsEqual(wKey, "USERTOOL_", 9) = true Then  ' is this a User Tool entry
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

      If WStrNIsEqual(wKey, "BUILD_", 6) = true Then  ' is this a BUILD entry
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

      If WStrNIsEqual(wKey, "THEME_", 6) = true Then  ' is this a THEME entry
         this.ImportTheme(wst, false)
         continue do
      End If

      If WStrNIsEqual(wKey, "MRU_", 4) = true Then  ' is this an MRU entry
         i = Val(Right(wKey,2))
         If (i >= 0) And (i <= 9) Then 
            this.MRU(i) = wData
            Continue Do
         End If
      End If
      
      If WStrNIsEqual(wKey, "MRUPROJECT_", 11) = true Then  ' is this an MRU Project entry
         i = Val(Right(wKey,2))
         If (i >= 0) And (i <= 9) Then 
            this.MRUProject(i) = wData
            Continue Do
         End If
      End If

      
      Select Case **wKey
         Case "WinFBEversion":        this.WinFBEversion        = wData
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
         case "WinFBXPath":           this.WinFBXPath           = wData

         Case "CompilerSwitches":     
            this.CompilerSwitches     = wData
         Case "CompilerHelpfile":     this.CompilerHelpfile     = wData
         Case "Win32APIHelpfile":     this.Win32APIHelpfile     = wData
         Case "WinFBXHelpfile":       this.WinFBXHelpfile       = wData
         case "RunViaCommandWindow":  this.RunViaCommandWindow  = nData
      End Select
   
   Loop
   pStream.Close
   
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

   
   dim DefaultThemes(1) as TYPE_THEMES
   with DefaultThemes(0)
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
      .colors(CLR_WINDOW).nFg       = BGR(255,255,255)    ' white
      .colors(CLR_WINDOW).nBg       = -1
   end with

   with DefaultThemes(1)
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
      .colors(CLR_WINDOW).nFg       = BGR(030,030,030)    ' light black
      .colors(CLR_WINDOW).nBg       = -1
   end with

   ' Create a default Theme if none exist
   if ubound(this.Themes) = -1 THEN
      redim this.Themes(1) 
      this.Themes(0) = DefaultThemes(0)   ' default theme
      this.Themes(1) = DefaultThemes(1)   ' visual studio dark
      this.SelectedTheme = this.Themes(0).id
   else
      ' As of v1.7.4 we are eliminating colors for Bookmarks and WinAPI Keywords. This
      ' means that the themes have to be re-created for at least the two base themes.
      ' Any user created themes will also be affected unfortunately. I should have made
      ' it easier to add/delete colors when this was first designed but I did not.
      if this.WinFBEversion < "1.7.4" then
         for i as long = lbound(this.Themes) to ubound(this.Themes)
            select case **this.Themes(i).wszDescription
               CASE "Default Theme"
                  this.Themes(0) = DefaultThemes(0)   
               case "Visual Studio Code (Dark)"
                  this.Themes(1) = DefaultThemes(1)   
            END SELECT
         NEXT
      end if
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

   dim as CWSTR cwzRelative, wszText 
   Dim As Long nCount 
   Dim As Long i, nActiveTab

   Dim pDoc As clsDocument Ptr
   dim as long idx = gApp.GetActiveProjectIndex()

   dim pStream as CTextStream  ' (utf16)
   if pStream.Create(gApp.Projects(idx).ProjectFilename, true, true) <> S_OK then return false
   
   pStream.WriteLine "' WINFBE PROJECT FILE"
   pStream.WriteLine "ProjectBuild=" &       gApp.Projects(idx).ProjectBuild
   pStream.WriteLine "ProjectOther32=" &     gApp.Projects(idx).ProjectOther32
   pStream.WriteLine "ProjectOther64=" &     gApp.Projects(idx).ProjectOther64
   pStream.WriteLine "ProjectCommandLine=" & gApp.Projects(idx).ProjectCommandLine

   ' Save all of the loaded tabs first
   nCount = TabCtrl_GetItemCount(gTTabCtl.hWindow) 
   nActiveTab = TabCtrl_GetCurSel( gTTabCtl.hWindow )

   pStream.WriteLine "ActiveTab=" & nActiveTab

   Dim TB As TC_ITEM 
   TB.Mask = TCIF_PARAM

   
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
            pStream.WriteLine "File=" & wszText
            pStream.WriteLine "FileType=" & pDoc->ProjectFileType
            pStream.WriteLine "TabIndex=" & -1     ' use -1 rather than True
            pStream.WriteLine "Bookmarks=" & pDoc->GetBookmarks()
            pStream.WriteLine "FirstLine=" & SendMessage( pDoc->hWindow(0), SCI_GETFIRSTVISIBLELINE, 0, 0) 
            pStream.WriteLine "Position=" & SendMessage( pDoc->hWindow(0), SCI_GETCURRENTPOS, 0, 0)  
            pStream.WriteLine "FirstLine1=" & SendMessage( pDoc->hWindow(1), SCI_GETFIRSTVISIBLELINE, 0, 0) 
            pStream.WriteLine "Position1=" & SendMessage( pDoc->hWindow(1), SCI_GETCURRENTPOS, 0, 0)  
            pStream.WriteLine "SplitPosition=" & pDoc->SplitY
            pStream.WriteLine "FocusEdit=" & iif(pDoc->hWndActiveScintilla = pDoc->hWindow(0), 0, 1)
            pStream.WriteLine "FileEnd=[-]" 
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
               pStream.WriteLine "File=" & wszText
               pStream.WriteLine "FileType=" & pDoc->ProjectFileType
               pStream.WriteLine "TabIndex=" & 0    
               pStream.WriteLine "FileEnd=[-]"
            End If
         end if
      End If
   Next

   ' If no project is active then save NOTES to config file.
   if gApp.IsProjectActive THEN
      dim wszText as CWSTR = wstr("NOTES-START") + vbcrlf + _
                             gApp.Projects(idx).ProjectNotes + vbcrlf + _
                             wstr("NOTES-END") + vbcrlf
      pStream.WriteLine ""
      pStream.WriteLine "[Notes]"
      pStream.WriteLine wszText                       
   END IF

   pStream.Close

   Function = True   ' successful save
End Function


' ========================================================================================
' Load all options for the Project from a diskfile
' ========================================================================================
Function clsConfig.ProjectLoadFromFile( byref wzFile as WSTRING) As BOOLEAN    

   Dim pDoc As clsDocument Ptr
   Dim as CWSTR wst, wKey, wData, wszFilename
      
   Dim sBookmarks As String
   Dim As Long nData, i, iTab, nActiveTab, nFocusEdit 
   dim as long nFirstLine, nPosition, nFirstLine1, nPosition1, nSplitPosition

   ' Info saved for each File and used to create the file once FileEnd is found
   Dim As Long nFileType, nTabIndex
   Dim As BOOLEAN bLoadInTab
   dim as Boolean bReadingNote
   dim as long idx    ' project index

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
      .ProjectPath = AfxStrPathname( "PATH", wzFile )
      .ProjectName = AfxStrPathname( "NAMEX", wzFile )
      .hExplorerRootNode = FF_TreeView_InsertItem (hTV, TVI_ROOT, UCase(.ProjectName),0,II4E_PROJECT,II4E_PROJECT)
      .ProjectNotes = ""
      Treeview_SetBold(hTV, .hExplorerRootNode, TRUE)
   end with
   
    
   dim pStream as CTextStream  ' (utf16)
   if pStream.OpenUnicode(gApp.Projects(idx).ProjectFilename) <> S_OK then return false

   ' Variable length array to hold sequence of TabCtrl tabs to open. We save the sequence
   ' here and open them only after all over documents have been loaded. This will look
   ' visually more appealing.
   Dim nTabs(any) As clsDocument Ptr

  
   ' Prevent the constant updating of the app caption bar with filename, (pDoc->DisplayStats)
   gProjectLoading = true 
   gApp.ProjectOverrideIndex = idx
   gApp.IsWindowIncludes = false
   gdb2.dbRewind4Proj(idx)
   gdb2.dbDeleteWinAPI
         
   ' Set InUse here prior to parsing the files of the project
   gApp.Projects(idx).InUse = true

   ' Prevent Treeview flickering during adding
   LockWindowUpdate(hTV)

   Do Until pStream.EOS
      AfxDoEvents()
      wst = pStream.ReadLine
      
      If Len(wst) = 0 Then Continue Do
      If WStrNIsEqual(wst, "'", 1) = true Then Continue Do
      If WStrNIsEqual(wst, "[", 1) = true Then Continue Do
      
      If WStrNIsEqual(wst, "NOTES-START", 11) = true Then  
         bReadingNote = true
         continue do
      end if   
      If WStrNIsEqual(wst, "NOTES-END", 9) = true Then  
         bReadingNote = false
         continue do
      end if
      if bReadingNote THEN
         gApp.Projects(idx).ProjectNotes = gApp.Projects(idx).ProjectNotes + wst + vbcrlf
         continue do
      END IF

      i = Instr(wst, "=")
      If i = 0 Then Continue Do
      
      wKey  = Left(wst, i-1)
      wData = Mid(**wst, i+1)
      
      Select Case UCase(wData)
         Case "TRUE":   nData = True
         Case "FALSE":  nData = False
         Case Else:     nData = Val(wData)
      End Select

      wData = ProcessFromCurdrive(wData)

      Select Case **wKey
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
                                       wszFilename, _  ' pwszName
                                       0, _
                                       0, _
                                       nFileType, _
                                       idx)  ' pDocIn
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
   pStream.Close
   
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


