'    LangEdit - Language Editor for the WinFBE Editor
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

#Define UNICODE
#Define _WIN32_WINNT &h0602  

#INCLUDE ONCE "Afx\CWSTR.inc"
#INCLUDE ONCE "Afx\CFindFile.inc"
#INCLUDE ONCE "Afx\AfxFile.inc"

#Define APPNAME       WStr("LangEdit")
#Define APPVERSION    WStr("1.0.0") 

#ifdef __FB_64BIT__
   #Define APPBITS WStr(" (64-bit)")
#else
   #Define APPBITS WStr(" (32-bit)")
#endif

#Define MAXPHRASES  400


#INCLUDE ONCE "clsLanguage.inc"
#INCLUDE ONCE "modRoutines.inc"
#INCLUDE ONCE "frmMain.inc"

  

''
''  
Application.Run(frmMain)
