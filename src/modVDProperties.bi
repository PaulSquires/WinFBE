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

#pragma once

Enum FontStyles
   Normal    = 0
   Bold      = 1
   Italic    = 2
   Strikeout = 4
   Underline = 8
End Enum

Enum FontCharset
   Default     = DEFAULT_CHARSET
   Ansi        = ANSI_CHARSET
   Arabic      = ARABIC_CHARSET
   Baltic      = BALTIC_CHARSET
   ChineseBig5 = CHINESEBIG5_CHARSET
   EastEurope  = EASTEUROPE_CHARSET
   GB2312      = GB2312_CHARSET
   Greek       = GREEK_CHARSET
   Hangul      = HANGUL_CHARSET
   Hebrew      = HEBREW_CHARSET
   Johab       = JOHAB_CHARSET
   Mac         = MAC_CHARSET
   OEM         = OEM_CHARSET
   Russian     = RUSSIAN_CHARSET
   Shiftjis    = SHIFTJIS_CHARSET
   Symbol      = SYMBOL_CHARSET
   Thai        = THAI_CHARSET
   Turkish     = TURKISH_CHARSET
   Vietnamese  = VIETNAMESE_CHARSET
End Enum

Declare Function DisplayPropertyDetails() as Long
Declare Function DisplayEventDetails() as Long
declare function GetRGBColorFromProperty( byref wszPropValue as wstring ) as COLORREF
declare function SetLogFontFromPropValue( byref wszPropValue as wstring ) as LOGFONT
declare function SetPropValueFromLogFont( byref lf as LOGFONT ) as CWSTR
Declare Function IsPropertyExists( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR ) as boolean
Declare Function IsEventExists( byval pCtrl as clsControl ptr, byval wszEventName as CWSTR ) as boolean
Declare Function GetControlPropertyPtr( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR ) as clsProperty Ptr
Declare Function GetControlProperty( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR ) as CWSTR
Declare Function SetControlProperty( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR, byval wszPropValue as CWSTR ) as long
Declare Function SetControlEvent( byval pCtrl as clsControl ptr, byval wszEventName as CWSTR, byval bIsSelected as boolean ) as long






