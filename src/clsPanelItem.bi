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

#include once "clsControl.bi"

type clsPanelItem
   private:
   
   public:
      wszName         as CWSTR
      wszText         as CWSTR
      wszTooltip      as CWSTR
      wszAlignment    as CWSTR = wstr("StatusBarPanelAlignment.Left")
      ' BorderStyle is deprecated as of v2.0.4 as it has not effect
      ' in WinFBE programs where Windows Themes are enabled.
      'wszBorderStyle as CWSTR = wstr("StatusBarPanelBorderStyle.Sunken")
      wszAutosize     as CWSTR = wstr("StatusBarPanelAutoSize.None")
      wszWidth        as CWSTR = wstr("100")
      wszMinWidth     as CWSTR = wstr("100")
      wszBackColor    as CWSTR = "SYSTEM|Control"
      wszBackColorHot as CWSTR = "SYSTEM|Control"
      wszForeColor    as CWSTR = "SYSTEM|ControlText"
      wszForeColorHot as CWSTR = "SYSTEM|ControlText"
      pProp           as clsProperty     ' for the panel image
      pPropColor      as clsProperty     ' for passing to color picker (see GetActivePropertyPtr)
      idColorCombo    as long            ' ctrl id of combobox that the pPropColor relates to.
end type

' Temporary PanelItem array to hold items while they are being
' edited in the StatusBar Editor. 
dim shared gPanelItems(any) as clsPanelItem

