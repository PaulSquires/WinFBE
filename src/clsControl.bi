'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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


enum PropertyType
   EditEnter = 1
   EditEnterNumericOnly
   TrueFalse
   ComboPicker
   ColorPicker
   FontPicker
   ImagePicker
   CustomDialog
end enum
   
type clsEvent
   private:
   public:
      wszEventName as CWSTR          ' Used for Get/Set of event value
      bIsSelected  as Boolean        ' User has selected this event to include into code
END TYPE

type clsProperty
   private:
   public:
      wszPropName      as CWSTR      ' Used for Get/Set of property value
      wszPropValuePrev as CWSTR      ' Previous property value (so that ApplyProperties will only act on changed properties)
      wszPropValue     as CWSTR
      wszPropDefault   as CWSTR
      PropType         as PropertyType 
END TYPE

type clsControl
   private:
   
   public:
      hWindow           as hwnd
      ControlType       as long 
      AfxButtonPtr      as CXPButton Ptr    ' we use XPButton rather than the built in Windows button
      AfxMaskedPtr      as CMaskedEdit Ptr 
      AfxPicturePtr     as CImageCtx Ptr 
      IsSelected        as Boolean
      IsActive          as Boolean
      SuspendLayout     as Boolean         ' prevent layout properties from being acted on individually (instead treat as a group)
      rcHandles(1 to 8) as RECT            ' 8 grab handles
      Properties(Any)   As clsProperty
      Events(Any)       As clsEvent
      hBackBrush        as HBRUSH            ' needed for STATIC/LABEL controls (destroyed in destructor)
      hImageList        as HANDLE            ' needed for TabControl
      Declare Destructor
END TYPE


