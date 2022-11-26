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

#include once "modVDControls.bi"
#include once "modVDDesignForm.bi"
#include once "frmMenuEditor.bi"
#include once "frmStatusBarEditor.bi"
#include once "frmToolBarEditor.bi"


' ========================================================================================
' Add a property to the incoming control and optionally a value
' ========================================================================================
function AddControlProperty( _
            byval pCtrl as clsControl ptr, _
            byref wszPropName as CWSTR, _
            byref wszPropValue as CWSTR, _
            byval nPropType as PropertyType _
            ) as Long
   dim as long ub = ubound(pCtrl->Properties) + 1
   redim preserve pCtrl->Properties(ub) as clsProperty
   pCtrl->Properties(ub).wszPropName  = wszPropName
   pCtrl->Properties(ub).wszPropValue = wszPropValue
   pCtrl->Properties(ub).wszPropDefault = wszPropValue
   pCtrl->Properties(ub).PropType = nPropType
   function = 0
end function


' ========================================================================================
' Attach default properties for the incoming control
' ========================================================================================
function AttachDefaultControlProperties( byval pCtrl as clsControl ptr ) as Long

   ' Only add default properties if the no properties already exist for the control
   if ubound(pCtrl->Properties) - lbound(pCtrl->Properties) + 1 > 0 then exit function
  
   ' Add the default same generic proprties that apply to all controls
   AddControlProperty(pCtrl, "Name", "", PropertyType.EditEnter)
   AddControlProperty(pCtrl, "Left", "0", PropertyType.EditEnterNumericOnly)
   AddControlProperty(pCtrl, "Top", "0", PropertyType.EditEnterNumericOnly)
   AddControlProperty(pCtrl, "Width", "0", PropertyType.EditEnterNumericOnly)
   AddControlProperty(pCtrl, "Height", "0", PropertyType.EditEnterNumericOnly)

   ' Add control specific properties
   select CASE pCtrl->ControlType
      case CTRL_FORM
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "ChildForm", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ChildFormParent", "", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "WindowState", "FormWindowState.Normal", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "StartPosition", "FormStartPosition.Manual", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "BorderStyle", "FormBorderStyle.Sizable", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "MinimizeBox", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MaximizeBox", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ControlBox", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "AcceptButton", "", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "KeyPreview", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "CancelButton", "", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Icon", "", PropertyType.ImagePicker)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MaximumHeight", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "MaximumWidth", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "MinimumHeight", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "MinimumWidth", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)

      case CTRL_LABEL
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BackColorHot", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "ForeColorHot", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.None", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "TextAlign", "LabelAlignment.TopLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
          
      case CTRL_BUTTON
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "AllowFocusRect", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BackColorDown", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BackColorHot", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "MultiLine", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TextForeColor", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "TextForeColorHot", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "TextBackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "TextForeColorDown", "SYSTEM|ControlText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "TextBackColorDown", "SYSTEM|Control", PropertyType.ColorPicker)
         ' The properties are automatically sorted when they are displayed, however we
         ' need to order them here correctly because with ApplyProperties is called we want
         ' to ensure that the properties are acted on in a specific order. The width and height
         ' need to be set before scaling is called.
         AddControlProperty(pCtrl, "Image", "", PropertyType.ImagePicker)
         AddControlProperty(pCtrl, "ImageAlign", "ImageAlignment.TopLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "ImageWidth", "16", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "ImageHeight", "16", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "ImageMargin", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "ImageHighDPI", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TextAlign", "ButtonAlignment.MiddleCenter", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "TextMargin", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ThemeSupport", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ToggleMode", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_TEXTBOX, CTRL_RICHEDIT
         AddControlProperty(pCtrl, "AcceptsReturn", iif(pCtrl->ControlType=CTRL_TEXTBOX,"False","True"), PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "AcceptsTab", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.Fixed3D", PropertyType.ComboPicker)
         if pCtrl->ControlType = CTRL_TEXTBOX then
         AddControlProperty(pCtrl, "CharacterCasing", "CharacterCase.Normal", PropertyType.ComboPicker)
         end if
         if pCtrl->ControlType = CTRL_TEXTBOX then
         AddControlProperty(pCtrl, "CueBannerText", "", PropertyType.EditEnter)
         end if
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "HideSelection", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         if pCtrl->ControlType = CTRL_TEXTBOX then
         AddControlProperty(pCtrl, "MaxLength", "32767", PropertyType.EditEnterNumericOnly)
         else
         AddControlProperty(pCtrl, "MaxLength", "64000", PropertyType.EditEnterNumericOnly)
         end if
         AddControlProperty(pCtrl, "Multiline", iif(pCtrl->ControlType=CTRL_TEXTBOX,"False","True"), PropertyType.TrueFalse)
         if pCtrl->ControlType = CTRL_TEXTBOX then
         AddControlProperty(pCtrl, "PasswordChar", "", PropertyType.EditEnter)
         end if
         AddControlProperty(pCtrl, "ReadOnly", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TextScrollBars", "ScrollBars.None", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TextAlign", "TextAlignment.Left", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "WordWrap", iif(pCtrl->ControlType=CTRL_TEXTBOX,"False","True"), PropertyType.TrueFalse)

      case CTRL_MASKEDEDIT
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.Fixed3D", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "DefaultCharacter", "_", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "HideSelection", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "InputString", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MaskString", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ReadOnly", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ValidCharacters", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_CHECKBOX
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "CheckState", "CheckBoxState.Unchecked", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TextAlign", "ButtonAlignment.MiddleLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "ThreeState", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_LISTBOX
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BackColorHot", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BackColorSelected", "SYSTEM|Highlight", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.Fixed3D", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "ColumnWidth", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "ForeColorHot", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "ForeColorSelected", "SYSTEM|HighlightText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "HorizontalExtent", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "IntegralHeight", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ItemHeight", "28", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MultiColumn", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ScrollAlwaysVisible", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "SelectionMode", "ListSelectionMode.One", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Sorted", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TextAlign", "LabelAlignment.MiddleLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "UseTabStops", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_COMBOBOX
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "DropDownStyle", "ComboBoxStyle.DropDownList", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "IntegralHeight", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Sorted", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_OPTION
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "Checked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "GroupName", "OptionGroup1", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "StartGroup", "StartGroup", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TextAlign", "ButtonAlignment.MiddleLeft", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_FRAME
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "UseMnemonic", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_PICTUREBOX
         AddControlProperty(pCtrl, "Text", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BackColorHot", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
         ' The properties are automatically sorted when they are displayed, however we
         ' need to order them here correctly because with ApplyProperties is called we want
         ' to ensure that the properties are acted on in a specific order. 
         AddControlProperty(pCtrl, "Image", "", PropertyType.ImagePicker)
         AddControlProperty(pCtrl, "ImageWidth", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "ImageHeight", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "ImageScaling", "ImageScale.AutoSize", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)

      case CTRL_PROGRESSBAR
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "Maximum", "100", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Minimum", "0", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "StepValue", "1", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Value", "0", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Marquee", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MarqueeAnimationSpeed", "30", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Vertical", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_LISTVIEW
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.Fixed3D", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "CheckBoxes", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "FullRowSelect", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "GridLines", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "HeaderThemed", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "HeaderStyle", "ColumnHeaderStyle.Clickable", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "HeaderHeight", "20", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "HeaderForeColor", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "HeaderBackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "HideSelection", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MultiSelect", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "OddRowColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "OddRowColorEnabled", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
      
      case CTRL_TREEVIEW
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Window", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "BorderStyle", "ControlBorderStyle.Fixed3D", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "CheckBoxes", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "FadeButtons", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "ForeColor", "SYSTEM|WindowText", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "FullRowSelect", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "HideSelection", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "HotTracking", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ItemHeight", "20", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Scrollable", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ShowLines", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ShowRootLines", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ShowPlusMinus", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Sorted", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
         
      case CTRL_MONTHCALENDAR
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "BackColor", "SYSTEM|Control", PropertyType.ColorPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ShortDayNames", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TodayCircle", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "SelectedDate", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "TodaySelector", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TrailingDates", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "WeekNumbers", "False", PropertyType.TrueFalse)

      case CTRL_DATETIMEPICKER
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "CalendarRightAlign", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "DateFormat", "DateTimePickerFormat.ShortDateCentury", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "FormatCustom", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "SelectedDate", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "SelectedTime", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ShowUpDown", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_TABCONTROL
         AddControlProperty(pCtrl, "(Custom)", "", PropertyType.CustomDialog)
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "AllowFocus", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "ResizeTabPages", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ButtonStyle", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "FixedWidthTabs", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ForceImageLeft", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "ForceLabelLeft", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "HotTracking", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabTopPadding", "4", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabSidePadding", "4", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "MultiLine", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabImageSize", "ControlImageSize.Size16", PropertyType.ComboPicker)
         AddControlProperty(pCtrl, "TabHeight", "24", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabWidth", "100", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_UPDOWN
         AddControlProperty(pCtrl, "AllowDrop", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Anchor", "", PropertyType.AnchorPicker)
         AddControlProperty(pCtrl, "Enabled", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Font", CreateDefaultFontPropValue, PropertyType.FontPicker)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "TabIndex", "0", PropertyType.EditEnterNumericOnly)
         AddControlProperty(pCtrl, "TabStop", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTip", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "ToolTipBalloon", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Visible", "True", PropertyType.TrueFalse)

      case CTRL_TIMER
         AddControlProperty(pCtrl, "AutoReset", "True", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Enabled", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Locked", "False", PropertyType.TrueFalse)
         AddControlProperty(pCtrl, "Tag", "", PropertyType.EditEnter)
         AddControlProperty(pCtrl, "Interval", "1000", PropertyType.EditEnterNumericOnly)

      case CTRL_HSCROLL
      case CTRL_VSCROLL
      case CTRL_SLIDER
      case CTRL_WEBBROWSER
      case CTRL_CUSTOM
      case CTRL_OCX
   END SELECT

   function = 0
end function


' ========================================================================================
' Attach default events for the incoming control
' ========================================================================================
private function AttachDefaultControlEvents( byval pCtrl as clsControl ptr ) as Long

   ' Only add default events if the no events already exist for the control
   if ubound(pCtrl->Events) - lbound(pCtrl->Events) + 1 > 0 then exit function
  
   ' Add the default AllEvents event that applies to all controls
   AddControlEvent(pCtrl, "AllEvents")

   ' Add control specific properties
   select CASE pCtrl->ControlType
      case CTRL_FORM
         AddControlEvent(pCtrl, "Load")
         AddControlEvent(pCtrl, "Initialize")
         AddControlEvent(pCtrl, "Shown")
         AddControlEvent(pCtrl, "Activated")
         AddControlEvent(pCtrl, "Deactivate")
         AddControlEvent(pCtrl, "Resize")
         AddControlEvent(pCtrl, "FormClosing")
         AddControlEvent(pCtrl, "FormClosed")
         AddControlEvent(pCtrl, "FormReady")
         AddControlEvent(pCtrl, "Move")
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
         AddControlEvent(pCtrl, "MessagePumpHook")
      
      case CTRL_LABEL
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
      
      case CTRL_BUTTON
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      
      case CTRL_TEXTBOX, CTRL_RICHEDIT
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "TextChanged")
         if pCtrl->ControlType = CTRL_RICHEDIT then
            AddControlEvent(pCtrl, "SelectionChanged")
         end if   
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      
      case CTRL_MASKEDEDIT
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "TextChanged")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      
      case CTRL_CHECKBOX, CTRL_UPDOWN
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      
      case CTRL_LISTBOX
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      
      case CTRL_OPTION
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
      
      case CTRL_FRAME
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
      
      case CTRL_COMBOBOX
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "DropDown")
         AddControlEvent(pCtrl, "DropDownClosed")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
         AddControlEvent(pCtrl, "TextChanged")
      
      case CTRL_PICTUREBOX
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
      
      case CTRL_PROGRESSBAR
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
      
      case CTRL_LISTVIEW
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "DoubleClick")
         AddControlEvent(pCtrl, "RightClick")
         AddControlEvent(pCtrl, "ColumnClick")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "ItemSelectionChanged")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         ' Do not add a MouseHover event for the ListView as it
         ' seems to invoke hot tracking and automatic row selection.
         'AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")

      case CTRL_TREEVIEW
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "BeforeSelect")
         AddControlEvent(pCtrl, "AfterSelect")
         AddControlEvent(pCtrl, "BeforeCollapse")
         AddControlEvent(pCtrl, "AfterCollapse")
         AddControlEvent(pCtrl, "BeforeExpand")
         AddControlEvent(pCtrl, "AfterExpand")
         AddControlEvent(pCtrl, "BeforeCheck")
         AddControlEvent(pCtrl, "AfterCheck")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")

      case CTRL_MONTHCALENDAR
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
         AddControlEvent(pCtrl, "SelectionChanged")

      case CTRL_DATETIMEPICKER
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "DateTimeChanged")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")

      case CTRL_TABCONTROL
         AddControlEvent(pCtrl, "Click")
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "DropFiles")
         AddControlEvent(pCtrl, "GotFocus")
         AddControlEvent(pCtrl, "LostFocus")
         AddControlEvent(pCtrl, "MouseMove")
         AddControlEvent(pCtrl, "MouseDown")
         AddControlEvent(pCtrl, "MouseUp")
         AddControlEvent(pCtrl, "MouseDoubleClick")
         AddControlEvent(pCtrl, "MouseEnter")
         AddControlEvent(pCtrl, "MouseHover")
         AddControlEvent(pCtrl, "MouseLeave")
         AddControlEvent(pCtrl, "KeyDown")
         AddControlEvent(pCtrl, "KeyPress")
         AddControlEvent(pCtrl, "KeyUp")
         AddControlEvent(pCtrl, "Selected")
         AddControlEvent(pCtrl, "Selecting")

      case CTRL_TIMER
         AddControlEvent(pCtrl, "Destroy")
         AddControlEvent(pCtrl, "Elapsed")
         
         
      case CTRL_HSCROLL
      case CTRL_VSCROLL
      case CTRL_SLIDER
      case CTRL_WEBBROWSER
      case CTRL_CUSTOM
      case CTRL_OCX
   END SELECT

   function = 0
end function

' ========================================================================================
' Respond to keyboard input and move selected controls by 1 pixel 
' ========================================================================================
public function KeyboardMoveControls( byval pDoc as clsDocument ptr, _
                                      byval vKeycode as long _
                                      ) as Long

   dim pCtrl as clsControl ptr
   dim as long nTemp
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType <> CTRL_FORM THEN
         if pCtrl->IsSelected then
            select CASE vKeycode
               case VK_UP
                  nTemp = val(GetControlProperty(pCtrl, "TOP")) - 1
                  SetControlProperty(pCtrl, "TOP", str(nTemp))
                  pDoc->UserModified = true
               case VK_DOWN
                  nTemp = val(GetControlProperty(pCtrl, "TOP")) + 1
                  SetControlProperty(pCtrl, "TOP", str(nTemp))
                  pDoc->UserModified = true
               case VK_LEFT
                  nTemp = val(GetControlProperty(pCtrl, "LEFT")) - 1
                  SetControlProperty(pCtrl, "LEFT", str(nTemp))
                  pDoc->UserModified = true
               case VK_RIGHT         
                  nTemp = val(GetControlProperty(pCtrl, "LEFT")) + 1
                  SetControlProperty(pCtrl, "LEFT", str(nTemp))
                  pDoc->UserModified = true
            END SELECT
            pCtrl->SuspendLayout = true
            ApplyControlProperties( pDoc, pCtrl )
            pCtrl->SuspendLayout = false
         end if
      end if
   next

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_SetStatusbar
   DisplayPropertyList(pDoc)

   function = 0
end function


' ========================================================================================
' Respond to keyboard input and resize selected controls by 1 pixel 
' ========================================================================================
public function KeyboardResizeControls( byval pDoc as clsDocument ptr, _
                                        byval vKeycode as long _
                                        ) as Long

   dim pCtrl as clsControl ptr
   dim as long nTemp
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl->ControlType <> CTRL_FORM THEN
         if pCtrl->IsSelected then
            select CASE vKeycode
               case VK_UP
                  nTemp = val(GetControlProperty(pCtrl, "HEIGHT")) - 1
                  SetControlProperty(pCtrl, "HEIGHT", str(nTemp))
                  pDoc->UserModified = true
               case VK_DOWN
                  nTemp = val(GetControlProperty(pCtrl, "HEIGHT")) + 1
                  SetControlProperty(pCtrl, "HEIGHT", str(nTemp))
                  pDoc->UserModified = true
               case VK_LEFT
                  nTemp = val(GetControlProperty(pCtrl, "WIDTH")) - 1
                  SetControlProperty(pCtrl, "WIDTH", str(nTemp))
                  pDoc->UserModified = true
               case VK_RIGHT         
                  nTemp = val(GetControlProperty(pCtrl, "WIDTH")) + 1
                  SetControlProperty(pCtrl, "WIDTH", str(nTemp))
                  pDoc->UserModified = true
            END SELECT
            pCtrl->SuspendLayout = true
            ApplyControlProperties( pDoc, pCtrl )
            pCtrl->SuspendLayout = false
         end if
      end if
   next

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_SetStatusbar
   DisplayPropertyList(pDoc)

   function = 0
end function


' ========================================================================================
' Respond to keyboard input to cycle amongst the selected controls in the group.
' ========================================================================================
public function KeyboardCycleActiveControls( byval pDoc as clsDocument ptr, _
                                             byval vKeycode as long _
                                             ) as Long

   dim pCtrl as clsControl ptr
   dim as long idxActive 
   dim as long idxPrev = -1
   dim as long idxNext = -1
   dim as Boolean bSelectedGroup = iif( pDoc->Controls.SelectedControlsCount > 1, true, false)
   
   ' Get the current active control in the group
   dim pCtrlActive as clsControl ptr = pDoc->Controls.GetActiveControl
   
   ' Determine the array position that the active control is at.
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      if pDoc->Controls.ItemAt(i) = pCtrlActive then
         idxActive = i: exit for
      end if
   next      
   
   ' The following will cycle amongst controls in a selected group. If there is
   ' no selected group then the focus will simply shift to the next/prev control.
   select CASE vKeycode
      case VK_UP, VK_LEFT    ' get the previous control in the group
         for i as long = (idxActive - 1) to pDoc->Controls.ItemFirst step -1
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->ControlType = CTRL_FORM THEN continue for
            if bSelectedGroup then
               if pCtrl->IsSelected then
                  idxPrev = i: exit for
               end if
            else
               idxPrev = i: exit for
            end if         
         next
         if idxPrev = -1 THEN  ' prev still not found. wrap to end of array
            for i as long = pDoc->Controls.ItemLast to idxActive step -1
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->ControlType = CTRL_FORM THEN continue for
               if bSelectedGroup then
                  if pCtrl->IsSelected then
                     idxPrev = i: exit for
                  end if   
               else
                  idxPrev = i: exit for
               end if   
            next
         end if
         if idxPrev <> -1 then
            if bSelectedGroup = false THEN pDoc->Controls.DeselectAllControls
            pCtrl = pDoc->Controls.ItemAt(idxPrev)
            pDoc->Controls.SetActiveControl(pCtrl->hWindow)
         end if
         
      case VK_DOWN, VK_RIGHT   ' get the next control in the gorup
         for i as long = (idxActive + 1) to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl->ControlType = CTRL_FORM THEN continue for
            if bSelectedGroup then
               if pCtrl->IsSelected then
                  idxNext = i: exit for
               end if   
            else
               idxNext = i: exit for
            end if      
         next
         if idxNext = -1 THEN  ' next still not found. wrap to start of array
            for i as long = pDoc->Controls.ItemFirst to idxActive 
               pCtrl = pDoc->Controls.ItemAt(i)
               if pCtrl->ControlType = CTRL_FORM THEN continue for
               if bSelectedGroup then
                  if pCtrl->IsSelected then
                     idxNext = i: exit for
                  end if   
               else    
                  idxNext = i: exit for
               end if   
            next
         end if
         if idxNext <> -1 then
            if bSelectedGroup = false THEN pDoc->Controls.DeselectAllControls
            pCtrl = pDoc->Controls.ItemAt(idxNext)
            pDoc->Controls.SetActiveControl(pCtrl->hWindow)
         END IF

   end select

   AfxRedrawWindow(pDoc->hWndForm)
   frmMain_SetStatusbar
   DisplayPropertyList(pDoc)

   function = 0
end function


' ========================================================================================
' Determine if the specified control name already exists on the incoming form.
' ========================================================================================
public function IsControlNameExists( byval pDoc as clsDocument ptr, _
                                     byref wszControlName as wstring _
                                     ) as boolean
   dim pCtrl as clsControl ptr 
   for ii as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(ii)
      if pCtrl->ControlType <> CTRL_FORM THEN
         if ucase(GetControlProperty(pCtrl, "NAME")) = ucase(wszControlName) THEN
            return true
         END IF
      END IF
   NEXT
   return false
end function


' ========================================================================================
' Determines if the control is locked or full form is locked
' ========================================================================================
public function IsControlLocked( byval pDoc as clsDocument ptr, _
                                 byval pCtrl as clsControl ptr _
                                 ) as boolean
   if pDoc->bLockControls then return true
   if GetControlProperty(pCtrl, "LOCKED") = "True" then return true
   return false
end function


' ========================================================================================
' Determine if the specified form name already exists.
' ========================================================================================
public function IsFormNameExists( byref wszFormName as wstring ) as boolean

   dim pDoc as clsDocument ptr = gApp.pDocList
   dim pCtrl as clsControl ptr 
   dim as Boolean bExists = false
   
   do until pDoc = 0
      for ii as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(ii)
         if pCtrl->ControlType = CTRL_FORM THEN
            if ucase(GetControlProperty(pCtrl, "NAME")) = ucase(wszFormName) THEN
               bExists = true
            END IF
            exit for
         END IF
      NEXT
      if bExists THEN exit do
      pDoc = pDoc->pDocNext
   loop
   
   return bExists
end function


' ========================================================================================
' Generate a default name for the new form
' ========================================================================================
private function GenerateDefaultFormName() as CWSTR

   ' Generate a default name for the form, then iterate the document collection
   ' to see if it already exists.
   dim as long NextFormNumber = 0
   dim as CWSTR wszDefaultName 
   
   do 
      NextFormNumber = NextFormNumber + 1
      wszDefaultName = "Form" & NextFormNumber
      if IsFormNameExists(wszDefaultName) = false then exit do
   loop 
   
   return wszDefaultName
end function


' ========================================================================================
' Generate a default name for the new control
' ========================================================================================
private function GenerateDefaultControlName( byval pDoc as clsDocument ptr, _
                                             byval nControlType as long _
                                             ) as CWSTR

   ' Generate a default name for the control, then iterate the control collection
   ' to see if it already exists.
   dim as long NextControlNumber = 0
   dim as CWSTR wszDefaultName 
   dim as CWSTR wszName 
   dim as Boolean bOK 
   dim pCtrl as clsControl ptr 
   
   wszName = GetControlName(nControlType)

   do 
      NextControlNumber = NextControlNumber + 1
      wszDefaultName = wszName & NextControlNumber
      bOK = true
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         if ucase(GetControlProperty(pCtrl, "NAME")) = ucase(wszDefaultName) THEN
            bOK = false: exit for
         END IF
      NEXT
   loop until bOK = true
   
   return wszDefaultName
end function


' ========================================================================================
' Generate a default TabIndex value for the new control
' ========================================================================================
private function GenerateDefaultTabIndex( byval pDoc as clsDocument ptr ) as CWSTR

   dim as long HighTabIndex = -1
   dim as long nPropValue 
   dim pCtrl as clsControl ptr 
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      nPropValue = val(GetControlProperty(pCtrl, "TABINDEX"))
      if nPropValue > HighTabIndex then HighTabIndex = nPropValue
   NEXT
   
   return wstr(HighTabIndex+1)
end function


' ========================================================================================
' Re-Create the specified control on the design form
' ========================================================================================
public function ReCreateToolboxControl( byval pDoc as clsDocument ptr, _
                                        byval pCtrl as clsControl ptr _
                                        ) as long

   ' Some controls need to be destroyed and re-created in order for certan properties 
   ' to be properly displayed in the visual designer.
   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
   dim pProp as clsProperty ptr
   
   dim as RECT rc
   dim as DWORD dwStyle, dwExStyle
   dim wszCustomFormat as wstring * MAX_PATH = ""
   dim as CWSTR wszPropValue
   dim as long CtrlID = GetDlgCtrlID(pCtrl->hWindow)
   
   GetWindowRect( pCtrl->hWindow, @rc )
   MapWindowPoints( HWND_DESKTOP, pWindow->hWindow, CAST(POINT PTR, @rc), 2)

   select case pCtrl->ControlType

      case CTRL_TABCONTROL
         if pCtrl->hImageList then ImageList_Destroy(pCtrl->hImageList)
         DestroyWindow( pCtrl->hWindow )

         dwStyle = WS_VISIBLE
         if GetControlProperty(pCtrl, "BUTTONSTYLE") = "True" then
            dwStyle = dwStyle or TCS_BUTTONS or TCS_FLATBUTTONS
         else
            dwStyle = dwStyle or TCS_TABS
         end if

         if GetControlProperty(pCtrl, "FIXEDWIDTHTABS") = "True" then
            dwStyle = dwStyle or TCS_FIXEDWIDTH
         else
            dwStyle = dwStyle or TCS_RAGGEDRIGHT
         end if

         if GetControlProperty(pCtrl, "MULTILINE") = "True" then
            dwStyle = dwStyle or TCS_MULTILINE
         end if

         if GetControlProperty(pCtrl, "FORCELABELLEFT") = "True" then
            ' This will force the label and icon left
            dwStyle = dwStyle or TCS_FORCELABELLEFT
         else
            if GetControlProperty(pCtrl, "FORCEIMAGELEFT") = "True" then
               dwStyle = dwStyle or TCS_FORCEICONLEFT
            end if
         end if
            
         pCtrl->hWindow = pWindow->AddControl( "TabControl", , CtrlID, "", _
                    pWindow->UnScaleX(rc.left), pWindow->UnScaleY(rc.top), _
                    pWindow->UnScaleX(rc.right-rc.left), pWindow->UnScaleY(rc.bottom-rc.top), _
                    dwStyle, dwExStyle, , _
                    CAST(SUBCLASSPROC, @Control_SubclassProc), CTRL_TABCONTROL, CAST(DWORD_PTR, pDoc))

         wszPropValue = GetControlProperty(pCtrl, "(CUSTOM)")
         frmVDTabChild_LoadTabPagesArray( wszPropValue )
         
         dim as boolean bHasImages 
         for i as long = 0 to ubound(gTabPages)
            dim as long nIndex = TabCtrl_AddTab( pCtrl->hWindow, 0, gTabPages(i).wszText, 0 )
            if gTabPages(i).IsActiveTab then
               TabCtrl_SetCurSel( pCtrl->hWindow, nIndex )
            end if   
            if len(gTabPages(i).wszImage) then bHasImages = true
         next 
            
         if bHasImages then
            dim as long nImageSize = 16
            wszPropValue = GetControlProperty(pCtrl, "TABIMAGESIZE")
            select case wszPropValue
               case "ControlImageSize.Size16": nImageSize = 16
               case "ControlImageSize.Size24": nImageSize = 24
               case "ControlImageSize.Size32": nImageSize = 32
               case "ControlImageSize.Size48": nImageSize = 48
            end select
            
            dim as long cx = nImageSize * (pWindow->DPI \ 96)
            pCtrl->hImageList = ImageList_Create(cx, cx, ILC_MASK Or ILC_COLOR32, 1, 0)
            TabCtrl_SetImageList( pCtrl->hWindow, pCtrl->hImageList )

            Dim As HICON hIcon
            dim as long nIndex
            for i as long = 0 to ubound(gTabPages)
               TabCtrl_SetImageIndex( pCtrl->hWindow, i, -1 )   
               if len(rtrim(gTabPages(i).wszImage)) then
                  dim pImageType as IMAGES_TYPE ptr = GetImagesTypePtr(gTabPages(i).wszImage)
                  if pImageType then
                     hIcon = AfxGdipIconFromFile( pImageType->wszFileName )
                     nIndex = ImageList_AddIcon( pCtrl->hImageList, hIcon )
                     TabCtrl_SetImageIndex( pCtrl->hWindow, i, nIndex )   
                  end if
               end if
            next
         end if

         if GetControlProperty(pCtrl, "FIXEDWIDTHTABS") = "True" then
            dim as long nWidth  = AfxScaleX(val(GetControlProperty(pCtrl, "TABWIDTH")))
            dim as long nHeight = AfxScaleY(val(GetControlProperty(pCtrl, "TABHEIGHT")))
            if (nWidth * nHeight) <> 0 then  ' neither can be zero
               TabCtrl_SetItemSize( pCtrl->hWindow, nWidth, nHeight )
            end if   
         end if
         
         ' Set the horizontal and vertical padding
         dim as long nHorizPadding = AfxScaleX(val(GetControlProperty(pCtrl, "TABSIDEPADDING")))
         dim as long nVertPadding  = AfxScaleY(val(GetControlProperty(pCtrl, "TABTOPPADDING")))
         TabCtrl_SetPadding( pCtrl->hWindow, nHorizPadding, nVertPadding )

         ' Reset the gTabPages array
         erase gTabPages

         
      case CTRL_DATETIMEPICKER
         DestroyWindow( pCtrl->hWindow )
         
         dwStyle = WS_VISIBLE
         if GetControlProperty(pCtrl, "SHOWUPDOWN") = "True" then
            dwStyle = dwStyle Or DTS_UPDOWN
         end if   
   
         select case GetControlProperty(pCtrl, "DATEFORMAT")
            case "DateTimePickerFormat.LongDate"
               dwStyle = dwStyle Or DTS_LONGDATEFORMAT
            case "DateTimePickerFormat.ShortDate"
               dwStyle = dwStyle Or DTS_SHORTDATEFORMAT
            case "DateTimePickerFormat.ShortDateCentury"
               dwStyle = dwStyle Or DTS_SHORTDATECENTURYFORMAT
            case "DateTimePickerFormat.TimeFormat"
               dwStyle = dwStyle Or DTS_TIMEFORMAT
            case "DateTimePickerFormat.CustomFormat"
               ' Set value after control is created
               wszCustomFormat = GetControlProperty(pCtrl, "FORMATCUSTOM")
         end select

         pCtrl->hWindow = pWindow->AddControl( "SysDateTimePick32", , CtrlID, "", _
                    pWindow->UnScaleX(rc.left), pWindow->UnScaleY(rc.top), _
                    pWindow->UnScaleX(rc.right-rc.left), pWindow->UnScaleY(rc.bottom-rc.top), _
                    dwStyle, dwExStyle, , _
                    CAST(SUBCLASSPROC, @Control_SubclassProc), CTRL_DATETIMEPICKER, CAST(DWORD_PTR, pDoc))

         DateTime_SetFormat( pCtrl->hWindow, @wszCustomFormat )      

         Dim As SYSTEMTIME pst
         dim as CWSTR wszDate = GetControlProperty(pCtrl, "SELECTEDDATE")
         pst.wYear  = val(left(wszDate, 4))
         pst.wMonth = val(mid(wszDate, 5, 2))
         pst.wDay   = val(right(wszDate, 2))
         dim as CWSTR wszTime = GetControlProperty(pCtrl, "SELECTEDTIME")
         pst.wHour   = val(left(wszTime, 2))
         pst.wSecond = val(mid(wszTime, 3, 2))
         pst.wMinute = val(right(wszTime, 2))
         DateTime_SetSystemTime( pCtrl->hWindow, GDT_VALID, @pst )

   end select
   
   ' Apply the correct font
   pProp = GetControlPropertyPtr( pCtrl, "FONT" )
   if pProp then
      dim as HFONT hFontOld = AfxGetWindowFont(pCtrl->hWindow)
      dim as LOGFONT lf
      lf = SetLogFontFromPropValue(pProp->wszPropValue)
      dim as HFONT hFont = CreateFontIndirect(@lf)
      AfxSetWindowFont(pCtrl->hWindow, hFont, true)
      DeleteFont(hFontOld)
   end if
   
   function = 0
end function


' ========================================================================================
' Create the specified control on the design form
' ========================================================================================
public function CreateToolboxControl( byval pDoc as clsDocument ptr, _
                                      byval ControlType as long, _
                                      byref rcDraw as RECT _
                                      ) as clsControl ptr

   Dim pWindow As CWindow Ptr = AfxCWindowPtr(pDoc->hWndForm)
   dim pCtrl as clsControl ptr = new clsControl
   dim as CWSTR wszDefaultControlName, wszDefaultGroupName 

   pDoc->Controls.DeselectAllControls
   pCtrl->ControlType = ControlType
   pCtrl->IsActive    = false
   
   ' Set default properties and events for the created control
   AttachDefaultControlProperties(pCtrl)
   AttachDefaultControlEvents(pCtrl)

   if ControlType = CTRL_FORM THEN
      dim pForm as CWindow ptr = New CWindow
      pForm->DPI = AfxCWindowPtr(pDoc->hWndFrame)->DPI

      wszDefaultControlName = GenerateDefaultFormName()
      pDoc->hWndForm = _
         pForm->Create( pDoc->hWndFrame, "Form1", @DesignerForm_WndProc, 0, 0, 0, 0, _
              WS_CHILD or WS_OVERLAPPED Or WS_OVERLAPPEDWINDOW or WS_VISIBLE or _
              WS_CLIPSIBLINGS or WS_CLIPCHILDREN, _
              WS_EX_CONTROLPARENT OR WS_EX_WINDOWEDGE)
      pForm->ClassStyle = CS_DBLCLKS 
      SetWindowLongPtr( pDoc->hWndForm, GWLP_ID, IDC_DESIGNFORM )
      ' We will set our own mouse pointer as needed
      SetClassLongPtr(pDoc->hWndForm, GCLP_HCURSOR, 0)   
      pCtrl->hWindow = pDoc->hWndForm
   
      ' Fake topmenu
      pDoc->hWndFakeMenu = _
      pForm->AddControl("LABEL", pDoc->hWndForm, IDC_LBLFAKEMAINMENU, "", 0, 0, 0, 0, _
           WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT Or SS_NOTIFY, _
           WS_EX_LEFT Or WS_EX_LTRREADING)
      dim ncm As NONCLIENTMETRICS
      ncm.cbSize = SizeOf(ncm)
      SystemParametersInfo(SPI_GETNONCLIENTMETRICS, SizeOf(ncm), @ncm, 0)
      pDoc->hFontFakeMenu = CreateFontIndirect(@ncm.lfMenuFont)
      SendMessage( pDoc->hWndFakeMenu, WM_SETFONT, cast(WPARAM, pDoc->hFontFakeMenu), cast(LPARAM, CTRUE))
   
      ' StatusBar
      pDoc->hWndStatusBar = _
      pForm->AddControl("STATUSBAR", pDoc->hWndForm, IDC_FAKESTATUSBAR, "", 0, 0, 0, 0, _
           WS_CHILD Or WS_CLIPCHILDREN OR WS_CLIPSIBLINGS OR CCS_BOTTOM OR SBARS_TOOLTIPS, _
           WS_EX_LEFT Or WS_EX_LTRREADING)
      
      ' ToolBar
      frmToolBarEditor_CreateFakeToolBar( pDoc )

      ' Create the four SnapLines 
      for i as long = 0 to 3
         pDoc->hSnapLine(i) = _
               pForm->AddControl("LABEL", pDoc->hWndForm, IDC_LBLSNAPLINE, "", 0, 0, 0, 0, _
               WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or SS_LEFT, WS_EX_LEFT Or WS_EX_LTRREADING)
      next

   
   else
      if pWindow = 0 THEN exit function
      wszDefaultControlName = GenerateDefaultControlName(pDoc, ControlType)
      select case ControlType 
         case CTRL_BUTTON
            ' Use the custom control CXpButton instead of the standard Windows button control.
            pCtrl->AfxButtonPtr = new CXpButton(pWindow, 100, wszDefaultControlName, _
                             rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top)
            pCtrl->hWindow = pCtrl->AfxButtonPtr->hWindow
            SetWindowSubclass(pCtrl->hWindow, CAST(SUBCLASSPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
         case CTRL_MASKEDEDIT
            pCtrl->AfxMaskedPtr = new CMaskedEdit(pWindow, 100, _
                             rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top)
            pCtrl->hWindow = pCtrl->AfxMaskedPtr->hWindow
            SetWindowSubclass(pCtrl->hWindow, CAST(SUBCLASSPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
         case CTRL_PICTUREBOX
            pCtrl->AfxPicturePtr = new CImageCtx(pWindow, 100, , _
                             rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top)
            pCtrl->hWindow = pCtrl->AfxPicturePtr->hWindow
            SetWindowSubclass(pCtrl->hWindow, CAST(SUBCLASSPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
         case CTRL_TIMER
            ' Readjust the rcDraw to be 16x16
            rcDraw.right = rcDraw.left + 16   ' do not dpi scale number
            rcDraw.bottom = rcDraw.top + 16   ' do not dpi scale number
            pCtrl->AfxPicturePtr = new CImageCtx(pWindow, 100, , _
                             rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top)
            pCtrl->hWindow = pCtrl->AfxPicturePtr->hWindow
            SetWindowSubclass(pCtrl->hWindow, CAST(SUBCLASSPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
            pCtrl->AfxPicturePtr->LoadImageFromResource( pWindow->InstanceHandle, "IMAGE_TIMERCONTROL" )
            pCtrl->AfxPicturePtr->SetImageAdjustment( GDIP_IMAGECTX_STRETCH, true )
         case else   
            dim as long dwExStyle = -1
            dim as long dwStyle = -1
            if ControlType = CTRL_LISTBOX then
               ' WinFBX defaults to listbox with integralheight. We want to make sure that
               ' our listboxes have NoIntegralHeight.
               dwStyle = WS_CHILD or WS_VISIBLE OR WS_HSCROLL OR WS_VSCROLL OR WS_BORDER OR _
                         LBS_HASSTRINGS OR LBS_NOTIFY or LBS_NOINTEGRALHEIGHT
            elseif ControlType = CTRL_LABEL then
               dwStyle = WS_CHILD or WS_VISIBLE OR SS_LEFT OR WS_GROUP OR SS_NOTIFY or SS_OWNERDRAW
            elseif ControlType = CTRL_COMBOBOX then
               ' WinFBX defaults to combobox with integralheight. We want to make sure that
               ' our comboboxes have NoIntegralHeight.
               dwStyle = WS_CHILD or WS_VISIBLE OR WS_HSCROLL OR WS_VSCROLL OR WS_BORDER OR _
                         CBS_HASSTRINGS or CBS_NOINTEGRALHEIGHT or CBS_DROPDOWNLIST
               ' Also subclass the combobox edit control
               dim as HWND hWndEditCtrl = ComboBox_GetEditBoxHandle(pCtrl->hWindow)
               if IsWindow(hWndEditCtrl) then
                  SetWindowSubclass(hWndEditCtrl, CAST(SUBCLASSPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))
               end if
               dim as HWND hWndListCtrl = ComboBox_GetListBoxHandle(pCtrl->hWindow)
               if IsWindow(hWndListCtrl) then
                  SetWindowSubclass(hWndListCtrl, CAST(SUBCLASSPROC, @Control_SubclassProc), CTRL_COMBOBOX, CAST(DWORD_PTR, pDoc))
               end if
            elseif ControlType = CTRL_FRAME then
               dwStyle = WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN or BS_GROUPBOX
               dwExStyle = WS_EX_TRANSPARENT
            elseif ControlType = CTRL_MONTHCALENDAR then
               dwStyle = WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN
               dwExStyle = 0
            elseif ControlType = CTRL_TABCONTROL then
               dwStyle = WS_CHILD or WS_VISIBLE or WS_CLIPSIBLINGS or WS_CLIPCHILDREN OR _
                         TCS_TABS OR TCS_SINGLELINE OR TCS_RAGGEDRIGHT or TCS_FORCEICONLEFT
               dwExStyle = WS_EX_CONTROLPARENT
            end if   

            ' Create the control
            pCtrl->hWindow = pWindow->AddControl( GetControlClassName(pCtrl), , 100, wszDefaultControlName, _
                             rcDraw.left, rcDraw.top, rcDraw.right-rcDraw.left, rcDraw.bottom-rcDraw.top, _
                             dwStyle, dwExStyle, , _
                             CAST(SUBCLASSPROC, @Control_SubclassProc), ControlType, CAST(DWORD_PTR, pDoc))

            if pCtrl->ControlType = CTRL_TABCONTROL then
               ReCreateToolboxControl( pDoc, pCtrl )
            end if

      end select
   END IF

   pDoc->Controls.Add(pCtrl)

   SetControlProperty(pCtrl, "NAME", wszDefaultControlName)
   SetControlProperty(pCtrl, "TEXT", wszDefaultControlName)
   SetControlProperty(pCtrl, "LEFT", str(rcDraw.left))
   SetControlProperty(pCtrl, "TOP", str(rcDraw.top))
   SetControlProperty(pCtrl, "WIDTH", str(rcDraw.right-rcDraw.left))
   SetControlProperty(pCtrl, "HEIGHT", str(rcDraw.bottom-rcDraw.top))
   
   ' Need to create fake statusbar here in order to ensure that all of the
   ' statusbar panels display correctly when the form is first shown.
   frmStatusBarEditor_CreateFakeStatusBar(pDoc)

   ' If this form/control is being created from reading an existing source file then 
   ' then we would wait until the all of the control's properties have been loaded. 
   ' See clsDocument.ParseDesignerString
   if pDoc->LoadingFromFile = false then
      ' By default, generate the next highest TabIndex value for this new control. 
      SetControlProperty(pCtrl, "TABINDEX", GenerateDefaultTabIndex(pDoc))
      pCtrl->SuspendLayout = true
      ApplyControlProperties( pDoc, pCtrl )
      pCtrl->SuspendLayout = false
      pDoc->bRegenerateCode = true
      pDoc->Controls.SelectControl(pCtrl->hWindow)
      pDoc->Controls.SetActiveControl(pCtrl->hWindow)
      DisplayPropertyList(pDoc)
   end if
   
   function = pCtrl
END FUNCTION


