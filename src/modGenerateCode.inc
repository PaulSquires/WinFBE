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


' ========================================================================================
' Retrieve the Form name for this document.
' ========================================================================================
function GetFormName( byval pDoc as clsDocument ptr ) as CWSTR
   if pDoc = 0 then exit function
   
   dim pCtrl as clsControl ptr
   dim as CWSTR wszFormName
   
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if (pCtrl <> 0) andalso (pCtrl->ControlType = CTRL_FORM) then
         wszFormName = GetControlProperty(pCtrl, "NAME")
         exit for
      end if
   next
   function = wszFormName   

end function


' ========================================================================================
' Generate (or regenerate) visual designer code and insert it into the code editor.
' ========================================================================================
function GenerateCode( byval pDoc as clsDocument ptr ) as long
   if pDoc = 0 then exit function
   if pDoc->IsDesigner = false then exit function
   if pDoc->bRegenerateCode = false then exit function
   
   dim pCtrl as clsControl ptr
   dim as CWSTR wszText, wszFormName, wszCtrlName, wszPropName, wszPropValue, wszPropDefault
   dim as CWSTR wszFunction, wszAllEvents, wszPrototype, wszAllDeclares
   dim as CWSTR wszCombinedName
   dim as string sStartTag, sEndTag, sReplaceText
   dim as CWSTR DQ = chr(34)
   dim as long nStartTag, nEndTag, nPropType
   dim pData as DB2_DATA ptr
   
   dim as string sp(5) 
   for i as long = 1 to ubound(sp)
      sp(i) = space(val(gConfig.TabSize) * i)
   NEXT
   
   wszFormName = GetFormName(pDoc)
   sStartTag = "' WINFBE_CODEGEN_START"
   sEndTag   = "' WINFBE_CODEGEN_END"

   ''
   ''  Generate the Declares for all of the Events for the Form & Controls
   ''
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl then
         wszCtrlName = GetControlProperty(pCtrl, "NAME")
         for ii as long = lbound(pCtrl->Events) to ubound(pCtrl->Events)
            if pCtrl->Events(ii).bIsSelected then
               ' Determine the name of the event function
               wszFunction = wszFormName
               if pCtrl->ControlType <> CTRL_FORM then wszFunction = wszFunction & "_" & wszCtrlName
               wszFunction = wszFunction & "_" & pCtrl->Events(ii).wszEventName 
               
               ' Define the function prototype/declaration
               wszPrototype = "Function " & wszFunction & "( ByRef sender As " & _
               GetWinformsXClassName(pCtrl->ControlType) & ", ByRef e As EventArgs) As LRESULT" & vbcrlf

               ' Add the event/function to the list of declares
               wszAllDeclares = wszAllDeclares & "Declare " & wszPrototype
               
               ' If the Event/function does not already exist in code then create it now.
               pData = gdb2.dbFindFunction(wszFunction)
               if pData = 0 THEN
                  wszAllEvents = wszAllEvents & "''" & vbcrlf & "''" & vbcrlf & _
                                 wszPrototype & _
                                 sp(1) & "Function = 0" & vbcrlf & _
                                 "End Function" & vbcrlf & vbcrlf
               end if
            end if   
         next
      END IF
   next   
   
   ''
   ''  Generate the TYPE definition for the Form
   ''
   wszText = wszText & wszAllDeclares & iif(len(wszAllDeclares), vbcrlf, "") & _
   "type " & wszFormName & "Type extends wfxForm" & vbcrlf & _
   sp(1) & "private:"              & vbcrlf & _
   sp(2) & "temp as byte"          & vbcrlf & _
   sp(1) & "public:"               & vbcrlf & _
   sp(2) & "declare constructor"   & vbcrlf & _
   sp(2) & "' Controls"            & vbcrlf 
      for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
         pCtrl = pDoc->Controls.ItemAt(i)
         if pCtrl then
            if pCtrl->ControlType = CTRL_FORM then continue for
            wszCtrlName = GetControlProperty(pCtrl, "NAME")
            wszText = wszText & sp(2) & wszCtrlName & " As " & GetWinformsXClassName(pCtrl->ControlType) & vbcrlf
         END IF
      next   
   wszText = wszText & _
   "end type" & vbcrlf & vbcrlf 


   ''
   ''  Generate the Constructor that defines the properties for the form and controls
   ''
   wszText = wszText & _
   "constructor " & wszFormName & "Type" & vbcrlf 
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)

      if pCtrl then
         wszCtrlName = GetControlProperty(pCtrl, "NAME")

         wszCombinedName = "this" & iif(pCtrl->ControlType = CTRL_FORM, "", "." & wszCtrlName.Utf8)
         if pCtrl->ControlType <> CTRL_FORM then 
            wszText = wszText & sp(1) & wszCombinedName & ".Parent = @this" & vbcrlf
         END IF

         dim as long nLeft, nTop, nWidth, nHeight
         for ii as long = lbound(pCtrl->Properties) to ubound(pCtrl->Properties)
            wszPropName    = pCtrl->Properties(ii).wszPropName
            wszPropValue   = pCtrl->Properties(ii).wszPropValue
            wszPropDefault = pCtrl->Properties(ii).wszPropDefault
            nPropType      = pCtrl->Properties(ii).PropType
            
            select case ucase(wszPropName)
               CASE "LEFT"
                  nLeft = wszPropValue.ValLong
                  continue for
               case "TOP"
                  nTop = wszPropValue.ValLong
                  continue for
               case "WIDTH"
                  nWidth = wszPropValue.ValLong
                  continue for
               case "HEIGHT"
                  nHeight = wszPropValue.ValLong
                  continue for
               case "NAME"  ' don't output NAME property b/c it can't be changed programmically
                  continue for
               case "TABINDEX"
                  continue for   
            END SELECT

            ' Only output properties that are different than the default WinFormsX values
            if ucase(wszPropValue) <> ucase(wszPropDefault) then
               select case ucase(wszPropName)
                  case "TEXT", "TAG", "PASSWORDCHAR"            
                     ' Ensure that any embedded double quotes are escaped
                     wszPropValue = AfxStrReplace(wszPropValue, DQ, DQ & " & chr(34) & " & DQ)
                     wszPropValue = DQ & wszPropValue & DQ
                  case "ACCEPTBUTTON", "CANCELBUTTON"
                     wszPropValue = "@this." & wszPropValue 
               end select
               select case nPropType
                  CASE PropertyType.ColorPicker
                     if left(wszPropValue, 7) = "SYSTEM|" then
                        wszPropValue = "Colors.System" & mid(wszPropValue, 8)
                     elseif left(wszPropValue, 7) = "COLORS|" then
                        wszPropValue = "Colors." & mid(wszPropValue, 8)
                     elseif left(wszPropValue, 7) = "CUSTOM|" then
                        wszPropValue = mid(wszPropValue, 8) & "  ' custom color"
                     end if   
                  case PropertyType.FontPicker
                     wszPropValue = "New wfxFont(" & SetFontClassFromPropValue(wszPropValue) & ")"
               END SELECT
               wszText = wszText & sp(1) & wszCombinedName & "." & wszPropName.Utf8 & " = " & wszPropValue.Utf8 & vbcrlf
            end if
         NEXT
         wszText = wszText & sp(1) & wszCombinedName & ".SetBounds(" & nLeft & "," & nTop & "," & nWidth & "," & nHeight & ")" & vbcrlf

         for ii as long = lbound(pCtrl->Events) to ubound(pCtrl->Events)
            if pCtrl->Events(ii).bIsSelected then
               ' Determine the name of the event function
               wszFunction = wszFormName
               if pCtrl->ControlType <> CTRL_FORM then wszFunction = wszFunction & "_" & wszCtrlName
               wszFunction = wszFunction & "_" & pCtrl->Events(ii).wszEventName 
               wszText = wszText & sp(1) & wszCombinedName & ".On" & pCtrl->Events(ii).wszEventName & " = @" & wszFunction & vbcrlf
            end if
         next   

      end if
   next   
   
   
   ''
   ''  Add the controls to the form collection (we first need to sort based on TabIndex).
   ''
   dim pProp as clsProperty ptr
   dim TabOrder(pDoc->Controls.Count) as clsTabOrder
   dim nTabIndex as Long
   
   ' Load the temporary array
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)
      if pCtrl then
         TabOrder(i).pCtrl = pCtrl
         pProp = GetControlPropertyPtr(pCtrl, "TABINDEX")
         if pProp then
            nTabIndex = val(pProp->wszPropValue)
            pProp = GetControlPropertyPtr(pCtrl, "TABSTOP")
            if pProp = 0 then nTabIndex = 999999
         else
            nTabIndex = 999999   
         end if
         TabOrder(i).TabIndex = nTabIndex
      end if   
   next

   ' Sort the array using a simple bubble sort
   Dim As Long lb = LBound(TabOrder)
   Dim As Long ub = UBound(TabOrder)
   Dim As boolean bHasChanged
   Do
      bHasChanged = false
      For i as long = lb To ub - 1
         If TabOrder(i).TabIndex > TabOrder(i+1).TabIndex Then
            Swap TabOrder(i), TabOrder(i+1)
            bHasChanged = true
         End If
      Next
   Loop Until bHasChanged = false
    
   ' Finally, output the controls 
   for i as long = lb to ub 
      pCtrl = TabOrder(i).pCtrl
      if pCtrl then
         if pCtrl->ControlType <> CTRL_FORM then 
            wszCtrlName = GetControlProperty(pCtrl, "NAME")
            wszText = wszText & _
            sp(1) & "this.Controls.Add(ControlType." & GetControlName(pCtrl->ControlType) & ", @this." & wszCtrlName.Utf8 & ")" & vbcrlf
         end if
      end if
   next
   
   
   ''
   ''  Add the form to the application
   ''
   wszText = wszText & _
   sp(1) & "Application.Forms.Add(ControlType.Form, @this)" & vbcrlf & _
   "end constructor" & vbcrlf & vbcrlf

   wszText = wszText & _ 
   "dim shared " & wszFormName & " as " & wszFormName & "Type" & vbcrlf
   
   ' Get the file positions of any existing start and end tags
   dim as hwnd hEdit = pDoc->hWindow(0)
   SciExec( hEdit, SCI_TARGETWHOLEDOCUMENT, 0, 0)
   nStartTag = SciExec( hEdit, SCI_SEARCHINTARGET, Len(sStartTag), strptr(sStartTag))
   SciExec( hEdit, SCI_TARGETWHOLEDOCUMENT, 0, 0)
   nEndTag   = SciExec( hEdit, SCI_SEARCHINTARGET, Len(sEndTag), strptr(sEndTag))
   
   sReplaceText = sStartTag & vbcrlf & _ 
                  wszText.Utf8 & _  ' Convert the unicode CWSTR to UTF-8
                  sEndTag & vbcrlf 
                     
   ' Save current position in order to restore it later
   dim as long curPos = SendMessage( hEdit, SCI_GETCURRENTPOS, 0, 0)
    
   if (nStartTag = -1) andalso (nEndTag = -1) then
      ' If this is the first time generating code then add the boilerplate code
      ' to start the application and show the form.
      dim as CWSTR wszHeader = _
      "#include once " & DQ & "WinFormsX\WinFormsX.bi" & DQ & vbcrlf 
      if gApp.IsProjectActive = false then
         wszHeader = wszHeader & _
         "' You should always include a resource file that references a valid manifest.xml" & vbcrlf & _
         "' file otherwise your application will not properly display Windows themed controls." & vbcrlf & _
         "' Sample resource.rc and manifest.xml files can be found in the WinFormsX download." & vbcrlf & _
         "' The following WinFBE directive includes the resource in your application. If you" & vbcrlf & _
         "' are using WinFBE's project management features then omit the following line because" & vbcrlf & _
         "' one will be generated automatically." & vbcrlf & _
         "'#RESOURCE " & DQ & "resource.rc" & DQ & vbcrlf & vbcrlf 
      END IF
      sReplaceText = wszHeader & sReplaceText & vbcrlf & _
      "''" & vbcrlf & _
      "''  Remove the following Application.Run code if it used elsewhere in your application." & vbcrlf & _
      "Application.Run(" & wszFormName & ")" & vbcrlf & vbcrlf
      pDoc->AppendText(sReplaceText)
   else
      ' Ensure that both tags were found
      if (nStartTag <> -1) andalso (nEndTag <> -1) then
         SciExec( hEdit, SCI_SETTARGETSTART, nStartTag, 0)
         SciExec( hEdit, SCI_SETTARGETEND, nEndTag + len(sEndTag) + 2, 0)   ' add 2 to account for vbcrlf
         SciExec( hEdit, SCI_REPLACETARGET, len(sReplaceText), strptr(sReplaceText))
      end if
   end if

   if len(wszAllEvents) then
      pDoc->AppendText(wszAllEvents)
      ParseDocument(gApp.GetActiveProjectIndex(), pDoc, "")
   END IF
   
   pDoc->bRegenerateCode = false
   
   function = 0
end function

