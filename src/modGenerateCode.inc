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
   dim as CWSTR wszText, wszFormName, wszCtrlName, wszPropName, wszPropValue
   dim as CWSTR wszFunction, wszAllEvents, wszPrototype, wszAllDeclares
   dim as string sStartTag, sEndTag, sReplaceText
   dim as CWSTR DQ = chr(34)
   dim as long nStartTag, nEndTag
   dim pData as DB2_DATA ptr
   

   wszFormName = GetFormName(pDoc)
   sStartTag = "' WINFBE_CODEGEN_START"
   sEndTag   = "' WINFBE_CODEGEN_END"

   ' Generate the Declares for all of the Events for the Form & Controls
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
               GetWinformsXClassName(pCtrl->ControlType) & ", ByRef e As EventArgs ) As LRESULT" & vbcrlf

               ' Add the event/function to the list of declares
               wszAllDeclares = wszAllDeclares & "Declare " & wszPrototype
               
               ' If the Event/function does not already exist in code then create it now.
               pData = gdb2.dbFindFunction(wszFunction)
               if pData = 0 THEN
                  wszAllEvents = wszAllEvents & "''" & vbcrlf & "''" & vbcrlf & _
                                 wszPrototype & _
                                 "   Function = 0" & vbcrlf & _
                                 "End Function" & vbcrlf & vbcrlf
               end if
            end if   
         next
      END IF
   next   
   
   wszText = wszText & wszAllDeclares & iif(len(wszAllDeclares), vbcrlf, "") & _
   "type " & wszFormName & "Type extends wfxForm" & vbcrlf & _
   "   private:"                            & vbcrlf & _
   "      temp as byte"                     & vbcrlf & _
   "   public:"                             & vbcrlf & _
   "      declare constructor"              & vbcrlf & _
   "      ' Controls"                       & vbcrlf 
         for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
            pCtrl = pDoc->Controls.ItemAt(i)
            if pCtrl then
               if pCtrl->ControlType = CTRL_FORM then continue for
               wszCtrlName = GetControlProperty(pCtrl, "NAME")
               wszText = wszText & space(6) & wszCtrlName & " As " & GetWinformsXClassName(pCtrl->ControlType) & vbcrlf
            END IF
         next   
   wszText = wszText & _
   "end type" & vbcrlf & vbcrlf 


   wszText = wszText & _
   "constructor " & wszFormName & "Type"      & vbcrlf & _
   "   Application.Forms.Add(ControlType.Form, @this)" & vbcrlf 
   for i as long = pDoc->Controls.ItemFirst to pDoc->Controls.ItemLast
      pCtrl = pDoc->Controls.ItemAt(i)

      if pCtrl then
         wszCtrlName = GetControlProperty(pCtrl, "NAME")
         wszText = wszText & "   with this" & iif(pCtrl->ControlType = CTRL_FORM, "", "." & wszCtrlName.Utf8) & vbcrlf
         if pCtrl->ControlType <> CTRL_FORM then 
            wszText = wszText & "      .Parent = @this" & vbcrlf
         END IF

         dim as long nLeft, nTop, nWidth, nHeight
         for ii as long = lbound(pCtrl->Properties) to ubound(pCtrl->Properties)
            wszPropName  = pCtrl->Properties(ii).wszPropName
            wszPropValue = pCtrl->Properties(ii).wszPropValue
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
               case "NAME"
                  wszCtrlName = wszPropValue
                  wszPropValue = DQ & wszPropValue & DQ
               case "TEXT"            
                  wszPropValue = DQ & wszPropValue & DQ
               case "STYLE", "EXSTYLE"
                  continue for   
            END SELECT
            wszText = wszText & "      ." & wszPropName.Utf8 & " = " & wszPropValue.Utf8 & vbcrlf
         NEXT
         wszText = wszText & "      .SetBounds(" & nLeft & "," & nTop & "," & nWidth & "," & nHeight & ")" & vbcrlf

         for ii as long = lbound(pCtrl->Events) to ubound(pCtrl->Events)
            if pCtrl->Events(ii).bIsSelected then
               ' Determine the name of the event function
               wszFunction = wszFormName
               if pCtrl->ControlType <> CTRL_FORM then wszFunction = wszFunction & "_" & wszCtrlName
               wszFunction = wszFunction & "_" & pCtrl->Events(ii).wszEventName 
               wszText = wszText & "      .On" & pCtrl->Events(ii).wszEventName & " = @" & wszFunction & vbcrlf
            end if
         next   

         wszText = wszText & "   end with" & vbcrlf 
         if pCtrl->ControlType <> CTRL_FORM then 
            wszText = wszText & _
               "   this.Controls.add(controltype." & GetControlName(pCtrl->ControlType) & ", @this." & wszCtrlName.Utf8 & ")" & vbcrlf
         end if

      end if
   next   
   wszText = wszText & _
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
      sReplaceText = _
      "#include once " & DQ & "WinFormsX\WinFormsX.bi" & DQ & vbcrlf & _
      "'#RESOURCE " & DQ & "resource.rc" & DQ & vbcrlf & vbcrlf & _
      sReplaceText & vbcrlf & _
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

