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


' WinFBE Version 3.0.2+ requires that older version form file formats get upgraded
' to the new json format and also separated out from the source code file. There
' will now be two files: (1) *.frm for the json form definitions, and (2) *.inc/bas
' for the actual form code.

#include once "mod302Upgrade.bi"

' ========================================================================================
' Upgrade pre-version 3.02 form files to new format
' ========================================================================================
public function FormUpgrade302Format( byval pDoc as clsDocument ptr ) as boolean

   if pDoc = 0 then exit function
   
   ' WinFBE Version 3.0.2+ requires that older version form file formats get upgraded
   ' to the new json format and also separated out from the source code file. There
   ' will now be two files: (1) *.frm for the json form definitions, and (2) *.inc/bas
   ' for the actual form code.
   if (ConvertWinFBEversion(pDoc->wszFormVersion) >= ConvertWinFBEversion("3.0.2")) orelse _
      (ConvertWinFBEversion(pDoc->wszFormVersion) = 0) then  'b/c upgraded form files do not have version loaded yet
      ' We are already using new form file format so simply exit
      return false
   end if
   
   gApp.PreventActivateApp = true
   
   ' Commence conversion to the new json file format
   dim as CWSTR wszCodeFilename = pDoc->DiskFilename
   dim as CWSTR wszFormFilename = wszCodeFilename & ".design"
   
   pDoc->DesignerFilename = wszFormFilename
   if pDoc->SaveFormJSONdata() = false then return false
   
   ' The new design Form file has now been created so we can remove all code in the
   ' existing code file.
   pDoc->wszFormMetaData = ""
   pDoc->wszFormCodeGen = ""
   pDoc->SaveFile
  
   gApp.PreventActivateApp = false

   return true  ' new format was created
end function


