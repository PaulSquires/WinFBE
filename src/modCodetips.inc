'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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
' Format the codetip prior to showing it
' ========================================================================================
function FormatCodetip( byval strCodeTip as string ) as STRING
   dim as long p

   ' remove multiple duplicate spaces
   strCodeTip = AfxStrReplace(strCodeTip, "  ", " ")

   ' If this is only a small CodeTip then don't bother trying
   ' to format it into multiple lines.
   if len(strCodeTip) < 75 THEN return strCodeTip

   IF INSTR(strCodeTip, "( ") THEN
      strCodeTip = AfxStrReplace(strCodeTip, "( ", "(  ")
   ELSEIF INSTR(strCodeTip, "(") THEN
      strCodeTip = AfxStrReplace(strCodeTip, "(", "(  ")
   END IF
   p = INSTR(strCodeTip, "(")
   IF p THEN
      IF MID(strCodeTip, p + 1, 1) <> ")" AND MID(strCodeTip, p + 1, 2) <> " )" THEN
         strCodeTip = AfxStrReplace(strCodeTip, "(", "( _" & vblf)
         IF INSTR(strCodeTip, " )") THEN
            strCodeTip = AfxStrReplace(strCodeTip, ")", "_" & vblf & ")")
         ELSEIF INSTR(strCodeTip, ")") THEN
            strCodeTip = AfxStrReplace(strCodeTip, ")", " _" & vblf & ")")
         END IF
      END IF
   END IF
   strCodeTip = AfxStrReplace(strCodeTip, ", ", ",")
   strCodeTip = AfxStrReplace(strCodeTip, " ,", ",")
   strCodeTip = AfxStrReplace(strCodeTip, ",", ", _" & vblf & "  ")
   
   FUNCTION = strCodeTip
end function


' ========================================================================================
' Show codetips
' ========================================================================================
function ShowCodetip() as BOOLEAN

   dim pDoc as clsDocument ptr
   dim pData as DBENGINE_DATA ptr

   dim as string buffer, codeTip
   dim as long curPos, tipPos, i

   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then exit function

   ' Retrieve the position
   curPos = SciExec(pDoc->hWindow, SCI_GETCURRENTPOS, 0, 0)
   tipPos = curPos - 1   ' Skip the "("

   ' Retrieve the word (keys must be searched using uppercase)
   buffer = TRIM(pDoc->GetWord(curPos - 1))
   IF LEN(buffer) THEN
      ' Search for a matching function name (user and builtin FB keywords)
      pData = gdb.dbFindFunctionName(buffer)
      if pData THEN
         ' Codetip is stored in data2
         codeTip = FormatCodetip(pData->data2)
         SciExec( pDoc->hWindow, SCI_CALLTIPSHOW, tipPos - LEN(buffer), strptr(codeTip))
         return TRUE
      END IF
       
      'TODO: Search the WinAPI declares if this is a Windows api program (ie. "windows.bi" exists)
      'TODO: Search depending on whether gApp.IsUnicodeCodetips is TRUE/FALSE
      
   END IF
   function = FALSE
END FUNCTION

' ========================================================================================
' Show Autocomplete list
' ========================================================================================
function ShowAutocompleteList() as BOOLEAN

   dim as hwnd hEdit
   dim as long curPos, nCount, i
   dim as string buffer
   static as string sList

   dim pDoc as clsDocument ptr
   dim pData as DBENGINE_DATA ptr
   
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then exit function

   hEdit = pDoc->hWindow
   
   ' Retrieve the position
   curPos = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)

   ' Retrieve the word
   buffer = TRIM(pDoc->GetWord(curPos - 1))
   IF LEN(buffer) THEN
      if SciExec(hEdit, SCI_AUTOCACTIVE, 0, 0) then
         SciExec(hEdit, SCI_AUTOCSHOW, len(buffer), strptr(sList))
      else
         ' Determine if we need to display the autocomplete list.
         if ucase(buffer) = "AS" THEN
            ' Create the space separated data type list
            sList = "Boolean Byte Double Integer Long LongInt Short Single " & _
                    "String UByte UInteger UShort WString ZString"
            'Select Case gConfig.KeywordCase
            '   Case 0:  sList = lcase(sList)
            '   Case 1:  sList = ucase(sList)
            '   Case 2:  ' already mixed case
            'End Select            
            ' Add all of the TYPE definitions that were found
            gdb.dbRewind()
            do 
               pData = gdb.dbGetNext()
               if pData = 0 THEN exit do
               if pData->dataType = DB_TYPE_NAME THEN
                  sList = sList & " " & pData->data1
               END IF
            loop
            ' If an autocomplete list is already active then continue to search the list
            SciExec(hEdit, SCI_AUTOCSETORDER, 1, 0)  ' List must be sorted alphabetically
            SciExec(hEdit, SCI_AUTOCSETIGNORECASE, 1, 0)
            SciExec(hEdit, SCI_AUTOCSETMAXHEIGHT, 15, 0)
            SciExec(hEdit, SCI_AUTOCSHOW, 0, strptr(sList))
         else
            sList = ""
         end if
      end if
   END IF

   function = FALSE
END FUNCTION



' ========================================================================================
' Show TYPE Autocomplete list
' ========================================================================================
function ShowTYPEAutocompleteList() as BOOLEAN
   dim as hwnd hEdit
   dim as long curPos, nCount, i
   dim as string buffer
   static as string sList

   dim pData as DBENGINE_DATA ptr
   dim pDoc as clsDocument ptr
   
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then exit function

   hEdit = pDoc->hWindow
     
   ' Retrieve the position
   curPos = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)
    
   ' Retrieve the TYPE variable name
   buffer = TRIM(pDoc->GetWord(curPos - 1), any " .")
   IF LEN(buffer) THEN
      if SendMessage(hEdit, SCI_AUTOCACTIVE, 0, 0) then
         SendMessage(hEdit, SCI_AUTOCSHOW, len(buffer), cast(LPARAM, strptr(sList)))
      else
         ' Need to determine what TYPE this variable was DIM'd as.
         ' Look up the variable
         pData = gdb.dbFindVariableName(buffer)
         if pData THEN
            pData = gdb.dbFindTypeName(pData->data2)
            if pData THEN
               ' Create the space separated data type list
               'sList = AfxStrReplace(pData->data2, DB_DELIM, " ")
               sList = pData->data2
               ' If an autocomplete list is already active then continue to search the list
               SciExec(hEdit, SCI_AUTOCSETSEPARATOR, 124, 0)  ' Pipe symbol as separator
               SciExec(hEdit, SCI_AUTOCSETORDER, 1, 0)  ' List must be sorted alphabetically
               SciExec(hEdit, SCI_AUTOCSETIGNORECASE, 1, 0)
               SciExec(hEdit, SCI_AUTOCSETMAXHEIGHT, 15, 0)
               SciExec(hEdit, SCI_AUTOCSHOW, 0, strptr(sList))
            else
               sList = ""
            end if
         END IF
      end if
   END IF

   function = FALSE
END FUNCTION
