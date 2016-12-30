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
' Show codetips
' ========================================================================================
function ShowCodetip() as BOOLEAN

   dim as hwnd hEdit
   dim as long curPos, tipPos, i
   Dim As Long nCount = gApp.GetDocumentCount()
   dim as string buffer, CodetipStr
   dim pDoc as clsDocument ptr
   Dim pList As FUNCTION_TYPE Ptr
   
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then exit function

   hEdit = pDoc->hWindow
   
   ' Retrieve the position
   curPos = SendMessage(hEdit, SCI_GETCURRENTPOS, 0, 0)
   tipPos = curPos - 1   ' Skip the "("

   ' Retrieve the word (keys must be searched using uppercase)
   buffer = ucase(TRIM(pDoc->GetWord(curPos - 1)))
   IF LEN(buffer) THEN
      ' Search the base builtin FB keywords first.
      if gCodetips.FindItem(buffer) then
         CodetipStr = gCodetips.FoundData
         SendMessage( hEdit, SCI_CALLTIPSHOW, tipPos - LEN(buffer), cast(lPARAM, strptr(CodetipStr)))
         return TRUE
      END IF
      ' Search all of the documents for user defined Sub/Functions   
      For i = 0 To nCount - 1
         pDoc = gApp.GetDocumentPtr(i)
         If pDoc Then 
            pList = pDoc->FnListPtr
            Do Until pList = 0
               if ucase(pList->zFnName) = buffer THEN
                  CodetipStr = pList->zFnName & pList->zFnDeclare
                  SendMessage( hEdit, SCI_CALLTIPSHOW, tipPos - LEN(buffer), cast(lPARAM, strptr(CodetipStr)))
                  return TRUE
               END IF
               pList = pList->pListNext
            Loop      
         End If   
      Next
   END IF
   function = FALSE
END FUNCTION

' ========================================================================================
' Show Autocomplete list
' ========================================================================================
function ShowAutocompleteList() as BOOLEAN

   dim as hwnd hEdit
   dim as long curPos
   dim as string buffer, sList
   dim pDoc as clsDocument ptr
   
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then exit function

   hEdit = pDoc->hWindow
   
   ' Retrieve the position
   curPos = SendMessage(hEdit, SCI_GETCURRENTPOS, 0, 0)

   ' Retrieve the word
   buffer = TRIM(pDoc->GetWord(curPos - 1))
   IF LEN(buffer) THEN
      ' Create the space separated data type list
      sList = "Boolean Byte Double Integer Long LongInt Short Single " & _
              "String UByte UInteger UShort WString ZString"
      Select Case gConfig.KeywordCase
         Case 0:  sList = lcase(sList)
         Case 1:  sList = ucase(sList)
      End Select            
      
      ' If an autocomplete list is alrady active then continue to search the list
      SendMessage(hEdit, SCI_AUTOCSETIGNORECASE, 1, 0)
      SendMessage(hEdit, SCI_AUTOCSETMAXHEIGHT, 15, 0)
      if SendMessage(hEdit, SCI_AUTOCACTIVE, 0, 0) then
         SendMessage(hEdit, SCI_AUTOCSHOW, len(buffer), cast(LPARAM, strptr(sList)))
      else
         ' Determine if we need to display the autocomplete list.
         if ucase(buffer) = "AS" THEN
            SendMessage(hEdit, SCI_AUTOCSHOW, 0, cast(LPARAM, strptr(sList)))
         end if
      end if
   END IF

   function = FALSE
END FUNCTION

