
' ========================================================================================
' Show codetips
' ========================================================================================
function ShowCodetip() as BOOLEAN

   dim as hwnd hEdit
   dim as long curPos, tipPos     
   dim as string buffer, CodetipStr
   dim pDoc as clsDocument ptr
   
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 Then exit function

   hEdit = pDoc->hWindow
   
   ' Retrieve the position
   curPos = SendMessage(hEdit, SCI_GETCURRENTPOS, 0, 0)
   tipPos = curPos - 1   ' Skip the "("

   ' Retrieve the word
   buffer = TRIM(pDoc->GetWord(curPos - 1))
'   IF LEN(buffer) THEN
'      'CodetipStr = TRIM$(CSED_GetCodetip(buffer))
'      IF LEN(CodetipStr) THEN
'         SendMessage( hEdit, SCI_CALLTIPSHOW, tipPos - LEN(buffer), cast(lPARAM, strptr(CodetipStr)))
'         FUNCTION = TRUE
'      END IF
'   END IF

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
      sList = "Byte Double Integer Long LongInt Short Single " & _
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

