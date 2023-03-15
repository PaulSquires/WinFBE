'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2023 Paul Squires, PlanetSquires Software
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

#include once "modCodetips.bi"
#include once "clsDocument.bi"


' ========================================================================================
' Format the codetip prior to showing it
' ========================================================================================
function FormatCodetip( byval strCodeTip as string ) as string
   dim as long p

   ' remove multiple duplicate spaces 
   strCodeTip = AfxStrReplace(strCodeTip, "  ", " ") 

   ' If this is only a small CodeTip then don't bother trying
   ' to format it into multiple lines.
   if len(strCodeTip) < 75 then return strCodeTip

   if instr(strCodeTip, "( ") then
      strCodeTip = AfxStrReplace(strCodeTip, "( ", "(  ")
   elseif instr(strCodeTip, "(") then
      strCodeTip = AfxStrReplace(strCodeTip, "(", "(  ")
   end if
   p = instr(strCodeTip, "(")
   if p then
      if mid(strCodeTip, p + 1, 1) <> ")" andalso mid(strCodeTip, p + 1, 2) <> " )" then
         strCodeTip = AfxStrReplace(strCodeTip, "(", "( _" & vblf)
         IF instr(strCodeTip, " )") then
            strCodeTip = AfxStrReplace(strCodeTip, ")", "_" & vblf & ")")
         elseif instr(strCodeTip, ")") then
            strCodeTip = AfxStrReplace(strCodeTip, ")", " _" & vblf & ")")
         end if
      end if
   end if
   strCodeTip = AfxStrReplace(strCodeTip, ", ", ",")
   strCodeTip = AfxStrReplace(strCodeTip, " ,", ",")
   strCodeTip = AfxStrReplace(strCodeTip, ",", ", _" & vblf & "  ")
   
   function = strCodeTip
end function


' ========================================================================================
' Change everything between quotes to # symbols
' ========================================================================================
function MaskStringCharacters( byval st as string) as string
   ' Iterate the line and change everything between quotes to # symbols. This
   ' ensures that we correctly deal with strings that have embedded single 
   ' quote characters.
   dim as long i 
   dim as Boolean binstring = false
   for i = 0 to len(st) - 1
      if st[i] = 34 then binstring = not(binstring)
      if binstring then  
         if st[i] <> 34 then st[i] = 35   ' # symbol
      end if
   NEXT
   function = st
end function


' ========================================================================================
' Removes a single line comment.
' ========================================================================================
function RemoveComments( byval st as string) as string
   function = AfxStrExtract( 1, st, "'")
end function  

' ========================================================================================
' Determine what variable relates to the "With" statement
' ========================================================================================
function DetermineWithVariable( byval pDoc as clsDocument ptr) as string
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
   dim as long nCurLine = pDoc->GetCurrentLineNumber - 1
   dim as string sLine, sWithVariable 
   
   for i as long = nCurLine to 0 step -1
      sLine = ltrim(pDoc->GetLine(i))
      ' Remove double spaces and replace TABs with single space
      sLine = AfxStrShrink(sLine, chr(32,9))
      if left(ucase(sLine), 5) = "WITH " then
         sWithVariable = AfxStrParse(sLine, 2, " ")
         exit for
      end if
   next
   
   function = sWithVariable
end function    



' ========================================================================================
' Take the current line and determine what variable is being referenced.
' Dereferences expressions like:  g.MyFunction, g.ot.rc, g->MyFunction
' Used for determining what popup autocomplete list or codetip to display.
' ========================================================================================
function DereferenceLine( _
            byval pDoc as clsDocument ptr, _
            byval sTrigger as String, _
            byval nPosition as long _
            ) as DB2_DATA ptr
   
   if pDoc = 0 then exit function
   
   dim pData     as DB2_DATA ptr
   dim pDataTYPE as DB2_DATA ptr
   dim pLastData as DB2_DATA ptr
   
   dim as string sCurrentFunction, sTypeName

   dim as hwnd hEdit   = pDoc->hWndActiveScintilla
   dim as long curPos  = nPosition  
   dim as long nCol    = SciExec(hEdit, SCI_GETCOLUMN, curPos, 0) 
   dim as string sLine = trim(left(pDoc->GetLine(pDoc->GetCurrentLineNumber), nCol+1), any chr(32,9))

   ' Comment out any comment line so the popup isn't activated when in a comment. 
   sLine = MaskStringCharacters(sLine)
   sLine = RemoveComments(sLine)

   ' Test to ensure that the incoming sTrigger character still exists after
   ' the commenting of the line. If it does then we know that the line
   ' is good and we can proceed to dereference elements on the line.
   ' The identifer that triggers an autocomplete list for a TYPE variable
   ' is either a dot "." or a pointer dereference "->"
   if right(sLine, len(sTrigger)) <> sTrigger then exit function
   sLine = rtrim(sLine, sTrigger)


   ' Proceed to parse the line. Find the beginning of the line in order to
   ' isolate the expression to dereference. The expression might contain an
   ' array reference so that would have to be removed, however the "(" can
   ' also signify the start position of the expression. For example:
   ' st = myFunction(myvar(3).myvalue)
   ' Simply reverse searching for the "(" would fail because the array "()"
   ' also needs to be taken into account. Therefore, search in reverse and
   ' take note of when the ")" is encountered in order to then match it with
   ' an "(".
   ' Need to alo handle situations where there are multiple () array references
   ' in the expression. For example:  ListView1.Item(10).rc(5)
   dim as long nStart, nEnd
   do
      do
         nEnd = instrRev(sLine, ")", -1)
         if nEnd = 0 then exit do
         nStart = instrRev(sLine, "(", nEnd)
         if nStart = 0 then exit do
         ' Remove any array parenthesis.
         sLine = left(sLine, nStart-1) & mid(sLine, nEnd+1)
      LOOP
      if (nStart = 0) or (nEnd = 0) then exit do
   loop
   
   ' All of the array parenthesis should now be removed so now we can simply
   ' search for the start of the line.
   dim as long i = instrRev(sLine, any " (*[&@", -1)
   if i then sLine = mid(sLine, i+1)
   
   ' Need to check if this line is part of a WITH / END WITH block.
   if (len(sLine) = 0) or (left(sLine, 1) = ".") then
      sLine = DetermineWithVariable(pDoc) & sLine
   end if
   
   ' Make it easier to parse by converting line to uppercase
   sLine = ucase(sLine)
   
   ' Convert all "->" pointer references to "." dots to make parsing easier
   sLine = AfxStrReplace(sLine, "->", ".")
   
   ' Determine what sub/function (if any) that we are in. This function will
   ' retrieve the current standalone function, or the fully qualified name of 
   ' a class/type function.
   ' eg.   clsType.MyFunctionName      ' inside a class/type function
   ' eg.   MyFunctionName              ' standalone function
   pDoc->GetCurrentFunctionName( sCurrentFunction, 0 )

   dim as long numParts = AfxStrParseCount(sLine, ".")
   dim parts(1 to numParts) as string
   for i as long = 1 to numParts
      parts(i) = AfxStrParse(sLine, i, ".")
   next

   dim as string sParent = sCurrentFunction
   dim as string sSearch
   dim as string sLookFor 
   dim as long curPart = 1

   do until curPart > numParts
      sLookFor = parts(curPart)
      
      if pLastData then
         sParent = pLastData->VariableType
      end if

      '' ATTEMPT TO MATCH LOCAL VARIABLE WITHIN CURRENT FUNCTION
      pData = gdb2.dbFindVariable( sParent, sLookFor )
      if pData then
         pLastData = pData
      else   
         '' ATTEMPT TO MATCH GLOBAL VARIABLE   
         pData = gdb2.dbFindVariable( "", sLookfor )
         if pData then
            pLastData = pData
         end if
      end if

      '' ATTEMPT TO MATCH FUNCTION NAME
      if pData = 0 then
         if len(sParent) then
            sSearch = sParent & "." & sLookFor
            pData = gdb2.dbFindFunctionTYPE( sParent, sSearch )
            if pData = 0 then
               ' Determine if there is a TYPE Extends that needs to be checked
               pData = gdb2.dbFindTYPE( sParent )
               if pData then
                  sParent = pData->TypeExtends
                  sSearch = sParent & "." & sLookFor
                  pData = gdb2.dbFindFunctionTYPE( sParent, sSearch )
               end if   
            end if   
         end if
         if pData then
            pLastData = pData
         end if
      end if

      if pData = 0 then
         pData = gdb2.dbFindFunction( sLookFor )
         if pData then
            pLastData = pData
         end if
      end if
     
      ' If we did not get anymore unresolved matches then return
      ' the most recent match.
      if pData = 0 then exit do

      curPart = curPart + 1
   loop

   return pLastData

end function

   
' ========================================================================================
' Show codetips
' ========================================================================================
function ShowCodetip( _
            byval pDoc as clsDocument ptr, _
            byval bCommaTrigger as Boolean = false _
            ) as boolean

   if gConfig.CodeTips = false then exit function
   
   dim pData as DB2_DATA ptr
   dim as string codeTip
   dim as hwnd hEdit  = pDoc->hWndActiveScintilla
   dim as long curPos = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0) 

   if bCommaTrigger then
      ' It is possible that a codetip had been displayed but the user moved off of the
      ' current line and thereby closed the codetip. If user goes back to the line and
      ' presses a comma to enter another parameter for function, then we should do a 
      ' test to see if we can redisplay the codetip popup again.
      ' Need to also handle situations where there are multiple () array references
      ' in the expression. For example:  ListView1.Item(10).rc(5)
      dim as long nCol = SciExec(hEdit, SCI_GETCOLUMN, curPos, 0) 
      dim as string sLine = rtrim(left(pDoc->GetLine(pDoc->GetCurrentLineNumber), nCol+1))
      dim as boolean bInClosed
      for i as long = len(sLine) -1 to 0 step -1
         if sLine[i] = asc(")") then 
            bInClosed = true: continue for
         end if
         if sLine[i] = asc("(") then
            if bInClosed = true then
               bInClosed = false
            else
               exit for
            end if
         end if         
         curPos = curPos - 1
      next
   end if

   pData = DereferenceLine(pDoc, "(", curPos-1)
   if pData then codeTip = pData->CallTip

   if len(codeTip) then 
      codeTip = FormatCodetip(codeTip)
      SciExec( hEdit, SCI_CALLTIPSHOW, curPos, strptr(codeTip))
      return true
   else
      pDoc->AutoCompleteType = AUTOCOMPLETE_NONE
   end if
   
   return false
end function

  
' ========================================================================================
' Display the actual Autocomplete popup list window
' ========================================================================================
function ShowAutoCompletePopup( _
            byval hEdit as hwnd, _
            byref sList as string _
            ) as Long
   if len(sList) = 0 then exit function
   dim as string sFillups = "(=."
                                
   SciExec(hEdit, SCI_AUTOCSETFILLUPS, 0, strptr(sFillups))  ' characters that also select and close the popup
   SciExec(hEdit, SCI_AUTOCSETSEPARATOR, 124, 0)             ' Pipe symbol as separator
   SciExec(hEdit, SCI_AUTOCSETORDER, 1, 0)                   ' List must be sorted alphabetically
   SciExec(hEdit, SCI_AUTOCSETIGNORECASE, 1, 0)
   SciExec(hEdit, SCI_AUTOCSETMAXHEIGHT, 9, 0)
   SciExec(hEdit, SCI_AUTOCSHOW, 0, strptr(sList))
   SciExec(hEdit, SCI_AUTOCSETOPTIONS, SC_AUTOCOMPLETE_FIXED_SIZE, 0)
                            
   function = 0
end function

       
' ========================================================================================
' Don't add duplicates in the Autocomplete list    
' ========================================================================================
function ExistsInAutocompleteList( _
            byref sList as string, _
            byref sMatch as string _
            ) as boolean
   if rtrim(sMatch) = "" then return true
   dim as string sLookFor = "|" & ucase(sMatch) & "|"
   if instr( "|" & ucase(sList) & "|", sLookFor ) then return true
   return false
end function


' ========================================================================================
' Show Autocomplete list    
' ========================================================================================
function ShowAutocompleteList( byval Notification as long = 0) as BOOLEAN
  
   IF gConfig.AutoComplete = false then exit function
   
   dim as long curPos, nCol, nLenMatchWord, ub, iNextType
   dim as string sWord, sList, st, sDot, sLookFor, sElement, sProp
   dim as Boolean bIsTHIS, bTypesOnly, bTypesNotPreloaded
   
   dim pDoc as clsDocument ptr
   dim pData as DB2_DATA ptr
   
   pDoc = gTTabCtl.GetActiveDocumentPtr()
   If pDoc = 0 then exit function
   
   dim as hwnd hEdit = pDoc->hWndActiveScintilla
             
   ' Retrieve the position
   curPos = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0) 
   nCol   = SciExec(hEdit, SCI_GETCOLUMN, curPos, 0) 
  
   dim as string sLine = left( pDoc->GetLine(pDoc->GetCurrentLineNumber), nCol )
   dim as string sLine_ucase

   ' Get the current entered word. We get the text between the "." or ">" and the current
   ' editing position. We do this rather than pDoc->GetWord because text may already
   ' exist after the current editing position.
   pDoc->sMatchWord = ""
   for i as long = len(sLine) to 1 step -1
      dim as string ch = mid(sLine, i, 1)
      select case ch
         case ".", ">", " "
            exit for
         case else
            pDoc->sMatchWord = ucase(ch) & pDoc->sMatchWord   
      end select
   next
   nLenMatchWord = len(pDoc->sMatchWord )
               
   ' If an autocomplete is active but now there is no match word then
   ' a space or tab must have been pressed so autocomplete cancel and now exit.
   if SciExec(hEdit, SCI_AUTOCACTIVE, 0, 0) then
      if nLenMatchWord = 0 then 
         SciExec( hEdit, SCI_AUTOCCANCEL, 0, 0)
         exit function
      end if   
   end if


   ' Get the styling of the current line to determine if we are in a 
   ' multiline or single line comment block then abort the autoinsert.
   select case SciExec(hEdit, SCI_GETSTYLEAT, curPos, 0)
      case SCE_B_MULTILINECOMMENT, SCE_B_COMMENT
         exit function
   end select

   ' If no active autocomplete then create the list based on the current underlying match word.
   ' also continue to display the popup if we have simply backspaced while a popup is open.
   dim as boolean bShowPopup = false
   if (notification = SCN_AUTOCCHARDELETED) then
      bShowPopup = true
   else   
      if (SciExec(hEdit, SCI_AUTOCACTIVE, 0, 0) = 0) andalso _
         (pDoc->AutoCompleteType = AUTOCOMPLETE_NONE) then
         bShowPopup = true
      end if
   end if

   if bShowPopup = true then     
      ' Comment out any comment line so the popup isn't activated when in a comment.
      sLine = MaskStringCharacters(sLine)
      sLine = RemoveComments(sLine)
      if (notification <> SCN_AUTOCCHARDELETED) then
         pDoc->sMatchWord = ""
         nLenMatchWord = 0         
      end if
      sLine_ucase = ucase(sLine)

      if right(sLine_ucase, 4) = " AS " then 
         ' okay
         bTypesOnly = false
         bTypesNotPreloaded = false  ' all types
         pDoc->AutoCompleteType = AUTOCOMPLETE_DIM_AS

      elseif right(sLine_ucase, 1) = "." then 
         sDot = "."  ' okay
         pDoc->AutoCompleteType = AUTOCOMPLETE_TYPE

      elseif right(sLine_ucase, 2) = "->" then
         sDot = "->"   
         pDoc->AutoCompleteType = AUTOCOMPLETE_TYPE

      elseif right(sLine_ucase, 1) = "(" then 
         ShowCodetip(pDoc)
         exit function 

      elseif right(sLine_ucase, 1) = "," then 
         ShowCodetip(pDoc, true)
         exit function 

      elseif right(sLine_ucase, 5) = " NEW " then 
         ' okay
         bTypesOnly = false
         bTypesNotPreloaded = false  ' all types
         pDoc->AutoCompleteType = AUTOCOMPLETE_DIM_AS

      elseif right(sLine_ucase, 12) = "CONSTRUCTOR " then 
         ' okay
         bTypesOnly = true
         bTypesNotPreloaded = true  ' only TYPES that are in our source.
         pDoc->AutoCompleteType = AUTOCOMPLETE_DIM_AS

      elseif right(sLine_ucase, 11) = "DESTRUCTOR " then 
         ' okay
         bTypesOnly = true
         bTypesNotPreloaded = true
         pDoc->AutoCompleteType = AUTOCOMPLETE_DIM_AS

      elseif right(sLine_ucase, 9) = " EXTENDS " then 
         ' okay
         bTypesOnly = true
         bTypesNotPreloaded = false  ' all types
         pDoc->AutoCompleteType = AUTOCOMPLETE_DIM_AS
      end if

      ' Check if it is a "FOR i AS LONG" type of statement
      if instr(sLine, "FOR ") then 
         SciExec( hEdit, SCI_AUTOCCANCEL, 0, 0)
         exit function
      end if
 
      if pDoc->AutoCompleteType <> AUTOCOMPLETE_NONE then
         ' If the file is dirty then reparse it before attempting the autocomplete
         If cbool(SciExec(pDoc->hWindow(0), SCI_GETMODIFY, 0, 0 )) = true orelse _
            (pDoc->UserModified = true) then
            pDoc->ParseDocument()
			end if   
      end if
   end if
   
   sList = "|"
         
   select case pDoc->AutoCompleteType
      
      case AUTOCOMPLETE_DIM_AS
         ' Create the space separated data type list
         ' Add all of the TYPE definitions that were found
         gdb2.dbRewind()
         do 
            pData = gdb2.dbGetNext()
            if pData = 0 then exit do

            if bTypesOnly = false then
               if pData->id = DB2_STANDARDDATATYPE then
                  st = pData->VariableType
                  if nLenMatchWord then
                     if left(ucase(st), nLenMatchWord) = pDoc->sMatchWord then
                        if instr(sList, st & "|") = 0 then
                           sList = sList & st & "|"
                        end if   
                     end if
                  else
                     if len(st) then      
                        if instr(sList, st & "|") = 0 then
                           sList = sList & st & "|"
                        end if   
                     end if   
                  end if
               end if   
            end if
    
            if pData->id = DB2_TYPE then
               if bTypesOnly then
                  if bTypesNotPreloaded then
                     ' Only look at TYPES that are loaded into the editor. These would
                     ' be entries that have a Filename attached to them. Preloaded TYPES
                     ' from WinAPI or Afx do not have Filenames.
                     if len(pData->fileName) = 0 then continue do
                  end if
               end if
 
               ' nLenMatchWord will allow partial matches to be added to the popup
               ' based on what the user has currently typed.
               st = rtrim(pData->VariableType)

               if nLenMatchWord then
                  if left(ucase(st), nLenMatchWord) = pDoc->sMatchWord then
                     if ExistsInAutocompleteList(sList, st) = false then
                        sList = sList & st & "|" 
                     end if
                  end if
               elseif len(st) then
                  if ExistsInAutocompleteList(sList, st) = false then
                     sList = sList & st & "|"
                  end if    
               end if
            end if
         loop


      case AUTOCOMPLETE_TYPE
         ' Attempt to popup a list of TYPE elements related the pData variableType (TYPE)
         pData = DereferenceLine( pDoc, sDot, curPos - 1)
         if pData = 0 then return false
         sLookFor = pData->VariableType    ' This is the TYPE structure
         if len(sLookFor) = 0 then return false

         sLookFor = ucase(sLookFor)

         do
            gdb2.dbRewind()
            do 
               pData = gdb2.dbGetnext()
               if pData = 0 then exit do
               if pData->deleted = true then continue do
               
               dim as boolean bAddToList = false
               
               if (pData->id = DB2_VARIABLE) andalso _
                  (ucase(pData->ParentName) = sLookFor) then
                  st = pData->ElementName
                  bAddToList = true
               elseif (pData->id = DB2_FUNCTION) andalso _
                  (ucase(pData->ParentName) = sLookFor) then   
                  st = AfxStrParse(pData->ElementName, 2, ".")
                  bAddToList = true
               end if
                        
               if bAddToList then
                  if nLenMatchWord then
                     if left(ucase(st), nLenMatchWord) = pDoc->sMatchWord then
                        if ExistsInAutocompleteList(sList, st) = false then
                           sList = sList & st & "|"
                        end if
                     end if
                  else
                     If ExistsInAutocompleteList(sList, st) = false then
                        sList = sList & st & "|"
                     end if
                  end if         
               end if   
            loop 
            
            dim pDataTYPE as DB2_DATA ptr
            pDataTYPE = gdb2.dbFindTYPE( sLookFor )
            if pDataTYPE then
               sLookFor = ucase(pDataType->TypeExtends)
            else
               exit do   
            end if
            if len(sLookFor) = 0 then exit do
         loop
      
   end select
   
   sList = trim(sList, any "| ")
   if len(sList) then
      ' Save the code editor starting position 
      pDoc->AutoCStartPos = SciExec(hEdit, SCI_GETCURRENTPOS, 0, 0)
      ShowAutoCompletePopup(hEdit, sList)
      return true
   else
      if (notification <> SCN_AUTOCCHARDELETED) then
         SciExec( hEdit, SCI_AUTOCCANCEL, 0, 0)
      end if   
   end if 

   function = true

end function



