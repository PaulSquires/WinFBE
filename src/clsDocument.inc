
''
''  clsDocument (Class to handle everything related to a Scintilla editing window)
''


''
''
Constructor clsDocument
'
End Constructor
   
''
''
Destructor clsDocument
   DestroyWindow m_hWindow       ' destroy the scintilla window
   this.DeallocateFunctionList   ' free allocated memory for the function list
End Destructor


''
''
Property clsDocument.hWindow( ByVal _HWnd as HWnd)
   m_hWindow = _HWnd
End Property

Property clsDocument.hWindow() as HWnd
   Property = m_hWindow
End Property

''
''
Property clsDocument.IsNewFlag( ByVal nValue As BOOLEAN)
   m_IsNewFlag = nValue
End Property

Property clsDocument.IsNewFlag() As BOOLEAN
   Property = m_IsNewFlag
End Property

''
''
Property clsDocument.IsProjectFile( ByVal nValue As BOOLEAN)
   m_IsProjectFile = nValue
End Property

Property clsDocument.IsProjectFile() As BOOLEAN
   Property = m_IsProjectFile
End Property

''
''
Property clsDocument.ProjectFileType( ByVal nValue As Long)
   m_ProjectFileType = nValue
End Property

Property clsDocument.ProjectFileType() As Long
   Property = m_ProjectFileType
End Property

''
''
Property clsDocument.DateFileTime( ByVal nValue As Double)
   m_DateFileTime = nValue
End Property

Property clsDocument.DateFileTime() As Double
   Property = m_DateFileTime
End Property

''
''
Property clsDocument.DiskFilename( ByVal wValue As WString Ptr)
   m_sFilename = *wValue
End Property

Property clsDocument.DiskFilename() As WString Ptr
   Property = @m_sFilename
End Property

''
''
Function clsDocument.CreateCodeWindow( ByVal hWndParent As HWnd, _
                                       ByVal IsNewFile  As BOOLEAN, _     
                                       ByVal IsTemplate As BOOLEAN = False, _
                                       ByVal pwszFile   As WString Ptr = 0 ) As HWnd   

   ' Creates a Scintilla editing window (initially not visible). Optionally, load a diskfile
   ' into the window and apply properties to it.
   m_hWindow = CreateWindowEx( 0, "Scintilla", "", _
                   WS_CHILD Or WS_TABSTOP Or WS_CLIPCHILDREN, _
                   0,0,0,0,HWndParent, _
                   Cast(HMENU, IDC_SCINTILLA), GetModuleHandle(Null), Null)

   ' Initialize our direct access to the Scintilla code windows. This is much faster than
   ' using SendMessage to the window. Only need to initialize once no matter how many
   ' code windows that are eventually opened.
   If IsWindow(m_hWindow) Then
      If SciMsg = 0 Then
         SciMsg = Cast( Scintilla_DirectFunction, SendMessage( m_hWindow, SCI_GETDIRECTFUNCTION, 0, 0 ) )
      End If

      ' Call the direct function for speed purposes rather than relying on the traditional SendMessage method.
      m_pSci = Cast(Any Ptr, SendMessage( m_hWindow, SCI_GETDIRECTPOINTER, 0, 0 )) 
   End If
   
   Static nNextFileNum As Long 
      
   ' If a disk file was specified then open it and load it into the editor
   this.IsNewFlag = IsNewFile
   If IsNewFile Then
      nNextFileNum = nNextFileNum + 1
      this.DiskFilename = "Untitled" & nNextFileNum
   End If

   If pwszFile Then   
      If Dir(*pwszFile) > "" Then   
         Dim As String st, sText
         Dim As Long f = Freefile
         Dim As Long idx
         
         If IsTemplate Then
            If Open(*pwszFile For Input As #f) = 0 Then
               ' Look at the first 4 lines
               ' Line 3 tells us the file type (bas or xml)
               Do Until Eof(f)
                  Line Input #f, st
                  idx = idx + 1
                  Select Case idx
                     Case 1
                     Case 2
                     Case 3: this.DiskFilename = *this.DiskFilename & Trim(st)    
                     Case 4 
                     Case Else
                        sText = sText & st & vbCrLf
                  End Select   
               Loop
               Close #f
               this.SetText( sText ) 
            End If
            ' Search for "|", replace it with an empty space "" and position the caret in that place
            this.FindReplace( "|", "" )
            ' don't set SAVEPOINT for newly loaded Template files because we want the document to display as dirty
         Else
            If Open(*pwszFile For Binary As #f) = 0 Then
               sText = Space(Lof(f))
               Get #f, , sText
               Close #f
               this.SetText( sText ) 
               this.DateFileTime = FileDateTime( *pwszFile )
            End If
            this.DiskFilename = *pwszFile
            SciMsg( m_pSci, SCI_SETSAVEPOINT, 0, 0)
            SciMsg( m_pSci, SCI_EMPTYUNDOBUFFER, 0, 0)
            ' Update the most recently used file list (only for non-IsNewFlag files)
            UpdateMRUList(pwszFile)
         End If
      End If   
   End If

   this.ApplyProperties()
   
   Function = m_hWindow
End Function

''
''
Function clsDocument.FindReplace( ByVal strFindText As String, ByVal strReplaceText As String ) As Long
   ' Return Value: New position if successful; -1 if text not found.
   Dim As Long findFlags = SCFIND_MATCHCASE Or SCFIND_WHOLEWORD
   Dim As Long startPos  = SciMsg( m_pSci, SCI_GETCURRENTPOS, 0, 0)
   Dim As Long endPos    = SciMsg( m_pSci, SCI_GETTEXTLENGTH, 0, 0)
   Dim As Long newPos

   ' Set the start and end positions, and search flags, and finally do the search
   SciMsg( m_pSci, SCI_SETTARGETSTART, startPos, 0)
   SciMsg( m_pSci, SCI_SETTARGETEND, endPos, 0)
   SciMsg( m_pSci, SCI_SETSEARCHFLAGS, findFlags, 0)

   ' Search the text to replace
   newPos = SciMsg( m_pSci, SCI_SEARCHINTARGET, Len(strFindText), Cast(LPARAM, Strptr(strFindText)) )
   
   ' Return -1 if not found
   If newPos = - 1 Then Return -1
   
   ' Position the caret and select the text
   SciMsg( m_pSci, SCI_SETCURRENTPOS, newPos, 0)
   SciMsg( m_pSci, SCI_GOTOPOS, newPos, 0)
   SciMsg( m_pSci, SCI_SETSELECTIONSTART, newPos, 0)
   SciMsg( m_pSci, SCI_SETSELECTIONEND, newPos + Len(strFindText), 0)

   ' Replace the selection (SCI_REPLACESEL fails if text is "" so use Cut instead for that scenario)                                
   If Len(strReplaceText) = 0 Then
      SciMsg( m_pSci, SCI_CUT, 0, 0 )
   Else
      SciMsg( m_pSci, SCI_REPLACESEL, 0, Cast(LPARAM, Strptr(strReplaceText)) )
   End If
   
   ' Return the new position
   Function = newPos

End Function

''
''
Function clsDocument.InsertFile() As BOOLEAN
   ' Display the Open File Dialog
   Dim pwszName As WString Ptr = AfxIFileOpenDialogW(HWND_FRMMAIN, IDM_INSERTFILE)
   If pwszName Then
      Dim As Long f = Freefile
      Dim As String sText
      If Open(*pwszName For Binary As #f) = 0 Then
         sText = Space(Lof(f))
         Get #f, , sText
         Close #f
      End If
      SciMsg( m_pSci, SCI_INSERTTEXT, -1, Cast(LPARAM, Strptr(sText)) )
      CoTaskMemFree pwszName
   End If
   
   Function = 0
End Function

''
''
Function clsDocument.SaveFile( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN

   Dim sText     As String
   Dim wFilename As WString * MAX_PATH
   Dim wText     As WString * MAX_PATH  
   Dim f         As Long 
   
   ' If this is a new Untitled document and it is dirty then
   ' set flag to display SaveAs dialog.
   If this.IsNewFlag Then 
      If SendMessageW( m_hWindow, SCI_GETMODIFY, 0, 0 ) Then bSaveAs = True
   End If
      
   wFilename = *this.Diskfilename
   
   If bSaveAs Then
      ' Display the Save File Dialog
      FF_PATHNAME( @WStr("NAMEX"), @wFilename, @wText)
      Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wText, "")
      If pwszName Then
         wFilename = *pwszName
         CoTaskMemFree(pwszName)
      Else
         Return False
      End If
   End If


   ' Save text buffer to disk by directly accessing buffer rather
   ' saving it to an intermediary string variable first.
   f = Freefile
   If Open( wFilename For Output As #f ) = 0 Then
      Dim As ZString Ptr psz = Cast( ZString Ptr, SendMessageW(this.hWindow, SCI_GETCHARACTERPOINTER, 0, 0) )
      Print #f, *psz
      Close #f
   Else
      ' error saving file
   End If         

'   ' Delete any existing file
'   If Dir(wFilename) > "" Then Kill wFilename
'   
'   ' Do the actual saving to disk
'   f = Freefile
'   If Open( wFilename For Binary As #f) = 0 Then
'      sText = this.GetText
'      Put #f, 1, sText
'      Close #f
'   Else
'      ' error saving file
'   End If   

   this.DiskFilename = @wFilename
   this.IsNewFlag    = False
   this.DateFileTime = FileDateTime( wFilename )
   
   ' Set the current state of the document to unmodified
   SciMsg( m_pSci, SCI_SETSAVEPOINT, 0, 0)
   
   Function = True
End Function


''
''
Function clsDocument.GetTextRange( ByVal cpMin As Long, ByVal cpMax As Long) As String
   Dim p      As Long
   Dim buffer As String
   Dim txtrg  As SCI_TEXTRANGE
   txtrg.chrg.cpMin = cpMin
   txtrg.chrg.cpMax = cpMax
   buffer = Space(cpMax - cpMin + 1)
   txtrg.lpstrText = Strptr(buffer)
   SendMessage(m_hWindow, SCI_GETTEXTRANGE, 0, Cast(LPARAM, @txtrg))
   p = Instr(buffer, Chr(0))
   If p Then buffer = Left(buffer, p - 1)
   Function = buffer
End Function

''
''
Function clsDocument.ChangeSelectionCase( ByVal fCase As Long) As Long 

   Dim startSelPos As Long     ' Starting position
   Dim endSelPos   As Long     ' Ending position
   Dim strText     As String   ' Selected text
   Dim i           As Long 

   ' fCase = 1 (upper case), 2 (lower case), 3 (mixed case)
   If (fCase < 1) Or (fCase > 3) Then Exit Function
   
   ' If startSelPos and endSelPos are the same there is not selection,
   startSelPos = SciMsg( m_pSci, SCI_GETSELECTIONSTART, 0, 0)
   endSelPos   = SciMsg( m_pSci, SCI_GETSELECTIONEND, 0, 0)
   If startSelPos = endSelPos Then Exit Function
   
   ' Retrieve the text
   strText = this.GetTextRange(startSelPos, endSelPos)
   
   ' Convert it to upper or lower case
   If fCase = 1 Then
      strText = Ucase(strText)
   ElseIf fCase = 2 Then
      strText = LCase(strText)
   ElseIf fCase = 3 Then
      ' Always uppercase the first character regardless
      Mid(strText,1,1) = Ucase(Left(strText,1))
      Do
          i = Instr(i+1, strText, " ")
          If i = 0 Then Exit Do
          Mid(strText,i+1,1) = Ucase(Mid(strText,i+1,1))
      Loop
   End If
   
   ' Replace the selected text
   SciMsg( m_pSci, SCI_REPLACESEL, 0, Cast(LPARAM, Strptr(strText)))

   Function = 0
End Function

''
''
Function clsDocument.GetCurrentLineNumber() As Long
   Dim nPos As Long = SciMsg( m_pSci, SCI_GETCURRENTPOS, 0, 0)
   Function = SciMsg( m_pSci, SCI_LINEFROMPOSITION, nPos, 0)
End Function

''
''
Function clsDocument.SelectLine( ByVal nLineNum As Long ) As Long
   ' Select the incoming nLineNum. If nLineNum is negative then select the current line
   nLineNum = Iif( nLineNum < 0, this.GetCurrentLineNumber, nLineNum)
   Dim nStartPos As Long = SciMsg( m_pSci, SCI_POSITIONFROMLINE, nLineNum, 0)
   Dim nEndPos   As Long = SciMsg( m_pSci, SCI_GETLINEENDPOSITION, nLineNum, 0)
   SciMsg( m_pSci, SCI_SETSELECTIONSTART, nStartPos, 0)
   SciMsg( m_pSci, SCI_SETSELECTIONEND, nEndPos, 0)
   Function = 0
End Function

''
''
Function clsDocument.GetLine( ByVal nLine As Long) As String
   Dim nLen   As Long
   Dim buffer As String
   nLen = SciMsg( m_pSci, SCI_LINELENGTH, nLine , 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen)
   SendMessage m_hWindow, SCI_GETLINE, nLine, Cast(LPARAM, Strptr(buffer))
   Function = Rtrim(buffer, Any Chr(13,10,0))
End Function

''
''
Function clsDocument.GetSelText() As String
   Dim nLen As Long
   Dim buffer As String
   nLen = SciMsg( m_pSci, SCI_GETSELTEXT, 0, 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen)
   SciMsg( m_pSci, SCI_GETSELTEXT, 0, Cast(LPARAM, Strptr(buffer)))
   Function = Trim(buffer, Chr(0))
End Function

''
''
Function clsDocument.GetText() As String
   Dim nLen   As Long
   Dim buffer As String
   nLen = SciMsg( m_pSci, SCI_GETLENGTH, 0 , 0)
   If nLen < 1 Then Exit Function
   buffer = Space(nLen+1)
   SciMsg( m_pSci, SCI_GETTEXT, nLen+1, Cast(LPARAM, Strptr(buffer)) )
   Function = Trim(buffer, Chr(0))
End Function

''
''
Function clsDocument.SetText( ByRef sText As Const String ) As Long 
   SciMsg( m_pSci, SCI_SETTEXT, 0, Cast(LPARAM, Strptr(sText)) )
   SciMsg( m_pSci, SCI_COLOURISE, 0, -1 )
   Function = 0
End Function

''
''
Function clsDocument.GetSelectedLineRange( ByRef startLine As Long, _
                                           ByRef endLine   As Long, _
                                           ByRef startPos  As Long, _
                                           ByRef endPos    As Long _
                                           ) As Long 
   
   startPos  = SciMsg( m_pSci, SCI_GETSELECTIONSTART, 0, 0) 
   endPos    = SciMsg( m_pSci, SCI_GETSELECTIONEND, 0, 0) 
   startLine = SciMsg( m_pSci, SCI_LINEFROMPOSITION, startPos, 0) 
   endLine   = SciMsg( m_pSci, SCI_LINEFROMPOSITION, endPos, 0) 

   Dim nCol As Long = SciMsg( m_pSci, SCI_GETCOLUMN, endPos, 0)
   If (nCol = 0) And (endLine > startLine) Then endLine = endLine - 1

   Function = 0
End Function

''
''
Function clsDocument.BlockComment( ByVal flagBlock As BOOLEAN ) As Long

   Dim i           As Long        ' Loop counter
   Dim startPos    As Long        ' Starting position
   Dim endPos      As Long        ' Ending position
   Dim startLine   As Long        ' Starting line
   Dim endLine     As Long        ' Ending line
   Dim nPos        As Long        ' Position
   Dim strText     As String      ' Portion of the line to replace
   Dim nCount      As Long        ' Number of "'" added or removed
   
   ' flagBlock = True for Blocking comment, False for UnBlocking comment
   this.GetSelectedLineRange( startLine, endLine, startPos, endPos )
   SciMsg( m_pSci, SCI_GOTOLINE, startLine, 0) 

   For i = startLine To endLine
      strText = this.GetLine(i)   
      If flagBlock = False Then   ' unblock comment
         If Left(strText, 1) <> "'" Then
            Continue For
         Else   
            strText = Iif(Len(strText) > 1, Mid(strText, 2), " ")
         End If   
      Else
         strText = "'" & strText
      End If
      If Len(strText) Then
         nPos = SciMsg( m_pSci, SCI_POSITIONFROMLINE, i, 0)   ' starting position of the line
         SciMsg( m_pSci, SCI_SETTARGETSTART, nPos, 0)
         SciMsg( m_pSci, SCI_SETTARGETEND, nPos + Len(strText) + Iif(flagBlock, -1, 1), 0)
         SciMsg( m_pSci, SCI_REPLACETARGET, Len(strText), Cast(LPARAM, Strptr(strText))) 
         nCount += 1
      End If
   Next

   If startPos <> endPos Then
      SciMsg( m_pSci, SCI_SETSELECTIONSTART, startPos, 0)
      SciMsg( m_pSci, SCI_SETSELECTIONEND, endPos + Iif(flagBlock, nCount, -nCount), 0)
   Else
      SciMsg( m_pSci, SCI_SETSELECTIONSTART, endPos + Iif(flagBlock, nCount, -nCount), 0)
      SciMsg( m_pSci, SCI_SETSELECTIONEND, endPos + Iif(flagBlock, nCount, -nCount), 0)
   End If
   
   Function = 0
End Function

''
''
Function clsDocument.CurrentLineUp() As Long
   Dim currentLine As Long = this.GetCurrentLineNumber()
   If (currentLine <> 0) Then
      SciMsg( m_pSci, SCI_BEGINUNDOACTION, 0, 0)
      currentLine = currentLine -1
      SciMsg( m_pSci, SCI_LINETRANSPOSE, 0, 0)
      SciMsg( m_pSci, SCI_GOTOLINE, currentLine, 0)
      SciMsg( m_pSci, SCI_ENDUNDOACTION, 0, 0)
   End If
   Function = 0
End Function

''
''
Function clsDocument.CurrentLineDown() As Long
   Dim currentLine As Long = this.GetCurrentLineNumber()
   If currentLine <> SciMsg( m_pSci, SCI_GETLINECOUNT, 0, 0) - 1 Then
      SciMsg( m_pSci, SCI_BEGINUNDOACTION, 0, 0)
      currentLine = currentLine + 1
      SciMsg( m_pSci, SCI_GOTOLINE, currentLine, 0)
      SciMsg( m_pSci, SCI_LINETRANSPOSE, 0, 0)
      SciMsg( m_pSci, SCI_ENDUNDOACTION, 0, 0)
   End If
   Function = 0
End Function
  
''
''
Function clsDocument.MoveCurrentLines( ByVal flagMoveDown As BOOLEAN ) As Long

   Dim startPos    As Long     ' Starting position
   Dim endPos      As Long     ' Ending position
   Dim startLine   As Long     ' Starting line
   Dim endLine     As Long     ' Ending line
   Dim line2swap   As Long 
   Dim nbChar      As Long 
   Dim nTextLen    As Long
   
   this.GetSelectedLineRange( startLine, endLine, startPos, endPos )

   Dim noSel      As Long = (startLine = endLine)
   Dim nbSelLines As Long = (endLine - startLine + 1)

   If flagMoveDown Then
      nTextLen  = SciMsg( m_pSci, SCI_GETTEXTLENGTH, 0, 0)
      If (startLine = -1) Or (endLine >= SciMsg( m_pSci, SCI_LINEFROMPOSITION, nTextLen, 0)) Then Exit Function
      line2swap = endLine + 1
      If (line2swap + 1) = SciMsg( m_pSci, SCI_GETLINECOUNT, 0, 0) Then nbChar += 2    ' cr/lf
   Else
      If (startLine <= 0) Then Exit Function
      line2swap = startLine - 1
   End If
   
   nbChar = nbChar + SciMsg( m_pSci, SCI_LINELENGTH, line2swap, 0)
   
   SciMsg( m_pSci, SCI_BEGINUNDOACTION, 0, 0)
   SciMsg( m_pSci, SCI_GOTOLINE, line2swap, 0)

   For i As Long = 0 To nbSelLines - 1
      If flagMoveDown Then
         this.CurrentLineUp()
      Else
         this.CurrentLineDown()
      End If   
   Next
   SciMsg( m_pSci, SCI_ENDUNDOACTION, 0, 0)

   If flagMoveDown Then
      SciMsg( m_pSci, SCI_SETSELECTIONSTART, startPos + nbChar, 0)
      SciMsg( m_pSci, SCI_SETSELECTIONEND, Iif(noSel, startPos + nbChar, endPos + nbChar), 0)
   Else
      SciMsg( m_pSci, SCI_SETSELECTIONSTART, startPos - nbChar, 0)
      SciMsg( m_pSci, SCI_SETSELECTIONEND, Iif(noSel, startPos - nbChar, endPos - nbChar), 0)
   End If
   
   Function = 0
End Function

''
''
Function clsDocument.ToggleBookmark( ByVal nLine As Long ) As Long
   Dim fMark As Long  ' 32 bit value
   fMark = SciMsg( m_pSci, SCI_MARKERGET, nLine, 0) 
   If Bit(fMark, 0) = -1 Then
      SciMsg( m_pSci, SCI_MARKERDELETE, nLine, 0)
   Else
      SciMsg( m_pSci, SCI_MARKERADD, nLine, 0)
   End If
   Function = 0
End Function

''
''
Function clsDocument.NextBookmark() As Long 
   Dim fMark As Long  ' 32 bit value
   Dim nLine As Long = this.GetCurrentLineNumber() + 1
   fMark = BitSet(fMark, 0)
   nLine = SciMsg( m_pSci, SCI_MARKERNEXT, nLine, fMark)
   If nLine > -1 Then
      SciMsg( m_pSci, SCI_GOTOLINE, nLine, 0)
   Else
      nLine = SciMsg( m_pSci, SCI_MARKERNEXT, nLine, fMark)
      If nLine > -1 Then
         SciMsg( m_pSci, SCI_GOTOLINE, nLine, 0)
      End If
   End If
   Function = 0 
End Function

''
''
Function clsDocument.PrevBookmark() As Long 
   Dim fMark  As Long  ' 32 bit value
   Dim nLine  As Long = this.GetCurrentLineNumber() - 1
   Dim nLines As Long = SciMsg( m_pSci, SCI_GETLINECOUNT, 0, 0) - 1
   fMark = BitSet(fMark, 0)
   nLine = SciMsg( m_pSci, SCI_MARKERPREVIOUS, nLine, fMark)
   If nLine > -1 Then
      SciMsg( m_pSci, SCI_GOTOLINE, nLine, 0)
   Else
      nLine = SciMsg( m_pSci, SCI_MARKERPREVIOUS, nLines, fMark)
      If nLine > -1 Then
         SciMsg( m_pSci, SCI_GOTOLINE, nLine, 0)
      End If
   End If
   Function = 0 
End Function

''
''   
Function clsDocument.FoldToggle( ByVal nLine As Long ) As Long
   Dim nFoldLevel As Long = SciMsg( m_pSci, SCI_GETFOLDLEVEL, nLine, 0)

   If (nFoldLevel And SC_FOLDLEVELHEADERFLAG) = 0 Then
      ' Get the number of the head line of the procedure or function
      nLine = SciMsg( m_pSci, SCI_GETFOLDPARENT, nLine, 0) 
   End If
   If nLine > -1 Then
      SciMsg( m_pSci, SCI_TOGGLEFOLD, nLine, 0) 
      SciMsg( m_pSci, SCI_GOTOLINE, nLine, 0)
   End If

   Function = nLine
End Function

''
''
Function clsDocument.FoldAll() As Long

   Dim i          As Long    
   Dim nLines     As Long    
   Dim nFoldLevel As Long
   
   ' Force the lexer to style the whole document
   SendMessage m_hWindow, SCI_COLOURISE, -1, 0

   nLines = SciMsg( m_pSci, SCI_GETLINECOUNT, 0, 0)

   For i = 0 To nLines
      ' If we are in the head line ...
      nFoldLevel = SciMsg( m_pSci, SCI_GETFOLDLEVEL, i, 0)
      If (nFoldLevel And SC_FOLDLEVELNUMBERMASK) = SC_FOLDLEVELBASE Then
         If SciMsg( m_pSci, SCI_GETFOLDEXPANDED, i, 0) Then
            SciMsg( m_pSci, SCI_TOGGLEFOLD, i, 0) 
         End If
      End If
   Next

   Function = 0
End Function

''
''
Function clsDocument.UnFoldAll() As Long

   Dim i          As Long    
   Dim nLines     As Long    
   Dim nFoldLevel As Long
   
   ' Force the lexer to style the whole document
   SendMessage m_hWindow, SCI_COLOURISE, -1, 0

   nLines = SciMsg( m_pSci, SCI_GETLINECOUNT, 0, 0)

   For i = 0 To nLines
      ' If we are in the head line ...
      nFoldLevel = SciMsg( m_pSci, SCI_GETFOLDLEVEL, i, 0)
      If (nFoldLevel And SC_FOLDLEVELNUMBERMASK) = SC_FOLDLEVELBASE Then
         If SciMsg( m_pSci, SCI_GETFOLDEXPANDED, i, 0) = 0 Then
            SciMsg( m_pSci, SCI_TOGGLEFOLD, i, 0) 
         End If
      End If
   Next

   Function = 0
End Function

''
''
Function clsDocument.FoldToggleOnwards( ByVal nLine As Long) As Long

   ' Toggles the curent fold point and, if it is a base level, all the base level fold points below.
   Dim i          As Long    
   Dim nLines     As Long    
   Dim nFoldLevel As Long
   Dim FoldState  As Long
   
   ' Force the lexer to style the whole document
   SendMessage m_hWindow, SCI_COLOURISE, -1, 0

   nLines = SciMsg( m_pSci, SCI_GETLINECOUNT, 0, 0)

   ' Toggle the first sub or function
   nLine = this.FoldToggle(nLine)
   
   ' Determine whether the fold is expanded or not
   FoldState = SciMsg( m_pSci, SCI_GETFOLDEXPANDED, nLine, 0)

   For i = nLine To nLines
      ' If we are in the head line
      nFoldLevel = SciMsg( m_pSci, SCI_GETFOLDLEVEL, i, 0)
      If (nFoldLevel And SC_FOLDLEVELNUMBERMASK) = SC_FOLDLEVELBASE Then
         ' If the state is different ...
         If SciMsg( m_pSci, SCI_GETFOLDEXPANDED, i, 0) <> FoldState Then
            SciMsg( m_pSci, SCI_TOGGLEFOLD, i, 0) 
         End If
      End If
   Next

   Function = 0
End Function

''
''
Function clsDocument.ConvertEOL( ByVal nMode As Long) As Long
   Function = SciMsg( m_pSci, SCI_CONVERTEOLS, nMode, 0)
End Function

''
''
Function clsDocument.TabsToSpaces() As Long

   Dim As Long i, n, nLen, nLines, TabSize, nLineNumber 
   Dim As String strText, strBuffer

   ' Get the current line
   nLineNumber = this.GetCurrentLineNumber
   ' Get the tab size
   TabSize = SciMsg( m_pSci, SCI_GETTABWIDTH, 0, 0)
   If TabSize < 1 Then Exit Function
   ' Get the length of the text
   nLen  = SciMsg( m_pSci, SCI_GETTEXTLENGTH, 0, 0) 
   ' Get the number of lines
   nLines = SciMsg( m_pSci, SCI_GETLINECOUNT, 0, 0)
   ' Remove tabs, line by line
   For i = 0 To nLines - 1
      strText = this.GetLine(i)
      n = 1
      Do
         n = Instr(n, strText, Chr(9))
         If n > 0 Then 
            strText = Left(strText, n - 1) & Space(TabSize) & Mid(strText, n + 1)
            n += 1
         End If   
      Loop Until n = 0
      strBuffer = strBuffer & strText & Chr(13,10)
   Next
   ' Set the new text
   this.SetText(strBuffer)
   ' Set the caret position
   SciMsg( m_pSci, SCI_GOTOLINE, nLineNumber, 0) 

   Function = 0
End Function

''
''
Function clsDocument.DisplayStats() As Long

   Dim As HWnd hStatusbar = GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_STATUSBAR)

   Dim As Long i, curPos, nLine, nCol, nLines, nTextLen, nCount, idx
   Dim As Long startPos, endPos, startLine, endLine, nSelLines
   Dim pDoc As clsDocument Ptr
   Dim TB As TC_ITEM 
   
   Static wText As WString * 1024   ' General purpose variable
   Static wTag  As WString * 10     ' * display if modified
   
   ' Retrieve the information and show it in the status bar
   nLines    = SciMsg( m_pSci, SCI_GETLINECOUNT, 0, 0) 
   nTextLen  = SciMsg( m_pSci, SCI_GETTEXTLENGTH, 0, 0) 
   curPos    = SciMsg( m_pSci, SCI_GETCURRENTPOS, 0, 0)
   nLine     = SciMsg( m_pSci, SCI_LINEFROMPOSITION, curPos, 0) 
   nCol      = SciMsg( m_pSci, SCI_GETCOLUMN, curPos, 0) 
   this.GetSelectedLineRange(startLine, endLine, startPos, endPos )

   If endPos - startPos Then  ' there is selected text
      startLine = Iif( nLine < startLine, nLine, startLine)
      endLine   = Iif( nLine > endLine, nLine, endLine)
      nSelLines = Iif(endLine = startLine, 0, endLine - startLine + 1)  
   End If
   
   wText = " Ln: " & Str(nLine+1) & "   Col: " & Str(nCol+1) & _
           "   Sel: " & Str(endPos - startPos) & " | " & Str(nSelLines)
   StatusBar_SetText(hStatusbar, 2, @wText)
   
   wText = "  Lines: " & Str(nLines) & "   Chars: " & Str(nTextLen) 
   StatusBar_SetText(hStatusbar, 3, @wText)

   ' Set the text that displays on the tab. Determine what tab this document belongs to. 
   Dim As HWnd hTabCtl = GetDlgItem(HWND_FRMMAIN, IDC_FRMMAIN_TOPTABCONTROL)
   nCount = TabCtrl_GetItemCount(hTabCtl) 
   For i = 0 To nCount - 1
      TB.Mask = TCIF_PARAM
      TabCtrl_GetItem( hTabCtl, i, @TB )
      pDoc = Cast(clsDocument Ptr, TB.lParam)
      If pDoc->hWindow = m_hWindow Then
         idx = i: Exit For
      EndIf
   Next      

   FF_PATHNAME( @WStr("NAMEX"), *this.DiskFilename, @wText)
   wTag = Iif( SendMessageW( m_hWindow, SCI_GETMODIFY, 0, 0 ), "*", "")
   wText = wText & wTag
   TB.Mask    = TCIF_TEXT 
   TB.pszText = @wText  
   TabCtrl_SetItem( hTabCtl, idx, @TB )
   
   ' Update the main window caption with the full filename and path
   wText = APPNAMESHORT 
   If gpApp->IsProjectActive Then wText = wText & " - [" & *gpApp->ProjectName & "]"
   wText = wText & " - [" & *this.DiskFilename & wTag & "]"
   SetWindowTextW( HWND_FRMMAIN, @wText )

   Function = 0
End Function


''
''
Function clsDocument.ApplyProperties() As Long

   Dim nCount      As Long 
   Dim i           As Long
   Dim nPixels     As Long  
   Dim bitsNeeded  As Long 
   Dim wFileExt    As WString * MAX_PATH
   Dim strFontName As String 
   Dim nFontSize   As Long  
   Dim nFontCase   As Long 
   Dim rxRatio     As Single = 1
   Dim ryRatio     As Single = 1
   
   ' Determine the pWindow parent of the Scintilla window in order
   ' to ensure that DPI ratios are correctly used.
   Dim pWindow As CWindow Ptr = AfxCWindowOwnerPtr(m_hWindow)
   If pWindow Then
      rxRatio = pWindow->rxRatio
      ryRatio = pWindow->ryRatio
   End If   
   
   strFontName = Str(gConfig.EditorFontname)
   nFontSize   = Val(gConfig.EditorFontsize)
   
   Select Case gConfig.KeywordCase
      Case 0:  nFontCase = SC_CASE_UPPER
      Case 1:  nFontCase = SC_CASE_LOWER
      Case 2:  nFontCase = SC_CASE_MIXED
   End Select            
   
   If m_pSci = 0 Then Exit Function

   SciMsg( m_pSci, SCI_STYLESETFONT, STYLE_DEFAULT, Cast(LPARAM, Strptr(strFontName)) )
   SciMsg( m_pSci, SCI_STYLESETSIZE, STYLE_DEFAULT, nFontSize )
   SciMsg( m_pSci, SCI_STYLESETCHARACTERSET, STYLE_DEFAULT, GetFontCharSetID(gConfig.EditorFontCharset) )

   SciMsg( m_pSci, SCI_STYLESETFORE, STYLE_DEFAULT, gConfig.clrTextFG )
   SciMsg( m_pSci, SCI_STYLESETBACK, STYLE_DEFAULT, gConfig.clrWindowFG )
   SciMsg( m_pSci, SCI_STYLECLEARALL, 0, 0 )  ' Copies global style to all others

   ''
   ''  MARGIN 0: Line Numbering (defaults to width 0)
   nPixels = SciMsg( m_pSci, SCI_TEXTWIDTH, 0, Cast(LPARAM, @"_99999"))
   SciMsg( m_pSci, SCI_SETMARGINTYPEN, 0, SC_MARGIN_NUMBER )
   SciMsg( m_pSci, SCI_STYLESETFORE, STYLE_LINENUMBER, gConfig.clrLineNumbersFG )
   SciMsg( m_pSci, SCI_STYLESETBACK, STYLE_LINENUMBER, gConfig.clrLineNumbersBG )
   SciMsg( m_pSci, SCI_SETMARGINWIDTHN, 0, Iif(gConfig.LineNumbering, nPixels, 0) )
   
   ''
   ''  MARGIN 1: Non-Folding symbols (defaults to width 16) (Bookmark symbol, etc) (will be same color as line numbering)
   SciMsg( m_pSci, SCI_SETMARGINTYPEN, 1, SC_MARGIN_SYMBOL )
   SciMsg( m_pSci, SCI_SETMARGINWIDTHN, 1, 16 * rxRatio )
   SciMsg( m_pSci, SCI_SETMARGINSENSITIVEN, 1, 1 )
   SciMsg( m_pSci, SCI_SETMARGINWIDTHN, 1, Iif(gConfig.LeftMargin, 16 * rxRatio, 0) )
              
   ''
   ''  MARGIN 2: Folding symbols (defaults to width 0)
   SciMsg( m_pSci, SCI_SETMARGINTYPEN, 2, SC_MARGIN_SYMBOL )
   SciMsg( m_pSci, SCI_SETMARGINMASKN, 2, SC_MASK_FOLDERS )
   SciMsg( m_pSci, SCI_SETMARGINWIDTHN, 2, 16 * rxRatio )
   SciMsg( m_pSci, SCI_SETFOLDMARGINCOLOUR, CTRUE, gConfig.clrFoldMarginFG )
   SciMsg( m_pSci, SCI_SETFOLDMARGINHICOLOUR, CTRUE, gConfig.clrFoldMarginFG )
   SciMsg( m_pSci, SCI_SETMARGINSENSITIVEN, 2, 1 )
   SciMsg( m_pSci, SCI_SETMARGINWIDTHN, 2, Iif(gConfig.FoldMargin, 16 * rxRatio, 0) )
   
   ''
   ''  MARGIN 3: Small margin to offset left margins from actual text (4 pixels)
   SciMsg( m_pSci, SCI_SETMARGINTYPEN, 3, SC_MARGIN_TEXT )
   SciMsg( m_pSci, SCI_SETMARGINWIDTHN, 3, 4 * rxRatio )
   SciMsg( m_pSci, SCI_SETMARGINTYPEN, 3, SC_MARGIN_FORE )
   SciMsg( m_pSci, SCI_SETMARGINTYPEN, 3, SC_MARGIN_BACK )

   '' 
   ''  FONT QUALITY
   SciMsg( m_pSci, SCI_SETFONTQUALITY, SC_EFF_QUALITY_ANTIALIASED, 0 )
   'SC_EFF_QUALITY_DEFAULT (backward compatible), SC_EFF_QUALITY_NON_ANTIALIASED, 
   'SC_EFF_QUALITY_ANTIALIASED, SC_EFF_QUALITY_LCD_OPTIMIZED
   
   SciMsg( m_pSci, SCI_SETTECHNOLOGY, SC_TECHNOLOGY_DIRECTWRITE, 0 )
   SciMsg( m_pSci, SCI_SETBUFFEREDDRAW, 0, 0 ) ' turn off b/c we are using Vista+ Direct Draw
   'SC_TECHNOLOGY_DEFAULT (0). 
   'On Windows Vista or later, 
   'SC_TECHNOLOGY_DIRECTWRITE (1), 
   'SC_TECHNOLOGY_DIRECTWRITERETAIN (2), or 
   'SC_TECHNOLOGY_DIRECTWRITEDC (3) 
   
   ''
   ''  CONFINE CARET TO TEXT
   If gConfig.ConfineCaret Then
      SciMsg( m_pSci, SCI_SETVIRTUALSPACEOPTIONS, SCVS_RECTANGULARSELECTION, 0 )
   Else
      SciMsg( m_pSci, SCI_SETVIRTUALSPACEOPTIONS, SCVS_RECTANGULARSELECTION Or SCVS_USERACCESSIBLE, 0 )
   End If
   
   ''
   ''  TABS AS SPACES
   If gConfig.TabIndentSpaces Then
      SciMsg( m_pSci, SCI_SETUSETABS, False, 0 )
   Else
      SciMsg( m_pSci, SCI_SETUSETABS, CTRUE, 0 )
   End If

   ''
   ''  SELECTIONS FILL ENTIRE SCREEN SPACE
   SciMsg( m_pSci, SCI_SETSELEOLFILLED, CTRUE, 0 )
       
   ''
   ''  TAB WIDTH 
   SciMsg( m_pSci, SCI_SETTABWIDTH, Val(gConfig.TabSize), 0 )
   SciMsg( m_pSci, SCI_SETINDENT, Val(gConfig.TabSize), 0 )
   
   ''
   ''  INDENTATION GUIDES
   If gConfig.IndentGuides Then
      SciMsg( m_pSci, SCI_SETINDENTATIONGUIDES, CTRUE, 0)
   Else
      SciMsg( m_pSci, SCI_SETINDENTATIONGUIDES, False, 0)
   End If
   SciMsg( m_pSci, SCI_SETSELFORE, STYLE_INDENTGUIDE, gConfig.clrIndentGuidesFG )
   SciMsg( m_pSci, SCI_SETSELBACK, STYLE_INDENTGUIDE, gConfig.clrIndentGuidesBG )
   
   ''
   ''  CARET
   SciMsg( m_pSci, SCI_SETCARETFORE, gConfig.clrCaretFG, 0 )
   SciMsg( m_pSci, SCI_SETCARETWIDTH, 2, 0 )       ' 2 pixels

   ''
   ''  SHOW CARET LINE
   If gConfig.HighlightCurrentLine Then
      SciMsg( m_pSci, SCI_SETCARETLINEVISIBLE, CTRUE, 0 )
   Else
      SciMsg( m_pSci, SCI_SETCARETLINEVISIBLE, False, 0 )
   End If

   ''
   ''  CARET LINE COLOR
   SciMsg( m_pSci, SCI_SETCARETLINEBACK, gConfig.clrHighlightedFG, 0 )
   
   ''
   ''  SELECTION COLORS
   SciMsg( m_pSci, SCI_SETSELFORE, CTRUE, gConfig.clrSelectionFG )
   SciMsg( m_pSci, SCI_SETSELBACK, CTRUE, gConfig.clrSelectionBG )
   
   ''
   ''  MULTIPLE SELECTIONS
   SciMsg( m_pSci, SCI_SETMULTIPLESELECTION, CTRUE, 0 ) 
            
   ''
   ''  ALWAYS KEEP THE CARET LINE VISIBLE
   SciMsg( m_pSci, SCI_SETCARETLINEVISIBLEALWAYS, CTRUE, 0 )
   
   ''
   ''  DISABLE RIGHT CLICK POPUP MENU
   SciMsg( m_pSci, SCI_USEPOPUP, False, 0 )

   ''
   ''  IDENTIFY CHARACTERS TO BE USED IN WORDS
   SciMsg( m_pSci, SCI_SETWORDCHARS, 0, Cast(LPARAM, @"~._:\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") )


   ''
   ''  APPLY ALL LANGUAGE SPECIFIC SYNTAX COLORING
   FF_Pathname( @WStr("EXTN"), this.DiskFilename, @wFileExt )
   wFileExt = Ucase(wFileExt)

   If cbool(wFileExt = ".BAS") OrElse cbool(wFileExt = ".INC") _
         OrElse cbool(wFileExt = ".BI") OrElse (this.IsNewFlag = True) Then
     
      SciMsg( m_pSci, SCI_SETLEXER, SCLEX_VB, 0 )
      bitsNeeded = SciMsg( m_pSci, SCI_GETSTYLEBITSNEEDED, 0, 0)
      SciMsg( m_pSci, SCI_SETSTYLEBITS, bitsNeeded, 0 )

      ' Set FreeBASIC Keywords
      If Len(gpApp->FBKeywords) Then
         SciMsg( m_pSci, SCI_SETKEYWORDS, 0, Cast(LPARAM, Strptr(gpApp->FBKeywords)) )
      End If
      
      If gConfig.SyntaxHighlighting Then
    
         ' Set the Comments style
         SciMsg( m_pSci, SCI_STYLESETFORE, SCE_B_COMMENT, gConfig.clrCommentsFG )
         SciMsg( m_pSci, SCI_STYLESETBACK, SCE_B_COMMENT, gConfig.clrCommentsBG )

         ' Set the Keywords style
         SciMsg( m_pSci, SCI_STYLESETFORE, SCE_B_KEYWORD, gConfig.clrKeywordFG )
         SciMsg( m_pSci, SCI_STYLESETBACK, SCE_B_KEYWORD, gConfig.clrKeywordBG )
         SciMsg( m_pSci, SCI_STYLESETCASE, SCE_B_KEYWORD, nFontCase )

         ' Set the Constants style
'            SciMsg( m_pSci, SCI_STYLESETFORE, SCE_B_CONSTANT, GetOptionVal("CLRKEYWORD_FG") )
'            SciMsg( m_pSci, SCI_STYLESETBACK, SCE_B_CONSTANT, GetOptionVal("CLRKEYWORD_BG") )

         ' Set the Identifiers style
         SciMsg( m_pSci, SCI_STYLESETFORE, SCE_B_IDENTIFIER, gConfig.clrTextFG )
         SciMsg( m_pSci, SCI_STYLESETBACK, SCE_B_IDENTIFIER, gConfig.clrTextBG )

         ' Set the Numbers style
         SciMsg( m_pSci, SCI_STYLESETFORE, SCE_B_NUMBER, gConfig.clrTextFG )
         SciMsg( m_pSci, SCI_STYLESETBACK, SCE_B_NUMBER, gConfig.clrTextBG )

         ' Set the Operators style
         SciMsg( m_pSci, SCI_STYLESETFORE, SCE_B_OPERATOR, gConfig.clrOperatorsFG )
         SciMsg( m_pSci, SCI_STYLESETBACK, SCE_B_OPERATOR, gConfig.clrOperatorsBG )

         ' Set the Preprocessor style
         SciMsg( m_pSci, SCI_STYLESETFORE, SCE_B_PREPROCESSOR, gConfig.clrPreprocessorFG )
         SciMsg( m_pSci, SCI_STYLESETBACK, SCE_B_PREPROCESSOR, gConfig.clrPreprocessorBG )

         ' Set the Strings style
         SciMsg( m_pSci, SCI_STYLESETFORE, SCE_B_STRING, gConfig.clrStringsFG )
         SciMsg( m_pSci, SCI_STYLESETBACK, SCE_B_STRING, gConfig.clrStringsBG )

      End If

   End If


   ''
   ''  CODE FOLDING
   If gConfig.LeftMargin Then
      ' Enable folding of the procedures and functions
      SciMsg( m_pSci, SCI_SETPROPERTY, Cast(WPARAM, @"fold"), Cast(LPARAM, @"1") )

      ' Initialize fold symbols for folding - Box tree
      SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN,    SC_MARK_BOXMINUS )
      SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDER,        SC_MARK_BOXPLUS )
      SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB,     SC_MARK_VLINE)
      SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL,    SC_MARK_LCORNER)
      SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND,     SC_MARK_BOXPLUSCONNECTED)
      SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)   ' SC_MARK_BOXMINUSCONNECTED
      SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER)

'         ' MacIntosh Arrows
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_ARROWDOWN)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_ARROW)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_EMPTY)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_EMPTY)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_EMPTY)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_EMPTY)

'         ' Plus/Minus
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_MINUS)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_PLUS)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_EMPTY)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_EMPTY)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_EMPTY)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_EMPTY)
'         SciMsg( m_pSci, SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_EMPTY)

      ' Draw line below if not expanded
      SciMsg( m_pSci, SCI_SETFOLDFLAGS, 16, 0 )

      ' Colors for folders closed and folders opened
      SciMsg( m_pSci, SCI_MARKERSETFORE, SC_MARKNUM_FOLDER,     gConfig.clrFoldSymbolFG )
      SciMsg( m_pSci, SCI_MARKERSETBACK, SC_MARKNUM_FOLDER,     gConfig.clrFoldSymbolBG )
      SciMsg( m_pSci, SCI_MARKERSETFORE, SC_MARKNUM_FOLDEROPEN, gConfig.clrFoldSymbolFG )
      SciMsg( m_pSci, SCI_MARKERSETBACK, SC_MARKNUM_FOLDEROPEN, gConfig.clrFoldSymbolBG )
   
   Else
      ' Disable folding of the procedures and functions
      SciMsg( m_pSci, SCI_SETPROPERTY, Cast(WPARAM, @"fold"), Cast(LPARAM, @"0") )
   End If
   
   Function = 0
End Function



''
''
Property clsDocument.FnListPtr( ByVal nValue As FUNCTION_TYPE Ptr)
   m_FnListPtr = nValue
End Property

Property clsDocument.FnListPtr() As FUNCTION_TYPE Ptr
   Property = m_FnListPtr
End Property


''
''
Function clsDocument.DeallocateFunctionList() As Long

   Dim pList As FUNCTION_TYPE Ptr
   Dim pListNext As FUNCTION_TYPE Ptr

   ' If a list already exists then we need to deallocate any memory
   If this.m_FnListPtr Then
      pList = this.m_FnListPtr
      Do Until pList = 0
         pListNext = pList->pListNext
         Delete pList
         pList = pListNext
      Loop
   End If
   this.m_FnListPtr = 0

   Function = 0
End Function
               
               
''
''
Function clsDocument.CreateFunctionList() As Long
   ' Generate a linked list of FUNCTION_TYPE structures that define all found
   ' Functions/Subs in the source code. This list is used to be displayed in
   ' the Function List.
 
   Dim As ZString Ptr pzByte = Cast( ZString Ptr, SendMessageW(m_hWindow, SCI_GETCHARACTERPOINTER, 0, 0) )
   Dim As Long nLen = SendMessageW(m_hWindow, SCI_GETTEXTLENGTH, 0, 0) 
   Dim As Long i = 0   ' current buffer position
   Dim As Long nCurPos, nEOL, nLineNum
   Dim As String st, st_ucase
   Dim As Long n, iParen
   Dim As BOOLEAN bFirstLine
   
   Dim pList As FUNCTION_TYPE Ptr
   Dim pListTemp As FUNCTION_TYPE Ptr

   ' If a list already exists then we need to deallocate any memory
   this.DeallocateFunctionList()
   
   ' Create the first function list node that holds the file name
   pList = New FUNCTION_TYPE
   pList->bIsHeader = True
   pList->nLineNumber = 0   'b/c Scintilla is zero based
   pList->pDoc = @this
   FF_PATHNAME( WStr("NAMEX"), this.DiskFilename, @pList->zFnName )
   pList->zFnName = Ucase(pList->zFnName)
   m_FnListPtr = pList

   Do While i < nLen - 1
      
      ' Get the next line. Up to the first LF or NUL (eof)
      nCurPos = i:  nEOL = -1
      For n = nCurPos To nLen - 1
         If pzByte[n] = 13 Then            ' end of line
            nEOL = n: i = n + 1            ' skip over the CR
            If pzByte[i] = 10 Then i += 1  ' skip over the LF
            Exit For
         End If   
      Next
      If nEOL = -1 Then 
         i = nLen: nEOL = nLen
      End If   
      st = Space(nEOL - nCurPos): memcpy( Strptr(st), @pzByte[nCurPos], nEOL-nCurPos )
      
      nLineNum += 1
      
      If Len(st) = 0 Then Continue Do
   
      ' Remove any comments
      n = Instr(st, "'")
      If n Then st = Left(st, n-1)
      
      st = Trim(st)
      If Len(st) = 0 Then Continue Do
   
      st_ucase = Ucase(st)
      
      If Left(st_ucase, 9)  = "FUNCTION " Then
         st = Mid(st, 10): bFirstLine = True
      ElseIf Left(st_ucase, 16) = "PUBLIC FUNCTION " Then
         st = Mid(st, 17): bFirstLine = True
      ElseIf Left(st_ucase, 17) = "PRIVATE FUNCTION " Then
         st = Mid(st, 18): bFirstLine = True
      End If   
      
      If Left(st_ucase, 4)  = "SUB " Then
         st = Mid(st, 5): bFirstLine = True
      ElseIf Left(st_ucase, 11) = "PUBLIC SUB " Then
         st = Mid(st, 12): bFirstLine = True
      ElseIf Left(st_ucase, 12) = "PRIVATE SUB " Then
         st = Mid(st, 13): bFirstLine = True
      End If   
         
   
      ' bypass any line with = embedded in it but need to check only up to the first
      ' open parenthesis because after that are the parameters and they could have equal
      ' signs for default parameter values.
      If bFirstLine Then
         n = Instr( st, "=") 
         If n Then 
            iParen = Instr(st, "(") 
            If (iParen = 0) OrElse (n < iParen) Then
               bFirstLine = False: Continue Do
            End If
         End If
         
         pList = New FUNCTION_TYPE
         ' Determine the name of this sub/function
         iParen = Instr(st, "(") 
         If iParen Then
            pList->zFnName = Left(st, iParen - 1)
         Else
            pList->zFnName = st
         End If
         pList->bIsHeader = False
         pList->nLineNumber = nLineNum - 1   'b/c Scintilla is zero based
         pList->pDoc = @this
            
         ' Need to add this function at the end of th linked list (maybe sorted order?)
         pListTemp = m_FnListPtr 
         Do While( pListTemp->pListNext <> 0 )
            pListTemp = pListTemp->pListNext
         Loop
         pListTemp->pListNext = pList
         
         bFirstLine = False
      EndIf
         
   Loop

   Function = 0
End Function


''
''
Function clsDocument.GetWord( ByVal curPos As Long = -1 ) As String

   ' Get word at the specified location or under the cursor
   Dim As Long x, y, p
   Dim As String buffer 

   ' Retrieve the current position
   If curPos = -1 Then curPos = SciMsg( m_pSci, SCI_GETCURRENTPOS, 0, 0)
   ' Retrieve the starting and ending position of the word
   x = SciMsg( m_pSci, SCI_WORDSTARTPOSITION, curPos, True)
   y = SciMsg( m_pSci, SCI_WORDENDPOSITION, curPos, False)
   If y > x Then
      ' Text range
      buffer = this.GetTextRange(x, y)
      ' Remove the $NUL
      p = Instr(buffer, Chr(0))
      If p Then buffer = Left(buffer, p - 1)
   End If
   buffer = FF_RemoveAny( buffer, Chr(13, 10, 34) & "()%," )
   Function = buffer

End Function
