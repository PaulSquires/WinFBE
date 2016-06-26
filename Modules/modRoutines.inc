
' ========================================================================================
' Process prefix {CURDRIVE} and convert to current drive letter.
' ========================================================================================
Function ProcessToCurdrive( ByVal pswzIn As WString Ptr ) As Long
   ' For each folder location determine if it resides on the same drive as
   ' the WinFBE application. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.

   ' It is assumed that the incoming ptr is to a WString of MAX_PATH
   Dim wCurDrive As WString * 10 = LCase(Left(Exepath, 3))  ' eg. D:\

   If LCase(Left(*pswzIn, 3)) = wCurDrive Then 
      *pswzIn = WStr("{CURDRIVE}") & Mid(*pswzIn, 2)
   End If

   Function = 0
End Function


' ========================================================================================
' Process current drive to prefix {CURDRIVE} 
' ========================================================================================
Function ProcessFromCurdrive( ByVal pswzIn As WString Ptr ) As Long
   ' For each folder location determine if it resides on the same drive as
   ' the WinFBE application. If it does then substitute the replaceable parameter
   ' {CURDRIVE} for the drive letter. This allows you to easily run the editor
   ' on different media (eg. thumb drive) that may be assigned a different
   ' drive letter.

   ' It is assumed that the incoming ptr is to a WString of MAX_PATH
   Dim wCurDrive As WString * 10 = LCase(Left(Exepath, 3))  ' eg. D:\

   If Ucase(Left(*pswzIn, 10)) = "{CURDRIVE}" Then 
      *pswzIn = Left(wCurDrive, 1) & Mid(*pswzIn, 11)
   End If

   Function = 0
End Function


' ========================================================================================
' Adds an item to a Treeview
' ========================================================================================
Function FF_TreeView_InsertItem( ByVal hWndControl As HWnd, _
                                 ByVal hParent As HANDLE, _
                                 ByRef TheText As WString, _
                                 ByVal lParam As Long = 0, _
                                 ByVal iImage As Long = 0, _
                                 ByVal iSelectedImage As Long = 0 _
                                 ) As HANDLE

   Dim uInsert As TV_INSERTSTRUCT

   uInsert.hInsertAfter            = TVI_LAST
   uInsert.item.mask               = TVIF_TEXT Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM
   uInsert.hParent                 = hParent
   uInsert.item.pszText            = Cast(WString Ptr, Strptr(TheText))
   uInsert.item.iImage             = iImage 
   uInsert.item.iSelectedImage     = iSelectedImage
   uInsert.item.lParam             = lParam 
   Function = TreeView_InsertItem( hWndControl, @uInsert )
    
End Function


' ========================================================================================
' Retrieve the lParam value from a Treeview
' ========================================================================================
Function FF_TreeView_GetlParam( ByVal hWndControl As HWnd, _
                                ByVal hItem As HANDLE _
                                ) As Long

    Dim ti As TV_ITEM

    ti.hItem = hItem                                                 
    ti.mask  = TVIF_HANDLE Or TVIF_PARAM  
    TreeView_GetItem(hWndControl, Varptr(ti))

    Function = ti.lParam
End Function


' ========================================================================================
' Return WSTRING text for a control
' ========================================================================================
Function FF_Control_GetTextW( ByVal hWndControl   As HWnd, _
                              ByVal pOutString    As WString Ptr, _
                              ByVal pOutStringLen As Long _
                              ) As Long
    
   Dim As Long nLen = GetWindowTextLengthW(hWndControl) + 2  ' add space for ending nuls
   If nLen > pOutStringLen Then nLen = pOutStringLen
   Function = GetWindowTextW( hWndControl, pOutString, nLen )
         
End Function


' ========================================================================================
' Return WSTRING text for a ComboBox
' ========================================================================================
Function FF_ComboBox_GetTextW( ByVal hWndControl   As HWnd, _
                               ByVal nIndex        As Long, _
                               ByVal pOutString    As WString Ptr, _
                               ByVal pOutStringLen As Long _
                               ) As Long

   Dim As Long nLen = SendMessageW(hWndControl, CB_GETLBTEXTLEN, nIndex, 0) + 2  ' add space for ending nuls
   If nLen > pOutStringLen Then Return 0  ' buffer not large enough
   Function = SendMessageW( hWndControl, CB_GETLBTEXT, nIndex, Cast(LPARAM, pOutString) ) 

End Function           


' ========================================================================================
' Displays the FileOpenDialog.
' The returned pointer must be freed with CoTaskMemFree
' ========================================================================================
Function AfxIFileOpenDialogW( ByVal hwndOwner As HWnd, ByVal idButton As Long) As WString Ptr

   Dim hr As Long
   Dim CLSID_FileOpenDialog As CLSID = (&hDC1C5A9C, &hE88A, &h4DDE, {&hA5, &hA1, &h60, &hF8, &h2A, &h20, &hAE, &hF7})
   Dim IID_IFileOpenDialog As GUID   = (&hD57C7288, &hD4AD, &h4768, {&hBE, &h02, &h9D, &h96, &h95, &h32, &hD9, &h60})

   ' Create an instance of the FileOpenDialog object
   Dim pofd As IFileOpenDialog Ptr
   hr = CoCreateInstance(@CLSID_FileOpenDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @pofd)
   If pofd = Null Then Return Null

   ' Set the file types depending on the button pushed that calls this open dialog
   Dim rgFileTypes(1 To 3) As COMDLG_FILTERSPEC

   Select Case idButton
      Case IDM_INSERTFILE
         rgFileTypes(1).pszName = @WSTR("FB code files")
         rgFileTypes(2).pszName = @WSTR("Resource files")
         rgFileTypes(3).pszName = @WSTR("All files")
         rgFileTypes(1).pszSpec = @WSTR("*.bas;*.inc;*.bi")
         rgFileTypes(2).pszSpec = @WSTR("*.rc")
         rgFileTypes(3).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 3, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, "Insert File")

      Case IDC_FRMOPTIONSCOMPILER_CMDFBWIN32, IDC_FRMOPTIONSCOMPILER_CMDFBWIN64
         rgFileTypes(1).pszName = @WSTR("FB Compiler")
         rgFileTypes(1).pszSpec = @WSTR("fbc.exe")
         rgFileTypes(2).pszName = @WSTR("All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, "Find FreeBASIC Compiler")
   
      Case IDC_FRMOPTIONSEDITOR_CMDLOCALIZATION
         rgFileTypes(1).pszName = @WSTR("Localization files")
         rgFileTypes(1).pszSpec = @WSTR("*.lang")
         rgFileTypes(2).pszName = @WSTR("All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, "Open Localization File")

      Case IDC_FRMOPTIONSCOMPILER_CMDFBHELPFILE, IDC_FRMOPTIONSCOMPILER_CMDAPIHELPPATH
         rgFileTypes(1).pszName = @WSTR("Help file")
         rgFileTypes(1).pszSpec = @WSTR("*.chm")
         rgFileTypes(2).pszName = @WSTR("All files")
         rgFileTypes(2).pszSpec = @WSTR("*.*")
         pofd->lpVtbl->SetFileTypes(pofd, 2, @rgFileTypes(1))
         ' Set the title of the dialog
         hr = pofd->lpVtbl->SetTitle(pofd, "Find Help File")
   End Select
   
   ' Display the dialog
   hr = pofd->lpVtbl->Show(pofd, hwndOwner)
   hr = pofd->lpVtbl->SetOptions(pofd, FOS_NOCHANGEDIR)

   ' Get the result
   Dim pItem As IShellItem Ptr
   Dim pwszName As WString Ptr
   If SUCCEEDED(hr) Then
      hr = pofd->lpVtbl->GetResult(pofd, @pItem)
      If SUCCEEDED(hr) Then
         hr = pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, @pwszName)
         Function = pwszName
      End If
   End If

   ' Cleanup
   If pItem Then pItem->lpVtbl->Release(pItem)
   If pofd Then pofd->lpVtbl->Release(pofd)

End Function


' ========================================================================================
' Displays the FileOpenDialog (multiple selection)
' Returns a pointer to the IShellItemArray collection.
' ========================================================================================
Function AfxIFileOpenDialogMultiple( ByVal hwndOwner As HWnd, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH) As IShellItemArray Ptr

   ' Create an instance of the FileOpenDialog interface
   Dim hr As Long
   Dim pofd As IFileOpenDialog Ptr
   hr = CoCreateInstance(@CLSID_FileOpenDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @pofd)
   If pofd = Null Then Return Null

   ' Set the file types
   Dim rgFileTypes(1 To 3) As COMDLG_FILTERSPEC
   rgFileTypes(1).pszName = @WSTR("FB code files")
   rgFileTypes(2).pszName = @WSTR("Resource files")
   rgFileTypes(3).pszName = @WSTR("All files")
   rgFileTypes(1).pszSpec = @WSTR("*.bas;*.inc;*.bi")
   rgFileTypes(2).pszSpec = @WSTR("*.rc")
   rgFileTypes(3).pszSpec = @WSTR("*.*")
   pofd->lpVtbl->SetFileTypes(pofd, 3, @rgFileTypes(1))

   ' Set the title of the dialog
   hr = pofd->lpVtbl->SetTitle(pofd, "Open source file")
   ' Set the default folder
   Dim pFolder As IShellItem Ptr
   SHCreateItemFromParsingName (Exepath, Null, @IID_IShellItem, @pFolder)
   If pFolder Then
      pofd->lpVtbl->SetDefaultFolder(pofd, pFolder)
      pFolder->lpVtbl->Release(pFolder)
   End If
   ' Allow multiselection
   hr = pofd->lpVtbl->SetOptions(pofd, FOS_ALLOWMULTISELECT Or FOS_NOCHANGEDIR)
   ' Display the dialog
   hr = pofd->lpVtbl->Show(pofd, hwndOwner)

   ' Get the result
   Dim pItemArray As IShellItemArray Ptr
   If SUCCEEDED(hr) Then
      hr = pofd->lpVtbl->GetResults(pofd, @pItemArray)
      Function = pItemArray
   End If

   If pofd Then pofd->lpVtbl->Release(pofd)

End Function


' ========================================================================================
' Displays the FileSaveDialog
' The returned pointer must be freed with CoTaskMemFree
' ========================================================================================
Function AfxIFileSaveDialog (ByVal hwndOwner As HWnd, ByVal pwszFileName As WString Ptr, _
ByVal pwszDefExt As WString Ptr, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH) As WString Ptr

   ' // Create an instance of the IFileSaveDialog interface
   Dim hr As Long
   Dim psfd As IFileSaveDialog Ptr
   hr = CoCreateInstance(@CLSID_FileSaveDialog, Null, CLSCTX_INPROC_SERVER, @IID_IFileSaveDialog, @psfd)
   If psfd = Null Then Return Null

   ' // Set the file types
   Dim rgFileTypes(1 To 3) As COMDLG_FILTERSPEC
   rgFileTypes(1).pszName = @WSTR("FB code files")
   rgFileTypes(2).pszName = @WSTR("Resource files")
   rgFileTypes(3).pszName = @WSTR("All files")
   rgFileTypes(1).pszSpec = @WSTR("*.bas;*.inc;*.bi")
   rgFileTypes(2).pszSpec = @WSTR("*.rc")
   rgFileTypes(3).pszSpec = @WSTR("*.*")
   psfd->lpVtbl->SetFileTypes(psfd, 3, @rgFileTypes(1))

   ' // Set the title of the dialog
   hr = psfd->lpVtbl->SetTitle(psfd, "File Save Dialog")
   ' // Set the file name
   hr = psfd->lpVtbl->SetFileName(psfd, pwszFileName)
   ' // Set the extension
   hr = psfd->lpVtbl->SetDefaultExtension(psfd, pwszDefExt)
   ' // Display the dialog
   hr = psfd->lpVtbl->Show(psfd, hwndOwner)

   ' // Set the default folder
   Dim pFolder As IShellItem Ptr
   SHCreateItemFromParsingName (Exepath, Null, @IID_IShellItem, @pFolder)
   If pFolder Then
      psfd->lpVtbl->SetDefaultFolder(psfd, pFolder)
      pFolder->lpVtbl->Release(pFolder)
   End If

   ' // Get the result
   Dim pItem As IShellItem Ptr
   Dim pwszName As WString Ptr
   If SUCCEEDED(hr) Then
      hr = psfd->lpVtbl->GetResult(psfd, @pItem)
      If SUCCEEDED(hr) Then
         hr = pItem->lpVtbl->GetDisplayName(pItem, sigdnName, @pwszName)
         Function = pwszName
      End If
   End If

   ' // Cleanup
   If pItem Then pItem->lpVtbl->Release(pItem)
   If psfd Then psfd->lpVtbl->Release(psfd)

End Function


' ========================================================================================
' FF_PATHNAME
' Parse a path/file name to extract component parts
' ========================================================================================
''  This function evaluates a text path/file text name, and returns a requested 
''  part of the name. The functionality is strictly one of string parsing alone. 
''  
''  sOption one of the following words which is used to specify the requested part:
''  FULL
''        Returns the full path/file name, just as given in the sFilespec parameter. 
''  PATH
''        Returns the path portion of the path/file Name. That is the text up to and 
''        including the last backslash (\) or colon (:).
''  NAME
''        Return the name portion of the path/file Name. That is the text To the right 
''        of the last backslash (\) or colon (:), ending just before the last period (.).
''  EXTN
''        Returns the extension portion of the path/file name. That is the last 
''        period (.) in the string plus the text to the right of it.
''  NAMEX
''        Returns the name and the EXTN parts combined.
'' 
Function FF_Pathname( ByVal wOption   As WString Ptr, _
                      ByVal InFilespec As WString Ptr, _
                      ByVal OutFilespec As WString Ptr _
                      ) As Long

    Dim nLen As Long = Len(*inFilespec)
    Dim s    As WString * MAX_PATH  = *InFilespec
    Dim st   As WString * MAX_PATH  = ""
    Dim i    As Long = 0

    If nLen = 0 Then 
       *OutFilespec = "": Exit Function
    End If   

    Select Case Ucase(Trim(*wOption))
       Case "FULL" 
          *OutFilespec = *InFilespec
          Exit Function
       
       Case "PATH"
          i = InStrRev(s, Any ":\/")
          If i Then 
             *OutFilespec = Left(s,i): Exit Function
          Else
             *OutFilespec = "": Exit Function
          End If
             
       Case "NAME"
          ' Get the full filename
          i = InStrRev(s, Any ":/\")
          st = Iif( i, Mid(s, i + 1), s )
          ' Get the name portion
          i = InStrRev(st, ".")
          st = Iif( i, Left(st, i - 1), st )
          *OutFilespec = st: Exit Function
       
       Case "EXTN"
          ' Get the full filename
          i = InStrRev(s, Any ":/\")
          st = Iif( i, Mid(s, i + 1), s )
          ' Get the extension
          i = InStrRev(st, ".")
          st = Iif( i, Mid(st, i), WStr("") )
          *OutFilespec = st: Exit Function
       
       Case "NAMEX"
          i = InStrRev(s, Any WStr(":/\"))
          If i Then  
             *OutFilespec = Mid(s, i + 1): Exit Function
          Else
             *OutFilespec = *InFilespec
          End If   
    End Select
    
    Function = 0
    
End Function

''
''  FF_REPLACE (ANSI ONLY)
''  Within a specified string, replace all occurrences of one string with another string.
''
''  Replaces all occurrences of MatchPattern in MainString with ReplaceWith.  
''  The replacement can cause MainString to grow or condense in size. 
''  This function is case-sensitive. 
''  When a match is found, the scan for the next match begins at the position 
''  immediately following the prior match.
''
Function FF_Replace( ByRef sMainString   As String, _
                     ByRef sMatchPattern As String, _
                     ByRef sReplaceWith  As String _
                     ) As String

    Dim i As Integer 
    Dim s As String 
   
    s = sMainString
    i = 1
    Do
        i = Instr(i, s, sMatchPattern)
        If i > 0 Then 
           s = Left(s, i - 1) & sReplaceWith & Mid(s, i + Len(sMatchPattern))
           i += Len(sReplaceWith)
        End If   
    Loop Until i = 0
   
    Return s

End Function

' ========================================================================================
' Enables the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
Function FF_Toolbar_EnableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
   ' Only enable the button if it is not already enabled. This should help reduce flicker.
   If SendMessageW(hToolBar, TB_ISBUTTONENABLED, Cast(WPARAM, idButton), 0 ) = 0 Then
      Function = SendMessageW(hToolBar, TB_ENABLEBUTTON, Cast(WPARAM, idButton), Cast(LPARAM, MAKELONG(CTRUE, 0)))
   End If   
End Function
' ========================================================================================

' ========================================================================================
' Disables the specified button in a toolbar.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
Function FF_Toolbar_DisableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
   ' Only disable the button if it is not already disabled. This should help reduce flicker.
   If SendMessageW(hToolBar, TB_ISBUTTONENABLED, Cast(WPARAM, idButton), 0 ) Then
      Function = SendMessageW(hToolBar, TB_ENABLEBUTTON, Cast(WPARAM, idButton), Cast(LPARAM, MAKELONG(False, 0)))
   End If
End Function


''
''  FF_PARSE_INTERNAL (ANSI ONLY)
''  Used by both FF_Parse and FF_ParseAny (internal function)
''
Function FF_Parse_Internal( ByRef sMainString   As String, _
                            ByRef sDelimiter    As String, _
                            ByRef nPosition     As Long, _
                            ByRef nIsAny        As BOOLEAN, _
                            ByRef nLenDelimiter As Long _
                            ) As String
                   
    ' Returns the nPosition-th substring in a string sMainString with separations sDelimiter (one or more charcters),
    ' beginning with nPosition = 1

    Dim As Long i, j, count
    Dim s As String 
    Dim fReverse As BOOLEAN = Iif(nPosition < 0, True, False)
    
    nPosition = Abs(nPosition)
    count = 0
    i = 1
    s = sMainString 
    
    If fReverse Then 
       ' Reverse search
       ' Get the start of the token (j) by searching in reverse
       If nIsAny Then
          i = InStrRev(sMainString, Any sDelimiter)
       Else
          i = InStrRev(sMainString, sDelimiter)
       End If
       Do While i > 0        ' if not found loop will be skipped
          j = i + nLenDelimiter
          count += 1
          i = i - nLenDelimiter
          If count = nPosition Then Exit Do
          If nIsAny Then
             i = InStrRev(sMainString, Any sDelimiter, i )
          Else
             i = InStrRev(sMainString, sDelimiter, i )
          End If   
       Loop
       If i = 0 Then j = 1

       ' Now continue forward to get the end of the token
       If nIsAny Then
          i = Instr(j, sMainString, Any sDelimiter)
       Else
          i = Instr(j, sMainString, sDelimiter)
       End If
       If (i > 0) Or (count = nPosition) Then
           If i = 0 Then
               s = Mid(sMainString, j)
           Else
               s = Mid(sMainString, j, i - j)
           End If
       End If
    
    Else              
    
       ' Forward search
       Do
           j = i
           If nIsAny Then
              i = Instr(i, sMainString, Any sDelimiter)
           Else
              i = Instr(i, sMainString, sDelimiter)
           End If   
           If i > 0 Then count += 1: i += nLenDelimiter
       Loop Until (i = 0) Or (count = nPosition)

       If (i > 0) Or (count = nPosition - 1) Then
           If i = 0 Then
               s = Mid(sMainString, j)
           Else
               s = Mid(sMainString, j, i - nLenDelimiter - j)
           End If
       End If
    End If
      
    Return s

End Function


''
''  FF_PARSE (ANSI ONLY)
''  Return a delimited field from a string expression.
''
''  Delimiter contains a string of one or more characters that must be fully matched to be successful.
''  If nPosition evaluates to zero or is outside of the actual field count, an empty string is returned.
''  If nPosition is negative then fields are searched from the right to left of the MainString.
''  Delimiters are case-sensitive,
''
Function FF_Parse( ByRef sMainString As String, _
                   ByRef sDelimiter  As String, _
                   ByVal nPosition   As Long _
                   ) As String
   ' The parse must match the entire deliminter string
   Function = FF_Parse_Internal( sMainString, sDelimiter, nPosition, False, Len(sDelimiter) )
End Function


''
''  FF_PARSECOUNT (ANSI ONLY)
''  Return the count of delimited fields from a string expression.
''
''  If MainString is empty (a null string) or contains no delimiter character(s), the string 
''  is considered to contain exactly one sub-field. In this case, FF_PARSECOUNT returns the value 1.
''  Delimiter contains a string (one or more characters) that must be fully matched.
''  Delimiters are case-sensitive,
''
Function FF_ParseCount( ByRef sMainString As String, _
                        ByRef sDelimiter  As String _
                        ) As Long
                   
    Dim As Long count = 1
    Dim As Long i = 1
   
    Do
        i = Instr(i, sMainString, sDelimiter)
        If i > 0 Then count += 1: i += Len(sDelimiter)
    Loop Until i = 0
   
    Return count
End Function


''
''  FF_REMOVE (ANSI ONLY)
''  Return a copy of a string with characters or strings removed
''
''  If MatchPattern is not present in MainString, all of MainString is returned intact.
''  The replacement can cause MainString to grow or condense in size. 
''  This function is case-sensitive. 
''
Function FF_Remove( ByRef sMainString   As String, _
                    ByRef sMatchPattern As String _
                    ) As String

    Dim i As Long 
    Dim s As String 
   
    If Len(sMainString) = 0 Then Return sMainString

    s = sMainString
    Do
        i = Instr(s, sMatchPattern)
        If i > 0 Then 
           s = Left(s, i - 1) & Mid(s, i + Len(sMatchPattern))
        End If   
    Loop Until i = 0
   
    Return s

End Function


''
''  FF_STRDELETE (ANSI ONLY)
''  Delete a specified number of characters from a string expression.
''
''  Returns a string based on MainString but with nCount characters deleted 
''  starting at position nStart. The first character in the string is position 1, etc.
'' 
Function FF_StrDelete( ByRef sMainString As String, _
                       ByRef nStart      As Long, _
                       ByRef nCount      As Long _
                       ) As String

   Dim nLenMain As Long = Len(sMainString)
   
   If nStart > nLenMain Then Return sMainString
   If nStart < 0        Then Return sMainString
   If nCount <= 0       Then Return sMainString
   If nStart = 0        Then nStart = nLenMain - nCount
      
   Dim s As String
   s = Left(sMainString, nStart-1) & Mid(sMainString, nStart+nCount+1)
   Function = s
End Function


'   Inserts an item at a specific location in the ListView.
'
'   hWndControl:    Handle of control (e.g. HWND_FORM1_LISTVIEW1)   
'   iRow:           The zero-based row to insert this item into.
'   iColumn:        The zero-based column to insert this item into.
'   TheText:        A string representing the text to display.   
'   lParam:         A 32-bit value to be stored with the item. (optional. Only for Col0)
'            
'   Returns:        If adding an item at Column 0 then the function
'                   will return the index of the new item. If inserting
'                   to other than Column 0 then the function returns
'                   TRUE if successful, FALSE if otherwise.
'
Function FF_ListView_InsertItem( ByVal hWndControl As HWnd, _
                                 ByVal iRow        As Long, _         
                                 ByVal iColumn     As Long, _
                                 ByVal pwszText    As WString Ptr, _
                                 ByVal lParam      As Long = 0 _
                                 ) As BOOLEAN

   Dim lvi As LVITEMW

   lvi.iItem     = iRow
   lvi.iSubItem  = iColumn 
   lvi.pszText   = pwszText
   lvi.lParam    = lParam
  
   If iColumn = 0 Then
      lvi.mask = LVIF_TEXT Or LVIF_PARAM Or LVIF_IMAGE 
      Function = SendMessageW( hWndControl, LVM_INSERTITEM, 0, Cast(LPARAM, @lvi) )
   Else 
      lvi.mask = LVIF_TEXT Or LVIF_IMAGE
      Function = SendMessageW( hWndControl, LVM_SETITEM, 0, Cast(LPARAM, @lvi) )
   End If
End Function


'   Retrieves the text of a ListView item.
'
'   hWndControl:  Handle of control (e.g. HWND_FORM1_LISTVIEW1)   
'   iRow:         The zero-based row to retrieve.
'   iColumn:      The zero-based column to retrieve.
'   pwszText:     Pointer to wString to hold the retrieved text,
'   nTextMax:     Size of text buffer pointed to by pwszText.           
'   Returns:      Returns TRUE if successful, FALSE otherwise.
'
Function FF_ListView_GetItemText( ByVal hWndControl As HWnd, _
                                  ByVal iRow As Long, _
                                  ByVal iColumn As Long, _
                                  ByVal pwszText As WString Ptr, _
                                  ByVal nTextMax As Long _
                                  ) As BOOLEAN

   If pwszText = 0 Then Return False
   If nTextMax = 0 Then Return False
   
   Dim lvi As LVITEMW

   lvi.mask       = LVIF_TEXT
   lvi.iItem      = iRow
   lvi.iSubItem   = iColumn 
   lvi.pszText    = pwszText
   lvi.cchTextMax = nTextMax
       
   Function = SendMessageW( hWndControl, LVM_GETITEM, 0, Cast(LPARAM, @lvi) )
End Function

' ========================================================================================
' Load a .lang localization file from disk and populate the localization array
' ========================================================================================
Function LoadLocalizationFile( ByVal pwszName As WString Ptr ) As BOOLEAN

   ' default that the file failed to load
   Function = False
   If pwszName = 0 Then Exit Function
   If Dir(*pwszName) = "" Then Exit Function

   Dim st    As WString * MAX_PATH
   Dim wKey  As WString * MAX_PATH
   Dim wData As WString * MAX_PATH
   Dim nKey  As Long
   Dim nData As Long  
   Dim i     As Long
    
   Dim f As Long = Freefile
   If Open( *pwszName For Input Encoding "utf16" As #f ) <> 0 Then Exit Function

   Do Until Eof(f)
      Line Input #f, st
      If Len(st) = 0 Then Continue Do
      If Left(st, 1) = "'" Then Continue Do
      
      i = Instr(st, ":")
      If i = 0 Then Continue Do
      
      wKey  = Left(st, i-1)
      wData = Mid(st, i+1)
      
      nKey  = Val(wKey)
      nData = Val(wData)

      If Ucase(wKey) = "MAXIMUM" Then
         ' resize the global dynamic array
         ReDim LL(nData) As WString * MAX_PATH
      Else
         ' this should be a key/value pair line in the format:
         ' 00001:value
         ' Ensure that we add the value to the array within the valid
         ' boundaries of the array.
         If (nKey >= LBound(LL)) AndAlso (nKey <= Ubound(LL)) Then
            LL(nKey) = wData
         End If
      End If   
         
   Loop
   Close #f

   Function = True
End Function

