

''
''
Constructor clsApp()
   m_IsProjectActive = False
   this.LoadKeywords
End Constructor


''
''
Destructor clsApp()
'
End Destructor 


''
''
Function clsApp.LoadKeywords() as Long
   Dim st as String 
   Dim f  as Long = Freefile
   
   If Open( Exepath & "\Settings\freebasic_keywords.txt" For Input As #f) = 0 Then
      this.FBKeywords  = ""
      Do Until EOF(f)
         Line Input #f, st
         If Len(st) Then this.FBKeywords = this.FBKeywords & st & " "
      Loop    
      this.FBKeywords = Trim(this.FBKeywords)
      Close #f
   End If
   
   Function = 0
End Function

''
''
Function clsApp.SaveKeywords() As Long
   Dim f As Long = Freefile
   
   If Open( Exepath & "\Settings\freebasic_keywords.txt" For Output As #f) = 0 Then
      Print #f, this.FBKeywords
      Close #f
   End If
   
   Function = 0
End Function

''
''
Property clsApp.ProjectThread( ByVal nValue As Long)
   m_ProjectThread = nValue
End Property
''
''
Property clsApp.ProjectThread() As Long
   Property = m_ProjectThread
End Property


Property clsApp.ProjectSuppressConsole( ByVal nValue As Long)
   m_ProjectSuppressConsole = nValue
End Property
''
''
Property clsApp.ProjectSuppressConsole() As Long
   Property = m_ProjectSuppressConsole
End Property


''
''
Property clsApp.SuppressNotify( ByVal nValue As BOOLEAN)
   m_SuppressNotify = nValue
End Property

Property clsApp.SuppressNotify() As BOOLEAN
   Property = m_SuppressNotify
End Property


''
''
Property clsApp.ProjectType( ByVal nValue As Long)
   m_ProjectType = nValue
End Property

Property clsApp.ProjectType() As Long
   Property = m_ProjectType
End Property

''
''
Property clsApp.ProjectDebug( ByVal nValue As Long)
   m_ProjectDebug = nValue
End Property

Property clsApp.ProjectDebug() As Long
   Property = m_ProjectDebug
End Property


''
''
Property clsApp.ProjectErrorOption( ByVal nValue As Long)
   m_ProjectErrorOption = nValue
End Property

Property clsApp.ProjectErrorOption() As Long
   Property = m_ProjectErrorOption
End Property


''
''
Property clsApp.IsProjectActive( ByVal nValue As BOOLEAN)
   m_IsProjectActive = nValue
   If m_IsProjectActive = False Then 
      SetWindowTextW( HWND_FRMMAIN, APPNAME )
   End If   
End Property

Property clsApp.IsProjectActive() As BOOLEAN
   Property = m_IsProjectActive
End Property


''
''
Property clsApp.IsNewProjectFlag( ByVal nValue As BOOLEAN)
   m_IsNewProjectFlag = nValue
End Property

Property clsApp.IsNewProjectFlag() As BOOLEAN
   Property = m_IsNewProjectFlag
End Property


''
''
Function clsApp.SaveProject( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN

   Dim wFilename As WString * MAX_PATH
   Dim wText     As WString * MAX_PATH  
   
   ' If this a new unsaved project set flag to display SaveAs dialog.
   If this.IsNewProjectFlag Then bSaveAs = True
      
   wFilename = this.ProjectFilename
   
   If bSaveAs Then
      ' Display the Save File Dialog
      wText = AfxStrPathname( "NAMEX", wFilename )
      Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wText, @WStr("wfbe"), IDM_PROJECTSAVE)
      If pwszName Then
         wFilename = *pwszName
         CoTaskMemFree(pwszName)
      Else
         Return False
      End If
   End If

   ' Delete any existing file
   If AfxFileExists(wFilename) Then Kill wFilename
   this.ProjectFilename = wFilename
   this.ProjectName = AfxStrPathname( "NAMEX", wFilename )
   gConfig.ProjectSaveToFile()
   
   this.IsNewProjectFlag = False

   ' Also need to add this new project name to the MRU list.
   UpdateMRUProjectList(@wFilename)
   
   ' Ensure that the caption bar and Explorer root node is updated with the name
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   frmMain_UpdateLineCol(HWND_FRMMAIN)    ' this will update the caption
   wFilename = ucase(gApp.ProjectName)
   TreeView_SetItemText( hTree, gApp.hExplorerRootNode, @wFilename)
   
   frmMain_PositionWindows(HWND_FRMMAIN)

   Function = True
End Function
      

''
''
Function clsApp.ProjectRemoveFile( ByVal pDoc As clsDocument Ptr ) As LRESULT

   ' Remove the incoming document from the project. We simply need to unset the flag for
   ' instance of the document and update the Explorer Treeview.
   pDoc->IsProjectFile = False
   ' Remove the document from the Explorer treeview (deletes the node)
   Treeview_DeleteItem( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), pDoc->hNodeExplorer)
   pDoc->hNodeExplorer = 0

   Function = 0
End Function


''
''
Function clsApp.ProjectSetFileType( ByVal pDoc As clsDocument Ptr, _
                                    ByVal nFileType As Long _       ' new filetype 
                                    ) As LRESULT

   ' If we are setting a "Main" project file then we need to toggle any other Main 
   ' file to "Normal". There can only be one Main file. Likewise, there can only be 
   ' one "Resource" file.
   If nFileType = FILETYPE_MAIN Then
      For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
         If m_arrDocuments(i)->ProjectFileType = FILETYPE_MAIN Then m_arrDocuments(i)->ProjectFileType = FILETYPE_NORMAL
      Next   
   End If      
   If nFileType = FILETYPE_RESOURCE Then
      For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
         If m_arrDocuments(i)->ProjectFileType = FILETYPE_RESOURCE Then m_arrDocuments(i)->ProjectFileType = FILETYPE_NORMAL
      Next   
   End If      

   pDoc->ProjectFileType = nFileType

   ' Refresh the statusbar to ensure that the file's type is displayed properly
   pDoc->DisplayStats
   
   Function = 0
End Function


''
''
Function clsApp.ProjectAddFile( ByVal pDoc As clsDocument Ptr ) As LRESULT
   If pDoc = 0 Then Exit Function
   
   Dim wzFileType As WString * MAX_PATH 
   Dim wzFilePath As WString * MAX_PATH 
   Dim wzFileExt  As WString * MAX_PATH 
   Dim wzFile     As WString * MAX_PATH
   
   ' Add the incoming document to the project. We simply need to set the flag for
   ' instance of the document and update the Project Manager listview.
   pDoc->IsProjectFile = True
   
   wzFile     = AfxStrPathname( "NAMEX", pDoc->DiskFilename )
   wzFilePath = AfxStrPathname( "PATH",  pDoc->DiskFilename )
   wzFileExt  = AfxStrPathname( "EXTN",  pDoc->DiskFilename )
 
   ' Determine if the document already exists in the project. If it does not then
   ' make a determination of a default file type for this new file being added to the project.
   Dim bFound As BOOLEAN = False
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      If m_arrDocuments(i) = pDoc Then
         If pDoc->IsProjectFile = False Then
            bFound = True: Exit For
         End If
      End If
   Next
   If bFound = False Then
      If pDoc->ProjectFileType = FILETYPE_UNDEFINED Then
         Select Case Ucase(wzFileExt)
            Case ".BAS": pDoc->ProjectFileType = FILETYPE_MODULE
            Case ".RC":  pDoc->ProjectFileType = FILETYPE_RESOURCE
            Case ".BI":  pDoc->ProjectFileType = FILETYPE_NORMAL
            Case Else:   pDoc->ProjectFileType = FILETYPE_NORMAL
         End Select   
      End If
   End If
   
   this.ProjectSetFileType(pDoc, pDoc->ProjectFileType)

   Function = 0
End Function


''
''
Function clsApp.GetDocumentCount() As Long
   Function = Ubound(m_arrDocuments) + 1 
End Function

''
''
Function clsApp.GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
   Dim lb As Long = LBound(m_arrDocuments) 
   Dim ub As Long = Ubound(m_arrDocuments) 
   If (idx >= lb) AndAlso (idx <= ub) Then
      Function = m_arrDocuments(idx) 
   End If
End Function

''
''
Function clsApp.AddDocument( ByVal pDoc As clsDocument Ptr ) As Long
   Dim ub as Long = UBound(m_arrDocuments) + 1
 
   ReDim Preserve m_arrDocuments(ub) As clsDocument Ptr             
   m_arrDocuments(ub) = pDoc

   ' Parse the document to create sub/function list
   pDoc->CreateFunctionList
   
   Function = 0
End Function

''
''
Function clsApp.RemoveDocumentFromArray( ByVal idx As Long) As Long
   ' This is an internal function that is called by RemoveDocument. It
   ' Simply resizes the class array that holds all of the pDoc documents.
   Dim i   as Long 
   Dim lb  As Long = LBound(m_arrDocuments)
   Dim ub  As Long = Ubound(m_arrDocuments) 
   If idx < lb Then Exit Function
   If idx > ub Then Exit Function
   
   Dim pDoc As clsDocument Ptr = m_arrDocuments(idx)
   
   ' remove the document from the array
   For i = idx To ub - 1
      m_arrDocuments(i) = m_arrDocuments(i+1)
   Next

   ' resize the array to the new smaller size
   If (ub-1) < 0 Then
      Erase m_arrDocuments
   Else
      ReDim Preserve m_arrDocuments(ub-1) As clsDocument Ptr             
   End If
   
   ' Finally, release memory associated with this pDoc
   Delete pDoc
 
   Function = 0
End Function

''
''
Function clsApp.RemoveDocument( ByVal pDoc As clsDocument Ptr ) As Long
   Dim idx as Long = -1
   Dim i   as Long 
   Dim lb  As Long = LBound(m_arrDocuments)
   Dim ub  As Long = Ubound(m_arrDocuments) 
   
   ' Find the pDoc in our application's document array
   For i = lb To ub
      If m_arrDocuments(i) = pDoc Then
         idx = i: Exit For
      End If   
   Next
   ' If not found then simply exit
   If idx = -1 Then Exit function

   this.RemoveDocumentFromArray(idx)
   
   Function = 0
End Function

''
''
Function clsApp.RemoveAllDocuments() As Long
   Dim lb  As Long = LBound(m_arrDocuments)
   Dim ub  As Long = Ubound(m_arrDocuments) 
   
   do until ub < lb
      this.RemoveDocumentFromArray(ub)
      ub = Ubound(m_arrDocuments) 
   loop
   
   Function = 0
End Function

''
''
Function clsApp.GetDocumentPtrByFilename( ByVal pswzName As WString Ptr ) As clsDocument Ptr
   If pswzName = 0 Then Return 0
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      If Ucase(m_arrDocuments(i)->DiskFilename) = Ucase(*pswzName) Then
         Return m_arrDocuments(i)
      End If
   Next
   function = 0
End Function


''
''
Function clsApp.GetMainDocumentPtr() As clsDocument Ptr
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      If m_arrDocuments(i)->ProjectFileType = FILETYPE_MAIN Then
         Return m_arrDocuments(i)
      End If
   Next
End Function

''
''
Function clsApp.GetResourceDocumentPtr() As clsDocument Ptr
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      If m_arrDocuments(i)->ProjectFileType = FILETYPE_RESOURCE Then
         Return m_arrDocuments(i)
      End If
   Next
End Function

''
''
'Function clsApp.Debug() As Long
'   Dim As Long lb = LBound(m_arrDocuments) 
'   Dim As Long ub = Ubound(m_arrDocuments)
'   
'   Print "--- DEBUG ---"
'   Print "lb = "; lb, "ub = "; ub
'   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
'      Print "Array[";i;"]", "hWindow="; m_arrDocuments(i)->hWindow
'   Next                          
'   Print:Print
'   
'   Function = 0
'End Function
'





















