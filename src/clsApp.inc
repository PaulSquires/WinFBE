

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
Property clsApp.IncludeFilename( ByVal sFilename As String)
   m_IncludeFilename = sFilename
End Property

Property clsApp.IncludeFilename() As String
   Property = m_IncludeFilename
End Property


Property clsApp.ProjectThread( ByVal nValue As Long)
   m_ProjectThread = nValue
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

Property clsApp.ProjectThread() As Long
   Property = m_ProjectThread
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
Property clsApp.ProjectOther( ByVal wValue As WString Ptr)
   m_wzProjectOther = *wValue
End Property

Property clsApp.ProjectOther() As WString Ptr
   Property = @m_wzProjectOther
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
Property clsApp.ProjectName( ByVal wValue As WString Ptr)
   m_wzProjectName = *wValue
End Property

Property clsApp.ProjectName() As WString Ptr
   Property = @m_wzProjectName
End Property


''
''
Property clsApp.ProjectDiskFilename( ByVal wValue As WString Ptr)
   m_wzProjectFilename = *wValue
End Property

Property clsApp.ProjectDiskFilename() As WString Ptr
   Property = @m_wzProjectFilename
End Property


''
''
Function clsApp.SaveProject( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN

   Dim wFilename As WString * MAX_PATH
   Dim wText     As WString * MAX_PATH  
   
   ' If this a new unsaved project set flag to display SaveAs dialog.
   If this.IsNewProjectFlag Then bSaveAs = True
      
   wFilename = *this.ProjectDiskfilename
   
   If bSaveAs Then
      ' Display the Save File Dialog
      FF_PATHNAME( @WStr("NAMEX"), @wFilename, @wText)
      Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wText, @WStr("wfbe"), IDM_PROJECTSAVE)
      If pwszName Then
         wFilename = *pwszName
         CoTaskMemFree(pwszName)
      Else
         Return False
      End If
   End If

   ' Delete any existing file
   If Dir(wFilename) > "" Then Kill wFilename
   this.ProjectDiskFilename = @wFilename
   gConfig.ProjectSaveToFile()
   
   this.IsNewProjectFlag    = False
   ' Ensure that the Project Manager title is correct
   frmProjectManager_PositionWindows(HWND_FRMPROJECTMANAGER)
   
   Function = True
End Function


''
''
Function clsApp.ProjectRemoveFile( ByVal pDoc As clsDocument Ptr ) As LRESULT

   ' Remove the incoming document from the project. We simply need to unset the flag for
   ' instance of the document and update the Project Manager listview.
   pDoc->IsProjectFile = False
   
   Dim As HWnd hLV = GetDlgItem( HWND_FRMPROJECTMANAGER, IDC_FRMPROJECTMANAGER_LISTVIEW)
   Dim As Long nCount = ListView_GetItemCount(hLV)
   Dim As Long i, nFound = -1

   For i = 0 To nCount - 1
       If FF_ListView_GetlParam( hLV, i ) = pDoc Then
          ' If the document is not currently also loaded in a Tab then we
          ' can now safely destroy the document and any allocated memory.
          If gTTabCtl.GetTabIndexByDocumentPtr(pDoc) >= 0 Then
             pDoc->Destructor    ' need to manually call the destructor
          End If
          nFound = i: Exit For
       End If   
   Next
   
   If nFound <> -1 Then ListView_DeleteItem( hLV, nFound)
   
   Function = 0
End Function


''
''
Function clsApp.ProjectSetFileType( ByVal pDoc As clsDocument Ptr, _
                                    ByVal nFileType As Long _       ' new filetype 
                                    ) As LRESULT

   Dim wzFileType As WString * MAX_PATH 
   Dim wzFilePath As WString * MAX_PATH 
   Dim wzFile     As WString * MAX_PATH
   
   Dim As HWnd hLV = GetDlgItem( HWND_FRMPROJECTMANAGER, IDC_FRMPROJECTMANAGER_LISTVIEW)
   Dim As Long nCount = ListView_GetItemCount(hLV)
   Dim As Long i, nFound = -1
   Dim pDocList As clsDocument Ptr
   
   ' Find the position in the Listview where this document is located. Also, if we are setting a "Main"
   ' project file then we need to toggle any other Main file to "Normal". There can only be one Main file.
   ' Likewise, there can only be one "Resource" file.
   For i = 0 To nCount - 1                     
       pDocList = Cast( clsDocument Ptr, FF_ListView_GetlParam(hLV, i) )
       If pDocList = pDoc Then nFound = i
       If nFileType = FILETYPE_MAIN Then
          If pDocList->ProjectFileType = FILETYPE_MAIN Then
             pDocList->ProjectFileType = FILETYPE_NORMAL
             wzFileType = ""  'L(210,"Normal")
             FF_ListView_SetItemText( hLV, i, 1, wzFileType, MAX_PATH )
          End If
       End If   
       If nFileType = FILETYPE_RESOURCE Then
          If pDocList->ProjectFileType = FILETYPE_RESOURCE Then
             pDocList->ProjectFileType = FILETYPE_NORMAL
             wzFileType = ""  'L(210,"Normal")
             FF_ListView_SetItemText( hLV, i, 1, wzFileType, MAX_PATH )
          End If
       End If   
   Next
   
   If nFound <> -1 Then
      pDoc->ProjectFileType = nFileType
      Select Case nFileType
         Case FILETYPE_MAIN:     wzFileType = L(212,"Main")
         Case FILETYPE_MODULE:   wzFileType = L(211,"Module")
         Case FILETYPE_NORMAL:   wzFileType = ""  'L(210,"Normal")
         Case FILETYPE_RESOURCE: wzFileType = L(213,"Resource")
      End Select
      FF_PATHNAME( @WStr("NAMEX"), pDoc->DiskFilename, @wzFile )
      FF_PATHNAME( @WStr("PATH"),  pDoc->DiskFilename, @wzFilePath )
      FF_ListView_SetItemText( hLV, nFound, 0, @wzFile, MAX_PATH )
      FF_ListView_SetItemText( hLV, nFound, 1, wzFileType, MAX_PATH )
      FF_ListView_SetItemText( hLV, nFound, 2, @wzFilePath, MAX_PATH )
   End If    
   
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
   
   FF_PATHNAME( @WStr("NAMEX"), pDoc->DiskFilename, @wzFile )
   FF_PATHNAME( @WStr("PATH"),  pDoc->DiskFilename, @wzFilePath )
   FF_PATHNAME( @WStr("EXTN"),  pDoc->DiskFilename, @wzFileExt )
 
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
   
   Dim As HWnd hLV = GetDlgItem( HWND_FRMPROJECTMANAGER, IDC_FRMPROJECTMANAGER_LISTVIEW)
   Dim As Long nCount = ListView_GetItemCount(hLV)
   
   FF_ListView_InsertItem( hLV, nCount, 0, wzFile, Cast(LPARAM, pDoc) )
   FF_ListView_InsertItem( hLV, nCount, 1, wzFileType )
   FF_ListView_InsertItem( hLV, nCount, 2, wzFilePath )
   
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
Function clsApp.RemoveDocument( ByVal pDoc As clsDocument Ptr ) As Long
   Dim idx as Long 
   Dim i   as Long 
   Dim lb  As Long = LBound(m_arrDocuments)
   Dim ub  As Long = Ubound(m_arrDocuments) 
   
   For i = lb To ub
      If m_arrDocuments(i) = pDoc Then
         pDoc->Destructor    ' need to manually call the destructor
         idx = i
         Exit For
      End If   
   Next

   ' remove the document from the array
   For i = idx To ub - 1
      m_arrDocuments(i) = m_arrDocuments(i+1)
   Next

   ' resize the array to the new smaller size
   If (ub-1) >= 0 Then  ' to avoid gpf
      ReDim Preserve m_arrDocuments(ub-1) As clsDocument Ptr             
   End If
   
   Function = 0
End Function


''
''
Function clsApp.GetDocumentPtrByFilename( ByVal pswzName As WString Ptr ) As clsDocument Ptr
   If pswzName = 0 Then Return 0
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      If Ucase(*m_arrDocuments(i)->DiskFilename) = Ucase(*pswzName) Then
         Return m_arrDocuments(i)
      End If
   Next
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



