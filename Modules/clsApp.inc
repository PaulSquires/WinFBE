

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
   
   If Open( Exepath & "\freebasic_keywords.txt" For Input as #f) = 0 Then
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
   
   If Open( Exepath & "\freebasic_keywords.txt" For Output As #f) = 0 Then
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


''
''
Property clsApp.IsProjectActive( ByVal nValue As BOOLEAN)
   m_IsProjectActive = nValue
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

   Dim sText     As String
   Dim wFilename As WString * MAX_PATH
   Dim wText     As WString * MAX_PATH  
   Dim f         As Long 
   
   ' If this a new unsaved project set flag to display SaveAs dialog.
   If this.IsNewProjectFlag Then bSaveAs = True
      
   wFilename = *this.ProjectDiskfilename
   
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

   ' Delete any existing file
   If Dir(wFilename) > "" Then Kill wFilename
   
   ' Do the actual saving to disk
   f = Freefile
   If Open( wFilename For Output As #f) = 0 Then
      Print #f, "*** PROJECT FILE ***"
      Close #f
   Else
      ' error saving file
   End If   

   this.ProjectDiskFilename = @wFilename
   this.IsNewProjectFlag    = False
   
   Function = True
End Function


''
''
Function clsApp.ProjectAddFile( ByVal pDoc As clsDocument Ptr ) As LRESULT
   If pDoc = 0 Then Exit Function
   
   Dim wzFileType As WString * MAX_PATH = "NORMAL"
   Dim wzFilePath As WString * MAX_PATH 
   Dim wzFile     As WString * MAX_PATH
   
   ' Add the incoming document to the project. We simply need to set the flag for
   ' instance of the document and update the Project Manager listview.
   pDoc->IsProjectFile = True
   
   FF_PATHNAME( @WStr("NAMEX"), pDoc->DiskFilename, @wzFile )
   FF_PATHNAME( @WStr("PATH"),  pDoc->DiskFilename, @wzFilePath )
 
   Dim As HWnd hLV = GetDlgItem( HWND_FRMPROJECTMANAGER, IDC_FRMPROJECTMANAGER_LISTVIEW)
   
   FF_ListView_InsertItem( hLV, 0, 0, wzFile, Cast(Long, pDoc) )
   FF_ListView_InsertItem( hLV, 0, 1, wzFileType )
   FF_ListView_InsertItem( hLV, 0, 2, wzFilePath )
   
   
   Function = 0
End Function

''
''
Function clsApp.ProjectRemoveFile( ByVal pDoc As clsDocument Ptr ) As LRESULT

   ' Remove the incoming document from the project. We simply need to unset the flag for
   ' instance of the document and update the Project Manager listview.
   pDoc->IsProjectFile = True

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
         DestroyWindow(pDoc->hWindow)     ' destroy the Scintilla window
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


'''
'''
'Function clsApp.GetDocument( ByVal hWindow as HWnd ) as DOCUMENT_TYPE Ptr
'   For i as Long = LBound(m_arrDocuments) To UBound(m_arrDocuments)
'      If m_arrDocuments(i).hWindow = hWindow Then
'         Function = VarPtr(m_arrDocuments(i))
'         Exit Function
'      End If   
'   Next
'End Function


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
