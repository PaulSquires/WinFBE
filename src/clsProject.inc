'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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

''
''
Function clsProject.SaveProject( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN

   Dim wFilename As WString * MAX_PATH
   Dim wText     As WString * MAX_PATH  
   
   ' If this a new unsaved project set flag to display SaveAs dialog.
   If gApp.IsNewProjectFlag Then bSaveAs = True
      
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
   
   ' Also need to add this new project name to the MRU list.
   UpdateMRUProjectList(@wFilename)
   
   ' Ensure that the caption bar and Explorer root node is updated with the name
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   frmMain_UpdateLineCol(HWND_FRMMAIN)    ' this will update the caption
   wFilename = ucase(this.ProjectName)
   TreeView_SetItemText( hTree, this.hExplorerRootNode, @wFilename)
   
   frmMain_PositionWindows(HWND_FRMMAIN)

   Function = True
End Function
      


''
''
Function clsProject.ProjectSetFileType( ByVal pDoc As clsDocument Ptr, _
                                        ByVal nFileType As Long _       ' new filetype 
                                        ) As LRESULT
   If pDoc = 0 Then Exit Function
   
   Dim wzFileExt As WString * MAX_PATH 

   wzFileExt  = AfxStrPathname( "EXTN",  pDoc->DiskFilename )
 
   ' Determine if the document already exists in the project. If it does not then
   ' make a determination of a default file type for this new file being added to the project.
   Dim bFound As BOOLEAN = False
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      If (m_arrDocuments(i) = pDoc) andalso (pDoc->ProjectFileType <> FILETYPE_UNDEFINED ) Then
         bFound = True: Exit For
      End If
   Next
   If bFound = False Then
      If pDoc->ProjectFileType = FILETYPE_UNDEFINED Then
         Select Case Ucase(wzFileExt)
            Case ".BAS":  nFileType = FILETYPE_MODULE
            Case ".RC":   nFileType = FILETYPE_RESOURCE
            Case ".BI":   nFileType = FILETYPE_NORMAL
            Case ".INC":  nFileType = FILETYPE_NORMAL
            Case Else:    nFileType = FILETYPE_NORMAL
         End Select   
      End If
   End If
   
   ' Do check to make sure that no other MAIN or RESOURCE exist because
   ' there can only be one unique MAIN and one unique RESOURCE per project.

   ' If we are setting a "Main" project file then we need to toggle any other Main 
   ' file to "Normal". There can only be one Main file. Likewise, there can only be 
   ' one "Resource" file.

   If nFileType = FILETYPE_MAIN Then
      For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
         If m_arrDocuments(i)->ProjectFileType = FILETYPE_MAIN Then 
            m_arrDocuments(i)->ProjectFileType = FILETYPE_NORMAL
            ChangeImage4ExplorerTreeview2Doc(m_arrDocuments(i))
         END IF
      Next   
   End If      
   If nFileType = FILETYPE_RESOURCE Then
      For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
         If m_arrDocuments(i)->ProjectFileType = FILETYPE_RESOURCE Then 
            m_arrDocuments(i)->ProjectFileType = FILETYPE_NORMAL
            ChangeImage4ExplorerTreeview2Doc(m_arrDocuments(i))
         end if
      Next   
   End If      

   pDoc->ProjectFileType = nFileType
   ChangeImage4ExplorerTreeview2Doc(pDoc)

   ' Refresh the statusbar to ensure that the file's type is displayed properly
   pDoc->DisplayStats

   Function = 0
End Function

''
''
Function clsProject.GetDocumentCount() As Long
   Function = Ubound(m_arrDocuments) + 1 
End Function

''
''
Function clsProject.GetDocumentPtr( ByVal idx As Long ) As clsDocument Ptr
   Dim lb As Long = LBound(m_arrDocuments) 
   Dim ub As Long = Ubound(m_arrDocuments) 
   If (idx >= lb) AndAlso (idx <= ub) Then
      Function = m_arrDocuments(idx) 
   End If
End Function

''
''
Function clsProject.AddDocument( ByVal pDoc As clsDocument Ptr ) As Long
   Dim ub as Long = UBound(m_arrDocuments) + 1
 
   ReDim Preserve m_arrDocuments(ub) As clsDocument Ptr             
   m_arrDocuments(ub) = pDoc

   Function = 0
End Function

''
''
Function clsProject.RemoveDocumentFromArray( ByVal idx As Long) As Long
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
Function clsProject.RemoveDocument( ByVal pDoc As clsDocument Ptr ) As Long
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
   
   ' Remove all references from the gdb2 database
   gdb2.dbDelete4File(idx,pDoc->DiskFilename)

   ' Update the TODO listview in the frmOutput form
   UpdateToDoListview
   
   ' Remove the document from the Explorer treeview (deletes the node)
   Treeview_DeleteItem( GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE), pDoc->hNodeExplorer)

   ' Now remove it from the array
   this.RemoveDocumentFromArray(idx)

   Function = 0
End Function

''
''
Function clsProject.RemoveAllDocuments() As Long
   Dim lb  As Long = LBound(m_arrDocuments)
   Dim ub  As Long = Ubound(m_arrDocuments) 
   
   do until ub < lb
      this.RemoveDocumentFromArray(ub)
      lb = Lbound(m_arrDocuments) 
      ub = Ubound(m_arrDocuments) 
   loop
   
   Function = 0
End Function

''
''
Function clsProject.GetDocumentPtrByFilename( ByVal pswzName As WString Ptr ) As clsDocument Ptr
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
Function clsProject.GetMainDocumentPtr() As clsDocument Ptr
   ' Get the Main document for the active project
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      If m_arrDocuments(i)->ProjectFileType = FILETYPE_MAIN Then
         Return m_arrDocuments(i)
      End If
   Next
   function = 0
End Function

''
''
Function clsProject.GetResourceDocumentPtr() As clsDocument Ptr
   ' Get the Resource document for the active project
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      If m_arrDocuments(i)->ProjectFileType = FILETYPE_RESOURCE Then
         Return m_arrDocuments(i)
      End If
   Next
   function = 0
End Function

''
''
Function clsProject.GetProjectCompiler() As long
   ' Get the compiler associated with this project
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      with gConfig.Builds(i)
         if .Id = this.ProjectBuild then
            if .Is32Bit THEN return IDM_32BIT
            if .Is64Bit THEN return IDM_64BIT
         end if
      end with
   next
   function = 0
End Function


''
''
Function clsProject.Debug() As Long
   Dim As Long lb = LBound(m_arrDocuments) 
   Dim As Long ub = Ubound(m_arrDocuments)
   
   Print "--- DEBUG ---"
   print "GetDocumentCount = "; this.GetDocumentCount
   Print "lb = "; lb, "ub = "; ub
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      Print "Array[";i;"]", "hWindow="; m_arrDocuments(i)->hWindow(0), m_arrDocuments(i)->DiskFilename
   Next                          
   Print:Print
   
   Function = 0
End Function

Function clsProject.SetActiveProject() As BOOLEAN
   Dim pDocMain As clsDocument Ptr
   with this
      .ProjectPath = AfxStrPathname("PATH", this.ProjectFilename)
      pDocMain = GetMainDocumentPtr()
      if pDocMain then
         .MainFolderPath = AfxStrPathname("PATH", pDocMain->DiskFilename)
      else
         .MainFolderPath = .ProjectPath  
      endif
      if .GetProjectCompiler()=IDM_32BIT then
         .UsedFbcPathName = gConfig.FBWINCompiler32
      else
         .UsedFbcPathName = gConfig.FBWINCompiler64
      end if
      .UsedCompilerPath = AfxStrPathName ("PATH",.UsedFbcPathName)
      .UsedIncludePath = .UsedCompilerPath & "Inc\"
      .UsedLibraryPath = .UsedCompilerPath & "Lib\"      
   end with
   function = true
end function

Sub clsProject.Format2FullPath(byref vFormatPath As CWSTR)
   '   Path acronymes: 
   '   - {A} Where WinFBE.exe is located(MainApp).
   '   - {C} Compiler path.
   '   - {H} Help path.
   '   - {P} Current project path.
   
   dim As WString * MAX_PATH AppPath = AfxStrReplace(vFormatPath,"/","\")
   dim As long index = IIF(INSTR(AppPath,"\") = 4,5,4),index2=any

   with this
      AppPath=LTrim(AppPath,ANY " /\")
      if len(AfxStrPathname("PATH", AppPath)) = 0 THEN
         AppPath= .MainFolderPath & AppPath
      ELSEif Mid(AppPath,2,2)<>":\" then   
         select case ucase(left(AppPath,3))
         case "..\"
             AppPath= .MainFolderPath & AppPath
         case "{A}" 'MainApp path
            AppPath= gConfig.MainAppPath  & mid(AppPath ,index)
         case "{C}"  'Help path
            AppPath= gConfig.MainHelpPath  & mid(AppPath ,index)
         case "{H}"  'Compiler path
            AppPath= .UsedCompilerPath  & mid(AppPath ,index)
         case "{P}"  'Current project path
            AppPath= .ProjectPath  & mid(AppPath ,index) 
         case else  
            If WStrNIsEqual(AppPath, ".\", 2) = True Then
               AppPath= .MainFolderPath & mid(AppPath ,3)
            end If
         end select
      end if
      do 
         index = INSTR(AppPath,"\..\")
         if index = 0 then exit do
         index2= INSTRrev(AppPath,"\",index-1)
         AppPath=left(AppPath, index2) & mid(AppPath, index+4)
      loop   
   end with
   vFormatPath=AppPath
End Sub