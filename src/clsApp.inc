'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
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

#include once "clsApp.bi"
#include once "frmExplorer.bi"


''
''
Function clsApp.GetDocumentCount() As Long
   dim pDoc as clsDocument ptr = this.pDocList
   dim nCount as Long = 0
   do until pDoc = 0
      nCount = nCount + 1
      pDoc = pDoc->pDocNext
   loop
   function = nCount
end function


''
''
Function clsApp.AddDocument( ByVal pDoc As clsDocument Ptr ) As Long
   ' Add it to the start of the linked list
   pDoc->pDocNext = this.pDocList
   this.pDocList = pDoc
   Function = 0
End Function


''
''
Function clsApp.RemoveDocument( ByVal pDoc As clsDocument Ptr ) As Long
   
   ' Remove from pDocList. Find the node that points to the incoming pDoc
   ' and then point that node to pDoc->pDocNext
   dim pDocSearch as clsDocument ptr = this.pDocList
   if pDocSearch = pDoc then
      this.pDocList = pDoc->pDocNext
   else
      do until pDocSearch = 0
         if pDocSearch->pDocNext = pDoc then
            pDocSearch->pDocNext = pDoc->pDocNext
            exit do
         end if
         pDocSearch = pDocSearch->pDocNext
      loop
   end if
   ' Release memory associated with this pDoc
   Delete pDoc

   Function = 0
End Function


''
''
Function clsApp.RemoveAllDocuments() As Long
   ' Remove from pDocList
   dim pDoc as clsDocument ptr = this.pDocList
   dim pDocNext as clsDocument ptr = this.pDocList

   do until pDoc = 0
      pDocNext = pDoc->pDocNext
      ' Release memory associated with this pDoc
      Delete pDoc
      pDoc = pDocNext
   loop
   this.pDocList = 0
   
   Function = 0
End Function


''
''
Function clsApp.GetDocumentPtrByFilename( ByVal pswzName As WString Ptr ) As clsDocument Ptr
   If pswzName = 0 Then Return 0
   dim pDoc as clsDocument ptr = this.pDocList
   do until pDoc = 0
      If Ucase(pDoc->DiskFilename) = Ucase(*pswzName) Then return pDoc
      pDoc = pDoc->pDocNext
   loop
   function = 0
End Function


''
''
Function clsApp.GetMainDocumentPtr() As clsDocument Ptr
   ' Get the Main document for the project
   dim pDoc as clsDocument ptr = this.pDocList
   do until pDoc = 0
      If pDoc->ProjectFileType = FILETYPE_MAIN Then return pDoc
      pDoc = pDoc->pDocNext
   loop
   function = 0
End Function


''
''
Function clsApp.GetResourceDocumentPtr() As clsDocument Ptr
   ' Get the Resource document for the active project
   dim pDoc as clsDocument ptr = this.pDocList
   do until pDoc = 0
      If pDoc->ProjectFileType = FILETYPE_RESOURCE Then return pDoc
      pDoc = pDoc->pDocNext
   loop
   function = 0
End Function


''
''
Function clsApp.GetHeaderDocumentPtr( byval pDocIn as clsDocument ptr ) As clsDocument Ptr
   ' Get the Header document related to the pDoc document. The Header
   ' is simply the same source file name but with a ".bi" extension.
   if pDocIn = 0 then exit function
   
   dim pDoc as clsDocument ptr
   
   dim as CWSTR wszFilename = _
                  AfxStrPathname( "PATH", pDocIn->Diskfilename ) & _
                  AfxStrPathname( "NAME", pDocIn->Diskfilename ) & ".bi"
   pDoc = this.GetDocumentPtrByFilename( wszFilename )
   
   if ( pDoc <> 0 ) andalso ( this.IsProjectActive = true ) then
      If pDoc->ProjectFileType <> FILETYPE_HEADER Then 
         pDoc = 0
      end if
   end if
   
   function = pDoc
End Function


''
''
Function clsApp.GetSourceDocumentPtr( byval pDocIn as clsDocument ptr ) As clsDocument Ptr
   ' Get the Source document related to the pDoc document. The Header
   ' is simply the same file name but with a ".bas" or ".inc" extension.
   if pDocIn = 0 then exit function
   
   dim as CWSTR wszFilename
   dim pDoc as clsDocument ptr
   
   wszFilename = AfxStrPathname( "PATH", pDocIn->Diskfilename ) & _
                 AfxStrPathname( "NAME", pDocIn->Diskfilename ) & ".bas"
   pDoc = this.GetDocumentPtrByFilename( wszFilename )
   
   if pDoc = 0 then
      wszFilename = AfxStrPathname( "PATH", pDocIn->Diskfilename ) & _
                    AfxStrPathname( "NAME", pDocIn->Diskfilename ) & ".inc"
      pDoc = this.GetDocumentPtrByFilename( wszFilename )
   end if
   
   if ( pDoc <> 0 ) andalso ( this.IsProjectActive = true ) then
      select case pDoc->ProjectFileType 
         case FILETYPE_MAIN, FILETYPE_MODULE, FILETYPE_NORMAL
         case else
            pDoc = 0
      end select
   end if
   
   function = pDoc
End Function


''
''
Function clsApp.GetProjectCompiler() As long
   ' Get the compiler associated with this project
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      if gConfig.Builds(i).Id = this.ProjectBuild then
         if gConfig.Builds(i).Is32Bit THEN return IDM_32BIT
         if gConfig.Builds(i).Is64Bit THEN return IDM_64BIT
      end if   
   NEXT
   function = 0
End Function


''
''
Function clsApp.SaveProject( ByVal bSaveAs As BOOLEAN = False ) As BOOLEAN

   Dim wFilename As WString * MAX_PATH
   Dim wText     As WString * MAX_PATH  
   
   wFilename = this.ProjectFilename
   
   If bSaveAs Then
      ' Display the Save File Dialog
      Dim pwszName As WString Ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wFilename, @WStr("wfbe"), IDM_PROJECTSAVE)
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
   UpdateMRUProjectList(wFilename)
   
   ' Ensure that the caption bar and Explorer root node is updated with the name
   Dim As Hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   frmMain_SetStatusbar    ' this will update the caption
   wFilename = ucase(this.ProjectName)
   TreeView_SetItemText( hTree, this.hExplorerRootNode, @wFilename)
   
   frmMain_PositionWindows

   Function = True
End Function
      

''
''
Function clsApp.ProjectSetFileType( ByVal pDoc As clsDocument Ptr, _
                                    ByVal nFileType As Long _       ' new filetype 
                                    ) As LRESULT
   If pDoc = 0 Then Exit Function
   
   Dim wzFileExt As WString * MAX_PATH 

   wzFileExt = AfxStrPathname( "EXTN",  pDoc->DiskFilename )
 
   ' Determine if the document already exists in the project. If it does not then
   ' make a determination of a default file type for this new file being added to the project.
   Dim bFound As BOOLEAN = False
   dim pDocSearch as clsDocument ptr 
   pDocSearch = this.pDocList
   do until pDocSearch = 0
      If (pDocSearch = pDoc) andalso (pDoc->ProjectFileType <> FILETYPE_UNDEFINED ) Then
         bFound = True: Exit do
      End If
      pDocSearch = pDocSearch->pDocNext
   loop
   If bFound = False Then
      If pDoc->ProjectFileType = FILETYPE_UNDEFINED Then
         Select Case Ucase(wzFileExt)
            Case ".BAS":  nFileType = FILETYPE_MODULE
            Case ".RC":   nFileType = FILETYPE_RESOURCE
            Case ".BI":   nFileType = FILETYPE_HEADER
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
      pDocSearch = this.pDocList
      do until pDocSearch = 0
         If pDocSearch->ProjectFileType = FILETYPE_MAIN Then 
            pDocSearch->ProjectFileType = FILETYPE_NORMAL
         END IF
         pDocSearch = pDocSearch->pDocNext
      loop
   End If      
   If nFileType = FILETYPE_RESOURCE Then
      pDocSearch = this.pDocList
      do until pDocSearch = 0
         If pDocSearch->ProjectFileType = FILETYPE_RESOURCE Then 
            pDocSearch->ProjectFileType = FILETYPE_NORMAL
         END IF
         pDocSearch = pDocSearch->pDocNext
      loop
   End If      

   pDoc->ProjectFileType = nFileType

   ' Refresh the statusbar to ensure that the file's type is displayed properly
   frmMain_SetStatusbar

   Function = 0
End Function


''
''
Function clsApp.RemoveAllSelectionAttributes() As long
   ' Remove all Attribute #8 selection highlights from the documents. This 
   ' occurs when the FindReplace dialog is closed. Therefore we need to apply
   ' it to all documents in all projects.
   dim pDoc as clsDocument ptr = this.pDocList
   dim as long startPos, endPos
   
   do until pDoc = 0
      SendMessage( pDoc->hWindow(0), SCI_INDICSETSTYLE, 8, INDIC_STRAIGHTBOX)
      SendMessage( pDoc->hWindow(0), SCI_SETINDICATORCURRENT, 8, 0)
      SendMessage( pDoc->hWindow(0), SCI_TARGETWHOLEDOCUMENT, 0, 0)
      startPos = SendMessage( pDoc->hWindow(0), SCI_GETTARGETSTART, 0, 0)
      endPos =  SendMessage( pDoc->hWindow(0), SCI_GETTARGETEND, 0, 0)
      SendMessage( pDoc->hWindow(0), SCI_INDICATORCLEARRANGE, startPos, Cast(LPARAM, endPos))
      pDoc = pDoc->pDocNext
   loop   
   function = 0
End Function

''
''
Function clsApp.GetDocumentPtrByWindow( byval hWindow as hwnd) As clsDocument ptr
   ' Find the pDoc pointer based on all the available hWindow for the any loaded
   ' document or visual designer.
   dim pDoc as clsDocument ptr = this.pDocList
   IF hWindow = NULL THEN EXIT FUNCTION
   
   do until pDoc = 0
      ' Determine if the incoming hWindow matches the clsDocument hWindow or
      ' is the Frame or Form windows.
      if (hWindow = pDoc->hWindow(0)) orelse _
         (hWindow = pDoc->hWindow(1)) orelse _
         (hWindow = pDoc->hWndFrame)  orelse _
         (hWindow = pDoc->hWndForm) THEN
         return pDoc
      end if   
      pDoc = pDoc->pDocNext
   loop   
   function = 0
End Function

''
''
function clsApp.AddQuickRunEXE( byref wszFilename as wstring ) as Long
   dim as long nFoundAt = -1
   if len(wszFilename) = 0 then exit function
   
   ' Scan array for an empty slot
   for i as long = lbound(m_arrQuickRun) to ubound(m_arrQuickRun)
      if len(m_arrQuickRun(i)) = 0 THEN
         m_arrQuickRun(i) = wszFilename
         exit function
      END IF
   NEXT
   if nFoundAt = -1 THEN
      redim preserve m_arrQuickRun(ubound(m_arrQuickRun) + 1)
      m_arrQuickRun(ubound(m_arrQuickRun)) = wszFilename
   END IF      

   function = 0
end function


''
''
function clsApp.CheckQuickRunEXE() as Long
   ' Scan all array items to see if can be deleted
   for i as long = lbound(m_arrQuickRun) to ubound(m_arrQuickRun)
      if len(m_arrQuickRun(i)) THEN
         if AfxFileExists(m_arrQuickRun(i)) THEN 
            DeleteFile(m_arrQuickRun(i))
         else
            m_arrQuickRun(i) = ""
         end if   
      END IF
   NEXT

   function = 0
end function
