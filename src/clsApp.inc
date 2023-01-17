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

#include once "clsApp.bi"
#include once "frmExplorer.bi"
#include once "modMRU.bi"

''
''
function clsApp.GetDocumentCount() as long
   dim pDoc as clsDocument ptr = this.pDocList
   dim nCount as long = 0
   do until pDoc = 0
      nCount = nCount + 1
      pDoc = pDoc->pDocNext
   loop
   function = nCount
end function


''
''
function clsApp.AddNewDocument() as clsDocument ptr 
   ' Add it to the start of the linked list
   dim pDoc as clsDocument ptr = new clsDocument
   pDoc->pDocNext = this.pDocList
   this.pDocList = pDoc
   function = pDoc
end function


''
''
function clsApp.RemoveDocument( byval pDoc as clsDocument ptr ) as long
   if pDoc = 0 then exit function
      
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
   if pDoc then delete(pDoc)

   function = 0
end function


''
''
function clsApp.RemoveAllDocuments() as long
   ' Remove from pDocList
   dim pDoc as clsDocument ptr = this.pDocList
   dim pDocNext as clsDocument ptr = this.pDocList

   do until pDoc = 0
      ' Remove any parsed data for this document from the in-memory database
      gdb2.dbDeleteByDocumentPtr(pDoc)
      pDocNext = pDoc->pDocNext
      ' Release memory associated with this pDoc
      delete pDoc
      pDoc = pDocNext
   loop
   this.pDocList = 0
   
   function = 0
end function


''
''
function clsApp.GetDocumentPtrByFilename( Byref wszName as wstring ) as clsDocument ptr
   if len(wszName) = 0 then return 0
   dim pDoc as clsDocument ptr = this.pDocList
   do until pDoc = 0
      if ucase(pDoc->DiskFilename) = ucase(wszName) then return pDoc
      pDoc = pDoc->pDocNext
   loop
   function = 0
end function


''
''
function clsApp.GetMainDocumentPtr() as clsDocument ptr
   ' Get the Main document for the project
   dim pDoc as clsDocument ptr = this.pDocList
   do until pDoc = 0
      if pDoc->ProjectFileType = FILETYPE_MAIN then return pDoc
      pDoc = pDoc->pDocNext
   loop
   function = 0
end function


''
''
function clsApp.GetResourceDocumentPtr() as clsDocument ptr
   ' Get the Resource document for the active project
   dim pDoc as clsDocument ptr = this.pDocList
   do until pDoc = 0
      if pDoc->ProjectFileType = FILETYPE_RESOURCE then return pDoc
      pDoc = pDoc->pDocNext
   loop
   function = 0
end function


''
''
function clsApp.GetHeaderDocumentPtr( byval pDocIn as clsDocument ptr ) as clsDocument ptr
   ' Get the Header document related to the pDoc document. The Header
   ' is simply the same source file name but with a ".bi" extension.
   if pDocIn = 0 then exit function
   
   dim pDoc as clsDocument ptr
   
   dim as CWSTR wszFilename = _
                  AfxStrPathname( "PATH", pDocIn->Diskfilename ) & _
                  AfxStrPathname( "NAME", pDocIn->Diskfilename ) & ".bi"
   pDoc = this.GetDocumentPtrByFilename( wszFilename )
   
   if ( pDoc <> 0 ) andalso ( this.IsProjectActive = true ) then
      if pDoc->ProjectFileType <> FILETYPE_HEADER then 
         pDoc = 0
      end if
   end if
   
   function = pDoc
end function


''
''
function clsApp.GetSourceDocumentPtr( byval pDocIn as clsDocument ptr ) as clsDocument ptr
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
end function


''
''
function clsApp.GetProjectCompiler() as long
   ' Get the compiler associated with this project
   for i as long = lbound(gConfig.Builds) to ubound(gConfig.Builds)
      if gConfig.Builds(i).Id = this.ProjectBuild then
         if gConfig.Builds(i).Is32Bit then return IDM_32BIT
         if gConfig.Builds(i).Is64Bit then return IDM_64BIT
      end if   
   NEXT
   function = 0
end function


''
''
function clsApp.SaveProject( byval bSaveAs as boolean = false ) as boolean

   dim wFilename as wstring * MAX_PATH
   dim wText     as wstring * MAX_PATH  
   
   wFilename = this.ProjectFilename
   
   if bSaveAs then
      ' Display the Save File Dialog
      dim pwszName as wstring ptr = AfxIFileSaveDialog(HWND_FRMMAIN, @wFilename, @wstr("wfbe"), IDM_PROJECTSAVE)
      if pwszName then
         wFilename = *pwszName
         CoTaskMemFree(pwszName)
      else
         return false
      end if
   end if

   ' delete any existing file
   if AfxFileExists(wFilename) then AfxDeleteFile(wFilename)
   this.ProjectFilename = wFilename
   this.ProjectName = AfxStrPathname( "NAMEX", wFilename )
   gConfig.ProjectSaveToFile()
   
   ' Also need to add this new project name to the MRU list.
   UpdateMRUProjectList(wFilename)
   
   frmMain_PositionWindows

   function = true
end function
      

''
''
function clsApp.ProjectSetFileType( _
            byval pDoc as clsDocument ptr, _
            byval wszFileType as CWSTR _       ' new filetype 
            ) as LRESULT

   if pDoc = 0 then exit function
   
   dim wzFileExt as wstring * MAX_PATH 

   wzFileExt = AfxStrPathname( "EXTN",  pDoc->DiskFilename )
 
   ' Determine if the document already exists in the project. If it does not then
   ' make a determination of a default file type for this new file being added to the project.
   dim bFound as boolean = false
   dim pDocSearch as clsDocument ptr 
   pDocSearch = this.pDocList
   do until pDocSearch = 0
      if (pDocSearch = pDoc) andalso (pDoc->ProjectFileType <> FILETYPE_UNDEFINED ) then
         bFound = true: exit do
      end if
      pDocSearch = pDocSearch->pDocNext
   loop
   if bFound = false then
      if pDoc->ProjectFileType = FILETYPE_UNDEFINED then
         select case ucase(wzFileExt)
            case ".BAS":  wszFileType = FILETYPE_NORMAL
            case ".RC":   wszFileType = FILETYPE_RESOURCE
            case ".BI":   wszFileType = FILETYPE_HEADER
            case ".INC":  wszFileType = FILETYPE_NORMAL
            case else:    wszFileType = FILETYPE_NORMAL
         end select   
      end if
   end if
   
   ' Do check to make sure that no other MAIN or RESOURCE exist because
   ' there can only be one unique MAIN and one unique RESOURCE per project.

   ' if we are setting a "Main" project file then we need to toggle any other Main 
   ' file to "Normal". There can only be one Main file. Likewise, there can only be 
   ' one "Resource" file.

   if wszFileType = FILETYPE_MAIN then
      pDocSearch = this.pDocList
      do until pDocSearch = 0
         if pDocSearch->ProjectFileType = FILETYPE_MAIN then 
            pDocSearch->ProjectFileType = FILETYPE_NORMAL
         end if
         pDocSearch = pDocSearch->pDocNext
      loop
   end if      
   if wszFileType = FILETYPE_RESOURCE then
      pDocSearch = this.pDocList
      do until pDocSearch = 0
         if pDocSearch->ProjectFileType = FILETYPE_RESOURCE then 
            pDocSearch->ProjectFileType = FILETYPE_NORMAL
         end if
         pDocSearch = pDocSearch->pDocNext
      loop
   end if      

   pDoc->ProjectFileType = wszFileType

   ' Refresh the statusbar to ensure that the file's type is displayed properly
   frmMain_SetStatusbar

   function = 0
end function


''
''
function clsApp.RemoveAllSelectionAttributes() as long
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
      SendMessage( pDoc->hWindow(0), SCI_INDICATORCLEARRANGE, startPos, cast(LPARAM, endPos))
      pDoc = pDoc->pDocNext
   loop   
   function = 0
end function

''
''
function clsApp.GetDocumentPtrByWindow( byval hWindow as hwnd) as clsDocument ptr
   ' Find the pDoc pointer based on all the available hWindow for the any loaded
   ' document or visual designer.
   dim pDoc as clsDocument ptr = this.pDocList
   if hWindow = null then exit function
   
   do until pDoc = 0
      ' Determine if the incoming hWindow matches the clsDocument hWindow or
      ' is the Frame or Form windows.
      if (hWindow = pDoc->hWindow(0)) orelse _
         (hWindow = pDoc->hWindow(1)) orelse _
         (hWindow = pDoc->hWndFrame)  orelse _
         (hWindow = pDoc->hWndForm) then
         return pDoc
      end if   
      pDoc = pDoc->pDocNext
   loop   
   function = 0
end function

''
''
function clsApp.AddQuickRunEXE( byref wszFilename as wstring ) as long
   dim as long nFoundAt = -1
   if len(wszFilename) = 0 then exit function
   
   ' Scan array for an empty slot
   for i as long = lbound(m_arrQuickRun) to ubound(m_arrQuickRun)
      if len(m_arrQuickRun(i)) = 0 then
         m_arrQuickRun(i) = wszFilename
         exit function
      end if
   NEXT
   if nFoundAt = -1 then
      redim preserve m_arrQuickRun(ubound(m_arrQuickRun) + 1)
      m_arrQuickRun(ubound(m_arrQuickRun)) = wszFilename
   end if      

   function = 0
end function


''
''
function clsApp.CheckQuickRunEXE() as long
   ' Scan all array items to see if can be deleted
   for i as long = lbound(m_arrQuickRun) to ubound(m_arrQuickRun)
      if len(m_arrQuickRun(i)) then
         if AfxFileExists(m_arrQuickRun(i)) then 
            AfxDeleteFile(m_arrQuickRun(i))
         else
            m_arrQuickRun(i) = ""
         end if   
      end if
   NEXT

   function = 0
end function




