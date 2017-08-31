'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2017 Paul Squires, PlanetSquires Software
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
Constructor clsApp()
   ' Array index 0 is "Project0" that always contains the FILES that
   ' are not actually related to any open project. If no projects are open
   ' then we always refer to Project0 for the file list.
   redim preserve this.Projects(0) as clsProject 
   this.Projects(0).InUse = true
End Constructor

''
''
Destructor clsApp()
'
End Destructor 


''
''
Function clsApp.GetProjectIndexByFilename( byref sFilename as wstring ) As long
   for idx as long = 0 to ubound(this.Projects)
      if this.Projects(idx).InUse = true THEN
         if ucase(this.Projects(idx).ProjectFilename) = ucase(sFilename) then
            return idx
         end if
      end if
   next   
   function = 0
End Function


''
''
Function clsApp.RemoveAllSelectionAttributes() As long
   ' Remove all Attribute #8 selection highlights from the documents. This 
   ' occurs when the FindReplace dialog is closed. Therefore we need to apply
   ' it to all documents in all projects.
   dim pDoc as clsDocument ptr
   dim as long startPos, endPos
   
   for idx as long = 0 to ubound(this.Projects)
      if this.Projects(idx).InUse = true THEN
         For i As Long = 0 To this.Projects(idx).GetDocumentCount - 1
            pDoc = this.Projects(idx).GetDocumentPtr(i)
            if pDoc THEN
               SendMessage( pDoc->hWindow, SCI_INDICSETSTYLE, 8, INDIC_STRAIGHTBOX)
               SendMessage( pDoc->hWindow, SCI_SETINDICATORCURRENT, 8, 0)
               SendMessage( pDoc->hWindow, SCI_TARGETWHOLEDOCUMENT, 0, 0)
               startPos = SendMessage( pDoc->hWindow, SCI_GETTARGETSTART, 0, 0)
               endPos =  SendMessage( pDoc->hWindow, SCI_GETTARGETEND, 0, 0)
               SendMessage( pDoc->hWindow, SCI_INDICATORCLEARRANGE, startPos, Cast(LPARAM, endPos))
            END IF
         Next
      end if
   next   
   function = 0
End Function

''
''
Function clsApp.GetDocumentPtrByWindow( byval hWindow as hwnd) As clsDocument ptr
   ' Find the pDoc pointer based on all the available hWindow for the any loaded
   ' document or visual designer.
   dim pDoc as clsDocument ptr
   IF hWindow = NULL THEN EXIT FUNCTION
   
   for idx as long = 0 to ubound(this.Projects)
      if this.Projects(idx).InUse = true THEN
         For i As Long = 0 To this.Projects(idx).GetDocumentCount - 1
            pDoc = this.Projects(idx).GetDocumentPtr(i)
            if pDoc THEN
               ' Determine if the incoming hWindow matches the clsDocument hWindow or
               ' is the Frame or Form windows.
               if (hWindow = pDoc->hWindow) orelse _
                  (hWindow = pDoc->hWndFrame ) orelse _
                  (hWindow = pDoc->hWndForm) THEN
                  return pDoc
               end if   
            END IF
         Next
      end if
   next   
   function = 0
End Function

''
''
Function clsApp.GetProjectCount() As long
   dim as long nCount = 0
   for i as long = 1 to ubound(this.Projects)
      if this.Projects(i).InUse = true THEN
         nCount = nCount + 1
      END IF
   NEXT
   function = nCount
end function

''
''
Function clsApp.GetActiveProjectIndex() As long
   ' First check to see if the index override has been set. This occurs if we want
   ' to do an action to a specific project rather than the active project.
   if gApp.ProjectOverrideIndex > 0 THEN return gApp.ProjectOverrideIndex
      
   ' Return the index of the project that has it's checkbox set
   dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   for i as long = 1 to ubound(this.Projects)
      if this.Projects(i).InUse = true THEN
         if TreeView_IsItemChecked(hTree, this.Projects(i).hExplorerRootNode) then
            return i
         end if   
      END IF
   NEXT
   function = 0    ' Use Project0 "FILES" by default
end function

''
''
Function clsApp.SetActiveProject( byval hNode As HTREEITEM) As long
   ' Set the treeview checkbox making it the active project
   dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   for i as long = 1 to ubound(this.Projects)
      if this.Projects(i).InUse = true THEN
         if hNode = gApp.Projects(i).hExplorerRootNode THEN
            Treeview_SetCheckState(hTree, gApp.Projects(i).hExplorerRootNode, TRUE)      
         else   
            Treeview_SetCheckState(hTree, gApp.Projects(i).hExplorerRootNode, false)      
         END IF
      end if
   NEXT
   function = 0
end function

''
''
Function clsApp.EnsureDefaultActiveProject(byval hNodeBypass as HTREEITEM = 0) As long
   ' Ensure that at least one of the treeview checkboxes are the active project
   ' If incoming hNodePass is set then exclude that node from our searches. This is
   ' the case where the node has just been clicked off so we don't want to set it 
   ' active immediately again.
   dim as hwnd hTree = GetDlgItem(HWND_FRMEXPLORER, IDC_FRMEXPLORER_TREE)
   dim as HTREEITEM hNodeActive = 0
   
   for i as long = 1 to ubound(this.Projects)
      if this.Projects(i).InUse = true THEN
         if gApp.Projects(i).hExplorerRootNode = hNodeBypass THEN continue for
         if TreeView_IsItemChecked(hTree, gApp.Projects(i).hExplorerRootNode) then
            hNodeActive = gApp.Projects(i).hExplorerRootNode
            exit for
         else
            if hNodeActive = 0 then hNodeActive = gApp.Projects(i).hExplorerRootNode
         end if
      end if
   next
   Treeview_SetCheckState(hTree, hNodeActive, TRUE)      
   function = 0
end function


''
''
Function clsApp.IsProjectActive() As BOOLEAN
   ' Determine is any projects are active. We cycle through the Projects to
   ' see if any of the slots are in use. Slot 0 is reserved for when no project
   ' is open so bypass that slot when searching.
   for i as long = 1 to ubound(this.Projects)
      if this.Projects(i).InUse = true THEN
         return true
      END IF
   NEXT
   return false
End Function


''
''
Function clsApp.GetNewProjectIndex() As Long
   ' Get a new project index 
   dim idx as long = -1

   for i as long = 1 to ubound(this.Projects)
      if this.Projects(i).InUse = false THEN
         this.Projects(i).InUse = true
         return i
      END IF
   NEXT
   idx = ubound(this.Projects) + 1
   redim preserve this.Projects(idx) as clsProject 

   return idx

end function


''
''
function clsApp.AddQuickRunEXE( byref wszFilename as wstring ) as Long
   dim as long nFoundAt = -1
   
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
