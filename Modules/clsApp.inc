

''
''
Constructor clsApp()
   m_ActiveProject = -1
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
      this.FBKeywords  = " "
      Do Until EOF(f)
         Line Input #f, st
         If Left(st, 1) = "'" Then Continue Do
         this.FBKeywords = this.FBKeywords & st & " "
      Loop
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


'''
'''
'Function clsApp.AddProject() as Long
'   Dim ub as Long = UBound(m_arrProjects) + 1
' 
'   ReDim Preserve m_arrProjects(ub) as clsProject             
' 
'   m_ActiveProject = ub 
'
'   Function = 0
'End Function
'
'
'''
'''
'Property clsApp.ActiveProject( ByVal nValue as Long )
'   m_ActiveProject = nValue
'End Property
'
'Property clsApp.ActiveProject() as Long
'   Property = m_ActiveProject
'End Property
'
'
'''
'''
'Sub clsApp.SetActiveProject( ByVal nValue as Long)
'End Sub
'
'Function clsApp.GetActiveProject() ByRef as clsProject
'   Function = m_arrProjects(m_ActiveProject)
'End Function


''
''
Function clsApp.GetDocumentCount() As Long
   Function = Ubound(m_arrDocuments) 
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
Function clsApp.Debug() As Long
   Dim As Long lb = LBound(m_arrDocuments) 
   Dim As Long ub = Ubound(m_arrDocuments)
   
   Print "--- DEBUG ---"
   Print "lb = "; lb, "ub = "; ub
   For i As Long = LBound(m_arrDocuments) To Ubound(m_arrDocuments)
      Print "Array[";i;"]", "hWindow="; m_arrDocuments(i)->hWindow
   Next                          
   Print:Print
   
   Function = 0
End Function

