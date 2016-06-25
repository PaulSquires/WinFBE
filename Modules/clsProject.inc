
Type clsProject
   Private:
      m_arrDocuments(Any) as DOCUMENT_TYPE
   
   Public:
      Declare Function AddDocument( ByVal hWndParent as HWnd ) ByRef as DOCUMENT_TYPE
      
End Type


''
''
Function clsProject.AddDocument( ByVal hWndParent as HWnd ) ByRef as DOCUMENT_TYPE
   Dim ub as Integer = UBound(m_arrDocuments) + 1
 
   ReDim Preserve m_arrDocuments(ub) as DOCUMENT_TYPE             
 
   Function = m_arrDocuments(ub)
End Function

