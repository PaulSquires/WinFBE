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

Type _HashItem
   sKey    As String
   sData1  As String
   sData2  As String
   sData3  As String
   sData4  As String   ' Will this one be needed...
   pNext   As _HashItem Ptr
End Type
' Structure for gCodeTips (FB's builtin keywords)
' sKey   = FB keyword
' sData1 = Syntax
' sData2 = <N/A>
' sData3 = <N/A>

' Structure for gFunctionDeclares
' sKey   = Function Name
' sData1 = Function Declare
' sData2 = Filename
' sData3 = Str(LineNumber)

' Structure for gFunctionLists
' sKey   = Full Disk Filename
' sData1 = DELIM separated list of function names
' sData2 = <N/A>
' sData3 = <N/A>

' Structure for gTYPES
' sKey   = TYPE name
' sData1 = DELIM separated list of Type elements
' sData2 = Filename
' sData3 = <N/A>

Type clsFastHash
   Private:
      m_Capacity  As LONG
      m_FoundNode As _HashItem Ptr
      
      Dim m_arrHash(Any) As _HashItem Ptr 
   
      Declare Function GetSlot( ByRef sKey As Const String) As LONG

   Public:
      declare Function AddItem( ByRef sKey As Const String, _
                                 ByRef sData1   As Const String, _
                                 ByRef sData2   As Const String = "", _
                                 ByRef sData3   As Const String = "", _
                                 ByRef sData4   As Const String = "" _
                                 ) As LONG   
      Declare Function FindItem( ByRef sKey As Const String) As boolean   
      Declare function ClearHash() as BOOLEAN 
      declare function DeleteItem(ByRef sKey As Const String) As boolean
      declare function DeleteAllData2(ByRef sData2 As Const String) as BOOLEAN
      declare function BuildKeyList() as string
      Declare Property FoundKey() As String
      Declare Property FoundData1() As String
      Declare Property FoundData2() As String
      Declare Property FoundData3() As String
      Declare Property FoundData4() As String
      Declare Function SaveToFile( ByRef sFilename As Const String, _
                                   ByRef sHeaderText As Const String = "", _
                                   ByRef sDelimiter As Const String = "=" _
                                   ) As LONG
      Declare Function LoadFromFile( ByRef sFilename As Const String, _
                                     ByRef sDelimiter As Const String = "=" _
                                     ) As LONG
      declare function CreateDeclaresFile( ByRef sFilename As Const String ) As LONG
      declare Sub Debug( ByRef sFilename As Const String)
      Declare Constructor ( ByVal nCapacity As LONG = -1)
      Declare Destructor
End Type

''
''
Constructor clsFastHash( ByVal nCapacity As LONG = -1)
   m_Capacity = Iif( nCapacity <= 0, 500, nCapacity )
   ReDim m_arrHash(m_Capacity) As _HashItem Ptr
End Constructor

''
''
Destructor clsFastHash()
   this.ClearHash
End Destructor

''
''
function clsFastHash.ClearHash() as BOOLEAN
   Dim _Item     As _HashItem Ptr
   Dim _ItemNext As _HashItem Ptr
   
   For i As LONG = LBound(m_arrHash) To Ubound(m_arrHash)
      If m_arrHash(i) <> 0 Then
         _Item = m_arrHash(i)
         Do Until _Item = 0
            _ItemNext = _Item->pNext
            Delete _Item
            _Item = _ItemNext
         Loop
      End If
   Next    
   Erase m_arrHash
   
   function = true
End function

''
''
Property clsFastHash.FoundKey() As String
   If m_FoundNode Then Property = m_FoundNode->sKey
End Property

''
''
Property clsFastHash.FoundData1() As String
   If m_FoundNode Then Property = m_FoundNode->sData1
End Property

''
''
Property clsFastHash.FoundData2() As String
   If m_FoundNode Then Property = m_FoundNode->sData2
End Property

''
''
Property clsFastHash.FoundData3() As String
   If m_FoundNode Then Property = m_FoundNode->sData3
End Property

''
''
Property clsFastHash.FoundData4() As String
   If m_FoundNode Then Property = m_FoundNode->sData4
End Property

''
''
Function clsFastHash.GetSlot( ByRef sKey As Const String ) As LONG
    
    Dim nLen  As LONG  = Len(sKey)
    Dim accul As ULONG = 0
    Dim last  As ULONG = 0
    Dim i     As LONG  = 0
    
    For i = 0 To nLen - 1
        last = accul
        accul = (accul Shl 5 )
        accul = accul - last + sKey[i]
    Next
    
    Function = (accul Mod m_Capacity)
End Function

''
''
Function clsFastHash.AddItem( ByRef sKey     As Const String, _
                              ByRef sData1   As Const String, _
                              ByRef sData2   As Const String = "", _
                              ByRef sData3   As Const String = "", _
                              ByRef sData4   As Const String = "" _
                              ) As LONG   
   
   If Ubound(m_arrHash) = -1 Then
      ReDim m_arrHash(m_Capacity) As _HashItem Ptr
   End If
   
   Dim nPosition As LONG = this.GetSlot(sKey)
   Dim _Item      As _HashItem Ptr
   Dim _ItemPrev  As _HashItem Ptr

   ' Determine if a linked list exists at this position.
   If m_arrHash(nPosition) = 0 Then
      m_arrHash(nPosition) = New _HashItem
      _Item = m_arrHash(nPosition)
   Else
      ' Determine if the key already exists. If yes, then simply update it's data.
      _Item = m_arrHash(nPosition)
      Do 
         If _Item->sKey = sKey Then
            Exit Do   ' update existing key/data
         End If   
         _ItemPrev = _Item
         _Item = _Item->pNext
      Loop Until _Item = 0
   End If

   ' Not found so create a new node and join it to the list
   If _Item = 0 Then
      _ItemPrev->pNext = New _HashItem
      _Item = _ItemPrev->pNext
   End If

   _Item->sKey   = sKey
   _Item->sData1 = sData1
   _Item->sData2 = sData2
   _Item->sData3 = sData3
   _Item->sData4 = sData4
   Function = 0
End Function

''
''
Function clsFastHash.FindItem( ByRef sKey As Const String) As boolean
   
   Dim nPosition As LONG = this.GetSlot(sKey)
   Dim _Item As _HashItem Ptr
   
   m_FoundNode = 0

   ' Determine if a linked list exists at this position.
   If m_arrHash(nPosition) = 0 Then Return false
   
   ' Add the key/data to the linked list of items
   _Item = m_arrHash(nPosition)
   Do 
      If _Item->sKey = sKey Then 
         m_FoundNode = _Item
         Return true
      End If
      If _Item->pNext = 0 Then Exit Do   
      _Item = _Item->pNext
   Loop
   
   ' not found
   Return false

End Function

''
''
function clsFastHash.DeleteItem(ByRef sKey As Const String) As boolean

   Dim nPosition As LONG = this.GetSlot(sKey)
   Dim _ItemPrev As _HashItem Ptr
   Dim _Item As _HashItem Ptr
   
   m_FoundNode = 0

   ' Determine if a linked list exists at this position.
   If m_arrHash(nPosition) = 0 Then Return false
   
   ' Key exists in this slot. Find it in the linked list 
   _Item = m_arrHash(nPosition)
   _ItemPrev = 0
   Do 
      If _Item->sKey = sKey Then 
         ' Key was located. Set the previous item to point to the
         ' item after this one and then deallocate this current node.
         if _ItemPrev = 0 THEN
            m_arrHash(nPosition) = _Item->pNext
         else
            _ItemPrev->pNext = _Item->pNext   
         END IF
         delete _Item
         return true
      End If
      If _Item->pNext = 0 Then Exit Do   
      _ItemPrev = _Item
      _Item = _Item->pNext
   Loop
   
   function = false
End function

''
''
function clsFastHash.DeleteAllData2(ByRef sData2 As Const String) as BOOLEAN
   Dim _Item     As _HashItem Ptr
   Dim _ItemNext As _HashItem Ptr
   Dim _ItemPrev As _HashItem Ptr
   
   ' Incoming sData2 is already UCASE

   ' Delete all nodes that match the Data2 field. This is currently used
   ' to match Filenames held in Data2 to delete all functions for a specific
   ' filename about to be parsed.
   function = false
   For i As LONG = LBound(m_arrHash) To Ubound(m_arrHash)
      If m_arrHash(i) <> 0 Then
         _Item = m_arrHash(i)
         _ItemPrev = 0
         Do Until _Item = 0
            If ucase(_Item->sData2) = sData2 Then 
               ' Matching Data2 was located. Set the previous item to point to the
               ' item after this one and then deallocate this current node.
               if _ItemPrev = 0 THEN
                  m_arrHash(i) = _Item->pNext
               else
                  _ItemPrev->pNext = _Item->pNext   
               END IF
               _ItemNext = _Item->pNext
               Delete _Item
               _Item = _ItemNext
               function = true
            else
               _ItemPrev = _Item
               _Item = _Item->pNext
            end if
         Loop
      End If
   Next    
   
End function


''
''
function clsFastHash.BuildKeyList() as string
   ' Build a space delimited list of all keys in this hash list. It is
   ' used for creating a list of autocomplete popup selections (gTYPES).
   Dim _Item    As _HashItem Ptr
   dim sKeyList as string 
   
   For i As LONG = LBound(m_arrHash) To Ubound(m_arrHash)
      If m_arrHash(i) <> 0 Then
         _Item = m_arrHash(i)
         Do Until _Item = 0
            sKeyList = sKeyList & " " & _Item->sKey
            _Item = _Item->pNext
         Loop
      End If
   Next    
   
   function = sKeyList
End function


''
''
Function clsFastHash.SaveToFile( ByRef sFilename As Const String, _
                                 ByRef sHeaderText As Const String = "", _
                                 ByRef sDelimiter As Const String = "=" _
                                 ) As LONG

   Dim f As LONG         
   Dim _Item As _HashItem Ptr
   
   f = Freefile
   Open sFilename For Output As #f

   If Len(sHeaderText) Then
      Print #f, sHeadertext
   End If
      
   For i As LONG = LBound(m_arrHash) To Ubound(m_arrHash)
      If m_arrHash(i) = 0 Then Continue For
      Print #f, m_arrHash(i)->sKey & sDelimiter & m_arrHash(i)->sData1
      _Item = m_arrHash(i)->pNext
      Do Until _Item = 0
         Print #f, _Item->sKey & sDelimiter & _Item->sData1
         _Item = _Item->pNext
      Loop
   Next    
   
   Close #f

   Function = 0
End Function


''
''
Function clsFastHash.CreateDeclaresFile( ByRef sFilename As Const String ) As LONG

   Dim f As LONG         
   Dim _Item As _HashItem Ptr
   dim pDoc as clsDocument ptr
      
   f = Freefile
   Open sFilename For Output As #f

' Structure for gFunctionDeclares
' sKey   = Function Name
' sData1 = Function Declare
' sData2 = Filename
' sData3 = Str(LineNumber)

   For i As LONG = LBound(m_arrHash) To Ubound(m_arrHash)
      If m_arrHash(i) = 0 Then Continue For

      ' Bypass any declares for Class subs/functions
      if instr(m_arrHash(i)->sKey, ".") = 0 THEN
         if gApp.GetDocumentPtrByFilename(wstr(m_arrHash(i)->sData2)) THEN 
            Print #f, "Declare " & AfxStrShrink(m_arrHash(i)->sData1, Chr(32,9,13,10))
         end if
      END IF

      _Item = m_arrHash(i)->pNext
      Do Until _Item = 0
         ' Bypass any declares for Class subs/functions
         if instr(_Item->sKey, ".") = 0 THEN
            if gApp.GetDocumentPtrByFilename(wstr(_Item->sData2)) THEN 
               Print #f, "Declare " & AfxStrShrink(_Item->sData1, Chr(32,9,13,10))
            end if
         END IF
         _Item = _Item->pNext
      Loop

   Next    
   
   Close #f

   Function = 0
End Function

''
''
Function clsFastHash.LoadFromFile( ByRef sFilename As Const String, _
                                   ByRef sDelimiter As Const String = "=" _
                                   ) As LONG
   Dim st     As String
   Dim sKey   As String
   Dim sData1 As String
   Dim f      As LONG         
   Dim i      As LONG
   
   f = Freefile
   If Open( sFilename For Input As #f ) = 0 Then
      Do Until Eof(f)
         Line Input #f, st
         
         st = Trim(st)
         If Len(st) = 0 Then Continue Do
         
         i = Instr(st, sDelimiter)
         If i Then
            sKey   = Left(st, i-1)
            sData1 = Mid(st, i+1)
            this.AddItem( sKey, sData1 ) 
         End If   
      Loop
      Close #f
   End If

   Function = 0
End Function

''
''
Sub clsFastHash.Debug( ByRef sFilename As Const String)
   Dim _Item As _HashItem Ptr

   Dim f As long = Freefile
   Open sFilename For Output As #f
   
   For i As long = LBound(m_arrHash) To Ubound(m_arrHash)
      If m_arrHash(i) = 0 Then
         Print #f, "Slot: "; i, "<Null>"
      Else   
         Print #f, "Slot: "; i, m_arrHash(i)->sKey, m_arrHash(i)->sData1   ', _gHashLang(i)->pNextItem
         _Item = m_arrHash(i)->pNext
         Do Until _Item = 0
            Print #f, "       ", _Item->sKey, _Item->sData1
            _Item = _Item->pNext
         Loop
      End If
   Next    
   
   Close #f
End Sub

' Create global hash variable to hold information for all FB language codetips
dim shared gCodetips as clsFastHash
' Create global hash variable to hold declares for all Subs/Functions
dim shared gFunctionDeclares as clsFastHash = 1000
' Create global hash variable to hold filenames and subs/functions within it
dim shared gFunctionLists as clsFastHash
' Create global hash variable to hold TYPE names and their elements
dim shared gTYPES as clsFastHash
