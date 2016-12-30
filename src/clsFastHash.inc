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
   sKey   As String
   sData  As String
   pNext  As _HashItem Ptr
End Type


Type clsFastHash
   Private:
      m_Capacity  As LONG
      m_FoundNode As _HashItem Ptr
      
      Dim m_arrHash(Any) As _HashItem Ptr 
   
      Declare Function GetSlot( ByRef sKey As Const String) As LONG

   Public:
      Declare Function AddItem( ByRef sKey As Const String, ByRef sData As Const String) As LONG   
      Declare Function FindItem( ByRef sKey As Const String) As boolean   
      Declare function ClearHash() as BOOLEAN 
      Declare Property FoundKey() As String
      Declare Property FoundData() As String
      Declare Function SaveToFile( ByRef sFilename As Const String, _
                                   ByRef sHeaderText As Const String = "", _
                                   ByRef sDelimiter As Const String = "=" _
                                   ) As LONG
      Declare Function LoadFromFile( ByRef sFilename As Const String, _
                                     ByRef sDelimiter As Const String = "=" _
                                     ) As LONG
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
Property clsFastHash.FoundData() As String
   If m_FoundNode Then Property = m_FoundNode->sData
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
Function clsFastHash.AddItem( ByRef sKey    As Const String, _
                              ByRef sData   As Const String _
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
      ' Add the key/data to the linked list 
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
   _Item->sData  = sData
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
      Print #f, m_arrHash(i)->sKey & sDelimiter & m_arrHash(i)->sData
      _Item = m_arrHash(i)->pNext
      Do Until _Item = 0
         Print #f, m_arrHash(i)->sKey & sDelimiter & m_arrHash(i)->sData
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
   Dim st    As String
   Dim sKey  As String
   Dim sData As String
   Dim f     As LONG         
   Dim i     As LONG
   
   f = Freefile
   If Open( sFilename For Input As #f ) = 0 Then
      Do Until Eof(f)
         Line Input #f, st
         
         st = Trim(st)
         If Len(st) = 0 Then Continue Do
         
         i = Instr(st, sDelimiter)
         If i Then
            sKey  = Left(st, i-1)
            sData = Mid(st, i+1)
            this.AddItem( sKey, sData ) 
         End If   
      Loop
      Close #f
   End If

   Function = 0
End Function

dim shared gCodetips as clsFastHash

