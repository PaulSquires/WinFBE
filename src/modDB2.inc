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
''  Application in-memory database
''

' VARIABLE NAMES
' ==============
' Every variable name is prefaced with its owner's name in order to differentiate
' it from other variables (Global variables have no owner name). Usually owners are class/TYPEs.
' The incoming parameter list for a Function is also processed as a variable.
' Therefore, the format is:  <OwnerName>.<FunctionName>.<VariableName>
'
' eg.  
'      TYPE MyType
'         PUBLIC:
'            st As String
'            Declare Function CalculateAmount( ByVal nDollars As Single ) As Single
'      END TYPE
'
'      DIM SHARED AS MyType g     '   ..g     ElementName = ..g   ElementValue = MyType
'
' eg. (Function as part of TYPE)
'      Function MyType.CalculateAmount( ByVal nDollars As Single ) As Single  
'         ' ElementName = MyType.CalculateAmount.nDollars   ElementValue = Single
'         Dim n As Long        ' ElementName = MyType.CalculateAmount.n   ElementValue = Long
'         Dim rc As RECT       ' ElementName = MyType.CalculateAmount.rc  ElementValue = RECT
'      End Function
'
' eg. (Standalone Function)
'      Function MyFunction( ByVal p As MyType Ptr ) As Long   ' ElementName = .MyFunction.p   ElementValue = MyType
'         Dim n As Long        ' ElementName = .MyFunction.n   ElementValue = Long
'         Dim m As MyType      ' ElementName = .MyFunction.m   ElementValue = MyType
'      End Function
'
'
' FUNCTION NAMES
' ==============
'    ElementName = Function name    (can also be fully qualified if from a TYPE.  eg.  MyType.CalculateAmount
'    ElementValue = The Calltip
'
' TYPE STRUCTURES
' ===============
'    ElementName  = The name of the CLASS/TYPE.
'    ElementValue = The true name of the CLASS/TYPE (b/c ElementName may be an ALIAS)
'
' ENUM STRUCTURES
' ===============
'    ElementName  = The name of the ENUM
'    ElementValue = The name of the ENUM
'
' TODO LABELS
' ===============
'    ElementValue = The text of the TODO as found in code.
'
' 

''
''
constructor clsDB2
   FirstNode4Proj=0
   LastNode4Proj=0
   m_ptrUnUsedNode4DB=0
   m_ptrRewindData4File=0
   m_ptrRewindData4Proj=0
   m_ptrRewindData4DB=0
   
   dbAdd4PublicProject()
   dbRewind4PublicProject()               
   ' Add the standard builtin data types
   dim parser as clsParser
   with parser
      .lineNum    = -1
      .idxProject = -1
      .Filename   = ""
      .OwnerNode4DB = 0
   
      .typeName = "Boolean"
      .Description = !"Data type that holds the value of True or False.\n"
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
                  
      .typeName = "Byte"      
      .Description = !"8-bit signed whole-number data type.\nValue range of -128 to 127."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  

      .typeName = "UByte"
      .Description = !"8-bit signed whole-number data type.\nValue range of 0 to 255."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
      
      .typeName = "Short"
      .Description = !"16-bit signed whole-number data type.\nValue range of -32768 to 32767."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
      
      .typeName = "UShort"
      .Description = !"16-bit signed whole-number data type.\nValue range of 0 to 65535."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  

      .typeName = "Integer"
      .Description = !"32-bit or 64-bit signed whole-number data type depending on the platform.\n"
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
      
      .typeName = "UInteger"
      .Description = !"32-bit or 64-bit unsigned whole-number data type depending on the platform.\n"
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
      
      .typeName = "Long"
      .Description = !"32-bit signed whole-number data type.\nValue range of -2147483648 to 2147483647."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  

      .typeName = "ULong"
      .Description = !"32-bit unsigned whole-number data type.\nValue range of 0 to 4294967295."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
      
      .typeName = "LongInt"
      .Description = !"A 64-bit signed whole-number data type.\nValues range of -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807"
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  

      .typeName = "ULongInt"
      .Description = !"A 64-bit unsigned whole-number data type.\nValues range of 0 to 18,446,744,073,709,551,615"
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  

      .typeName = "Single"
      .Description = !"32-bit floating-point data type.\nContains at most 24 bits of precision, or about 6 decimal digits."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  

      .typeName = "Double"
      .Description = !"64-bit floating-point data type.\nContains at most 53 bits of precision, or about 15 decimal digits."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  

      .typeName = "String"
      .Description = "A String is an array of characters."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
      
      .typeName = "WString"
      .Description = !"A fixed-size array of wide-chars that never overflows if the size is known at compile-time.\n"
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  
      
      .typeName = "ZString"
      .Description = "A C-style fixed-size array of chars."
      .CurrentNode4DB = this.dbAdd4DB(parser, DB2_STANDARDDATATYPE)  

   end with
   this.dbRewind4DB
end constructor

''
''
Destructor clsDB2
   while FirstNode4Proj <> 0
     dbDelete4Proj(FirstNode4Proj)
   wend
   dbDelete4DB(m_ptrUnUsedNode4DB)  
end Destructor

''
''
function clsDB2.dbAdd4PublicProject()  as boolean 
   dbAdd4Proj(defPUBLICPROJECTDATA,-1) 
   dbRewind4Proj(defPUBLICPROJECTDATA,-1)
   dbAdd4File(defCodetips4FB)
   return dbRewind4File(defCodetips4FB)<>0
end function

''
''
function clsDB2.dbRewind4PublicProject() as boolean 
  return dbRewind4File(defPUBLICPROJECTDATA,defCodetips4FB,-1)<>0 
end function

''
''
Property clsDB2.ptrRewindData4File() As File_DATA ptr
   return m_ptrRewindData4File
end property

''
''
Property clsDB2.ptrRewindData4Proj() As Project_DATA ptr
   return m_ptrRewindData4Proj
end property

''
''
function clsDB2.dbFreeIndex4Proj() as long
   dim as Long index=-1
   dim as Project_DATA ptr ptrData4Proj=FirstNode4Proj
   while ptrData4Proj<>0
      if ptrData4Proj->ProjIndex>index then
         index=ptrData4Proj->ProjIndex
      end if
      ptrData4Proj=ptrData4Proj->NextNode
   wend 
   return index+1  
end function

''
''
function clsDB2.dbFindIndex4Proj( byref wszProjName as wstring) as long
   dim as Project_DATA ptr ptrData4Proj=dbFind4Proj(wszProjName,-1)
   if ptrData4Proj<>0 then
      return ptrData4Proj->ProjIndex
   else
      return -2
   end if       
end function

''
''
function clsDB2.dbAdd4Proj( byref wszProjName as wstring, byval ProjIndex as long=-1) as Project_DATA ptr
     dim as Project_DATA ptr ptrData4Proj = dbFind4Proj(wszProjName, ProjIndex)
     if ptrData4Proj = 0 then
         ptrData4Proj = new Project_DATA
         ptrData4Proj->ProjIndex = ProjIndex
         ptrData4Proj->ProjName = wszProjName
         ptrData4Proj->FirstNode4File = 0
         ptrData4Proj->LastNode4File = 0
         ptrData4Proj->PrevNode = 0
         ptrData4Proj->NextNode = 0
         
         if FirstNode4Proj = 0 then
            FirstNode4Proj = ptrData4Proj      
         eLSEif LastNode4Proj <> 0 then 
            ptrData4Proj->PrevNode = LastNode4Proj
            LastNode4Proj->NextNode = ptrData4Proj
         endif   
         LastNode4Proj = ptrData4Proj
     end if
     dbRewind4Proj(wszProjName, ProjIndex)
     return ptrData4Proj
end function

''
''
function clsDB2.dbDelete4Proj Overload( byval pProjectDATA as Project_DATA ptr ) as boolean
   dim as File_DATA Ptr ptrData4File = any
   dim as Project_DATA ptr tmpProjectDATA = any
   
   IF pProjectDATA THEN
      tmpProjectDATA = FirstNode4Proj
      do
        select case tmpProjectDATA
        case 0
            return false
        case pProjectDATA
            exit do
        end select
        tmpProjectDATA = tmpProjectDATA->NextNode
      loop       

      while pProjectDATA->FirstNode4File <> 0
         dbDelete4File(pProjectDATA->FirstNode4File)
      wend
  
      if FirstNode4Proj = LastNode4Proj then
         if pProjectDATA = FirstNode4Proj then
            FirstNode4Proj = 0
            LastNode4Proj = 0
         else
            return false
         end if
      elseif pProjectDATA = FirstNode4Proj then       
         FirstNode4Proj = pProjectDATA->NextNode
         if FirstNode4Proj <> 0 then 
            FirstNode4Proj->PrevNode = 0
         else
            LastNode4Proj=0
         end if
      elseif pProjectDATA = LastNode4Proj then       
         LastNode4Proj = pProjectDATA->PrevNode
         if LastNode4Proj <> 0 then 
            LastNode4Proj->NextNode = 0
         else
            FirstNode4Proj = 0
         end if
      else
         tmpProjectDATA = pProjectDATA->PrevNode
         if tmpProjectDATA then
          tmpProjectDATA->NextNode = pProjectDATA->NextNode
         end if
        
         tmpProjectDATA= pProjectDATA->NextNode
         if tmpProjectDATA <> 0 then
            tmpProjectDATA->PrevNode = pProjectDATA->PrevNode
         end if
      end if
      
      if m_ptrRewindData4Proj = pProjectDATA then
         m_ptrRewindData4File = 0
         m_ptrRewindData4Proj = 0
      endif
      
      dbRewind4DB()
      
      delete pProjectDATA 
      RETURN TRUE
   ELSE
      RETURN FALSE
   end if 
end function

''
'' 
function clsDB2.dbDelete4Proj Overload( byval ProjIndex as long ) as boolean
   RETURN dbDelete4Proj(dbFind4Proj(ProjIndex))
end function

''
'' 
function clsDB2.dbDelete4Proj Overload( byref wszProjName as wstring, byval ProjIndex as long=-1 ) as boolean
   RETURN dbDelete4Proj(dbFind4Proj(wszProjName, ProjIndex))
end function

''
'' 
function clsDB2.dbFind4Proj Overload( byref wszProjName as wstring, byval ProjIndex as long=-1) as Project_DATA ptr
   dim as Project_DATA ptr ptrData4Proj = FirstNode4Proj
   while ptrData4Proj <> 0
      if (WStrIsEqual(wszProjName,ptrData4Proj->ProjName,false) = true) and ((ptrData4Proj->ProjIndex = ProjIndex) or (ProjIndex = -1)) = true then
         exit while
      end if
      ptrData4Proj = ptrData4Proj->NextNode
   wend   
   return  ptrData4Proj
   
end function

''
''
function clsDB2.dbFind4Proj Overload( byval ProjIndex as long) as Project_DATA ptr
   dim as Project_DATA ptr ptrData4Proj=FirstNode4Proj
   while ptrData4Proj<>0
      if (ptrData4Proj->ProjIndex=ProjIndex) then
         exit while
      end if
      ptrData4Proj=ptrData4Proj->NextNode
   wend
   
   return  ptrData4Proj
end function

''
''
function clsDB2.dbRewind4Proj Overload() as Project_DATA ptr
   m_ptrRewindData4Proj=FirstNode4Proj   
   if m_ptrRewindData4Proj then
      m_ptrRewindData4File=m_ptrRewindData4Proj->FirstNode4File
   else
      m_ptrRewindData4File=0
   end if
   dbRewind4DB()
   return m_ptrRewindData4Proj  
end function

''
''
function clsDB2.dbRewind4Proj Overload(byval ProjIndex as long) as Project_DATA ptr
   m_ptrRewindData4Proj=dbFind4Proj(ProjIndex)   
   if m_ptrRewindData4Proj then
      m_ptrRewindData4File=m_ptrRewindData4Proj->FirstNode4File
   else
      m_ptrRewindData4File=0
   end if
   dbRewind4DB()
   return m_ptrRewindData4Proj  
end function

''
''
function clsDB2.dbRewind4Proj Overload(byref wszProjName as wstring, byval ProjIndex as long=-1) as Project_DATA ptr
   m_ptrRewindData4Proj=dbFind4Proj(wszProjName,ProjIndex) 
   if m_ptrRewindData4Proj then
      m_ptrRewindData4File=m_ptrRewindData4Proj->FirstNode4File
   else
      m_ptrRewindData4File=0
   end if
   dbRewind4DB()
   return m_ptrRewindData4Proj  
end function

''
''
function clsDB2.dbGetNext4Proj() as Project_DATA ptr
   if m_ptrRewindData4Proj then
      m_ptrRewindData4Proj = m_ptrRewindData4Proj->NextNode
   end if
   return m_ptrRewindData4Proj
end function 

''
''
function clsDB2.dbAdd4File( byref wszFilename as wstring) as File_DATA ptr
   dim as File_DATA Ptr ptrData4File = 0
   if m_ptrRewindData4Proj then
      ptrData4File = dbFind4File(wszFilename)
      if ptrData4File = 0 then
         ptrData4File = new File_DATA
         ptrData4File->FileName = wszFilename
         ptrData4File->pProjectDATA = m_ptrRewindData4Proj
         ptrData4File->FirstOwnerNode4DB = 0
         ptrData4File->LastOwnerNode4DB = 0 
         ptrData4File->PrevNode = 0
         ptrData4File->NextNode = 0
         
         if m_ptrRewindData4Proj->FirstNode4File = 0 then
            m_ptrRewindData4Proj->FirstNode4File = ptrData4File      
         end if
       
         if m_ptrRewindData4Proj->LastNode4File <> 0 then 
            ptrData4File->PrevNode = m_ptrRewindData4Proj->LastNode4File
            m_ptrRewindData4Proj->LastNode4File->NextNode = ptrData4File
         endif
            
         m_ptrRewindData4Proj->LastNode4File = ptrData4File
      end if
      dbRewind4File(wszFilename)
   end if
   return ptrData4File
end function

''
''
function clsDB2.dbAdd4File Overload( byval ProjIndex as long, byref wszFilename as wstring) as File_DATA ptr 
   dbRewind4Proj(ProjIndex) 
   return dbAdd4File(wszFilename)
end function

function clsDB2.dbAdd4File Overload( byref wszProjName as wstring, byref wszFilename as wstring, byval ProjIndex as long=-1) as File_DATA ptr
   dbRewind4Proj(wszProjName,ProjIndex)
   return dbAdd4File(wszFilename)
end function

''
''
function clsDB2.dbDelete4File Overload( byref wszFilename as wstring ) as boolean
   return dbDelete4File(dbFind4File(wszFilename))       
end function

''
''
function clsDB2.dbDelete4File Overload( byval ProjIndex as long, byref wszFilename as wstring ) as boolean
   dbRewind4Proj(ProjIndex)
   return dbDelete4File(wszFilename) 
end function 

''
''
function clsDB2.dbDelete4File Overload( byref wszProjName as wstring, byref wszFilename as wstring, byval ProjIndex as long=-1 ) as boolean
   dbRewind4Proj(wszProjName,ProjIndex)
   return dbDelete4File(wszFilename)       
end function

''
''
function clsDB2.dbDelete4File Overload( byval pDelFileDATA as File_DATA ptr ) as boolean 
   dim as File_DATA Ptr tmpNode4File=any
   dim as Project_DATA ptr pProjectDATA=any
   dim as DB2_DATA Ptr tmpNode4DB=any
   if pDelFileDATA then
      pProjectDATA = pDelFileDATA->pProjectDATA
      if pProjectDATA then
         tmpNode4File=pProjectDATA->FirstNode4File
         do
           select case tmpNode4File
           case 0
               return false
           case pDelFileDATA
               exit do
           end select
           tmpNode4File=tmpNode4File->NextNode
         loop
                 
         tmpNode4DB = pDelFileDATA->FirstOwnerNode4DB
         while tmpNode4DB<>0
            tmpNode4DB->deleted=true
            tmpNode4DB=tmpNode4DB->NextNode
         wend
          
         if m_ptrUnUsedNode4DB then
            if pDelFileDATA->LastOwnerNode4DB then
              pDelFileDATA->LastOwnerNode4DB->NextNode=m_ptrUnUsedNode4DB
              m_ptrUnUsedNode4DB->PrevNode=pDelFileDATA->LastOwnerNode4DB
              m_ptrUnUsedNode4DB=pDelFileDATA->FirstOwnerNode4DB
              IF m_ptrUnUsedNode4DB THEN
                 m_ptrUnUsedNode4DB->PrevNode = 0
              end if
            endif
         else
            m_ptrUnUsedNode4DB=pDelFileDATA->FirstOwnerNode4DB
            if m_ptrUnUsedNode4DB then m_ptrUnUsedNode4DB->PrevNode = 0
         end if
         
         if pProjectDATA->FirstNode4File=pProjectDATA->LastNode4File then
            if pDelFileDATA=pProjectDATA->FirstNode4File then
               pProjectDATA->FirstNode4File=0
               pProjectDATA->LastNode4File=0
            else
               return false
            end if
         elseif pDelFileDATA=pProjectDATA->FirstNode4File then
            tmpNode4File=pProjectDATA->FirstNode4File
            tmpNode4File=tmpNode4File->NextNode
            if tmpNode4File then 
               tmpNode4File->PrevNode=0
            else
               pProjectDATA->LastNode4File=0  
            end if
            pProjectDATA->FirstNode4File = tmpNode4File
         elseif pDelFileDATA=pProjectDATA->LastNode4File then
            tmpNode4File=pProjectDATA->LastNode4File         
            tmpNode4File=tmpNode4File->PrevNode
            if tmpNode4File then 
               tmpNode4File->NextNode=0
            else
               pProjectDATA->FirstNode4File=0
            end if
            pProjectDATA->LastNode4File = tmpNode4File
         else
            tmpNode4File= pDelFileDATA->PrevNode
            if tmpNode4File then
               tmpNode4File->NextNode=pDelFileDATA->NextNode
            end if
            
            tmpNode4File= pDelFileDATA->NextNode
            if tmpNode4File then
              tmpNode4File->PrevNode=pDelFileDATA->PrevNode
            end if
         endif
          
         if m_ptrRewindData4File=pDelFileDATA then
            m_ptrRewindData4File=0
         endif
         
         dbRewind4DB()
         
         tmpNode4File=pProjectDATA->FirstNode4File
      
         delete pDelFileDATA 
         RETURN TRUE
      endif
   end if 
   RETURN FALSE
end function

''
''
function clsDB2.dbFind4File( byref wszFilename as wstring) as File_DATA ptr
   dim as File_DATA Ptr ptrData4File=0
   if m_ptrRewindData4Proj then
      ptrData4File=m_ptrRewindData4Proj->FirstNode4File
      while ptrData4File<>0
         if WStrIsEqual(wszFilename,ptrData4File->FileName,false) then
            exit while
         end if 
         ptrData4File = ptrData4File->NextNode
      wend
   end if
   return ptrData4File
end function

''
''
function clsDB2.dbRewind4File Overload() as File_DATA ptr 
   if m_ptrRewindData4Proj then
      m_ptrRewindData4File=m_ptrRewindData4Proj->FirstNode4File
   else
      m_ptrRewindData4File=0
   end if
   dbRewind4DB()
   return m_ptrRewindData4File   
end function

''
''
function clsDB2.dbRewind4File Overload(byref wszFilename as wstring) as File_DATA ptr 
   if m_ptrRewindData4Proj then
      m_ptrRewindData4File=dbFind4File(wszFilename)
   else
      m_ptrRewindData4File=0
   end if
   dbRewind4DB()
   return m_ptrRewindData4File   
end function

''
''
function clsDB2.dbRewind4File Overload(byval ProjIndex as long, byref wszFilename as wstring ) as File_DATA ptr
   dbRewind4Proj(ProjIndex)
   return dbRewind4File(wszFilename)  
end function

''
''  
function clsDB2.dbRewind4File Overload(byref wszProjName as wstring, byref wszFilename as wstring, byval ProjIndex as long=-1) as File_DATA ptr
   dbRewind4Proj(wszProjName,ProjIndex)
   return dbRewind4File(wszFilename) 
end function

''
''  
function clsDB2.dbGetNext4File() as File_DATA ptr
   if m_ptrRewindData4File then
      m_ptrRewindData4File = m_ptrRewindData4File->NextNode
   end if
   return m_ptrRewindData4File
end function

''
''
function clsDB2.dbNew( ) as DB2_DATA ptr
   DIM as DB2_DATA ptr tmpNode4DB=any
   static as DB2_DATA vNode4DB 
   if m_ptrUnUsedNode4DB then
      tmpNode4DB=dbPopUnUsedNode4DB(m_ptrUnUsedNode4DB)
      if tmpNode4DB <> 0 then
         if tmpNode4DB=m_ptrUnUsedNode4DB then
            m_ptrUnUsedNode4DB=m_ptrUnUsedNode4DB->NextNode
            if m_ptrUnUsedNode4DB then
               m_ptrUnUsedNode4DB->PrevNode=0
            endif 
         endif
         *tmpNode4DB= vNode4DB
         RETURN tmpNode4DB 
      endif
   end if
   function = new DB2_DATA
end function

''
''
function clsDB2.dbAdd4DB Overload( byref db as DB2_DATA ) as boolean
   dim as DB2_DATA ptr ptrdb=new DB2_DATA
   Poke DB2_DATA, ptrdb, db
   return dbAdd4DB(ptrdb)
end function

''
''
function clsDB2.dbAdd4DB Overload( byval ptrdb as DB2_DATA ptr ) as boolean
   return dbAdd4DB(ptrdb,0)
end function

''
'
function clsDB2.dbAdd4DB Overload( byval ptrdb as DB2_DATA ptr ,byval prevptrdb as DB2_DATA ptr ) as boolean
   if ptrdb<>0 then
      ptrdb->PrevNode = 0            
      ptrdb->NextNode = 0
      if prevptrdb<>0 then
         if ptrdb->OwnerNode4DB = prevptrdb then
            prevptrdb = prevptrdb->ChildNode4DB
            IF prevptrdb=0 THEN
               ptrdb->OwnerNode4DB->ChildNode4DB= ptrdb
               return true
            ELSE
               prevptrdb=dbGetLastSibling4DB(prevptrdb)
            ENDIF 
         end if
         ptrdb->PrevNode = prevptrdb
         prevptrdb->NextNode = ptrdb
         return true
      elseif m_ptrRewindData4File then
         if m_ptrRewindData4File->FirstOwnerNode4DB = 0 then
            m_ptrRewindData4File->FirstOwnerNode4DB = ptrdb
            m_ptrRewindData4File->LastOwnerNode4DB = ptrdb
         else
            ptrdb->PrevNode = m_ptrRewindData4File->LastOwnerNode4DB
            m_ptrRewindData4File->LastOwnerNode4DB->NextNode = ptrdb
            m_ptrRewindData4File->LastOwnerNode4DB = ptrdb
         end if
         return true 
      end if 
   end if
   return false  
end function

''
''
function clsDB2.dbAdd4DB Overload( byref parser as clsParser, _
                       byref id as long _
                       ) as DB2_DATA ptr                       
   if (m_ptrRewindData4File<>0) and (m_ptrRewindData4Proj<>0) then   
      dim as DB2_DATA ptr ptrdb=dbNew()
      if ptrdb then
         with parser
            ptrdb->deleted       = false
            ptrdb->id            = id                ' the type of database record that we are storing
            ptrdb->nStartLineNum = .lineNum
            ptrdb->nEndLineNum   = .lineNum
            ptrdb->pFileDATA     = m_ptrRewindData4File
            ptrdb->pProjectDATA  = m_ptrRewindData4Proj
            ptrdb->OwnerNode4DB  = .OwnerNode4DB

            select CASE id
               case DB2_STANDARDDATATYPE
                  ptrdb->ElementName  = .typeName
                  ptrdb->ElementValue = .typeName
                  ptrdb->IsWinApi     = .IsWinApi
                  if ptrdb->ElementName = "" then exit function

               case DB2_TYPE    ' this handles ENUM also
                  ' If the type has already been added skip adding it again. 
                  if this.dbFindTYPE( .typeName) then return 0
                  ptrdb->nStartLineNum= .funcLineNum
                  ptrdb->ElementName  = .typeName
                  ptrdb->ElementValue = .typeAlias     ' same as typeName unless it was an ALIAS (always search using this name)
                  ptrdb->IsWinApi     = .IsWinApi
                  ptrdb->SubType      = .SubType
                  ptrdb->TypeExtends  = .TypeExtends
                  if ptrdb->ElementName = "" then exit function

               case DB2_TODO
                  ptrdb->ElementName  = ""
                  ptrdb->ElementValue = .todoText

               case DB2_FUNCTION, DB2_SUB, DB2_PROPERTY   
                  ptrdb->nStartLineNum= .funcLineNum
                  ptrdb->ElementName  = .funcName
                  ptrdb->ElementValue = .funcParams   ' Calltip
                  ptrdb->IsWinApi     = .IsWinApi
                  if ptrdb->ElementName = "" then exit function

               case DB2_VARIABLE
                  ' Variables are always in the format:
                  ' <OwnerName>.<FunctionName>.<VariableName>
                  if instr(.funcName, ".") = 0 then
                     ptrdb->ElementName = .typeName & "." & .funcName & "." & .varName
                  Else
                     ' funcName is already mytype.MyFunctionName
                     ptrdb->ElementName = .funcName & "." & .varName
                  end if
                  ptrdb->ElementValue = .varType 
                  ptrdb->IsPrivate = iif(.bInTypePublic, false, true)
                  ptrdb->IsWinApi     = .IsWinApi
                  if ptrdb->ElementName = "" then exit function
            END SELECT
            IF dbAdd4db(ptrdb,.CurrentNode4DB)=true THEN
               return ptrdb
            else
               dbPushUnUsedNode4DB(ptrdb)
            end if  
         end with
      end if
   end if
   function = 0

end function

''
''
function clsDB2.dbPushUnUsedNode4DB Overload( byval pDB2DATA as DB2_DATA ptr ) as boolean  
   dim as File_DATA Ptr pFileDATA=any
   DIM as DB2_DATA ptr tmpNode4DB=any
    if pDB2DATA then
      IF pDB2DATA->deleted=TRUE THEN return true
      pFileDATA=pDB2DATA->pFileDATA
      if pFileDATA then
         if pDB2DATA=pFileDATA->FirstOwnerNode4DB then
             if pFileDATA->FirstOwnerNode4DB=pFileDATA->LastOwnerNode4DB then
               pFileDATA->FirstOwnerNode4DB=0
               pFileDATA->LastOwnerNode4DB=0 
             else
               pFileDATA->FirstOwnerNode4DB=pFileDATA->FirstOwnerNode4DB->NextNode
               if pFileDATA->FirstOwnerNode4DB then
                  pFileDATA->FirstOwnerNode4DB->PrevNode=0
               else
                  pFileDATA->LastOwnerNode4DB=0
               end if
             endif
         elseif pDB2DATA=pFileDATA->LastOwnerNode4DB then
             if pFileDATA->FirstOwnerNode4DB=pFileDATA->LastOwnerNode4DB then
                  pFileDATA->FirstOwnerNode4DB=0
                  pFileDATA->LastOwnerNode4DB=0 
             else
                  pFileDATA->LastOwnerNode4DB=pFileDATA->LastOwnerNode4DB->PrevNode
                  if pFileDATA->LastOwnerNode4DB then
                     pFileDATA->LastOwnerNode4DB->NextNode=0
                  else
                     pFileDATA->FirstOwnerNode4DB=0
                  end if
             endif
         else
            tmpNode4DB= pDB2DATA->PrevNode
            if tmpNode4DB then
              tmpNode4DB->NextNode=pDB2DATA->NextNode
            end if

            tmpNode4DB= pDB2DATA->NextNode
            if tmpNode4DB then
              tmpNode4DB->PrevNode=pDB2DATA->PrevNode
            end if
         end if
         
         pDB2DATA->deleted=true
         pDB2DATA->PrevNode=0
         if m_ptrUnUsedNode4DB then
            pDB2DATA->NextNode=m_ptrUnUsedNode4DB
            m_ptrUnUsedNode4DB->PrevNode=pDB2DATA
            m_ptrUnUsedNode4DB=pDB2DATA
         else
            pDB2DATA->NextNode=0
            m_ptrUnUsedNode4DB=pDB2DATA
         end if
      endif      
      return true
    end if
    return false
end function

function clsDB2.dbPushUnUsedNode4DB Overload( byref wszDB2Name as wstring, byval Action as long, byval Index as Long = 1 ) as boolean
   return dbPushUnUsedNode4DB(dbSeek(wszDB2Name,Action,Index))
end function

''
''
function clsDB2.dbPopUnUsedNode4DB(byval MainNode4DB as DB2_DATA ptr) as DB2_DATA ptr
   DIM as DB2_DATA ptr tmpNode4DB = any, CurrentNode=0 
   if MainNode4DB then
      tmpNode4DB = MainNode4DB->ChildNode4DB
      if tmpNode4DB<>0 then
         CurrentNode = dbPopUnUsedNode4DB(tmpNode4DB)
         if CurrentNode = tmpNode4DB then
            MainNode4DB->ChildNode4DB = 0
         endif
      else
         CurrentNode = MainNode4DB->NextNode
         if CurrentNode <> 0 then
            tmpNode4DB=CurrentNode->NextNode
            MainNode4DB->NextNode=tmpNode4DB
            IF tmpNode4DB THEN
               tmpNode4DB->PrevNode=MainNode4DB
            EnDIF
         else
            CurrentNode = MainNode4DB
         endif
      end if
   end if
   RETURN CurrentNode
end function 
 
''
''
function clsDB2.dbDelete4DB(byval MainNode4DB as DB2_DATA ptr) as boolean
   DIM as DB2_DATA ptr CurrentNode=0 
   if MainNode4DB then
      CurrentNode= MainNode4DB->ChildNode4DB
      if CurrentNode then
         dbDelete4DB(CurrentNode)
      else
         while MainNode4DB <> 0
            CurrentNode = MainNode4DB            
            MainNode4DB = MainNode4DB->NextNode
            delete CurrentNode
         wend 
      end if
   end if
   RETURN true
end function

''
''
function clsDB2.dbRewind4DB() as DB2_DATA ptr
   ' Set index pointer to immediately before first array index 
   if m_ptrRewindData4File then
      m_ptrRewindData4DB=m_ptrRewindData4File->FirstOwnerNode4DB
   else
      m_ptrRewindData4DB=0
   end if
   function = m_ptrRewindData4DB
END FUNCTION

''
''
function clsDB2.dbDeleteWinAPI() as boolean
   dim as File_DATA Ptr l_ptrRewindData4File=m_ptrRewindData4File
   dim as Project_DATA ptr l_ptrRewindData4Proj=m_ptrRewindData4Proj 
   dim as DB2_DATA ptr pData=any
   dbRewind4PublicProject()
   IF this.dbRewind4DB() THEN
      do 
         pData = this.dbGetNext4DB()
         if pData = 0 THEN exit do
         if pData->IsWinApi = true THEN
            dbPushUnUsedNode4DB(pData)
         END IF
      loop
   ENDIF
   m_ptrRewindData4File=l_ptrRewindData4File
   m_ptrRewindData4Proj=l_ptrRewindData4Proj
   dbRewind4DB()
   function = 0
function = true
end function

''
''
function clsDB2.dbGetNext4DB() as DB2_DATA ptr
   ' Set index pointer to next array index that is not deleted
   dim as DB2_DATA ptr pData= m_ptrRewindData4DB
   IF pData THEN
      while pData<>0
         if pData->deleted THEN 
            pData=pData->NextNode
         else
            exit while   
         END IF      
      wend 
      IF pData THEN
         m_ptrRewindData4DB=pData->NextNode
      ELSE
         m_ptrRewindData4DB=0
      end if
   ENDIF
   function = pData
END FUNCTION

function clsDB2.dbGetChild4DB(byval ptrdb as DB2_DATA ptr ) as DB2_DATA ptr
  IF ptrdb THEN
     return ptrdb->ChildNode4DB
  end if
  function = 0
end function

function clsDB2.dbGetOwner4DB(byval ptrdb as DB2_DATA ptr ) as DB2_DATA ptr
  IF ptrdb THEN
     return ptrdb->OwnerNode4DB
  end if
  function = 0
end function

function clsDB2.dbGetFirstSibling4DB(byval ptrdb as DB2_DATA ptr ) as DB2_DATA ptr
  while ptrdb<>0
     if ptrdb->PrevNode=0 then exit while
     ptrdb = ptrdb->PrevNode   
  wend
  function = ptrdb
end function

function clsDB2.dbGetLastSibling4DB(byval ptrdb as DB2_DATA ptr ) as DB2_DATA ptr
  while ptrdb<>0
     if ptrdb->NextNode=0 then exit while
     ptrdb = ptrdb->NextNode   
  wend
  function = ptrdb
end function

function clsDB2.dbGetPrevSibling4DB(byval ptrdb as DB2_DATA ptr ) as DB2_DATA ptr
  IF ptrdb THEN
     return ptrdb->PrevNode
  end if
  function = 0
end function

function clsDB2.dbGetNextSibling4DB(byval ptrdb as DB2_DATA ptr ) as DB2_DATA ptr
  IF ptrdb THEN
     return ptrdb->NextNode
  end if
  function = 0
end function
   
function clsDB2.dbSeek4DB Overload( byval sLookFor as string, byval Action as long, byref SequentialIndex as Long = 1 ) as DB2_DATA ptr
   dim as DB2_DATA ptr pData = dbRewind4DB
   return dbSeek4DB( pData, sLookFor, Action, SequentialIndex )
END FUNCTION

function clsDB2.dbSeek4DB Overload( byval ptrdb as DB2_DATA ptr, byval sLookFor as string, byval Action as long, byref SequentialIndex as Long = 1 ) as DB2_DATA ptr
   dim pData as DB2_DATA ptr
   'ELF_CurrentFile
   while ptrdb<>0
      if ptrdb->id = Action THEN
         if WStrIsEqual(sLookFor, ptrdb->ElementName, false)=true THEN
            SequentialIndex-=1
            if SequentialIndex<=0 then return ptrdb
         end if
      elseif ptrdb->ChildNode4DB then
         pData=dbSeek4DB(ptrdb->ChildNode4DB,sLookFor,Action,SequentialIndex)
         if pData then return pData
      END IF 
      ptrdb = ptrdb->NextNode
   wend 
   function = 0
END FUNCTION
      
function clsDB2.dbSeek4File( byval sLookFor as string, byval Action as long, byref SequentialIndex as Long = 1 ) as DB2_DATA ptr
   dim as DB2_DATA ptr ptrDB2DATA=0
   dim as File_DATA Ptr l_ptrRewindData4File=any
   
   if m_ptrRewindData4File<>0 and m_ptrRewindData4Proj<>0 then
      'ELF_CurrentFile
      ptrDB2DATA=dbSeek4DB( sLookFor,Action,SequentialIndex) 
      if ptrDB2DATA=0 then 'ELF_CurrentProject
         l_ptrRewindData4File=m_ptrRewindData4File
         m_ptrRewindData4File=m_ptrRewindData4Proj->FirstNode4File
         while m_ptrRewindData4File
            if m_ptrRewindData4File<>l_ptrRewindData4File then
               ptrDB2DATA=dbSeek4DB( sLookFor,Action,SequentialIndex)
               if ptrDB2DATA then exit while
            endif
            m_ptrRewindData4File=m_ptrRewindData4File->NextNode
         wend
         m_ptrRewindData4File = l_ptrRewindData4File
         dbRewind4DB()
      endif
   end if
   function = ptrDB2DATA
END FUNCTION
   
function clsDB2.dbSeek4Proj( byval sLookFor as string, byval Action as long, byval SequentialIndex as Long = 1 ) as DB2_DATA ptr
   dim as File_DATA Ptr l_ptrRewindData4File=m_ptrRewindData4File
   dim as Project_DATA ptr l_ptrRewindData4Proj=m_ptrRewindData4Proj
   dim as Project_DATA ptr l_ptrRewind4PublicProject=any,ptrData4Proj=any
   dim as DB2_DATA ptr ptrDB2DATA=0
   
   'ELF_CurrentFile & ELF_CurrentProject
   ptrDB2DATA=dbSeek4File( sLookFor,Action,SequentialIndex) 
   if ptrDB2DATA=0 then  'ELF_PublicProject
      dbRewind4PublicProject()
      l_ptrRewind4PublicProject = m_ptrRewindData4Proj
      ptrDB2DATA=dbSeek4File( sLookFor,Action,SequentialIndex) 
      if ptrDB2DATA=0 then  'ELF_AllProject
         m_ptrRewindData4Proj=FirstNode4Proj
         while m_ptrRewindData4Proj
            select case m_ptrRewindData4Proj
               case  l_ptrRewindData4File, l_ptrRewind4PublicProject
               case else
                  m_ptrRewindData4File=m_ptrRewindData4Proj->FirstNode4File
                  ptrDB2DATA=dbSeek4File( sLookFor,Action,SequentialIndex)
                  if ptrDB2DATA then exit while
            end select
            m_ptrRewindData4Proj=m_ptrRewindData4Proj->NextNode
         wend            
      endif
   endif
   m_ptrRewindData4File=l_ptrRewindData4File
   m_ptrRewindData4Proj=l_ptrRewindData4Proj
   dbRewind4DB()
   function = ptrDB2DATA
END FUNCTION
  
''
''
function clsDB2.dbSeek( byval sLookFor as string, byval Action as long, byval Type4LookFor as eLookFor = ELF_CurrentProject, byval SequentialIndex as Long = 1 ) as DB2_DATA ptr
   ' GENERIC SEEK FUNCTION THAT OTHER FUNCTIONS CALL TO DO THE HARD WORK
   ' Find the array element that contains the function name beng searched for
   dim as File_DATA Ptr l_ptrRewindData4File=m_ptrRewindData4File
   dim as Project_DATA ptr l_ptrRewindData4Proj=m_ptrRewindData4Proj
   dim as DB2_DATA ptr ptrDB2DATA=0
   select case Type4LookFor
      case ELF_CurrentFile
         ptrDB2DATA=dbSeek4DB( sLookFor,Action,SequentialIndex)
      case ELF_CurrentProject
         ptrDB2DATA=dbSeek4File( sLookFor,Action,SequentialIndex)
      case ELF_PublicProject
         dbRewind4PublicProject()
         ptrDB2DATA=dbSeek4File( sLookFor,Action,SequentialIndex)
      case else
         if (Type4LookFor and ELF_CurrentAndPublicProject) = ELF_CurrentAndPublicProject then
            ptrDB2DATA=dbSeek4File( sLookFor,Action,SequentialIndex)
            if ptrDB2DATA=0 then
               dbRewind4PublicProject()
               ptrDB2DATA=dbSeek4File( sLookFor,Action,SequentialIndex)
            endif
         else
            ptrDB2DATA=dbSeek4Proj( sLookFor,Action,SequentialIndex)
         end if      
   end select
   m_ptrRewindData4File=l_ptrRewindData4File
   m_ptrRewindData4Proj=l_ptrRewindData4Proj
   dbRewind4DB()
   return ptrDB2DATA
end function


''
''
function clsDB2.dbFindFunction( byref sFunctionName as string, byval Type4LookFor as eLookFor= ELF_CurrentProject, byval SequentialIndex as Long = 1 ) as DB2_DATA ptr
   return this.dbSeek(sFunctionName, DB2_FUNCTION, Type4LookFor, SequentialIndex) 
end function

''
''
function clsDB2.dbFindSub( byref sFunctionName as string, byval Type4LookFor as eLookFor= ELF_CurrentProject, byval SequentialIndex as Long = 1 ) as DB2_DATA ptr
   return this.dbSeek(sFunctionName, DB2_SUB, Type4LookFor, SequentialIndex) 
end function

''
''
function clsDB2.dbFindProperty( byref sFunctionName as string, byval Type4LookFor as eLookFor= ELF_CurrentProject, byval SequentialIndex as Long = 1 ) as DB2_DATA ptr
   return this.dbSeek(sFunctionName, DB2_PROPERTY, Type4LookFor, SequentialIndex) 
end function

''
''
function clsDB2.dbFindVariable( byref sVariableName as string, byval Type4LookFor as eLookFor= ELF_CurrentProject, byval SequentialIndex as Long = 1 ) as DB2_DATA ptr
   return this.dbSeek(sVariableName, DB2_VARIABLE, Type4LookFor, SequentialIndex)      
end function

''
''
function clsDB2.dbFindTYPE( byref sTypeName as string, byval Type4LookFor as eLookFor= ELF_CurrentProject, byval SequentialIndex as Long = 1 ) as DB2_DATA ptr
   dim pData as DB2_DATA ptr       
   pData = this.dbSeek(sTypeName, DB2_TYPE, Type4LookFor, SequentialIndex) 
   if pData = 0 then return 0
   ' If this Type is an ALIAS for another Type then get that real type
   if WStrIsEqual(pData->ElementName, pData->ElementValue, false)=false then
      return this.dbSeek(pData->ElementValue, DB2_TYPE, Type4LookFor, SequentialIndex) 
   END IF   
   return pData
end function

''
''
function clsDB2.dbFilenameExists Overload( byref wszFilename as wstring ) as boolean
   ' Search to see if the incoming filename has already been parsed.
   dim as File_DATA Ptr FirstNode4File=any
   if m_ptrRewindData4Proj then
      FirstNode4File=m_ptrRewindData4Proj->FirstNode4File
      while FirstNode4File<>0
         if WStrIsEqual(wszFilename, FirstNode4File->fileName, false)=true THEN
            return true
         end if
         FirstNode4File=FirstNode4File->NextNode
      wend
   endif
   return false
end function

''
''
function clsDB2.dbFilenameExists Overload(byref wszProjName as wstring, byref wszFilename as wstring, byval ProjIndex as long=-1 ) as boolean
   ' Search to see if the incoming filename has already been parsed.
   dim as File_DATA Ptr FirstNode4File=any
   dim as Project_DATA ptr ptrData4Proj=dbFind4Proj(wszProjName, ProjIndex)
   if ptrData4Proj then
      FirstNode4File=ptrData4Proj->FirstNode4File
      while FirstNode4File<>0
         if WStrIsEqual(wszFilename, FirstNode4File->fileName, false)=true THEN
            return true
         end if
         FirstNode4File=FirstNode4File->NextNode
      wend
   endif
   return false
end function

'''
'''
function clsDB2.dbDebug() as long
   dim pStream as CTextStream
   pStream.Create("_debug.txt")

   dim pData as DB2_DATA ptr
    

   ' PRINT ALL WINAPI/WINFBX FUNCTION DEFINITIONS
   pStream.WriteLine "WINAPI/WINFBX FUNCTIONS"

   dbRewind4DB()
   do 
      pData = this.dbGetNext4DB
      if pData = 0 THEN exit do
      'if pData->projectIndex <> -1 THEN continue do
      if pData->deleted = true THEN continue do
      'if pData->id <> DB2_FUNCTION then continue do
      pStream.WriteLine "   fileName: " & pData->pFileDATA->fileName
      pStream.WriteLine "   ElementName: " & pData->ElementName
      pStream.WriteLine "   ElementValue: " & pData->ElementValue
      pStream.WriteLine ""
   loop
   pStream.WriteLine ""
   
   pStream.Close
   function = 0
end function


dim shared gdb2 as clsDB2

