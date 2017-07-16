''
''  Application in-memory database
''

const DB_DELIM = "|"

const DB_FUNCTION_NAME    = 1
const DB_TYPE_NAME        = 2
const DB_TODO_ITEM        = 3
const DB_FILENAME_PARSE   = 4
const DB_DIM_SHARED       = 5

type DBENGINE_DATA
   deleted      as BOOLEAN = true  ' True/False
   dataType     as LONG            ' See below
   projectIndex as long            ' idx of project
   nLineNum     as long            ' Location in the file where found
   diskFilename as string          ' Full disk filename
   data1        as string          ' Value dependant on DataType
   data2        as string          ' Value dependant on DataType
   idUnicode    as string          ' U, A, B  (unicode, ansi, both)
END TYPE

TYPE clsDBEngine
   private:
      m_arrData(any) as DBENGINE_DATA
      m_index as LONG
      
   public:   
      declare function dbAdd( byref DiskFilename as string, _
                              byref projectIndex as long, _
                              byref DataType as long, _
                              byref nLineNum as long, _
                              byref Data1    as string, _
                              byref Data2    as string, _
                              byref idUnicode    as string = "" _
                              ) as LONG
      declare function dbDelete( byref Diskfilename as string ) as long
      declare function dbDeleteByProject( byval idx as long ) as boolean
      declare function dbRewind() as LONG
      declare function dbGetNext() as DBENGINE_DATA ptr
      declare function dbFindFunctionName( byref sFunctionName as string ) as DBENGINE_DATA ptr
      declare function dbFindVariableName( byref sVariableName as string ) as DBENGINE_DATA ptr
      declare function dbFindTypeName( byref sTypeName as string ) as DBENGINE_DATA ptr
      declare function dbFilenameExists( byref sFilename as string ) as boolean

      declare function dbDebug() as long
      declare function dbStats() as long
      declare constructor
END TYPE


' Function names
'    DataType = DB_FUNCTION_NAME 
'    nLineNum = Line number in file being parsed.
'    Data1    = Function name.
'    Data2    = Function declare parameter list.

' TYPE
'    DataType = DB_TYPE_NAME
'    Data1    = Type name.
'    Data2    = Deliminated Type elements and sub/functions.

' ToDo item
'    DataType = DB_TODO_ITEM
'    nLineNum = Line number in file being parsed.
'    Data1    = Description.
'    Data2    = Not used.

' Shared global variable
'    DataType = DB_DIM_SHARED
'    Data1    = Variable name.
'    Data2    = TYPE name related to variable. 

''
''
constructor clsDBEngine
   redim m_arrData(500) as DBENGINE_DATA
   this.dbRewind
end constructor

''
''
function clsDBEngine.dbAdd( byref DiskFilename as string, _
                            byref projectIndex as long, _
                            byref DataType     as long, _
                            byref nLineNum     as long, _
                            byref Data1        as string, _
                            byref Data2        as string, _
                            byref idUnicode    as string = "" _
                            ) as LONG
   
   dim as long nSlot = -1

   ' Find a slot to put the data in
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).deleted THEN
         nSlot = i: exit for
      END IF   
   NEXT
   
   ' If no slot found then resize the array to make room
   if nSlot = -1 THEN
      nSlot = ubound(m_arrData) + 1
      redim preserve m_arrData(nSlot+100) as DBENGINE_DATA
   END IF
   
   with m_arrData(nSlot)
      .deleted      = false
      .diskFilename = DiskFilename
      .dataType     = DataType
      .projectIndex = projectIndex
      .nLineNum     = nLineNum
      .data1        = Data1
      .data2        = Data2
      .idUnicode    = idUnicode
   end with
   
   function = nSlot               
end function                

''
''
function clsDBEngine.dbDelete( byref Diskfilename as string ) as long
   dim nCount as long 
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).deleted = true then continue for
      if ucase(m_arrData(i).diskFilename) = ucase(DiskFilename) THEN
         m_arrData(i).deleted = true
         nCount = nCount + 1
      END IF   
   NEXT
   function = nCount
end function

''
''
function clsDBEngine.dbDeleteByProject( byval idx as long ) as boolean
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      if m_arrData(i).projectIndex = idx THEN
         m_arrData(i).deleted = true
         function = true
      END IF
   NEXT
end function

''
''
function clsDBEngine.dbRewind() as long
   ' Set index pointer to immediately before first array index 
   m_index = lbound(m_arrData) - 1
   function = 0
END FUNCTION

''
''
function clsDBEngine.dbGetNext() as DBENGINE_DATA ptr
   ' Set index pointer to next array index that is not deleted
   dim as long ub = ubound(m_arrData)
   do
      m_index = m_index + 1
      if m_index > ub THEN return 0
      if m_arrData(m_index).deleted THEN 
         continue do
      else
         exit do   
      END IF
   loop
   function = @m_arrData(m_index)
END FUNCTION

''
''
function clsDBEngine.dbFindFunctionName( byref sFunctionName as string ) as DBENGINE_DATA ptr
   ' Find the array element that contains the function name beng searched for
   dim pData as DBENGINE_DATA ptr
   this.dbRewind()
   do 
      pData = this.dbGetNext()
      if pData = 0 THEN exit do
      if pData->dataType = DB_FUNCTION_NAME THEN
         if _stricmp(sFunctionName, pData->data1) = 0 THEN
            return @m_arrData(m_index)
         END IF
      END IF
   loop
   function = 0
end function

''
''
function clsDBEngine.dbFindVariableName( byref sVariableName as string ) as DBENGINE_DATA ptr
   ' Find the array element that contains the variable name beng searched for
   dim pData as DBENGINE_DATA ptr
   this.dbRewind()
   do 
      pData = this.dbGetNext()
      if pData = 0 THEN exit do
      if pData->dataType = DB_DIM_SHARED THEN
         if _stricmp(sVariableName, pData->data1) = 0 THEN
            return @m_arrData(m_index)
         END IF
      END IF
   loop
   function = 0
end function

''
''
function clsDBEngine.dbFindTypeName( byref sTypeName as string ) as DBENGINE_DATA ptr
   ' Find the array element that contains the variable name beng searched for
   dim pData as DBENGINE_DATA ptr
   this.dbRewind()
   do 
      pData = this.dbGetNext()
      if pData = 0 THEN exit do
      if pData->dataType = DB_TYPE_NAME THEN
         if _stricmp(sTypeName, pData->data1) = 0 THEN
            return @m_arrData(m_index)
         END IF
      END IF
   loop
   function = 0
end function

''
''
function clsDBEngine.dbFilenameExists( byref sFilename as string ) as boolean
   ' Search to see if the incoming filename has already been parsed.
   dim pData as DBENGINE_DATA ptr
   this.dbRewind()
   do 
      pData = this.dbGetNext()
      if pData = 0 THEN exit do
      if _stricmp(sFilename, pData->diskFilename) = 0 THEN
         return true
      END IF
   loop
end function

''
''
function clsDBEngine.dbDebug() as long
   dim as long f = freefile
   open "_debug.txt" for output as #f
   dim pData as DBENGINE_DATA ptr
   
   this.dbRewind()
   do 
      pData = this.dbGetNext
      if pData = 0 THEN exit do
         
      'print #f, "index: "; i
      print #f, "deleted: "; pData->deleted
      print #f, "diskFilename: "; pData->diskFilename
      print #f, "projectIndex: "; pData->projectIndex
      print #f, "dataType: "; pData->dataType
      print #f, "lineNum: "; pData->nLineNum
      print #f, "data1: "; pData->data1
      print #f, "data2: "; pData->data2
      print #f, ""
   loop

   close #f
   function = 0
end function

''
''
function clsDBEngine.dbStats() as long
   dim as long numTotal, numDeleted, numActive 
   dim as long numKeywords, numFunctions, numTYPEs
   dim as long numToDo, numFilenames
   
   for i as long = lbound(m_arrData) to ubound(m_arrData)
      numTotal += 1
      if m_arrData(i).deleted THEN 
         numDeleted += 1
      else
         numActive += 1
         select CASE m_arrData(i).dataType
            case DB_FUNCTION_NAME 
               if m_arrData(i).projectIndex = -1 then
                  numKeywords += 1
               else   
                  numFunctions += 1
               end if   
            case DB_TYPE_NAME     :  numTYPEs += 1  
            case DB_TODO_ITEM     :  numToDo += 1 
            case DB_FILENAME_PARSE:  numFilenames += 1 
            CASE ELSE
         END SELECT
      END IF
   next

   print "gdb stats"
   print "numTotal = "; numTotal
   print "numDeleted = "; numDeleted
   print "numActive = "; numActive
   print "numKeywords = "; numKeywords
   print "numFunctions = "; numFunctions
   print "numTYPEs = "; numTYPEs
   print "numToDo = "; numToDo
   print "numFilenames = "; numFilenames

   function = 0
END FUNCTION


dim shared gdb as clsDBEngine


