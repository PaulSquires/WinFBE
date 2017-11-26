'#CONSOLE ON
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor
' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Retrieve the number of columns
DIM numCols AS SHORT = pStmt.NumResultCols
PRINT "Number of columns:" & STR(numCols)
IF numCols = 0 THEN PRINT "There are no columns": SLEEP : END
' // Retrieve the names of the fields (columns)
FOR idx AS LONG = 1 TO numCols
   PRINT "Field #" & STR(idx) & " name: " & pStmt.ColName(idx)
NEXT
' // Parse the result set
DO
   ' // Fetch the record
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   FOR idx AS LONG = 1 TO numCols
      PRINT pStmt.GetData(idx)
   NEXT
LOOP

PRINT
PRINT "Press any key..."
SLEEP
