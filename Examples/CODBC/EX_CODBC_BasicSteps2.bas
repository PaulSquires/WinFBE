'#CONSOLE ON
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Error callback procedure
SUB ODBC_ErrorCallback (BYVAL nResult AS SQLRETURN, BYREF wszSrc AS WSTRING, BYREF wszErrorMsg AS WSTRING)
   PRINT "Error: " & STR(nResult) & " - Source: " & wszSrc
   IF LEN(wszErrorMsg) THEN PRINT wszErrorMsg
END SUB

' // Create a connection object and connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

' // Set the address of the error callback for the connection object
pDbc.SetErrorProc(@ODBC_ErrorCallback)

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Unable to create the statement handle" : SLEEP : END

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors ORDER BY Author")

' // Parse the result set
DIM cwsOutput AS CWSTR
DO
   ' // Fetch the record
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ' // Get the values of the columns and display them
   cwsOutput = ""
   cwsOutput += pStmt.GetData(1) & " "
   cwsOutput += pStmt.GetData(2) & " "
   cwsOutput += pStmt.GetData(3)
   PRINT cwsOutput
LOOP

PRINT
PRINT "Press any key..."
SLEEP
