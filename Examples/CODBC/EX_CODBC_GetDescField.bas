'#CONSOLE ON
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

' // Cursor type
pStmt.SetMultiuserKeysetCursor

' // Generate a result set
pStmt.ExecDirect ("SELECT TOP 20 * FROM Titles ORDER BY Title")

' // Get the current setting or values of the descriptor record for the 9th field ("Price")
PRINT pStmt.GetImpRowDescFieldName(9)
PRINT STR(pStmt.GetImpRowDescFieldPrecision(9))
PRINT STR(pStmt.GetImpRowDescFieldOctetLength(9))

PRINT
PRINT "Press any key..."
SLEEP
