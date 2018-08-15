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
DIM wszName AS WSTRING * 260, StringLength AS SHORT
DIM nType AS SHORT, nSubType AS SHORT, nLength AS LONG
DIM nPrecision AS SHORT, nScale AS SHORT, nNullable AS SHORT

pStmt.GetImpRowDescRec(9, @wszName, 260, @StringLength, @nType, @nSubType, @nLength, @nPrecision, @nScale, @nNullable)

PRINT "Name: " & wszName
PRINT "String length: " & STR(StringLength)
PRINT "Type: " & STR(nType)
PRINT "Subtype: " & STR(nSubType)
PRINT "Length: " & STR(nLength)
PRINT "Precision: " & STR(nPrecision)
PRINT "Scale: " & STR(nScale)
PRINT "Nullable: " & STR(nNullable)

PRINT
PRINT "Press any key..."
SLEEP
