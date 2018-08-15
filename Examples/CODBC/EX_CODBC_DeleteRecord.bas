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

' // Bind the columns
DIM AS LONG lAuId, cbAuId
pStmt.BindCol(1, @lAuId, @cbAuId)
DIM wszAuthor AS WSTRING * 256, cbAuthor AS LONG
pStmt.BindCol(2, @wszAuthor, 256, @cbAuthor)
DIM iYearBorn AS SHORT, cbYearBorn AS LONG
pStmt.BindCol(3, @iYearBorn, @cbYearBorn)

' // Generate a result set
pStmt.ExecDirect ("SELECT * FROM Authors WHERE Au_Id=999")

' // Fetch the record
pstmt.Fetch

' // Fill the values of the binded application variables and its sizes
cbAuID = SQL_COLUMN_IGNORE
wszAuthor = "Félix Lope de Vega Carpio"
cbAuthor = LEN(wszAuthor) * 2   ' Unicode uses 2 bytes per character
iYearBorn = 1562
cbYearBorn = SIZEOF(iYearBorn)

' // Delete the record
pStmt.DeleteRecord
IF pStmt.Error = FALSE THEN PRINT "Record deleted" ELSE PRINT pStmt.GetErrorInfo

PRINT
PRINT "Press any key..."
SLEEP
