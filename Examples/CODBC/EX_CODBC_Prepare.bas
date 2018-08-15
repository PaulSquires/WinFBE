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

' // Prepare the statement
pStmt.Prepare("UPDATE Authors SET Author=? WHERE Au_ID=?")
' // Bind the columns
DIM wszAuthor AS WSTRING * 256, cbAuthor AS LONG
DIM lAuId AS LONG, cbAuId AS LONG
pStmt.BindParameter(1, SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 255, 0, @wszAuthor, SIZEOF(wszAuthor), @cbAuthor)
pStmt.BindParameter(2, SQL_PARAM_INPUT, SQL_C_LONG, SQL_INTEGER, 4, 0, @lAuId, 0, @cbAuID)
' // Fill the parameter value
lAuId = 999 : cbAuId = 4
wszAuthor = "William Shakespeare" : cbAuthor = LEN(wszAuthor)
' // Execute the prepared statement
pStmt.Execute
IF pStmt.Error = FALSE THEN  PRINT "Record updated" ELSE PRINT pStmt.GetErrorInfo

PRINT
PRINT "Press any key..."
SLEEP
