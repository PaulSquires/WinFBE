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
pStmt.ExecDirect ("SELECT TOP 20 * FROM Authors ORDER BY Author")

' -------------------------------------------------------------------------------------
' Use DescribeCol to retrieve information about column 2
' -------------------------------------------------------------------------------------
DIM wszColName AS WSTRING * 260
DIM iNameLength AS SHORT
DIM iDataType AS SHORT
DIM dwColumnSize AS DWORD
DIM iDecimalDigits AS SHORT
DIM iNullable AS SHORT

pStmt.DescribeCol(2, @wszColName, 260, @iNameLength, @iDataType, @dwColumnSize, @iDecimalDigits, @iNullable)

? "Column name: " & wszColName
? "Name length: " & STR(iNameLength)
? "Data type: " & STR(iDataType)
? "Column size: " & STR(dwColumnSize)
? "Decimal digits: " & STR(iDecimalDigits)
? "Nullable: " & STR(iNullable) & " - " & IIF(iNullable, "TRUE", "FALSE")

PRINT
PRINT "Press any key..."
SLEEP
