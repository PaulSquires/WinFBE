'#CONSOLE ON
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Connect with the database
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr

' // Allocate an statement object
DIM pStmt AS COdbcStmt = @pDbc
IF pStmt.Handle = NULL THEN PRINT "Failed to allocate the statement handle": SLEEP : END

DIM cbbytes AS LONG
DIM szCatalogName AS ZSTRING * 256
DIM szSchemaName AS ZSTRING * 256
DIM szTableName AS ZSTRING * 129
DIM szTableType AS ZSTRING * 129
DIM szRemarks AS ZSTRING * 256

' // Note: Despite calling the SQLTablesW function, the GetTables
' // method fails if we use unicode variables instead of ansi.
szTableType = "TABLE"
pStmt.GetTables(szCatalogName, LEN(szCatalogName), szSchemaName, LEN(szSchemaName), _
   szTableName, LEN(szTableName), szTableType, LEN(szTableType))
print pStmt.GetErrorInfo

pStmt.BindCol(1, @szCatalogName, SIZEOF(szCatalogName), @cbbytes)
pStmt.BindCol(2, @szSchemaName, SIZEOF(szSchemaName), @cbbytes)
pStmt.BindCol(3, @szTableName, SIZEOF(szTableName), @cbbytes)
pStmt.BindCol(4, @szTableType, SIZEOF(szTableType), @cbbytes)
pStmt.BindCol(5, @szRemarks, SIZEOF(szRemarks), @cbbytes)
DO
   IF pStmt.Fetch = FALSE THEN EXIT DO
   PRINT "szCatalogName: "; szCatalogName
   PRINT "szSchemaName: "; szSchemaName
   PRINT "szTableName: "; szTableName
   PRINT "szTableType: "; szTableType
   PRINT "szRemarks: "; szRemarks
   PRINT
   PRINT "Press any key..."
   SLEEP
   CLS
LOOP

PRINT
PRINT "Press any key..."
SLEEP
