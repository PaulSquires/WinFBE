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
DIM iNonUnique AS SHORT
DIM szIndexQualifier AS ZSTRING * 129
DIM szIndexName AS ZSTRING * 129
DIM iInfoType AS SHORT
DIM iOrdinalPosition AS SHORT
DIM szColumnName AS ZSTRING * 129
DIM szAscOrDesc AS ZSTRING * 2
DIM lCardinality AS LONG
DIM lPages AS LONG
DIM szFilterCondition AS ZSTRING * 129

' // Note: Despite calling the SQLStatisticsW function, the GetStatistics
' // method fails if we use unicode variables instead of ansi.
szTableName = "Authors"
pStmt.GetStatistics(szCatalogName, LEN(szCatalogName), szSchemaName, LEN(szSchemaName), _
   szTableName, LEN(szTableName), SQL_INDEX_ALL, SQL_ENSURE)

print pStmt.GetLastResult
pStmt.BindCol( 1, @szCatalogName, SIZEOF(szCatalogName), @cbBytes)
pStmt.BindCol( 2, @szSchemaName, SIZEOF(szSchemaName), @cbbytes)
pStmt.BindCol( 3, @szTableName, SIZEOF(szTableName), @cbbytes)
pStmt.BindCol( 4, @iNonUnique, @cbbytes)
pStmt.BindCol( 5, @szIndexQualifier, SIZEOF(szIndexQualifier), @cbbytes)
pStmt.BindCol( 6, @szIndexName, SIZEOF(szIndexName), @cbbytes)
pStmt.BindCol( 7, @iInfoType, @cbbytes)
pStmt.BindCol( 8, @iOrdinalPosition, @cbbytes)
pStmt.BindCol( 9, @szColumnName, SIZEOF(szColumnName), @cbbytes)
pStmt.BindCol(10, @szAscOrDesc, SIZEOF(szAscOrDesc), @cbbytes)
pStmt.BindCol(11, @lCardinality, @cbbytes)
pStmt.BindCol(12, @lPages, @cbbytes)
pStmt.BindCol(13, @szFilterCondition, SIZEOF(szFilterCondition), @cbbytes)
DO
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ? "Table catalog name: " & szCatalogName
   ? "Table schema name: " & szSchemaName
   ? "Table name: " & szTableName
   ? "Non unique: " & STR(iNonUnique)
   ? "Index qualifier: " & szIndexQualifier
   ? "Index name: " & szIndexName
   ? "Info type: " & STR(iInfoType)
   ? "Ordinal position: " & STR(iOrdinalPosition)
   ? "Column name: " & szColumnName
   ? "Asc or desc: " & szAscOrDesc
   ? "Cardinality: " & STR(lCardinality)
   ? "Pages: " & STR(lPages)
   ? "Filter condition: " & szFilterCondition
   PRINT
   PRINT "Press any key..."
   SLEEP
   CLS
LOOP

PRINT
PRINT "Press any key..."
SLEEP
