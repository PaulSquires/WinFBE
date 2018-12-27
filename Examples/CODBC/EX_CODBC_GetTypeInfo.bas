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
DIM wszTypeName AS WSTRING * 129
DIM iDataType AS SHORT
DIM lColumnSize AS LONG
DIM wszIntervalPrefix AS WSTRING * 129
DIM wszIntervalSuffix AS WSTRING * 129
DIM wszCreateParams AS WSTRING * 129
DIM iNullable AS SHORT
DIM iCaseSensitive AS SHORT
DIM iSearchable AS SHORT
DIM iUnsignedAttribute AS SHORT
DIM iFixedPrecScale AS SHORT
DIM iAutoUniqueValue AS SHORT
DIM wszLocalTypeName AS WSTRING * 129
DIM iMinimumScale AS SHORT
DIM iMaximumScale AS SHORT
DIM iSqlDataType AS SHORT
DIM iSqlDatetimeSub AS SHORT
DIM lNumPrecRadix AS LONG
DIM iIntervalPrecision AS SHORT

pStmt.GetTypeInfo(SQL_ALL_TYPES)
pStmt.BindCol( 1, @wszTypeName, SIZEOF(wszTypeName), @cbbytes)
pStmt.BindCol( 2, @iDataType, @cbbytes)
pStmt.BindCol( 3, @lColumnSize, @cbbytes)
pStmt.BindCol( 4, @wszIntervalPrefix, SIZEOF(wszIntervalPrefix), @cbbytes)
pStmt.BindCol( 5, @wszIntervalSuffix, SIZEOF(wszIntervalSuffix), @cbbytes)
pStmt.BindCol( 6, @wszCreateParams, SIZEOF(wszCreateParams), @cbbytes)
pStmt.BindCol( 7, @iNullable, @cbbytes)
pStmt.BindCol( 8, @iCasesensitive, @cbbytes)
pStmt.BindCol( 9, @iSearchable, @cbbytes)
pStmt.BindCol(10, @iUnsignedAttribute, @cbbytes)
pStmt.BindCol(11, @iFixedPrecScale, @cbbytes)
pStmt.BindCol(12, @iAutoUniqueValue, @cbbytes)
pStmt.BindCol(13, @wszLocalTypeName, SIZEOF(wszLocalTypeName), @cbbytes)
pStmt.BindCol(14, @iMinimumScale, @cbbytes)
pStmt.BindCol(15, @iMaximumScale, @cbbytes)
pStmt.BindCol(16, @iSqlDataType, @cbbytes)
pStmt.BindCol(17, @iSqlDateTimeSub, @cbbytes)
pStmt.BindCol(18, @lNumPrecRadix, @cbbytes)
pStmt.BindCol(19, @iIntervalPrecision, @cbbytes)

DO
   IF pStmt.Fetch = FALSE THEN EXIT DO
   ? "Type name: " & wszTypeName
   ? "Data type: " & STR(iDataType)
   ? "Column size: " & STR(lColumnSize)
   ? "Interval prefix: " & wszIntervalPrefix
   ? "Interval suffix: " & wszIntervalSuffix
   ? "Create params: " & wszCreateParams
   ? "Nullable: " & STR(iNullable)
   ? "Case sensitive: " & STR(iCaseSensitive)
   ? "Searchable: " & STR(iSearchable)
   ? "Unsigned attribute: " & STR(iUnsignedAttribute)
   ? "Fixed prec scale: " & STR(iFixedPrecScale)
   ? "Auto unique value: " & STR(iAutoUniqueValue)
   ? "Local type name: " & wszLocalTypeName
   ? "Minimum scale: " & STR(iMinimumScale)
   ? "Maximum scale: " & STR(iMaximumScale)
   ? "SQL data type: " & STR(iSqlDataType)
   ? "SQL Datetime sub: " & STR(iSqlDatetimeSub)
   ? "Num prec radix: " & STR(lNumPrecRadix)
   ? "Interval precision: " & STR(iIntervalPrecision)
   PRINT
   PRINT "Press any key..."
   SLEEP
   CLS
LOOP

PRINT
PRINT "Press any key..."
SLEEP
