'#CONSOLE ON
#include once "Afx/COdbc/COdbc.inc"
USING Afx

' // Create a connection object and connect with the database
' // We need to call it to create the environment handle
DIM wszConStr AS WSTRING * 260 = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=biblio.mdb"
DIM pDbc AS CODBC = wszConStr
IF pDbc.Handle = NULL THEN PRINT "Unable to create the connection handle" : SLEEP : END

DIM wDirection AS WORD
DIM cwsDriverAttributes AS CWSTR
DIM cwsDriverDesc AS CWSTR

wDirection = SQL_FETCH_FIRST
DO
   cwsDriverDesc = ""
   cwsDriverAttributes = ""
   IF SQL_SUCCEEDED(pDbc.GetDrivers(wDirection, cwsDriverDesc, cwsDriverAttributes)) = 0 THEN EXIT DO
   ? "Driver description: " & cwsDriverDesc
   ? "Driver attributes: " & cwsDriverAttributes
   wDirection = SQL_FETCH_NEXT
LOOP

PRINT
PRINT "Press any key..."
SLEEP
