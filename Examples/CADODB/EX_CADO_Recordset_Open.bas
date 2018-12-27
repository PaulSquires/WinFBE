' ========================================================================================
' ADO test
' ========================================================================================

'#CONSOLE ON
#include "Afx/CADODB/CADODB.inc"
USING Afx

' // Create a Connection object
DIM pConnection AS CAdoConnection
' // Create a Recordset object
DIM pRecordset AS CAdoRecordset

'/////////////////////////////////////////////////////////////////////////////////
'/// Connection
'/////////////////////////////////////////////////////////////////////////////////

' // Open the connection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' --> Alternate way <--
'DIM cvConStr AS CVAR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"
'pConnection.Open cvConStr

'/////////////////////////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////////////////////////
'/// Open recordset
'/////////////////////////////////////////////////////////////////////////////////

' // Open the recordset
DIM cvSource AS CVAR = "SELECT TOP 20 * FROM Authors ORDER BY Author"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

'/////////////////////////////////////////////////////////////////////////////////

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
'   DIM cvRes AS CVA = pRecordset.Collect("Author")
'   PRINT cvRes
   PRINT pRecordset.Collect("Author")
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP

' // Close the recordset
'IF pRecordset.State = adStateOpen THEN pRecordset.Close
'' // Close the connection
'IF pConnection.State = adStateOpen THEN pConnection.Close

' The recordset and connection will be closed by the destructors of the classes.

PRINT
PRINT "Press any key..."
SLEEP
