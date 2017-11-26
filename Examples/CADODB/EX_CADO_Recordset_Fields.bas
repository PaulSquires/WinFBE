' ========================================================================================
' ADO test
' ========================================================================================

'#CONSOLE ON
#include "Afx/CADODB/CADODB.inc"
USING Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CVAR = "SELECT * FROM Publishers ORDER BY PubID"
DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents of the fields
   ' // Note: Instead of pField.Attach / pField.Value, we can use pRecordset.Collect
   DIM pField AS CAdoField
   pField.Attach(pFields.Item("PubID"))
   DIM cvRes1 AS CVAR = pField.Value
   pField.Attach(pFields.Item("Name"))
   DIM cvRes2 AS CVAR = pField.Value
   pField.Attach(pFields.Item("Company Name"))
   DIM cvRes3 AS CVAR = pField.Value
   PRINT cvRes1; " "; cvRes2; " "; cvRes3
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP

PRINT
PRINT "Press any key..."
SLEEP
