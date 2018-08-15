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

' // Parse the Properties collection
DIM pProperties AS CAdoProperties = pRecordset.Properties
DIM nCount AS LONG = pProperties.Count
FOR i AS LONG = 0 TO nCount - 1
   DIM pProperty AS CAdoProperty = pProperties.Item(i)
   PRINT "Property name: "; pProperty.Name; " - Attributes: "; WSTR(pProperty.Attributes)
NEXT

PRINT
PRINT "Press any key..."
SLEEP
