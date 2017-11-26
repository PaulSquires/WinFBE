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
DIM nCount AS LONG = pFields.Count
IF nCount THEN
   PRINT
   PRINT "Nullable fields:"
   PRINT "================"
   PRINT
   FOR i AS LONG = 0 TO nCount - 1
      DIM pField AS CAdoField
      pField.Attach(pFields.Item(i))
      ' // Get the attributes of the field
      DIM lAttr AS LONG = pField.Attributes
      ' // Display fields that are nullable
      IF (lAttr AND adFldIsNullable) = adFldIsNullable THEN
         PRINT "Field name: "; pField.Name
      END IF
   NEXT
END IF

PRINT
PRINT "Press any key..."
SLEEP
