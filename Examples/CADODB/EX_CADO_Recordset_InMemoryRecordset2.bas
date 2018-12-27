' ========================================================================================
' ADO test: In-memory recordset
' ========================================================================================
#include "Afx/CADODB/CADODB.inc"
using Afx

' // Create an instance of the CAdoRecorset class
DIM pRecordset AS CAdoRecordset
' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields

pFields.Append("Key", adVarChar, 10)
pFields.Append("Item", adVarChar, 20)

pRecordset.CursorLocation = adUseClient
pRecordset.Open(adOpenKeyset, adLockOptimistic)

pRecordset.AddNew
   pRecordset.Collect("Key") = "One"
   pRecordset.Collect("Item") = "Item one"

pRecordset.AddNew
   pRecordset.Collect("Key") = "Two"
   pRecordset.Collect("Item") = "Item two"

pRecordset.AddNew
   pRecordset.Collect("Key") = "Three"
   pRecordset.Collect("Item") = "Item three"

' // Start searching from the first record (adBookmarkFirst)
' // You can also start from the current record (the default value, adBookmarkCurrent)
' // or from the last record (adBookmarkLast). The default search direction is
' // adSearchForward. If you want to perform a backwards search, you will have to
' // specify adSearchBackward in the search direction parameter.
IF pRecordset.Find("Key = 'Two'", adBookmarkFirst) = S_OK THEN
   AfxMsg pRecordset.Collect("Item").ToStr
ELSE
   AfxMsg "Record not found"
END IF

PRINT
PRINT "Press any key..."
SLEEP
