' ========================================================================================
' ADO test : In-memory recordset
' ========================================================================================

'#CONSOLE ON
#include "Afx/CADODB/CADODB.inc"
USING Afx

' // Create an instance of the CAdoRecorset class
DIM pRecordset AS CAdoRecordset
' // Get a reference to the Fields collection
DIM pFields AS CAdoFields = pRecordset.Fields

pFields.Append("Key", adVarChar, 10)
pFields.Append("Item", adVarChar, 20)

pRecordset.CursorType = adOpenKeyset
pRecordset.CursorLocation = adUseClient
pRecordset.LockType = adLockOptimistic
pRecordset.Open(adOpenKeyset, adLockOptimistic)

print "Record count: ", pRecordset.Recordcount

pRecordset.AddNew
   pRecordset.Collect("Key") = "One"
   pRecordset.Collect("Item") = "Item one"
'pRecordset.Update
' Don't call Update or it will add an additional empty record

pRecordset.AddNew
   pRecordset.Collect("Key") = "Two"
   pRecordset.Collect("Item") = "Item two"
'pRecordset.Update

pRecordset.AddNew
   pRecordset.Collect("Key") = "Three"
   pRecordset.Collect("Item") = "Item three"
'pRecordset.Update

print "New record count: ", pRecordset.Recordcount


'pRecordset.MoveFirst
pRecordset.AbsolutePosition = 3
DO
   IF pRecordset.EOF THEN EXIT DO
   PRINT pRecordset.Collect("Key")
   PRINT pRecordset.Collect("Item")
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP

pRecordset.MoveFirst
AfxMsg pRecordset.GetString

PRINT
PRINT "Press any key..."
SLEEP
