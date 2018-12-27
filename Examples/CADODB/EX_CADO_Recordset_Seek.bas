' ########################################################################################
' Microsoft Windows
' Contents: ADO - Seek example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates the use of the Seek method.
' Note: Error checking ommited for brevity.
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define UNICODE
#include "Afx/CADODB/CADODB.inc"
USING Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Set the cursor location
DIM pRecordset AS CAdoRecordset
pRecordset.CursorLocation = adUseServer

' // Open the recordset
DIM cvSource AS CVAR = "Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdTableDirect)

' // Set the index
pRecordset.Index = "PrimaryKey"

' // Seek the record 70
pRecordset.Seek 70, 1

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the contents
   DIM cvRes1 AS CVAR = pRecordset.Collect("PubID")
   DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
   DIM cvRes3 AS CVAR = pRecordset.Collect("Company Name")
   PRINT cvRes1 & " " & cvRes2 & " " & cvRes3
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP


PRINT
PRINT "Press any key..."
SLEEP
