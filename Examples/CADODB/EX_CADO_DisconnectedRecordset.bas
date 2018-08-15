' ########################################################################################
' Microsoft Windows
' Contents: ADO - Disconnectewd recordset example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates how to create a disconnected recordset.
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
DIM pConnection AS CAdoConnection PTR = NEW CAdoConnection
pConnection->Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
' // Setting the cursor location to client side is important to get a disconnected recordset
pRecordset.CursorLocation = adUseClient
' // Open the recordset
DIM cvSource AS CVAR = "SELECT * FROM Authors"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Disconnect the recordset by setting its active connection to null.
' // Casting to Afx_ADOConnection PTR is needed to get the correct overloaded method called;
' // otherwise, the CVAR version will be called and it will fail.
pRecordset.ActiveConnection = cast(Afx_ADOConnection PTR, NULL)

' // Close and release the connection
Delete pConnection

' // Parse the recordset
DO
   ' // While not at the end of the recordset...
   IF pRecordset.EOF THEN EXIT DO
   ' // Get the content of the "Author" column
'   DIM cvRes AS CVARIANT = pRecordset.Collect("Author")
'   PRINT cvRes
   PRINT pRecordset.Collect("Author")
   ' // Fetch the next row
   IF pRecordset.MoveNext <> S_OK THEN EXIT DO
LOOP

PRINT
PRINT "Press any key..."
SLEEP
