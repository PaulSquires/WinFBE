' ########################################################################################
' Microsoft Windows
' Contents: ADO - Save example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates the use of the Save method.
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

' // Create an scoped instance of the CAdoRecordset class
SCOPE
   DIM pRecordset AS CAdoRecordset
   ' // Set the cursor location
   pRecordset.CursorLocation = adUseClient

   ' // Open the recordset
   DIM cvSource AS CVAR = "SELECT * FROM Publishers"
   pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

   ' // Save the recordset as XML
   DIM wszFileName AS WSTRING * MAX_PATH = "Publishers.dat"
   IF AfxFileExists(wszFileName) THEN KILL wszFileName
   IF pRecordset.Save(wszFileName, adPersistADTG) = S_OK THEN
      PRINT "Recordset saved"
   ELSE
      PRINT "Save failed"
   END IF
END SCOPE

IF AfxFileExists("Publishers.dat") THEN
   ' // Reopen the recordset
   DIM pRecordset AS CAdoRecordset
   DIM cvSource AS CVAR = "Publishers.dat"
   DIM hr AS HRESULT = pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdFile)
   IF hr = S_OK THEN
      ' // Parse the recordset
      DO
         ' // While not at the end of the recordset...
         IF pRecordset.EOF THEN EXIT DO
         ' // Get the contents
         DIM cvRes1 AS CVAR = pRecordset.Collect("PubID")
         DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
         DIM cvRes3 AS CVAR = pRecordset.Collect("Company Name")
         PRINT cvRes1 + " " + cvRes2 + " " + cvRes3
         ' // Fetch the next row
         IF pRecordset.MoveNext <> S_OK THEN EXIT DO
      LOOP
   END IF
END IF

PRINT
PRINT "Press any key..."
SLEEP
