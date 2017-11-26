' ########################################################################################
' Microsoft Windows
' Contents: ADO - Save example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates the use of the Save method to save a recordset as XML.
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

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cvSource AS CBSTR = "SELECT * FROM Publishers"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // Save the recordset as XML
DIM wszFileName AS WSTRING * MAX_PATH = "Publishers.xml"
IF AfxFileExists(wszFileName) THEN KILL wszFileName
IF pRecordset.Save(wszFileName, adPersistXML) = S_OK THEN
   PRINT "Recordset saved"
ELSE
   PRINT "Save failed"
END IF

PRINT
PRINT "Press any key..."
SLEEP
