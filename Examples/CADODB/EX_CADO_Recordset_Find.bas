' ########################################################################################
' Microsoft Windows
' Contents: ADO - Find example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates the use of the Find method.
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
DIM cvSource AS CVAR = "SELECT * FROM Publishers ORDER By PubID"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

pRecordset.Find "PubID = #70#"
DIM cvRes1 AS CVAR = pRecordset.Collect("PubID")
DIM cvRes2 AS CVAR = pRecordset.Collect("Name")
PRINT cvRes1 & " " & cvRes2

PRINT
PRINT "Press any key..."
SLEEP
