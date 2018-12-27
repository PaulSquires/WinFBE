' ########################################################################################
' Microsoft Windows
' Contents: ADO - Delete_ example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates the use of the Delete_ method.
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
DIM cvSource AS CVAR = "SELECT * FROM Publishers WHERE PubID=10000"
pRecordset.Open(cvSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

DIM cvRes AS CVAR = pRecordset.Collect("PubID")
IF cvRes.ValInt = 10000 THEN
   pRecordset.Delete_ adAffectCurrent
   PRINT "Record deleted"
ELSE
   PRINT "Record not found"
END IF

' We can also use:
'IF pRecordset.Collect("PubID").ValInt = 10000 THEN
'   pRecordset.Delete_ adAffectCurrent
'   PRINT "Record deleted"
'ELSE
'   PRINT "Record not found"
'END IF

PRINT
PRINT "Press any key..."
SLEEP
