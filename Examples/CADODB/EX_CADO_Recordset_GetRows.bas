' ########################################################################################
' Microsoft Windows
' Contents: ADO - GetRows example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates the use of the GetRows method.
' Note: Error checking ommited for brevity.
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define UNICODE
#include "Afx/CADODB/CADODB.inc"
#include "Afx/CSafeArray.inc"
USING Afx

' // Open the connection
DIM pConnection AS CAdoConnection
pConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb"

' // Open the recordset
DIM pRecordset AS CAdoRecordset
DIM cbsSource AS CBSTR = "SELECT TOP 20 * FROM Publishers ORDER BY Name"
DIM hr AS HRESULT = pRecordset.Open(cbsSource, pConnection, adOpenKeyset, adLockOptimistic, adCmdText)

' // GetRows returns a pointer to a two-dimensional safe array
' // that we are going to attach to an instance of the CSafeArray class
DIM csa AS CSafeArray = CSafeArray(pRecordset.GetRows, TRUE)

' // Print the contents of the safe array
FOR j AS LONG = csa.LBound(2) TO csa.UBound(2)
   FOR i AS LONG = csa.LBound(1) TO csa.UBound(1)
      PRINT csa.GetVar(i, j)
   NEXT
   PRINT
NEXT

PRINT
PRINT "Press any key..."
SLEEP
