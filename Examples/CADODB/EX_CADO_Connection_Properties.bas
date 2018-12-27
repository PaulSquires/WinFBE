' ########################################################################################
' Microsoft Windows
' Contents: ADO - Properties collection example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates the use of the Properties collection.
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

' // Create an instance of the CAdoProperties class
' // with a reference to the Peoperties collection.
DIM pProperties AS CAdoProperties = pConnection.Properties
PRINT "Number of properties: "; pProperties.Count

' // Create an instance of the CAdoProperty class
' // with a reference to a Property object.
DIM pProperty AS CAdoProperty = pProperties.Item("DBMS Version")

' // Print the value of the property
PRINT "DBMS version : "; pProperty.Value

' // Alternate version using a compound syntax
'PRINT CAdoProperty(CAdoProperties(pConnection.Properties).Item("DBMS Version")).Value

PRINT
PRINT "Press any key..."
SLEEP
