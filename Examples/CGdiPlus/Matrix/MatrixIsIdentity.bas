' ########################################################################################
' Microsoft Windows
' File: MatrixIsIdentity.bas
' Contents: GDI+ - MatrixIsIdentity example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define UNICODE
#INCLUDE ONCE "Afx/CGdiPlus/CGdiPlus.inc"
USING Afx

' // Must be constructed with NEW to we able to delete it before the call to AfxGdipShutdown
DIM pMatrix AS CGpMatrix PTR = NEW CGpMatrix(1.0, 0.0, 0.0, 1.0, 0.0, 2.0)

IF pMatrix->IsIdentity THEN
   PRINT "The matrix is the identity matrix"
ELSE
   PRINT "The matrix is not the identity matrix"
END IF

' // Must be deleted before the call to AfxGdipShutdown
Delete pMatrix

PRINT
PRINT "Press any key"
SLEEP
