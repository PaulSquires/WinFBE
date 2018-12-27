' ########################################################################################
' Microsoft Windows
' File: MatrixIsInvertible.bas
' Contents: GDI+ - MatrixIsInvertible example
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
DIM pMatrix AS CGpMatrix PTR = NEW CGpMatrix(3.0, 0.0, 0.0, 2.0, 20.0, 10.0)

IF pMatrix->IsInvertible THEN
   PRINT "The matrix is invertible"
ELSE
   PRINT "The matrix is not invertible"
END IF

' // Must be deleted before the call to AfxGdipShutdown
Delete pMatrix

PRINT
PRINT "Press any key"
SLEEP
