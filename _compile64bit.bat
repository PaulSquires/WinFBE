set FBLIBDIR=win64
set FBCPATH=x:\FB\FreeBASIC-1.06.0-win64

cd src
%FBCPATH%\fbc WinFBE.bas -x ..\WinFBE64.exe WinFBE.rc -s gui
cd ..

@if ErrorLevel 1 goto compileerror
@echo SUCCESS!! WinFBE (64 bit) built.

goto exit

:compileerror
@echo *** COMPILE ERROR ***

:exit
@pause

