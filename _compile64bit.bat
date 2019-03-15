set FBLIBDIR=win64
set FBCPATH=X:\PSS121\FB\WinFBE_Suite\FreeBASIC-1.06.0

cd src
%FBCPATH%\fbc64 WinFBE.bas -x ..\WinFBE64.exe WinFBE.rc -s gui
cd ..


