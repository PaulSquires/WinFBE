set FBLIBDIR=win32
set FBCPATH=%~d0\PSS121\FB\WinFBE_Suite\Toolchains\FreeBASIC-1.10.0-winlibs-gcc-9.3.0

cd src
%FBCPATH%\fbc32 WinFBE.bas -x ..\WinFBE32.exe WinFBE.rc -s gui
cd ..
