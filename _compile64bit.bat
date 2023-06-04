set FBLIBDIR=win64
set FBCPATH=%~d0\PSS121\FB\WinFBE_Suite\Toolchains\FreeBASIC-1.10.0-winlibs-gcc-9.3.0

cd src
%FBCPATH%\fbc64 WinFBE.bas -x ..\WinFBE64.exe WinFBE.rc -s gui -gen gcc -O 2
cd ..
