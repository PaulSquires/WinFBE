#include once "token.bi"
#include once "scanner.bi"
#include once "parser.bi"

dim as string	inFileName, outFileName

cls

if( __FB_ARGC__ = 2 ) then
	inFileName = command(1)
	dim as integer	dotPos = inStrRev( inFileName, "." ) 
	outFileName = left( inFileName, dotPos ) + "json"
elseif( __FB_ARGC__ = 3 ) then
		inFileName = command(1)
	outFileName = command(2)
end if

if( len( dir( inFileName ) ) < 1 ) then
	print "File " + inFileName + " isn't exist! Exit~"
	sleep
	end
end if
/'
inFileName = "main.bas"
outFileName = "main.json"
'/
dim as Scanner				tokenScanner
dim as DynamicArray ptr		list = tokenScanner.scanFile( inFileName )	' Save the Scanned tokens
dim as CParser				parser = CParser( list )					' Create the parse

parser.parse( inFileName )												' Start parse!
'parser.printAST()
parser.save( outFileName )


print "****************************************************"
print "Press Any Key to Continue..."

sleep
end