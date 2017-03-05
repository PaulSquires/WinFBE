#include once "parserFB\scanner.bi"


Constructor Scanner()
End Constructor

Constructor Scanner( fullPath as string )

	scanFile( fullPath )
	
End Constructor

destructor Scanner()

	if( tokenList <> 0 ) then delete tokenList

End destructor

sub Scanner.scan( fileData as string )

	if( fileData = "" ) then exit sub
	
	fileData = fileData + chr(10)
	
	dim as boolean		bStringFlag, bCommentFlag, bCommentBlockFlag
	dim as string		identifier
	dim as integer		lineNum = 1, commentCount, fileDataLength = len( fileData ), tokenIndex
	redim as TokenUnit	results(fileDataLength)
	
	for i as integer = 0 to fileDataLength - 1
		if( bCommentBlockFlag = true ) then
			do
				 'Check if /'
				if( chr( fileData[i] ) = "/" AND chr( fileData[i+1] ) = "'" ) then
						commentCount += 1
						'continue do
				elseif( chr( fileData[i] ) = "'" AND chr( fileData[i+1] ) = "/" ) then
						commentCount -= 1
						if( commentCount = 0 ) then
							bCommentBlockFlag = false
							i += 1
							exit do
						end if
				elseif( fileData[i] = 10 ) then
					lineNum += 1
				end if

				i += 1
			loop while( i < fileDataLength )

			'Continue the main for-loop
			continue for
		else
			if( bStringFlag = false ) then
				if( chr( fileData[i] ) = "/" AND chr( fileData[i+1] ) = "'" ) then
					bCommentBlockFlag = true
					commentCount = 1
					continue for
				end if
			end if
		end if
		
		'Comment Line
		if( bStringFlag = false ) then
			if( chr( fileData[i] ) = "'" ) then
				i += 1
				do while( i < fileDataLength )
					'Ascii 10 = '\n'
					if( fileData[i] = 10 ) then exit do
					i += 1
				loop
			end if
		end if

		if( bStringFlag = true ) then
			if( fileData[i] = 34 ) then 'Ascii 34 = "
				bStringFlag = false
				identifier += chr( fileData[i] )
				
				dim t as TokenUnit '= ( TOK.Tstrings, identifier, lineNum )
				t.tok = TOK.Tstring
				t.identifier = identifier
				t.lineNumber = lineNum
				
				results(tokenIndex) = t
				tokenIndex += 1
				identifier = ""
				continue for
			end if

			identifier += chr( fileData[i] )
			continue for
		else
			if( bCommentBlockFlag = False ) then
				if( fileData[i] = 34 ) then 'Ascii 34 = "
					bStringFlag = true
					identifier += chr( fileData[i] )
					continue for
				end if
			end if
		end if
		
		'print i
		'print " " + chr( fileData[i] ) 
		
		'Change Line
		select case fileData[i]
			case 35 '#
				if( fileData[i+1] = 35 ) then
					identifier += "##"
					i += 1
					continue for
				end if
				
				dim t as TokenUnit '= ( TOK.Tpound, "#", lineNum )
				t.tok = TOK.Tpound
				t.identifier = "#"
				t.lineNumber = lineNum
				
				results(tokenIndex) = t
				tokenIndex += 1
				identifier = ""

			case 45 '-
				if( results(tokenIndex).tok = TOK.Tassign ) then
					identifier += chr( fileData[i] )
					continue for
				end if

			case 44, 43, 42, 47, 58, 40, 41, 91, 93, 60, 62, 61 ' ',', '+', '*', ' /', ':', '(', ')', '[', ']', '>', '<', '='
				dim as TOK _token = identToTOK( identifier )
				if( _token > -1 ) then
					if( len( identifier ) > 0 ) then
						dim t as TokenUnit '= ( _token, identifier, lineNum )
						t.tok = _token
						t.identifier = identifier
						t.lineNumber = lineNum
								
						results(tokenIndex) = t
						tokenIndex += 1
					end if
				else
					if( len( identifier ) > 0 ) then
						dim t as TokenUnit
						
						if( identifier[0] >= 48 AND identifier[0] <= 57 ) then t.tok = TOK.Tnumbers else t.tok = TOK.Tidentifier

						if( len( identifier ) > 1 ) then

							if( identifier[0] = 45 AND identifier[1] >=48 AND identifier[1] <= 57 ) then t.tok = TOK.Tnumbers
							if( len( identifier ) > 2 ) then
								if( lcase( left( identifier, 2 ) ) = "&h" ) then t.tok = TOK.Tnumbers
								if( lcase( left( identifier, 3 ) ) = "-&h" ) then t.tok = TOK.Tnumbers
							end if									
						end if

						t.identifier = identifier
						t.lineNumber = lineNum
						results(tokenIndex) = t
						tokenIndex += 1
					end if
				end if

				dim as string s = chr( fileData[i] )
				dim as TokenUnit t '= ( identToTOK( s ), s, lineNum )
				t.tok = identToTOK( s )
				t.identifier = s
				t.lineNumber = lineNum
				
				results(tokenIndex) = t
				tokenIndex += 1
				identifier = ""

			case 13 ' CR
				' Do Nothing
				
			case 9, 32, 10 ' '\t', ' ', '\n':
				dim as TOK _token = identToTOK( identifier )

				if( _token > -1 ) then
					if( len( identifier ) > 0 ) then
						dim t as TokenUnit '= ( _token, identifier, lineNum )
						t.tok = _token
						t.identifier = identifier
						t.lineNumber = lineNum
						
						results(tokenIndex) = t
						tokenIndex += 1
					end if
				else
					if( len( identifier ) > 0 ) then
						dim t as TokenUnit
						
						if( identifier[0] >= 48 AND identifier[0] <= 57 ) then t.tok = TOK.Tnumbers else t.tok = TOK.Tidentifier

						if( len( identifier ) > 1 ) then

							if( identifier[0] = 45 AND identifier[1] >=48 AND identifier[1] <= 57 ) then t.tok = TOK.Tnumbers
							if( len( identifier ) > 2 ) then
								if( lcase( left( identifier, 2 ) ) = "&h" ) then t.tok = TOK.Tnumbers
								if( lcase( left( identifier, 3 ) ) = "-&h" ) then t.tok = TOK.Tnumbers
							end if									
						end if

						t.identifier = identifier
						t.lineNumber = lineNum
						results(tokenIndex) = t
						tokenIndex += 1
					end if
				end if

				if( fileData[i] = 10 ) then
					if( tokenIndex > 0 ) then
						if( results(tokenIndex-1).tok <> TOK.Tunderline ) then
							' Keep the TOK.Teol just only one
							if( results(tokenIndex-1).tok <> TOK.Teol ) then
								dim as TokenUnit t '= ( TOK.Teol, chr( 10 ), lineNum )
								t.tok = TOK.Teol
								t.identifier = chr( 10 )
								t.lineNumber = lineNum
								
								results(tokenIndex) = t
								tokenIndex += 1
							end if
						else
							tokenIndex -= 1
						end if
					end if
		
					lineNum += 1
					
				end if

				identifier = ""

			case 46 '.
				if( len( identifier ) > 0 ) then
					if( identifier[0] < 48 OR identifier[0] > 57 ) then

						dim as TOK _token = identToTOK( identifier )
						
						if( _token > -1 ) then
							dim as TokenUnit t '= ( identToTOK( identifier ), identifier, lineNum )
							t.tok = identToTOK( identifier )
							t.identifier = identifier
							t.lineNumber = lineNum
								
							results(tokenIndex) = t
							tokenIndex += 1
						else
							dim as TokenUnit t '= ( TOK.Tidentifier, identifier, lineNum )
							t.tok = TOK.Tidentifier
							t.identifier = identifier
							t.lineNumber = lineNum
							
							results(tokenIndex) = t
							tokenIndex += 1
						end if

						identifier = ""
						dim as TokenUnit t '= ( TOK.Tdot, ".", lineNum )
						t.tok = TOK.Tdot
						t.identifier = "."
						t.lineNumber = lineNum
							
						results(tokenIndex) = t
						tokenIndex += 1
						exit select
					end if
				end if
				
				identifier += chr( fileData[i] )

			case else
				identifier += chr( fileData[i] )
				
		end select
	next
	
	redim preserve results(tokenIndex)
	
	tokenList = new DynamicArray( tokenIndex )
	
	for i as integer = 0 to tokenIndex - 1
		'print results(i).lineNumber;
		'print "  " + results(i).identifier
		'tokenVector.push( results(tokenIndex) )
		tokenList->getMemoryPointer[i] = results(i)
	next

end sub

function Scanner.scanFile( fullPath as string ) as DynamicArray ptr

	dim fileLength as integer
	dim	fileData as string
	
	open fullpath For input As #1
	
	fileLength = lof( 1 )
	
	if fileLength = 0 then
		print "File " + fullPath + " open error!"
		return false
	end if
	
	fileData = string( fileLength, 0 )
	
	if get( #1, ,fileData ) <> 0 then
		print "File " + fullPath + " read error!"
		return false
	end if

	scan( trim( fileData ) )
	
	close
	
	return tokenList

end function

function Scanner.getTokenCount() as integer
	
	if( tokenList <> 0 )	then return tokenList->getSize
	
	return -1
	
end function

function Scanner.getTokenList() as DynamicArray ptr

	return tokenList

end function

