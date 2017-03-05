#include once "parser.bi"

Constructor CParser( _pTokenList as DynamicArray ptr )
	
	pTokenList = _pTokenList
	lineStack = new Stack(50)

end Constructor

Destructor CParser()
	
	if( pTokenList <> 0 ) then delete pTokenList
	if( lineStack <> 0 ) then delete lineStack

end Destructor

function CParser.token() as TokenUnit

	if( tokenIndex < pTokenList->getSize ) then return pTokenList->getMemoryPointer[tokenIndex]
	return nullTokenUnit

end function

function CParser.next1() as TokenUnit 
	
	if( tokenIndex < pTokenList->getSize - 1 ) then return pTokenList->getMemoryPointer[tokenIndex+1]
	return nullTokenUnit

end function

function CParser.next2() as TokenUnit
	
	if( tokenIndex < pTokenList->getSize - 2 ) then return pTokenList->getMemoryPointer[tokenIndex+2]
	return nullTokenUnit

end function

sub CParser.parseToken( t as TOK = TOK.Tidentifier )
	
	if( tokenIndex < pTokenList->getSize ) then tokenIndex += 1
	
end sub


private function CParser.parsePreprocessor() as boolean

	static as string	_type
	
	parseToken( TOK.Tpound )

	select case( token().tok )
		
		case TOK.Tinclude:
			parseToken( TOK.Tinclude )
			if( token().tok = TOK.Tonce ) then parseToken( TOK.Tonce )

			if( token().tok = TOK.Tstrings ) then 
				dim as TokenUnit t = token()
				parseToken( TOK.Tstring )
				JSONast->addChild( mid( str_replace( t.identifier, "\", "/" ), 2, len( trim( t.identifier ) ) - 2 ) , str( B_INCLUDE ), "", _type, "", str( t.lineNumber ) )
				'activeASTnode.addChild( Util.replace( t.identifier, '\\', '/' ) , B_INCLUDE, null, _type, null, t.lineNumber );
			end if

		case TOK.Tmacro:
			parseToken( TOK.Tmacro )
			if( token().tok = TOK.Tidentifier ) then

				dim as string		_name = token().identifier
				dim as integer		lineNumber = token().lineNumber
				dim as string		paramName, paramString
				
				parseToken( TOK.Tidentifier )

				JSONast->addChild( _name, str( B_MACRO ), "", paramString, "", str( lineNumber ) )
				JSONast->toChildLayer()
				'activeASTnode = activeASTnode.addChild( name, B_MACRO, null, paramString, null, lineNumber );
				lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
				
				dim as integer	nearestTypeInsertPos = JSONast->changeMemberValue( "type", "" )
				
				
				do while( token().tok <> TOK.Teol AND token().tok <> TOK.Tcolon )

					paramString += token().identifier
					
					if( token().tok = TOK.Topenparen ) then
						parseToken( TOK.Topenparen )
					elseif( token().tok = TOK.Tcloseparen ) then
						parseToken( TOK.Tcloseparen )
						exit do
					elseif( token().tok = TOK.Tcomma ) then
						JSONast->addChild( paramName, str( B_PARAM ), "", "", "", str( lineNumber ) )
						'activeASTnode.addChild( paramName, B_PARAM, null, null, null, lineNumber );
						paramName = ""
						parseToken()
					else
						paramName += token().identifier
						parseToken()
					end if
				loop

				JSONast->insertText( nearestTypeInsertPos, paramString )
				'JSONast->changeMemberValue( "type", paramString )
			end if

		case TOK.Tdefine:
			parseToken( TOK.Tdefine )

			if( token().tok = TOK.Tidentifier ) then
				
				dim as string	_name = token().identifier
				dim as integer	lineNumber = token().lineNumber
				parseToken( TOK.Tidentifier )

				if( token().tok <> TOK.Teol AND token().tok <> TOK.Tcolon ) then
					if( token().tok = TOK.Tnumbers  ) then

						if( InStr( token().identifier, "." ) > 0 ) then _type = "single" else _type = "integer"
						JSONast->addChild( _name, str( B_VARIABLE OR B_DEFINE ), "", _type, "", str( lineNumber ) )
						'activeASTnode.addChild( name, B_VARIABLE | B_DEFINE, null, type, null, lineNumber );
						
					elseif( token().tok = TOK.Tstrings  ) then
					
						JSONast->addChild( _name, str( B_VARIABLE OR B_DEFINE ), "", "string", "", str( lineNumber ) )
						'activeASTnode.addChild( name, B_VARIABLE | B_DEFINE, null, "string", null, lineNumber );
						
					elseif( token().tok = TOK.Topenparen ) then
						
						' #define GTK_VSCALE(obj) G_TYPE_CHECK_INSTANCE_CAST((obj), GTK_TYPE_VSCALE, GtkVScale)
						dim as string	param

						do while( token().tok <> TOK.Teol )
							param += token().identifier
							if(  token().tok = TOK.Tcloseparen ) then
								parseToken()
								exit do
							end if
							parseToken()
						loop
						JSONast->addChild( _name, str( B_FUNCTION OR B_DEFINE ), "", param, "", str( lineNumber ) )
						'activeASTnode.addChild( name, B_FUNCTION | B_DEFINE, null, param, null, lineNumber );	
					
					elseif( token().tok = TOK.Tidentifier ) then
						
						JSONast->addChild( _name, str( B_VARIABLE OR B_DEFINE ), "", "", "", str( lineNumber ) )
						'activeASTnode.addChild( name, B_VARIABLE | B_DEFINE, null, null, null, lineNumber );	
					end if
				end if
			end if

		case TOK.Tifdef:
			parseToken( TOK.Tifdef )

			_type = ucase( token().identifier )
			parseToken()

		case TOK.Telseif, TOK.Telse:
			parseToken()

			if( len( _type ) > 0 ) then
				if( chr( _type[0] ) = "!" ) then _type = ucase( right( _type, len( _type ) - 1 ) ) else	_type = ucase( "!" + _type )
			end if

		case TOK.Tendif:
			parseToken( TOK.Tendif )
			_type = ""

		case TOK.Tifndef:
			parseToken( TOK.Tifndef )
			_type = "!" + token().identifier
			parseToken()
			
		case else
	end select
	
	return true
	
end function

private function CParser.getVariableType() as string

	if( token().tok = TOK.Tunsigned ) then
		select case( next1().tok )
			case TOK.Tbyte:
				parseToken()
				return "ubyte"
				
			case TOK.Tshort:
				parseToken()
				return "ushort"
			
			case TOK.Tinteger
				parseToken()
				return "uinteger"
			
			case TOK.Tlongint
				parseToken()
				return "ulongint"

			case else
				parseToken()
		end select
	end if
	
	select case( token().tok )
		case TOK.Tbyte, TOK.Tubyte, TOK.Tshort,TOK.Tushort, TOK.Tinteger, TOK.Tuinteger, TOK.Tlongint, TOK.Tulongint, TOK.Tsingle, TOK.Tdouble, TOK.Tstring, TOK.Tzstring, TOK.Twstring, TOK.Tany
			return lcase( token.identifier )
		
		case TOK.Tidentifier
			dim as string _type
			do while( next1().tok = TOK.Tdot )
				_type += token.identifier
				parseToken( TOK.Tidentifier )
				_type += token.identifier
				parseToken( TOK.Tdot )
			loop

			if( len( _type ) < 1 ) then _type = token.identifier else _type += token.identifier
			return _type

		case else
	end select 
	
	return ""
	
end function

private function CParser.parseParam( bDeclare as boolean ) as string

	dim as string		_param = "("

	if( token().tok = TOK.Topenparen ) then

		parseToken( TOK.Topenparen )
		
		if( token().tok = TOK.Tcloseparen ) then
			parseToken( TOK.Tcloseparen )
			return "()"
		end if
				
		do while( token().tok <> TOK.Tcloseparen )
		
			dim as string	_name, _type
			dim as integer	_lineNum

			if( token().tok = TOK.Tbyref OR token().tok = TOK.Tbyval ) then parseToken()

			_name = token().identifier
			_lineNum = token().lineNumber

			parseToken()

			' Array
			if( token().tok = TOK.Topenparen ) then _name += parseArray()

			_param += ( _name + " " )
			
			if( token().tok = TOK.Tas ) then

				_param += ( token().identifier + " " )
				parseToken( TOK.Tas )

				' Function pointer
				if( token().tok = TOK.Tfunction OR token().tok = TOK.Tsub ) then
					dim as integer		_kind
					dim as string		__param
					
					if( token().tok = TOK.Tfunction ) then _kind = B_FUNCTION else _kind = B_SUB
					
					_param += token().identifier
					parseToken()

					if( token().tok = TOK.Tstdcall OR token().tok = TOK.Tcdecl OR token().tok = TOK.Tpascal ) then parseToken()

					' like " Declare Function app_oninit_cb WXCALL () As wxBool "
					if( token().tok = TOK.Tidentifier ) then
						_param += ( " " + token().identifier )
						parseToken( TOK.Tidentifier )
					end if

					' Overload
					if( token().tok = TOK.Toverload ) then parseToken( TOK.Toverload )

					' Alias "..."
					if( token().tok = TOK.Talias ) then

						parseToken( TOK.Talias )
						if( token.tok = TOK.Tstrings ) then parseToken( TOK.Tstrings ) else return ""
						
					end if

					dim as string		_returnType
					if( token().tok = TOK.Topenparen ) then __param = parseParam( false )

					_param += __param

					if( token().tok = TOK.Tas ) then
						_param += ( " " + token().identifier )
						parseToken( TOK.Tas )

						if( token().tok = TOK.Tconst ) then
							_param += ( " " + token().identifier )
							parseToken( TOK.Tconst )
						end if

						_returnType = getVariableType()
						if( len( _returnType ) > 0 ) then 
							parseToken()
							do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
								_returnType += "*"
								parseToken()
							loop

							' Mother param strings
							_param += ( " " + _returnType )

							_type = _returnType + __param
							JSONast->addChild( _name, str( _kind ), "", _type, "", str( _lineNum ) )
							'activeASTnode.addChild( _name, _kind, null, _type, null, _lineNum );
						end if
					end if

					if( token().tok = TOK.Tstatic OR token().tok = TOK.Texport  ) then parseToken()

					' SUB
					if( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
						JSONast->addChild( _name, str( _kind ), "", __param, "", str( _lineNum ) )
						'activeASTnode.addChild( _name, _kind, null, __param, null, _lineNum );
					end if	
				end if

				if( token().tok = TOK.Tconst ) then
					_param += ( token().identifier + " " )
					parseToken( TOK.Tconst )
				end if

				_type = getVariableType()
				if( len( _type ) > 0 ) then
					
					parseToken()
					do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
						_type += "*"
						parseToken()
					loop
					_param += ( _type + " " )

					if( token().tok = TOK.Tassign ) then
						parseToken( TOK.Tassign )

						dim as integer countParen
						do
							if( token().tok = TOK.Topenparen ) then
								countParen += 1
							elseif( token().tok = TOK.Tcloseparen ) then
								countParen -= 1
								if( countParen = -1 ) then exit do 'Param Tail
							elseif( token().tok = TOK.Tcomma ) then
								if( countParen = 0 ) then exit do
							end if
							
							if( tokenIndex >= pTokenList->getSize ) then exit do
							parseToken()
							
						loop while( token().tok <> TOK.Teol )
					end if							

				
					if( token().tok = TOK.Tcomma OR token().tok = TOK.Tcloseparen ) then
						_param = trim( _param )
						
						if( bDeclare = false ) then
							JSONast->addChild( _name, str( B_PARAM ), "", _type, "", str( _lineNum ) )
							'activeASTnode.addChild( _name, B_PARAM, null, _type, null, _lineNum );
						end if
						
						if( token().tok = TOK.Tcomma ) then
							_param += ( token().identifier )
							parseToken( TOK.Tcomma )
						end if
					end if
				else
					exit do
				end if
			end if
		loop
	end if
	
	if( token().tok = TOK.Tcloseparen ) then
		parseToken( TOK.Tcloseparen )
		if( right( _param, 1 ) = " " ) then _param = left( _param, len( _param ) - 1 )

		_param = _param + ")"
	end if

	return _param

end function

private function CParser.parseArray() as string

	dim as string		result

	if( token().tok = TOK.Topenparen ) then
		
		parseToken( TOK.Topenparen )
		result = "("
		if( token().tok = TOK.Tcloseparen ) then
			parseToken( TOK.Tcloseparen )
			return "()"
		end if
		
		dim as integer	countOpenParen = 1
		do
			select case token().tok
				case TOK.Tcloseparen
					countOpenParen -= 1
					parseToken( TOK.Tcloseparen )
					result += ")"
					if( countOpenParen = 0 ) then return result

				case TOK.Topenparen
					countOpenParen += 1
					parseToken( TOK.Topenparen )
					result += "("

				case TOK.Tto:
					result += " to "
					parseToken()

				case TOK.Teol, TOK.Tcolon
					parseToken()
					return ""

				case else
					result += token.identifier
					parseToken()
			end select
		loop while( tokenIndex < pTokenList->getSize )
	end if
	
	return result
	
end function

function CParser.parseVariable() as boolean

	dim as boolean		bConst
		
	if( token().tok = TOK.Tdim OR token().tok = TOK.Tstatic OR token().tok = TOK.Tcommon OR token().tok = TOK.Tconst OR token().tok = TOK.Tredim ) then
		
		if( token().tok = TOK.Tconst ) then bConst = true

		if( token().tok = TOK.Tredim AND next1().tok = TOK.Tpreserve ) then parseToken( TOK.Tredim )

		parseToken()

		dim as string		_type, _name, _protection
		dim as integer		_lineNum

		if( token().tok = TOK.Tshared ) then
			parseToken( TOK.Tshared )
			_protection = "shared"
		end if
		
		if( token().tok = TOK.Tas ) then

			parseToken( TOK.Tas )
			if( token().tok = TOK.Tconst ) then parseToken( TOK.Tconst )
			_type = getVariableType()
			
			if( len( _type ) > 0 ) then
				parseToken()

				if( _type = "string" OR _type = "zstring" OR _type = "wstring" ) then
					if( token().tok = TOK.Ttimes ) then
						parseToken( TOK.Ttimes )
						parseToken() ' TOK.Tnumber or TOK.Tidentifier
					end if
				end if

				do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
					_type += "*"
					parseToken()
				loop

				do while( token().tok = TOK.Tidentifier )

					_lineNum = token().lineNumber
					_name	= token().identifier
					
					parseToken( TOK.Tidentifier )

					' Array
					if( token().tok = TOK.Topenparen ) then _name += parseArray()
					
					if( token().tok = TOK.Tassign ) then

						parseToken( TOK.Tassign )

						dim as integer		countCurly, countParen

						select case( token().tok )
							case TOK.Topencurly
								do
									if( token().tok = TOK.Topencurly ) then
										countCurly += 1
									elseif( token().tok = TOK.Tclosecurly ) then
										countCurly -= 1
										if( countCurly = 0 ) then
											parseToken( TOK.Tclosecurly )
											exit do
										end if
									elseif( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
										exit do
									end if

									parseToken()
								loop while( tokenIndex < pTokenList->getSize )

							case TOK.Topenparen
								do
									if( token().tok = TOK.Topenparen ) then
										countParen += 1
									elseif( token().tok = TOK.Tcloseparen ) then
										countParen -= 1
										if( countParen = 0 ) then
											parseToken( TOK.Tcloseparen )
											exit do
										end if
									elseif( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
										exit do
									end if

									parseToken()
								loop while( tokenIndex < pTokenList->getSize )

							case else
								do while( token().tok <> TOK.Teol AND token().tok <> TOK.Tcolon )
									if( token().tok = TOK.Tcomma ) then
										if( countParen = 0 AND countCurly = 0 ) then exit do
									end if
									parseToken()
									if( tokenIndex >= pTokenList->getSize ) then exit do
								loop
						end select
					end if

					if( token().tok = TOK.Tcomma ) then
						JSONast->addChild( _name, str( B_VARIABLE ), _protection, _type, "", str( _lineNum ) )
						'activeASTnode.addChild( _name, B_VARIABLE, _protection, _type, null, _lineNum );
						parseToken( TOK.Tcomma )
						
					elseif( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
						JSONast->addChild( _name, str( B_VARIABLE ), _protection, _type, "", str( _lineNum ) )
						'activeASTnode.addChild( _name, B_VARIABLE, _protection, _type, null, _lineNum );
						parseToken()
						exit do
					else
						return false
					end if
				loop
			end if				
		
		elseif( token().tok = TOK.Tidentifier ) then
		
			do while( token().tok = TOK.Tidentifier )
				_name = token().identifier
				_lineNum = token().lineNumber
				parseToken( TOK.Tidentifier )

				if( bConst = true ) then
					if( token().tok = TOK.Tassign ) then
						parseToken( TOK.Tassign )
						if( token().tok = TOK.Tnumbers OR token().tok = TOK.Tidentifier ) then
							parseToken()
							if( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
								parseToken( )
								JSONast->addChild( _name, str( B_VARIABLE ), "", "", "", str( _lineNum ) )
								'activeASTnode.addChild( _name, B_VARIABLE, null, null, null, _lineNum );
								return true
							else
								return false
							end if
						else
							return false
						end if
					end if
				end if
				
				' Array
				if( token().tok = TOK.Topenparen ) then _name += parseArray()
			
				if( token().tok = TOK.Tas ) then
					parseToken( TOK.Tas )

					if( token().tok = TOK.Tconst ) then parseToken( TOK.Tconst )

					_type = getVariableType()
					if( len( _type ) > 0 ) then
						parseToken()

						if( _type = "string" OR _type = "zstring" OR _type = "wstring" ) then
							if( token().tok = TOK.Ttimes ) then
								parseToken( TOK.Ttimes )
								parseToken() 'TOK.Tnumber or TOK.Tidentifier
							end if
						end if

						do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
							_type += "*"
							parseToken()
						loop
						
						JSONast->addChild( _name, str( B_VARIABLE ), "", _type, "", str( _lineNum ) )
						'activeASTnode.addChild( _name, B_VARIABLE, null, _type, null, _lineNum );

						if( token.tok <> TOK.Tcomma ) then exit do else parseToken( TOK.Tcomma )
					end if
				end if
			loop
		end if
		
		return true
		
	end if

	return false

end function

function CParser.parserVar() as boolean

	dim as string		_name
	dim as integer	 	_lineNum
	
	parseToken( TOK.Tvar )

	if( token().tok = TOK.Tshared ) then parseToken( TOK.Tshared )

	_name = token().identifier
	_lineNum = token().lineNumber
	parseToken()

	if( token().tok = TOK.Tassign ) then
		parseToken( TOK.Tassign )
		if( token().tok = TOK.Tcast ) then

			parseToken( TOK.Tcast )
			
			if( token().tok = TOK.Topenparen ) then
				parseToken( TOK.Topenparen )

				JSONast->addChild( _name, str( B_VARIABLE ), "", lcase( token().identifier ), "", str( _lineNum ) )
				'activeASTnode.addChild( _name, B_VARIABLE, null, toLower( token().identifier ), null, _lineNum );

				do while( token().tok <> TOK.Teol AND token().tok <> TOK.Tcolon )
					if( tokenIndex < pTokenList->getSize ) then parseToken() else return false
					parseToken()
				loop
				parseToken()
			end if
		else
			dim as string	_typeName = token().identifier
			parseToken()
			JSONast->addChild( _name, str( B_VARIABLE ), "", "Var(" + _typeName + ")", "", str( _lineNum ) )
			'activeASTnode.addChild( _name, B_VARIABLE, null, "Var(" ~ _typeName ~ ")", null, _lineNum );
			do while( token().tok <> TOK.Teol AND token().tok <> TOK.Tcolon )
				if( tokenIndex < pTokenList->getSize ) then parseToken() else return false
			loop
			parseToken()
		end if
	end if

	return true

end function

function CParser.parseNamespace() as boolean

	parseToken( TOK.Tnamespace )
	
	if( token().tok = TOK.Tidentifier ) then
		JSONast->addChild( token().identifier, str( B_NAMESPACE ), "", "", "", str( token().lineNumber ) )
		JSONast->toChildLayer()
		'activeASTnode = activeASTnode.addChild( token().identifier, B_NAMESPACE, null, null, null, token().lineNumber );
		lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
		parseToken( TOK.Tidentifier )
		return true
	end if

	return false
	
end function

function CParser.parseOperator( bDeclare as boolean, _protection as string ) as boolean

	dim as string		_returnType, _name, _kind, _param
	dim as integer		_lineNum, opType
	
	parseToken( TOK.Toperator )

	' Function Name
	if( bDeclare = false ) then
		if( token().tok = TOK.Tidentifier AND next1().tok = TOK.Tdot  ) then
			_name = token().identifier
			parseToken()

			_name += token().identifier
			parseToken( TOK.Tdot )
		end if
	end if

	select case( token().tok )

		case TOK.Tcast, TOK.Tat
		
			_name += token().identifier
			_lineNum = token().lineNumber
			parseToken()
			
			if( token().tok = TOK.Topenparen AND next1().tok = TOK.Tcloseparen ) then
				parseToken( TOK.Topenparen )
				parseToken( TOK.Tcloseparen )

				if( token().tok = TOK.Tbyref ) then
					parseToken( TOK.Tbyref )
				elseif( token().tok = TOK.Tbyval ) then
					parseToken( TOK.Tbyval )
				end if

				if( token().tok = TOK.Tas ) then
					parseToken( TOK.Tas )
					_returnType = getVariableType()
					if( len( _returnType ) > 0 ) then
						parseToken()
						do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
							_returnType += "*"
							parseToken()
						loop
					end if

					JSONast->addChild( _name, str( B_OPERATOR ), _protection, _returnType + "()", "", str( _lineNum ) )
					JSONast->toChildLayer()
					'activeASTnode = activeASTnode.addChild( _name, B_OPERATOR, _protection, _returnType ~ "()", null, _lineNum );
					lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
					if( bDeclare = true ) then
						JSONast->toFatherLayer()
						JSONast->insertText( lineStack->pop, str( token().lineNumber ) )
					end if
					
					return true
				end if
			end if

		case TOK.Topenbracket
			if( next1().tok = TOK.Tclosebracket ) then
				_lineNum = token().lineNumber
				
				parseToken( TOK.Topenbracket )
				parseToken( TOK.Tclosebracket )
				_name += "[]"
				
				if( token().tok = TOK.Topenparen ) then
					JSONast->addChild( _name, str( B_OPERATOR ), _protection, "", "", str( _lineNum ) )
					JSONast->toChildLayer()
					lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
					
					dim as integer	nearestTypeInsertPos = JSONast->changeMemberValue( "type", "" )
					
					_param = parseParam( bDeclare )

					if( token().tok = TOK.Tbyref ) then
						parseToken( TOK.Tbyref )
					elseif( token().tok = TOK.Tbyval ) then
						parseToken( TOK.Tbyval )
					end if

					if( token().tok = TOK.Tas ) then 
						parseToken( TOK.Tas )
						_returnType = getVariableType()
						if( len( _returnType ) > 0 ) then
							parseToken()
							do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
								_returnType += "*"
								parseToken()
							loop
						end if

						JSONast->insertText( nearestTypeInsertPos, _returnType + _param )
						'activeASTnode = activeASTnode.addChild( _name, B_OPERATOR, _protection, _returnType ~ _param, null, _lineNum );
						if( bDeclare = true ) then
							JSONast->toFatherLayer
							JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
						end if
						
						return true
					end if
				end if
			end if
			
		case TOK.Tnew, TOK.Tdelete ' "new" and "new[]" and "delete" and "delete[]"
			_lineNum = token().lineNumber
			_name += token().identifier
			parseToken()

			if( token().tok = TOK.Topenbracket AND next1().tok = TOK.Tclosebracket ) then
				parseToken( TOK.Topenparen )
				parseToken( TOK.Tcloseparen )	
				_name += "[]"
			end if

			if( token().tok = TOK.Topenparen ) then
				JSONast->addChild( _name, str( B_OPERATOR ), _protection, "", "", str( _lineNum ) )
				JSONast->toChildLayer()
				lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
				
				dim as integer	nearestTypeInsertPos = JSONast->changeMemberValue( "type", "" )	
				
				_param = parseParam( bDeclare )
				
				if( left( lcase( _name ), 3 ) = "new" ) then
					if( token().tok = TOK.Tas ) then
						parseToken( TOK.Tas )
						_returnType = getVariableType()
						if( len( _returnType ) > 0 ) then
							parseToken()
							do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
								_returnType += "*"
								parseToken()
							loop
						end if

						JSONast->insertText( nearestTypeInsertPos, _returnType + _param )
						'activeASTnode = activeASTnode.addChild( _name, B_OPERATOR, _protection, _returnType ~ _param, null, _lineNum );
						if( bDeclare = true ) then
							JSONast->toFatherLayer()
							JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
						end if
							
					end if
					
				else 'delete

					JSONast->insertText( nearestTypeInsertPos, _param )
					'activeASTnode = activeASTnode.addChild( _name, B_OPERATOR, _protection, _param, null, _lineNum );
					if( bDeclare = true ) then
						JSONast->toFatherLayer()
						JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
					end if
					
				end if
				
				return true
			end if

		case TOK.Tlet
			if( next1().tok = TOK.Topenparen ) then opType = 1 else return false 'Check if assignment_op, p.s. opType = 1

		case TOK.Tplus, TOK.Tminus, TOK.Ttimes, TOK.Tandsign, TOK.Tdiv, TOK.Tintegerdiv, TOK.Tmod, TOK.Tshl, TOK.Tshr, TOK.Tand, TOK.Tor, TOK.Txor, TOK.Timp, TOK.Teqv, TOK.Tcaret:
			if( opType = 0 ) then
				if( next1().tok = TOK.Tassign ) then
					opType = 1
				elseif( next1().tok = TOK.Topenparen ) then
					opType = 3
				elseif( token().tok = TOK.Tminus AND next1().tok = TOK.Tgreater ) then
					_name += token().identifier
					parseToken()
					opType = 2						
				end if		
			end if
			
		case TOK.Tassign
			if( opType = 0 ) then
				if( next1().tok = TOK.Topenparen ) then opType = 3 else return false
			end if

		case TOK.Tless
			if( opType = 0 ) then
				if( next1().tok = TOK.Tgreater OR next1().tok = TOK.Tassign ) then
					_name += token().identifier
					parseToken()
					opType = 3
				elseif( next1().tok = TOK.Topenparen ) then
					opType = 3
				else
					return false
				end if
			end if

		case TOK.Tgreater
			if( opType = 0 ) then
				if( next1().tok = TOK.Tassign ) then
					_name += token().identifier
					parseToken()
					opType = 3
				elseif( next1().tok = TOK.Topenparen ) then
					opType = 3
				else
					return false
				end if
			end if

		case TOK.Tnot
			if( opType = 0 ) then
				if( next1().tok = TOK.Topenparen ) then opType = 2 else return false
			end if
			
		case TOK.Tidentifier
			if( opType = 0 ) then
				select case( lcase( token().identifier ) )
					case "abs", "sgn", "fix", "frac", "int", "exp", "log", "sin", "asin", "cos", "acos", "tan", "atn", "len"
						if( next1().tok = TOK.Topenparen ) then opType = 2

					case else
						return false
				end select
			end if

		case else
			return false
	end select
	
	select case ( opType )
		case 1
			_name += token().identifier
			_lineNum = token().lineNumber
			parseToken()
			if( token().tok = TOK.Tassign ) then
				parseToken( TOK.Tassign )
				_name += "="
			end if

			if( token().tok = TOK.Topenparen ) then	
				JSONast->addChild( _name, str( B_OPERATOR ), _protection, "", "", str( _lineNum ) )
				JSONast->toChildLayer()
				lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
				
				dim as integer	nearestTypeInsertPos = JSONast->changeMemberValue( "type", "" )
				
				_param = parseParam( bDeclare )

				JSONast->insertText( nearestTypeInsertPos, _param )
				'activeASTnode = activeASTnode.addChild( _name, B_OPERATOR, _protection, _param, null, _lineNum );
				if( bDeclare = true ) then
					JSONast->toFatherLayer()
					JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
				end if
				
				if( token().tok = TOK.Texport ) then parseToken( TOK.Texport )
				
				return true
			end if

		case 2, 3
			_name += token().identifier
			_lineNum = token().lineNumber
			parseToken()

			if( token().tok = TOK.Topenparen ) then
				JSONast->addChild( _name, str( B_OPERATOR ), _protection, "", "", str( _lineNum ) )
				JSONast->toChildLayer()
				lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
				
				dim as integer	nearestTypeInsertPos = JSONast->changeMemberValue( "type", "" )			
			
				_param = parseParam( bDeclare )
				
				if( token().tok = TOK.Tbyref ) then
					parseToken( TOK.Tbyref )
				elseif( token().tok = TOK.Tbyval ) then
					parseToken( TOK.Tbyval )
				end if

				if( token().tok = TOK.Tas ) then
					parseToken( TOK.Tas )
					_returnType = getVariableType()
					if( len( _returnType ) > 0 ) then
						parseToken()
						do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
							_returnType += "*"
							parseToken()
						loop
					end if

					JSONast->insertText( nearestTypeInsertPos, _returnType + _param )
					'activeASTnode = activeASTnode.addChild( _name, B_OPERATOR, _protection, _returnType ~ _param, null, _lineNum );
					if( bDeclare = true ) then
						JSONast->toFatherLayer()
						JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
					end if
					
					if( token().tok = TOK.Texport ) then parseToken( TOK.Texport )
					
					return true
				end if
			end if

		case else
			'return false
	end select


	return false
end function

function CParser.parseProcedure( bDeclare as boolean, _protection as string ) as boolean

	dim as integer		_kind
	
	if( next1().tok = TOK.Tassign ) then
		parseToken()
		return true
	end if
	
	select case( token().tok )
		case TOK.Tfunction
			_kind = B_FUNCTION
		case TOK.Tsub
			_kind = B_SUB
		case TOK.Tproperty
			_kind = B_PROPERTY
		case TOK.Tconstructor
			_kind = B_CTOR
		case TOK.Tdestructor
			_kind = B_DTOR
	end select
	
	if( token().tok = TOK.Tfunction OR token().tok = TOK.Tsub OR token().tok = TOK.Tproperty OR token().tok = TOK.Tconstructor OR token().tok = TOK.Tdestructor ) then
		if( bDeclare = true ) then
			if( token().tok <> TOK.Tconstructor AND token().tok <> TOK.Tdestructor ) then
				parseToken()
				if( token().tok = TOK.Tstdcall OR token().tok = TOK.Tcdecl OR token().tok = TOK.Tpascal ) then parseToken()
			end if
		else
			parseToken()
		end if

		' Function Name
		dim	as boolean		bFunction
		if( bDeclare = true ) then
			if( token().tok = TOK.Tconstructor OR token().tok = TOK.Tdestructor ) then bFunction = true
		end if
		
		if( cbool( token().tok = TOK.Tidentifier ) OR ( bDeclare = true AND cbool( token().tok = TOK.Tconstructor OR token().tok = TOK.Tdestructor ) ) ) then
			
			dim as string		_name, _param, _returnType
			dim as integer		_lineNum

			_name = token().identifier
			_lineNum = token().lineNumber
			parseToken()

			' method
			if( token().tok = TOK.Tdot AND next1().tok = TOK.Tidentifier ) then

				_name += token().identifier
				parseToken( TOK.Tdot )

				_name += token().identifier
				parseToken( TOK.Tidentifier )
			end if

			if( token().tok = TOK.Tstdcall OR token().tok = TOK.Tcdecl OR token().tok = TOK.Tpascal ) then parseToken()

			' like " Declare Function app_oninit_cb WXCALL () As wxBool "
			if( token().tok = TOK.Tidentifier ) then parseToken( TOK.Tidentifier )

			' Overload
			if( token().tok = TOK.Toverload ) then parseToken( TOK.Toverload )

			' Alias "..."
			if( token().tok = TOK.Talias ) then
				parseToken( TOK.Talias )
				if( token.tok = TOK.Tstrings ) then parseToken( TOK.Tstrings ) else return false
			end if

			JSONast->addChild( _name, str( _kind ), _protection, "", "", str( _lineNum ) )
			JSONast->toChildLayer()
			'activeASTnode = activeASTnode.addChild( _name, _kind, _protection, null, null, _lineNum );
			lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
			
			dim as integer	nearestTypeInsertPos = JSONast->changeMemberValue( "type", "" )
			
			if( token().tok = TOK.Topenparen ) then _param = parseParam( bDeclare )

			if( token().tok = TOK.Tas ) then
				parseToken( TOK.Tas )

				if( token().tok = TOK.Tconst ) then parseToken( TOK.Tconst )

				_returnType = getVariableType()
				if( len( _returnType ) > 0 ) then
					parseToken()
					do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
						_returnType += "*"
						parseToken()
					loop
				
					
					JSONast->insertText( nearestTypeInsertPos, _returnType + _param )
					'JSONast->changeMemberValue( "type", _returnType + _param )
					'activeASTnode.type = _returnType ~ _param;
					if( bDeclare = true ) then
						JSONast->toFatherLayer() ' activeASTnode = activeASTnode.getFather();
						JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
					end if

					if( token().tok = TOK.Tstatic ) then parseToken( TOK.Tstatic )
					
					return true
				end if
			end if

			if( token().tok = TOK.Tstatic OR token().tok = TOK.Texport ) then parseToken()
			
			' SUB
			if( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then

				JSONast->insertText( nearestTypeInsertPos, _param )
				'JSONast->changeMemberValue( "type", _param )
				if( bDeclare = true ) then
					JSONast->toFatherLayer() ' activeASTnode = activeASTnode.getFather()
					JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
				end if
				return true
			end if					
		end if
	end if


	return false
	
end function


function CParser.parseTypeBody( B_KIND as integer ) as boolean

	dim as string		_protection

	select case( B_KIND )
		case B_TYPE
			B_KIND = TOK.Ttype
		case B_UNION
			B_KIND = TOK.Tunion
		case B_ENUM
			B_KIND = TOK.Tenum
		case else
			B_KIND = TOK.Ttype
	end select
	
	do while( token().tok <> TOK.Tend AND next1().tok <> B_KIND )
	
		dim as string		_type, _name
		dim as integer		_lineNum
		
		select case( token().tok )
			case TOK.Tpublic
				parseToken( TOK.Tpublic )
				if( token().tok = TOK.Tcolon ) then
					parseToken( TOK.Tcolon )
					_protection = "public"
				end if

			case TOK.Tprivate
				parseToken( TOK.Tprivate )
				if( token().tok = TOK.Tcolon ) then
					parseToken( TOK.Tcolon )
					_protection = "private"
				end if

			case TOK.Tprotected
				parseToken( TOK.Tprivate )
				if( token().tok = TOK.Tcolon ) then
					parseToken( TOK.Tcolon )
					_protection = "protected"
				end if
				
			case TOK.Tdim
				parseVariable()
			
			case TOK.Tas
				parseToken( TOK.Tas )
				_type = getVariableType()

				if( len( _type ) > 0 ) then

					parseToken()

					if( _type = "string" OR _type = "zstring" OR _type = "wstring" ) then

						if( token().tok = TOK.Ttimes ) then
							parseToken( TOK.Ttimes )
							parseToken() ' TOK.Tnumber or TOK.Tidentifier
						end if
					end if

					do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
						_type += "*"
						parseToken()
					loop

					do while( token().tok = TOK.Tidentifier )

						_lineNum	= token().lineNumber
						_name		= token().identifier
						
						parseToken( TOK.Tidentifier )

						' Array
						if( token().tok = TOK.Topenparen ) then _name += parseArray()

						' As DataType fieldname : bits [= initializer], ...
						if( token().tok = TOK.Tcolon ) then
							parseToken( TOK.Tcolon )
							if( token().tok = TOK.Tnumbers ) then parseToken( TOK.Tnumbers ) else return false
						end if

						' [= initializer]
						if( token().tok = TOK.Tassign ) then
						
							parseToken( TOK.Tassign )

							select case ( token().tok )
								case TOK.Topencurly
									dim as integer	countCurly

									do
										if( token().tok = TOK.Topencurly ) then
											countCurly += 1
										elseif( token().tok = TOK.Tclosecurly ) then
											countCurly -= 1
											if( countCurly = 0 ) then
												parseToken( TOK.Tclosecurly )
												exit do
											end if
										elseif( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
											exit do
										end if

										parseToken()
									loop while( tokenIndex < pTokenList->getSize )

								case TOK.Topenparen
									dim as integer	countParen

									do
										if( token().tok = TOK.Topenparen ) then
											countParen += 1
										elseif( token().tok = TOK.Tcloseparen ) then
											countParen -= 1
											if( countParen = 0 ) then
												parseToken( TOK.Tcloseparen )
												exit do
											end if
										elseif( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
											exit do
										end if
										
										parseToken()
									loop while( tokenIndex < pTokenList->getSize )

								case else
									do while( token.tok <> TOK.Teol AND token.tok <> TOK.Tcolon )
										parseToken()
										if( tokenIndex >= pTokenList->getSize ) then exit do
									loop
							end select
						end if

						if( token().tok = TOK.Tcomma ) then
							JSONast->addChild( _name, str( B_VARIABLE ), _protection, _type, "", str( _lineNum ) )
							'activeASTnode.addChild( _name, B_VARIABLE, _protection, _type, null, _lineNum );
							parseToken( TOK.Tcomma )
						elseif( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
							JSONast->addChild( _name, str( B_VARIABLE ), _protection, _type, "", str( _lineNum ) )
							'activeASTnode.addChild( _name, B_VARIABLE, _protection, _type, null, _lineNum );
							parseToken()
							exit do
						else
							return false
						end if
					loop
				else
					parseToken()
				end if

			case TOK.Tstatic
				parseToken( TOK.Tstatic )

			case TOK.Tdeclare:
				parseToken( TOK.Tdeclare )

				if( token().tok = TOK.Tstatic OR token().tok = TOK.Tconst ) then parseToken()

				if( token().tok = TOK.Tfunction OR token().tok = TOK.Tsub OR token().tok = TOK.Tproperty ) then
					parseProcedure( true, _protection )
				elseif( token().tok = TOK.Tconstructor OR token().tok = TOK.Tdestructor ) then
					parseProcedure( true, _protection )
				elseif( token().tok = TOK.Toperator ) then
					parseOperator( true, _protection )
				end if

			case TOK.Teol, TOK.Tcolon
				parseToken()

			case TOK.Tunion, TOK.Ttype:
				dim as TOK		nestUnnameTOK = token().tok
				dim as integer	B_VALUE
				_lineNum = token().lineNumber

				if( next1().tok = TOK.Teol OR next1().tok = TOK.Tcolon ) then
					parseToken()
					parseToken()
					
					if( nestUnnameTOK = TOK.Tunion ) then B_VALUE = B_UNION else B_VALUE = B_TYPE
					
					JSONast->addChild( "", str( B_VALUE ), "", "", "", str( _lineNum ) )
					JSONast->toChildLayer()
					'activeASTnode = activeASTnode.addChild( null, ( nestUnnameTOK == TOK.Tunion ? B_UNION : B_TYPE ), null, null, null, _lineNum );
					lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
					
					parseTypeBody( B_VALUE )
				end if

			'case TOK.Tidentifier:
			case else
				_name = token().identifier
				_lineNum = token().lineNumber
				parseToken()

				' Array
				if( token().tok = TOK.Topenparen ) then _name += parseArray()

				' fieldname : bits As DataType [= initializer]
				if( token().tok = TOK.Tcolon ) then
					parseToken( TOK.Tcolon )
					if( token().tok = TOK.Tnumbers ) then parseToken( TOK.Tnumbers ) else return false
				end if
			
				if( token().tok = TOK.Tas ) then

					parseToken( TOK.Tas )

					if( token().tok = TOK.Tconst ) then parseToken( TOK.Tconst )

					_type = getVariableType()
					if( len( _type ) > 0 ) then

						parseToken()
						do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
							_type += "*"
							parseToken()
						loop

						' [= initializer]
						if( token().tok = TOK.Tassign ) then
							parseToken( TOK.Tassign )
							if( token().tok = TOK.Tstring OR token().tok = TOK.Tidentifier OR token().tok = TOK.Tnumbers ) then parseToken() else return false
						end if
						JSONast->addChild( _name, str( B_VARIABLE ), _protection, _type, "", str( _lineNum ) )
						'activeASTnode.addChild( _name, B_VARIABLE, _protection, _type, null, _lineNum );
					end if
				end if

		end select
	loop

	' After exit loop......
	if( token().tok = TOK.Tend AND next1().tok = B_KIND ) then
		parseToken()
		parseToken()
		JSONast->toFatherLayer()
		JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
		return true
	end if

	return false

end function

function CParser.parseType( bClass as boolean = false ) as boolean

	dim as string		_name, _param, _type, _base
	dim as integer		_lineNum, _kind

	
	if( cbool( token().tok = TOK.Ttype ) OR cbool( token().tok = TOK.Tunion ) OR ( bClass = true AND cbool( token().tok = TOK.Tclass ) ) ) then
		select case( token().tok )
			case TOK.Tunion
				_kind = B_UNION
			case TOK.Ttype
				_kind = B_TYPE
			case TOK.Tclass
				_kind = B_CLASS
			case else
		end select
		
		parseToken()

		_name = token().identifier
		_lineNum = token().lineNumber

		parseToken( TOK.Tidentifier )

		if( token().tok = TOK.Tas ) then 
			parseToken( TOK.Tas )

			' Function pointer
			if( token().tok = TOK.Tfunction OR token().tok = TOK.Tsub ) then
				
				if( token().tok = TOK.Tfunction ) then _kind = B_FUNCTION else _kind = B_SUB
				
				parseToken()

				if( token().tok = TOK.Tstdcall OR token().tok = TOK.Tcdecl OR token().tok = TOK.Tpascal ) then parseToken()

				' like " Declare Function app_oninit_cb WXCALL () As wxBool "
				if( token().tok = TOK.Tidentifier ) then parseToken( TOK.Tidentifier )

				' Overload
				if( token().tok = TOK.Toverload ) then parseToken( TOK.Toverload )

				' Alias "..."
				if( token().tok = TOK.Talias ) then
					parseToken( TOK.Talias )
					if( token.tok = TOK.Tstrings ) then parseToken( TOK.Tstrings ) else return false
				end if

				dim as string	_returnType

				if( token().tok = TOK.Topenparen ) then _param = parseParam( true )

				if( token().tok = TOK.Tas ) then
					parseToken( TOK.Tas )

					if( token().tok = TOK.Tconst ) then parseToken( TOK.Tconst )

					_returnType = getVariableType()
					
					if( len( _returnType ) > 0 ) then
						parseToken()
						do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
							_returnType += "*"
							parseToken()
						loop

						_type = _returnType + _param
						JSONast->addChild( _name, str( _kind ), "", _type, "", str( _lineNum ) )
						'activeASTnode.addChild( _name, _kind, null, _type, null, _lineNum );
						
						return true
					end if
				end if

				if( token().tok = TOK.Tstatic OR token().tok = TOK.Texport  ) then parseToken()
				
				if( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then 'SUB
					JSONast->addChild( _name, str( _kind ), "", _param, "", str( _lineNum ) )
					'activeASTnode.addChild( _name, _kind, null, _param, null, _lineNum );
					return true
				end if		
			end if

			if( token().tok = TOK.Tconst ) then parseToken( TOK.Tconst )

			_type = getVariableType()
			if( len( _type ) > 0 ) then 
				parseToken()

				do while( token().tok = TOK.Tptr OR token().tok = TOK.Tpointer )
					_type += "*"
					parseToken()
				loop
			end if

			if( token().tok = TOK.Tcolon OR token().tok = TOK.Teol ) then
				JSONast->addChild( _name, str( B_ALIAS ), "", _type, "", str( _lineNum ) )
				'activeASTnode.addChild( _name, B_ALIAS, null, _type, null, _lineNum );
				return true
			end if
		else
			if( token().tok = TOK.Textends ) then
				parseToken( TOK.Textends )
				_base = token().identifier
				parseToken( TOK.Tidentifier )
			end if

			if( token().tok = TOK.Tfield ) then
				if( next1().tok = TOK.Tassign ) then 
					parseToken( TOK.Tfield )
					parseToken( TOK.Tassign )
					parseToken( TOK.Tnumbers )
				else
					return false
				end if
			end if

			if( token().tok = TOK.Teol ) then
				if( bClass = true ) then
					JSONast->addChild( _name, str( B_CLASS ), "", "", _base, str( _lineNum ) )
					JSONast->toChildLayer()
					'activeASTnode = activeASTnode.addChild( _name, B_CLASS, null, null, _base, _lineNum )
					lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
				else
					JSONast->addChild( _name, str( _kind ), "", "", _base, str( _lineNum ) )
					JSONast->toChildLayer()
					'activeASTnode = activeASTnode.addChild( _name, _kind, null, null, _base, _lineNum );
					lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
				end if
				parseToken( TOK.Teol )
				parseTypeBody( _kind )
			end if
		end if
		
		return true
	end if

	return false

end function 

function CParser.parseEnum() as boolean

	dim as string	_name, _param, _type, _base
	dim as integer	_lineNum

	if( token().tok = TOK.Tenum ) then

		parseToken( TOK.Tenum )

		if( token().tok = TOK.Tidentifier ) then
			_name = token().identifier
			parseToken( TOK.Tidentifier )
		end if

		_lineNum = token().lineNumber

		if( token().tok = TOK.Teol ) then

			JSONast->addChild( _name, str( B_ENUM ), "", "", _base, str( _lineNum ) )
			JSONast->toChildLayer()
			'activeASTnode = activeASTnode.addChild( _name, B_ENUM, null, null, _base, _lineNum );
			lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
			parseToken( TOK.Teol )

			do while( token().tok <> TOK.Tend AND next1().tok <> TOK.Tenum )
				
				if( token().tok = TOK.Tidentifier ) then
					_name = token().identifier
					_lineNum = token().lineNumber
					parseToken( TOK.Tidentifier )

					if( token().tok = TOK.Tassign ) then parseToken( TOK.Tassign )

					' Pass the maybe complicated express
					do while( token().tok <> TOK.Teol AND token().tok <> TOK.Tcolon )
						parseToken()
					loop

					if( token().tok = TOK.Teol OR token().tok = TOK.Tcolon ) then
						JSONast->addChild( _name, str( B_ENUMMEMBER ), "", "", "", str( _lineNum ) )
						'activeASTnode.addChild( _name, B_ENUMMEMBER, null, null, null, _lineNum );
						parseToken()
					end if
				else
					exit do
				end if
			loop
		end if
		
		return true
		
	end if

	return false

end function

function CParser.parseScope() as boolean

	if( next1().tok = TOK.Teol OR next1().tok = TOK.Tcolon ) then
		JSONast->addChild( "", str( B_SCOPE ), "", "", "", str( token().lineNumber ) )
		JSONast->toChildLayer()
		'activeASTnode = activeASTnode.addChild( null, B_SCOPE, null, null, null, token().lineNumber );
		lineStack->push( JSONast->changeMemberValue( "tail", "" ) )
		parseToken( TOK.Tscope )
	else
		return false
	end if

	return true
	
end function	

function CParser.parseEnd() as boolean

	parseToken( TOK.Tend )
	
	select case( token().tok )
		case TOK.Tsub, TOK.Tfunction, TOK.Tproperty, TOK.Toperator, TOK.Tconstructor, TOK.Tdestructor, TOK.Ttype, TOK.Tenum, TOK.Tunion, TOK.Tnamespace, TOK.Tscope
			parseToken()
			if( JSONast->getLayer() > 0 ) then
				JSONast->toFatherLayer()
				JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
			end if
			'if( activeASTnode.getFather() !is null ) activeASTnode = activeASTnode.getFather();
			return true

		case else
			'parseToken()
			
	end select

	return false

end function

function CParser.parse( fullPath as string ) as boolean

	JSONast = new ToJSON( fullPath )
	lineStack->push( JSONast->changeMemberValue( "tail", "" ) )

	do while( tokenIndex < pTokenList->getSize() )
		
		select case ( pTokenList->getValueByIndex( tokenIndex ).tok  )
			case TOK.Tprivate
				parseToken( TOK.Tprivate )
				if( token().tok = TOK.Tsub OR token().tok = TOK.Tfunction ) then parseProcedure( false, "private" )

			case TOK.Tprotected
				parseToken( TOK.Tprotected )
				if( token().tok = TOK.Tsub OR token().tok = TOK.Tfunction ) then parseProcedure( false, "protected" )
				
			case TOK.Tpublic
				parseToken( TOK.Tpublic )
				if( token().tok = TOK.Tsub OR token().tok = TOK.Tfunction ) then parseProcedure( false, "public" )
			
			case TOK.Tpound
				parsePreprocessor()

			case TOK.Tconst, TOK.Tcommon, TOK.Tstatic, TOK.Tdim, TOK.Tredim
				parseVariable()

			case TOK.Tvar
				parserVar()

			case TOK.Tfunction, TOK.Tsub, TOK.Tproperty, TOK.Tconstructor, TOK.Tdestructor
				parseProcedure( false, "" )
				
			case TOK.Toperator
				parseOperator( false, "" )

			case TOK.Tend
				parseEnd()

			case TOK.Ttype
				parseType()

			case TOK.Tclass
				parseType( true )

			case TOK.Tunion
				parseType()
				
			case TOK.Tenum
				parseEnum()

			case TOK.Tscope
				parseScope()

			case TOK.Tnamespace
				parseNamespace()
				
			case TOK.Tdeclare
				parseToken( TOK.Tdeclare )
				if( token().tok = TOK.Tfunction OR token().tok = TOK.Tsub OR token().tok = TOK.Tconstructor OR token().tok = TOK.Tdestructor OR token().tok = TOK.Tproperty ) then 
					parseProcedure( true, "" )
				elseif( token().tok = TOK.Toperator ) then
					parseOperator( true, "" )
				end if

			case TOK.Tendmacro
				if( JSONast->getLayer > 0 ) then
					JSONast->toFatherLayer()
					JSONast->insertText( lineStack->pop(), str( token().lineNumber ) )
				end if
				'activeASTnode = activeASTnode.getFather();					

			case else
				tokenIndex += 1
				'Stdout( tokenIndex );
				'Stdout( "   " ~ token().identifier ).newline;
		end select
		
	loop

	'printAST( head );
	
	' Write to head's tail the max integer value
	JSONast->insertText( lineStack->pop(), "2147483647" )
	return true
end function

sub CParser.printAST()

	print JSONast->getDocument()

end sub

sub CParser.save( fullPath as string )

	open fullPath for output As #1
    print #1, JSONast->getDocument()
    close #1

end sub