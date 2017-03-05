#Include Once "parserFB\token.bi"
#include once "parserFB\tools.bi"


type CParser
	private:
		as DynamicArray	ptr		pTokenList
		as integer				tokenIndex, nearestTailInsertPos
		as TokenUnit			nullTokenUnit
		as Stack ptr			lineStack
		as ToJSON ptr			JSONast

		declare function parsePreprocessor() as boolean		
		declare function getVariableType() as string
		declare function parseParam( bDeclare as boolean ) as string
		declare function parseArray() as string
		declare function parseVariable() as boolean
		declare function parserVar() as boolean
		declare function parseNamespace() as boolean
		declare function parseOperator( bDeclare as boolean, _protection as string ) as boolean
		declare function parseProcedure( bDeclare as boolean, _protection as string ) as boolean
		declare function parseTypeBody( B_KIND as integer ) as boolean
		declare function parseType( bClass as boolean = false ) as boolean
		declare function parseEnum() as boolean
		declare function parseScope() as boolean
		declare function parseEnd() as boolean
		
		
	public:
		declare Constructor( _pTokenList as DynamicArray ptr )
		declare Destructor()	

		declare function token() as TokenUnit
		declare function next1() as TokenUnit
		declare function next2() as TokenUnit
		declare sub parseToken( t as TOK = TOK.Tidentifier )

		declare function parse( fullPath as string ) as boolean
		declare	sub printAST()
		declare sub save( fullPath as string )

end type
