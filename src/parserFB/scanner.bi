#Include Once "parserFB\token.bi"
#include once "parserFB\tools.bi"


type Scanner
	private:
		as DynamicArray	ptr		tokenList
		
		declare sub scan( fileData as string )
	
	public:
		declare Constructor()
		declare Constructor( fullPath as string )
		declare Destructor()
		
		declare function scanFile( fullPath as string ) as DynamicArray ptr
		declare function getTokenCount() as integer
		declare function getTokenList() as DynamicArray ptr
end type