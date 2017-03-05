#include once "parserFB\token.bi"

/'
Class DynamicArray
'/
type DynamicArray
	private:
		as integer			size
		as TokenUnit ptr	pMemory
	
	public:
		Declare Constructor()
		Declare Constructor( size as integer )
		Declare Destructor()
		Declare function init( size as integer ) as TokenUnit ptr
		Declare function getValueByIndex( index as integer ) as TokenUnit
		Declare function getAddressByIndex( index as integer ) as TokenUnit ptr
		Declare function value( index as integer ) as TokenUnit
		Declare function address( index as integer ) as TokenUnit ptr
		Declare function getMemoryPointer() as TokenUnit ptr
		Declare function getSize() as integer
		
end type


/'
Class Stack
'/
type Stack
	private:
		as integer			index
		as integer			container(any)
	
	public:
		Declare Constructor()
		Declare Constructor( _size as uinteger )
		Declare Destructor()
		Declare sub push( value as integer )
		declare function pop() as integer
		declare function size() as integer
end type



/'
Class ToJSON
'/
type ToJSON
	private:
		as string			document
		as integer			layer
		
		Declare function createObject( _name as string, _kind as string, _prot as string, _type as string, _base as string, _line as string ) as string
		
	public:
		Declare Constructor()
		Declare Constructor( fullPath as string )

		Declare function toChildLayer() as integer
		Declare function toFatherLayer() as integer		
		Declare sub addChild( _name as string, _kind as string, _prot as string, _type as string, _base as string, _line as string )
		Declare function changeMemberValue( _string as string, _value as string ) as integer
		Declare sub insertText( insertPos as integer, _insertText as string )
		'declare static function str_replace( byref haystack as string, byref needle1 as string, byref needle2 as string ) as string
		Declare function getLayer() as integer
		Declare function getDocument() as string
end type

' This function by counting_pine
declare function str_replace( byref haystack as string, byref needle1 as string, byref needle2 as string ) as string
