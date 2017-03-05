#include once "parserFB\tools.bi"

/'
Member function of DynamicArray
'/
constructor DynamicArray()
end constructor

constructor DynamicArray( _size as integer )

	init( _size )
	
end constructor

destructor DynamicArray()

	deallocate( pMemory )

end destructor

function DynamicArray.init( _size as integer ) as TokenUnit ptr

	size		= _size
	pMemory		= CAllocate( _size, sizeof( TokenUnit ) )
	
	return pMemory

end function

function DynamicArray.getValueByIndex( index as integer ) as TokenUnit
	
	dim as TokenUnit		emptyType
	
	if( index < size ) then
		if( pMemory <> 0 ) then return pMemory[index]
	end if
	
	return emptyType

end function

function DynamicArray.getAddressByIndex( index as integer ) as TokenUnit ptr
	
	if( index < size ) then
		if( pMemory <> 0 ) then return ( pMemory + index )
	end if
	
	return 0

end function

function DynamicArray.value( index as integer ) as TokenUnit
	
	return getValueByIndex( index )

end function

function DynamicArray.address( index as integer ) as TokenUnit ptr
	
	return getAddressByIndex( index )

end function	


function DynamicArray.getMemoryPointer() as TokenUnit ptr

	return pMemory
	
end function

function DynamicArray.getSize() as integer

	return size
	
end function





/'
Stack
'/
constructor Stack()
end constructor

constructor Stack( _size as uinteger )

	redim container(1 to _size)

end constructor

destructor Stack()

	erase container
	
end destructor

sub Stack.push( value as integer )
	
	index += 1
	if( index > ubound( container ) ) then redim preserve container(index+99)
	
	container(index) = value

end sub

function Stack.pop() as integer

	if( index > 0 ) then 
		dim as integer tempValue = container( index )
		index -= 1
		return tempValue
	end if
	
	return -1

end function

function Stack.size() as integer

	return index
	
end function




/'
Member function of ToJSON
'/
private function ToJSON.createObject( _name as string, _kind as string, _prot as string, _type as string, _base as string, _line as string ) as string

	dim as string		result
	
	result =	"{" + chr(10)
	result +=	"""line"": " + """" + _line + """," + chr(10)
	result +=	"""tail"": " + """" + _line + """," + chr(10)	
	result +=	"""name"": " + """" + _name + """," + chr(10)
	result +=	"""kind"": " + """" + _kind + """," + chr(10)
	result +=	"""prot"": " + """" + _prot + """," + chr(10)
	result +=	"""type"": " + """" + _type + """," + chr(10)
	result +=	"""base"": " + """" + _base + """," + chr(10)
	result +=	"""sons"": " + "[]" + chr(10)
	result +=	"}"
	
	return result

end function

constructor ToJSON()
end constructor

constructor ToJSON( fullPath as string )

	dim as string	fileTYPE
	
	if( lcase( right( fullPath, 4 ) ) = ".bas" ) then 
		fileTYPE = str( B_BAS )
	elseif ( lcase( right( fullPath, 3 ) ) = ".bi" ) then
		fileTYPE = str( B_BI )
	end if
	
	document = createObject( fullPath, fileTYPE, "", "", "", "" )
	
	layer = 1	
	
end constructor

function ToJSON.toChildLayer() as integer

	layer += 1
	return layer

end function

function ToJSON.toFatherLayer() as integer

	layer -= 1
	return layer

end function

sub ToJSON.addChild( _name as string, _kind as string, _prot as string, _type as string, _base as string, _line as string )
	
	dim as integer		closeBracketCount, index, documentLength = len( document )
	dim as string		leftString, rightString
	
	for index = documentLength - 1 to 0 step -1
		if( chr( document[index] ) = "]" ) then
			closeBracketCount += 1
			if( closeBracketCount = layer ) then exit for
		end if
	next
	
	if( index > 0 ) then
		leftString	= left( document, index )
		rightString	= right( document, documentLength - index )
		
		if( chr( document[index-1] ) = "}" ) then
			' Not first son
			document = leftString + "," + chr(10) + createObject( _name, _kind, _prot, _type, _base, _line ) + rightString
		else
			' First Son
			document = leftString + chr(10) + createObject( _name, _kind, _prot, _type, _base, _line ) + rightString
		end if
		
		
	end if

end sub

function ToJSON.changeMemberValue( _string as string, _value as string ) as integer
	
	dim as integer		index = InStrRev( document, """" + _string + """: " )
	dim as integer		documentLength = len( document )
	dim as integer		result = index + 8
	
	if( index > 0 ) then
		index += 8
		dim as string		leftString	= left( document, index - 1 ) + """" + _value + """"
		
		do until( chr( document[index] ) = chr(10) )
			index += 1
		loop
		
		dim as string		rightString	= right( document, documentLength - index + 1 )
		
		document = leftString + rightString
	end if
	
	return result

end function

sub ToJSON.insertText( insertPos as integer, _insertText as string )

	dim as integer		documentLength = len( document )
	dim as string		leftString	= left( document, insertPos )
	dim as string		rightString	= right( document, documentLength - insertPos )
	
	document = leftString + _insertText + rightString

end sub


function ToJSON.getLayer() as integer

	return layer

end function

function ToJSON.getDocument() as string

	return document
	
end function







' This function by counting_pine
function str_replace(byref haystack as string, byref needle1 as string, byref needle2 as string) as string
   
    dim as integer len1 = len(needle1), len2 = len(needle2)
    dim as integer i
   
    dim as string haystack_ = haystack
   
    i = instr(haystack_, needle1)
    while i
       
        haystack_ = left(haystack_, i - 1) _
                  & needle2 _
                  & mid(haystack_, i + len1)
       
        i = instr(i + len2, haystack_, needle1)
       
    wend
   
    function = haystack_
   
end function