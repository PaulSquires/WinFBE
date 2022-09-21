/'	
	This Source Code Form is subject to the terms of the Mozilla Public
	License, v. 2.0. If a copy of the MPL was not distributed with this
	file, You can obtain one at http://mozilla.org/MPL/2.0/. 
'/
#include once "crt.bi"

enum jsonDataType
	malformed = -1
	jsonNull = 0
	jsonObject
	jsonArray
	jsonNumber
	jsonString
	jsonBool
end enum

namespace fbJsonInternal

enum jsonError
	arrayNotClosed
	objectNotClosed
	stringNotClosed
	invalidValue
	invalidEscapeSequence
	invalidNumber
	expectedKey
	expectedValue
	unexpectedToken
	invalidCodepoint
end enum

enum jsonToken
	tab = 9
	newLine = 10
	space = 32
	quote = 34
	comma = 44
	colon = 58
	squareOpen = 91
	backSlash = 92
	forwardSlash = 47
	squareClose = 93
	curlyOpen = 123
	curlyClose = 125
	minus = 45
	plus = 43
end enum

enum parserState
	none = 0
	keyToken = 1
	keyTokenClosed = 2
	valueToken = 3
	valueTokenClosed = 4
	nestEnd = 5
	resetState
end enum

' Allows us to interact directly with the FB-Internal string-structure.
' Don't use it, unless you know what you're doing.
type fbString
    dim as byte ptr stringData
    dim as integer length
    dim as integer size
end type

const replacementChar as string  = "?"

'declare function validateCodepoint(byref codepoint as ubyte) as boolean
declare sub FastSpace(byref destination as string, length as uinteger)
declare sub FastLeft(byref destination as string, length as uinteger)
declare sub FastMid(byref destination as string, byref source as byte ptr, start as uinteger, length as uinteger)
declare function isInString(byref target as string, query as byte) as boolean
declare sub LongToUft8(byref codepoint as long, byref result as string) 
declare sub SurrogateToUtf8(surrogateA as long, surrogateB as long, byref result as string)
declare function areEqual(byref stringA as string, byref stringB as string) as boolean
declare function DeEscapeString(byref escapedString as string) as boolean

sub FastSpace(byref destination as string, length as uinteger)
	dim as fbString ptr destinationPtr = cast(fbString ptr, @destination)
	if ( destinationPtr->size < length ) then 
		deallocate destinationptr->stringdata
		destinationPtr->stringData = allocate(length+1)
	end if
    memset(destinationPtr->stringData, 32, length)
    destinationPtr->length = length
end sub

sub FastCopy(byref destination as string, byref source as string)
	dim as fbString ptr destinationPtr = cast(fbString ptr, @destination)
	dim as fbString ptr sourcePtr = cast(fbString ptr, @source)
	if (sourcePtr->length = 0 and destinationPtr->length = 0) then return
	if ( destinationPtr->size < sourcePtr->size ) then 
		deallocate destinationptr->stringdata
		destinationPtr->length = sourcePtr->length
		destinationPtr->size = sourcePtr->length
		destinationPtr->stringData = allocate(sourcePtr->length+1)
		destinationPtr->stringData[sourcePtr->length] = 0
	end if
	
	' We allocate an extra byte here because FB tries to write into that extra byte when doing string copies.
	' The more "correct" mitigation would be to allocate up to the next blocksize (32 bytes), but that's slow.
	memcpy( destinationPtr->stringData, sourcePtr->stringData, destinationPtr->size)
	
end sub


sub FastLeft(byref destination as string, length as uinteger)
	dim as fbString ptr destinationPtr = cast(fbString ptr, @destination)
	dim as any ptr oldPtr = destinationPtr->stringData
	destinationPtr->length = IIF(length < destinationPtr->length, length, destinationPtr->length)
	'destinationPtr->size = destinationPtr->length
end sub

sub FastMid(byref destination as string, byref source as byte ptr, start as uinteger, length as uinteger)
	dim as fbString ptr destinationPtr = cast(fbString ptr, @destination)
	
	' Setting the length and size of the string, so the runtime knows how to handle it properly.
	if ( destinationPtr->size < length ) then 
		if ( destinationPtr->size ) then deallocate destinationPtr->stringData

		' We allocate an extra byte here because FB tries to write into that extra byte when doing string copies.
		' The more "correct" mitigation would be to allocate up to the next blocksize (32 bytes), but that's slow.
		destinationPtr->stringData = allocate(length+1)
		destinationPtr->stringData[length] = 0
	end if
	destinationPtr->length = length
	destinationPtr->size = length
	memcpy( destinationPtr->stringData, source+start, destinationPtr->size )
end sub

function isInString(byref target as string, query as byte) as boolean
	dim as fbstring ptr targetPtr = cast(fbstring ptr, @target)
	if ( targetPtr->size = 0 ) then return false
	
	return memchr( targetPtr->stringData, query, targetPtr->size ) <> 0
end function

sub LongToUft8(byref codepoint as long, byref result as string)	
	if codePoint <= &h7F then
		fastSpace(result, 1)
		result[0] = codePoint
	endif
	
	if (&hD800 <= codepoint AND codepoint <= &hDFFF) OR _
		(codepoint > &h10FFFD) then
		result = replacementChar
		return
	end if
	
	if (codepoint <= &h7FF) then
		fastSpace(result, 2)
		result[0] = &hC0 OR (codepoint SHR 6) AND &h1F 
		result[1] = &h80 OR codepoint AND &h3F
		return
	end if
	if (codepoint <= &hFFFF) then
		fastSpace(result, 3)
        result[0] = &hE0 OR codepoint SHR 12 AND &hF
        result[1] = &h80 OR codepoint SHR 6 AND &h3F
        result[2] = &h80 OR codepoint AND &h3F
        return
    end if
	
	fastSpace(result, 4)
	result[0] = &hF0 OR codepoint SHR 18 AND &h7
	result[1] = &h80 OR codepoint SHR 12 AND &h3F
	result[2] = &h80 OR codepoint SHR 6 AND &h3F
	result[3] = &h80 OR codepoint AND &h3F
    
	return
end sub

sub SurrogateToUtf8(surrogateA as long, surrogateB as long, byref result as string)
	dim as long codepoint = 0
    if (&hD800 <= surrogateA and surrogateA <= &hDBFF) then
		if (&hDC00 <= surrogateB and surrogateB <= &hDFFF) then
			codepoint = &h10000
			codepoint += (surrogateA and &h03FF) shl 10
			codepoint += (surrogateB and &h03FF)
		end if
	end if
	
	
	if ( codePoint = 0 ) then
		result = replacementChar
	end if
	FastSpace(result, 4)
	result[0] = &hF0 OR codepoint SHR 18 AND &h7
	result[1] = &h80 OR codepoint SHR 12 AND &h3F
	result[2] = &h80 OR codepoint SHR 6 AND &h3F
	result[3] = &h80 OR codepoint AND &h3F
end sub

function areEqual(byref stringA as string, byref stringB as string) as boolean
	dim as fbString ptr A = cast(fbString ptr, @stringA)
	dim as fbString ptr B = cast(fbString ptr, @stringB)

	if (A->length <> B->length) then
		return false
	end if
	
	if (A = B) then
		return true
	end if
	return strcmp(A->stringData, B->stringData) = 0
end function

function DeEscapeString(byref escapedString as string) as boolean
	dim as uinteger length = len(escapedString)-1

	dim as uinteger trimSize = 0
	dim as boolean isEscaped
	for i as uinteger = 0 to length 
		' 92 is backslash
		
		if ( escapedString[i] = 92 andAlso isEscaped = false) then
			isEscaped = true
			if ( i < length ) then
				select case as const escapedString[i+1]
				case 34, 92, 47: ' " \ /
					' Nothing to do here.
				case 98 ' b
					escapedString[i+1] = 8 ' backspace
				case 102 ' f
					escapedString[i+1] = 12
				case 110 ' n
					escapedString[i+1] = 10
				case 114 ' r
					escapedString[i+1] = 13
				case 116 ' t
					escapedString[i+1] = 9 ' tab
				case 117 ' u
					'magic number '6': 2 for "\u" and 4 digit.
					if (i+5 > length) then
						return false
					end if
					dim sequence as string = mid(escapedString, i+3, 4)

					dim pad as integer
					dim as string glyph
					dim as long codepoint = strtoull(sequence, 0, 16)
					if (&hD800 <= codepoint and codepoint <= &hDFFF) then
						dim secondSurrogate as string = mid(escapedString, i+7+2, 4)
						if (len(secondSurrogate) = 4) then
							SurrogateToUtf8(codepoint, strtoull(secondSurrogate, 0, 16), glyph)
							pad = 12 - len(glyph)
						else
							return false
						end if
					elseif (codepoint > 0 orElse sequence = "0000") then
						LongToUft8(codepoint, glyph)
						pad = 6 - len(glyph)
					end if
					
					if (len(glyph) = 0) then
						return false
					end if
					
					for j as integer = 0 to len(glyph)-1
						escapedString[i+j+pad] = glyph[j]
					next
					i += pad -1
					trimSize += pad -1
				case else
					return false
				end select
				trimSize+=1
			end if
		elseif ( trimSize > 0 ) then
			isEscaped = false
			escapedString[i-trimsize] = escapedString[i]
		end if
	next
	if ( trimSize > 0 ) then
		fastleft(escapedString, length - trimSize+1)
	end if
	return true
end function

function isValidDouble(byref value as string) as boolean	
	' Note to reader: 
	' This function is strictly for validation as far as the IETF rfc7159 is concerned.
	' This might be more restrictive than you need it to be outside JSON use.
	
	' It's also a bit nuts. Callgrind is such a fascinating thing.
	

	dim as fbString ptr valuePtr = cast(fbString ptr, @value)
	
	if (value[0] = 48) then
		if ( valuePtr->length > 2) then
			select case value
				' Shorthands for "0" that won't pass this validation otherwise.
				case  "0e1","0e+1","0E1", "0E+1", "0e-1", "0E-1"
					value = "0"
					return true
				end 
			end select
		elseif ( valuePtr->length = 1) then
			return true
		end if
	end if
	
	dim as integer period = 0, exponent = 0, sign = 0
	
	' Yay for manual loop-unrolling.
	select case as const value[0]
		case 48: ' 0. No leading zeroes allowed.
			if (valuePtr->length > 1 and value[1] <> 101 and value[1] <> 69  and value[1] <> 46 ) then
				return false
			end if
		case 49 to 57 ' 1 - 9
			' do nothing
		case 101, 69: 'e, E
			return false
		case 46: ' .
			return false
		case 45: ' -
			sign += 1
			if (valuePtr->length = 1) then
				return false
			elseif (valuePtr->length >= 3) then
				if (value[1] = 48 andAlso (value[2] >= 48 and value[2] <= 57)) then
					return false
				end if
			end if
		case else
			return false
	end select
	
	
	for i as integer = 1 to valuePtr->length-1
		select case as const value[i]
			case 48 to 57 ' 0 - 9
				' do nothing
			case 101, 69: 'e, E
				if (i = valuePtr->length-1) then
					return false
				end if
				if (exponent > 0) then
					return false
				end if
				if (value[i-1] = 46) then
					return false
				end if
				exponent += 1
			case 46: ' .
				if (i = valuePtr->length-1) then
					return false
				end if
				if (period > 0 or exponent > 0 ) then
					return false
				end if
				if ( value[i-1] = 45) then return false
				period += 1
			case 45, 43 ' -, +
				if ((value[i-1] <> 101 and value[i-1] <> 69)) then
					return false
				end if
				if (i = valuePtr->length-1) then
					return false
				end if
			
			case else
				return false
		end select
	next
	
	if (exponent = 0 and period = 0 and (sign = 0 orElse valuePtr->length > 1) and valuePtr->length < 309) then
		return true
	end if
	value = str(cdbl(value))
	return not(valuePtr->length = 1 andAlso (value = "0"))
end function


sub FastTrimWhitespace(byref destination as string)
	dim as uinteger start, i, strLen = len(destination)
	if (strLen = 0) then return
	
	while start < strLen-1 andAlso destination[start] = 32 orElse destination[start] = 9 orElse destination[start] = 10  orElse destination[start] = 13
		start += 1
	wend
	if start = strLen -1 then 
		destination = ""
		return
	end if
	
	i = strLen -1
	while i > 1 andAlso destination[i] = 32 orElse destination[i] = 9 orElse destination[i] = 10  orElse destination[i] = 13
		i -= 1
	wend
	if (start = 0 and i = strLen-1) then return
	
	fastMid(destination, cast(byte ptr,strptr(destination)), start, i - start + 1)
end sub

end namespace

type JsonBase extends object
	protected:
		_dataType as jsonDataType = jsonNull
		_value as string
		_error as string
		_children as JsonBase ptr ptr = 0
		_parent as JsonBase ptr = 0
		_key as string
		_count as integer = -1
		
		declare sub Parse(jsonString as ubyte ptr, endIndex as integer)
		declare sub SetMalformed()
		declare function AppendChild(newChild as JsonBase ptr, override as boolean = false) as boolean
		declare sub setErrorMessage(errorCode as fbJsonInternal.jsonError, jsonstring as byte ptr, position as uinteger)
	public:
		declare static function ParseJson(inputString as string) byref as JsonBase
		declare constructor()
		declare constructor(byref jsonString as string)
		
		declare destructor()
				
		declare property Count() as integer
		declare property DataType() as jsonDataType
		
		declare property Parent() byref as JsonBase
		
		declare operator LET(A as JsonBase)
		
		declare sub Parse(byref jsonString as string)
			
		declare function ContainsKey(byref key as string) as boolean
		
		declare function getError() as string
end type

type JsonItem extends JsonBase
	public:
		declare constructor()
		declare constructor(byref jsonString as string)
		
		declare destructor()
		
		declare property Key () as string
		declare property Key (value as string)
		
		declare property Value(byref newValue as string)
		declare property Value() as string
		
		declare property Count() as integer
		declare property DataType() as jsonDataType
		
		declare operator [](key as string) byref as JsonItem
		declare operator [](index as integer) byref as JsonItem
				
		'declare operator LET(A as JsonItem)
		
		declare function ToString(level as integer = 0) as string
		
		declare function AddItem(key as string, value as string) as boolean
		declare function AddItem(key as string, item as JsonItem) as boolean
		
		declare function AddItem(value as string) as boolean
		declare function AddItem(item as JsonItem) as boolean
					
		declare function RemoveItem(key as string) as boolean
		declare function RemoveItem(index as integer) as boolean
end type

sub JsonBase.setErrorMessage(errorCode as fbJsonInternal.jsonError, jsonstring as byte ptr, position as uinteger) 
	using fbJsonInternal
	dim as integer lineNumber = 1
	dim as integer linePosition = 0
	dim as integer lastBreak = 0
	dim as string lastLine
	
	for j as integer = 0 to position
		if ( jsonString[j] = 10 ) then
			lineNumber +=1
			linePosition = 0
			lastBreak = j
		end if
		linePosition +=1
	next
	
	select case as const errorCode
		case arrayNotClosed:
			this._error = "Array was not properly closed. Expected ']' at position "& linePosition &" on line "& lineNumber &", found '"& chr(jsonstring[position]) &"' instead."
		case objectNotClosed:
			this._error = "Object was not properly closed. Expected '}' at position "& linePosition &" on line "& lineNumber &", found '"& chr(jsonstring[position]) &"' instead."
		case stringNotClosed:
			this._error = "String value or key was not closed. Expected '""', found "& chr(jsonstring[position]) &" at position "& linePosition &" on line "& lineNumber &"."
		case invalidValue:
			this._error = "Invalid value '"& this._value &"' encountered at position "& linePosition &" on line "& lineNumber &"."
		case invalidEscapeSequence:
			this._error = "Could not de-escape '"& this._value &"' encountered at position "& linePosition &" on line "& lineNumber &"."
		case invalidNumber:
			this._error = "Invalid number '"& this._value &"' encountered at position "& linePosition &" on line "& lineNumber &"."
		case expectedKey:
			this._error = "Expected a key at position "& linePosition &" on line "& lineNumber &", found '"& chr(jsonstring[position]) &"' instead."
		case expectedValue:
			this._error = "Expected a value at position "& linePosition &" on line "& lineNumber &", found '"& chr(jsonstring[position]) &"' instead."
		case unexpectedToken:
			this._error = "Unexpected token '"& chr(jsonstring[position]) &"' at "& linePosition &" on line "& lineNumber &"."
		case invalidCodepoint
			this._error = "Invalid UTF-8 bytesequence encountered: "& hex(jsonstring[position])  &"' at "& linePosition &" on line "& lineNumber &"."
	end select
	
	this._value = ""
	this.SetMalformed()
	#ifdef fbJSON_debug
		print "fbJSON Error: "& this._error
		'end -1
	#endif
end sub

constructor JsonBase()
	' Nothing to do
end constructor

constructor JsonBase(byref jsonString as string)
	fbJsonInternal.FastTrimWhitespace(jsonstring)
	this.Parse(strptr(jsonString), len(jsonstring)-1)
end constructor

destructor JsonBase()
	if (this._count >= 0 and this._children <> 0) then
		for i as integer = 0 to this._count
			delete this._children[i]
		next
		this._count = -1
		deallocate(this._children)
	end if
end destructor

operator JsonBase.LET(copy as JsonBase)
	this.destructor()
	fbJsonInternal.FastCopy(this._key, copy._key)
	fbJsonInternal.FastCopy(this._value, copy._value)
	
	this._dataType = copy._dataType
	this._error = copy._error
	this._count = copy._count
	
	if ( copy._count >= 0) then
		this._children = allocate(sizeOf(JsonBase ptr) * (copy._count+1))
		for i as integer = 0 to copy._count
			this._children[i] = callocate(sizeOf(JsonBase))
			*this._children[i] = *copy._children[i]
		next
	end if
end operator

property JsonBase.Parent() byref as JsonBase
	if ( this._parent <> 0 ) then
		return *this._parent
	end if
	
	return *new JsonBase()
end property

property JsonBase.Count() as integer
	return this._count + 1
end property

property JsonBase.DataType() as jsonDataType
	return this._datatype
end property

function JsonBase.AppendChild(newChild as JsonBase ptr, override as boolean = false) as boolean
	if ( newChild = 0 ) then return false	
	if ( this._datatype = jsonObject ) then 
		for i as integer = 0 to this._count
			if ( fbJsonInternal.AreEqual(this._children[i]->_key, newChild->_key) ) then
				if (override) then
					delete this._children[i]
					newChild->_parent = @this
					this._children[i] = newChild
					return true
				else 
					return false
				end if
			end if
		next
	end if
	
	newChild->_parent = @this
	this._count += 1
	
	' I think allocating 2 elements at a time is a decent compromise between memory and speed.
	' And it does cut the number of reallocations in half.
	if (this._count mod 2 = 0) then
		if this._children = 0 then
			this._children = allocate(sizeof(JsonBase ptr) * (this._count+2))
		else
			this._children = reallocate(this._children, sizeof(JsonBase ptr) * (this._count+2))
		end if
		
		if this._children = 0 then
			this.setMalformed()
			return false
		end if
	end if
	this._children[this._count] = newChild
	
	if ( newChild->_datatype = malformed ) then
		this.SetMalformed()
	end if
	return true
end function

function JsonBase.ContainsKey(byref newKey as string) as boolean
	if ( this._datatype <> jsonObject ) then return false
	
	for i as integer = 0 to this._count
		if ( fbJsonInternal.areEqual(this._children[i]->_key, newKey ) ) then
			return true
		end if
	next
	return false
end function

sub JsonBase.SetMalformed()
	this._datatype = malformed
	if (this._parent <> 0) then
		dim item as JsonBase ptr = this._parent
		item->_error = this._error
		do
			if (item->_parent <> 0) then
				item->_parent->_error = item->_error
			end if
			item->_datatype = malformed
			item = item->_parent
			
		loop until item = 0
	end if
end sub

sub JsonBase.Parse(jsonString as ubyte ptr, endIndex as integer) 
	using fbJsonInternal
	if (endIndex < 0) then
		this.setMalformed()
		this._error = "No data!"
		return
	end if
	
	dim currentType as jsonDataType = -2
	
	' Objects we will work with:
	dim currentItem as JsonBase ptr = @this
	dim as JsonBase ptr child
	
	' key states and variables for the main parsing:
	dim as uinteger parseStart = 1, parseEnd = endIndex -1
	dim as integer valueStart
	dim as parserState state
	dim as boolean isStringOpen
	dim as boolean stringIsEscaped = false
	dim as byte unicodeSequence
	dim as boolean isEscaped = false
	
	' To handle trimming, we use these:
	dim as integer valueEnd
	dim as boolean trimLeftActive = false
	
	if ( jsonstring[0] = jsonToken.CurlyOpen ) then
		if ( jsonString[endIndex] = jsonToken.CurlyClose ) then
			currentItem->_datatype = jsonObject
			state = parserState.none
		else
			currentItem->setErrorMessage(objectNotClosed, jsonstring, endIndex)
			return
		end if
	elseif ( jsonstring[0] = jsonToken.SquareOpen ) then
		if (jsonString[endIndex] = jsonToken.SquareClose ) then
			currentItem->_dataType = jsonArray
			valueStart = 1
			trimLeftActive = true
			state = valueToken
			currentType = -2
		else
			currentItem->setErrorMessage(arrayNotClosed, jsonstring, endIndex)
			return
		end if
	else
		parseStart = 0
		parseEnd = endIndex
		state = valueToken
	end if
			
	' Skipping the opening and closing brackets makes things a bit easier.
	for i as integer = parseStart to parseEnd
		select case as const jsonstring[i]		
			' These codepoints are straight up invalid no matter what:
			case 192, 193, 245 to 255:
				currentItem->setErrorMessage(invalidCodepoint, jsonstring, i)
				goto cleanup
			case 237
				' TODO Validate against surrogate pairs, which are invalid in UTF-8.
				currentItem->setErrorMessage(invalidCodepoint, jsonstring, i)
				goto cleanup

			case jsonToken.Quote
				if (isEscaped = false) then
					isStringOpen = not(isStringOpen)
					if ( currentItem->_datatype = jsonObject ) then
						if state = none then
							state = keyToken
							valueStart = i+1
						elseif state = keyToken then
							child = new JsonBase()
							fastmid(child->_key, jsonString, valuestart,  i - valueStart)
							if ( isInString(child->_key, jsonToken.backslash) <> 0 ) then
								if ( DeEscapeString(child->_key) = false ) then
									child->setErrorMessage(invalidEscapeSequence, jsonstring, i)
								end if
							end if
							state = keyTokenClosed
						end if
					end if
				end if
				goto utf8Validation ' Fall-through to the else-case:
			case else
				utf8Validation:
				' UTF-8 length validation:
				if ( jsonstring[i] > 127 andAlso jsonstring[i] SHR 6 = &b10 ) then
					unicodeSequence -= 1
					if (unicodeSequence < 0 ) then
						currentItem->setErrorMessage(invalidCodepoint, jsonstring, i)
						goto cleanup
					end if
				else
					if (unicodeSequence > 0) then
						currentItem->setErrorMessage(invalidCodepoint, jsonstring, i)
						goto cleanup
					end if
					select case as const jsonString[i] ' shr 4 
						case 192 to 223
							unicodeSequence = 1
						case 224 to 239
							unicodeSequence = 2
						case 240 to 247
							unicodeSequence = 3
						case else
							unicodeSequence = 0
					end select
				end if
		end select
		' When not in a string, we can handle the complicated suff:
		if ( isStringOpen = false ) then
			' Note: Not a single string-comparison in here.
			select case as const jsonstring[i]
				case jsonToken.BackSlash
					currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
					goto cleanup

				case jsonToken.Colon:
					if ( state = keyTokenClosed ) then 
						state = valueToken
						currentType = -2
						trimLeftActive = true
						valueStart = i+1
					else
						currentItem->setErrorMessage(expectedKey, jsonstring, i)
						goto cleanup
					end if
					
				case jsonToken.Comma:
					if (i = parseEnd) then
						currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
						goto cleanup
					end if

					if ( state = valueToken ) then
						state = valueTokenClosed
						if valueEnd = 0 then valueEnd = i
					elseif ( state = nestEnd ) then
						state = resetState
					else 
						currentItem->setErrorMessage(expectedKey, jsonstring, i)
						goto cleanup
					end if
					
				case jsonToken.CurlyOpen:
					if ( state = valueToken ) then
						if (child = 0) then child = new JsonBase()
						child->_datatype = jsonobject
						currentItem->AppendChild( child , true)
						currentItem = child
						state = resetState
					else
						currentItem->setErrorMessage(expectedKey, jsonstring, i)
						goto cleanup
					end if
					
				case jsonToken.SquareOpen:
					if ( state = valueToken andAlso valueStart = i ) then
						if (child = 0) then child = new JsonBase()
						child->_datatype = jsonArray
						currentItem->AppendChild( child , true)
						currentItem = child
						state = resetState
					else
						currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
						goto cleanup
					end if
					
				case jsonToken.CurlyClose:
					if (currentItem->_datatype = jsonObject) then
						if (currentItem = 0 or currentItem->_parent = 0) then
							this.setMalformed()
							goto cleanup
						end if
						state = nestEnd
						currentItem->AppendChild(child, true)
						
						currentItem = currentItem->_parent
					else
						currentItem->setErrorMessage(arrayNotClosed, jsonstring, i)
						goto cleanup
					end if
							
				case jsonToken.SquareClose:
					if (currentItem->_datatype = jsonArray ) then
						if state = valueToken andAlso valueEnd = 0 andAlso valueStart <> i then 
							valueEnd = i
							if (child = 0) then 
								child = new jsonBase()
							end if
						end if
						if (currentItem = 0 or currentItem->_parent = 0) then
							this.setMalformed()
							goto cleanup
						end if
						state = nestEnd
						
						currentItem->AppendChild(child, true)
						currentItem = currentItem->_parent
					else
						currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
						goto cleanup
					end if
					
				case jsonToken.Space, jsonToken.Tab, jsonToken.NewLine, 13
					' Here, we count the left trim we need. This is faster than using TRIM() even for a single space in front of the value
					' And most important: It's not slower if we have no whitespaces.
					if ( state = valueToken ) then
						if( trimLeftActive) then
							valueStart = i+1
						else
							if valueEnd = 0 then valueEnd = i
						end if
					end if
					
				case jsonToken.Quote
					' The closing quote get's through to here. We treat is as part of a value, but without throwing errors.
					if ( state = valueToken ) then
						trimLeftActive = false
						valueEnd = i
					end if
				case asc("n"),asc("-"),asc("e"),asc("t"),asc("r"),asc("u"),asc("l"),asc("f"),asc("a"),asc("s"), 48 to 57, asc("E"), asc("+"), asc(".")
					' If we are currently parsing values, add up the length and abort the trim-counting.
					if ( state = valueToken andAlso valueEnd = 0 ) then
						trimLeftActive = false
					else
						currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
						goto cleanup
					end if
					select case currentType
					case -2:
						select case as const jsonString[i]:
							case asc("n"):
								if (jsonString[i+1] = asc("u") _
								    andAlso jsonString[i+2] = asc("l") _
								    andAlso jsonstring[i+3] = asc("l") ) then
									currentType = jsonNull
									i+=3
								else
									currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
									goto cleanup
								end if
								
							case asc("t")
								if (jsonString[i+1] = asc("r") _
								    andAlso jsonString[i+2] = asc("u") _
								    andAlso jsonstring[i+3] = asc("e") ) then
									currentType = jsonBool
									i+=3
								else
									currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
									goto cleanup
								end if
								
							case asc("f"):
								if (jsonString[i+1] = asc("a") _
								    andAlso jsonString[i+2] = asc("l") _
								    andAlso jsonString[i+3] = asc("s") _
								    andAlso jsonstring[i+4] = asc("e") ) then
									currentType = jsonBool
									i+=4
								else
									currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
									goto cleanup
								end if
							case jsonToken.minus, 48 to 57:
								currentType = jsonNumber
							case else							
								currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
								goto cleanup
						end select
					case jsonNull, jsonBool
						currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
						goto cleanup
					end select

				case else:
					
					currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
					goto cleanup
			end select
		else
			if (state = valueToken) then
				currentType = jsonDataType.jsonString
			end if
			select case as const jsonString[i]
				case 0 to 31
					currentItem->setErrorMessage(unexpectedToken, jsonstring, i)
					goto cleanup
				case 92
					stringIsEscaped = true
					if (isEscaped = false) then
						isEscaped = true
					else
						isEscaped = false
					end if
				case else
					if (isEscaped) then isEscaped = false
			end select
			continue for
		end if	
		
		if ( i = parseEnd andAlso state = valueToken) then
			if valueEnd = 0 then valueEnd = i +1
			state = valueTokenClosed
		end if
		
		select case as const state
			case valueTokenClosed, nestEnd
				' because we already know how long the string we are going to parse is, we can skip if it's 0.
				if ( valueEnd > 0 andAlso (child <> 0 or state = valueTokenClosed)) then
					if (child = 0) then child = new JsonBase()
					' The time saved with this is miniscule, but reliably measurable.
					select case as const currentType
						case jsonDataType.jsonString
							FastMid(child->_value, jsonString, valuestart+1, valueEnd - valueStart-1)
							child->_dataType = jsonDataType.jsonString
							if ( stringIsEscaped ) then
								if ( len(child->_value) > 0 andAlso DeEscapeString(child->_value) = false ) then
									child->setErrorMessage(invalidEscapeSequence, jsonstring, i)
								end if
								stringIsEscaped = false
							end if
							
						case jsonNull
							child->_dataType = jsonNull
							FastMid(child->_value, jsonString, valuestart, valueEnd - valueStart)	
						case jsonBool
							' Nesting "select-case" isn't pretty, but fast. Saw this first in the .net compiler.
							FastMid(child->_value, jsonString, valuestart, valueEnd - valueStart)
							
							child->_datatype = jsonBool
							
						case jsonNumber
							fastMid(child->_value, jsonString, valuestart, valueEnd - valueStart)
							if ( isValidDouble(child->_value) ) then
								child->_dataType = jsonDataType.jsonNumber
							else
								child->setErrorMessage(invalidNumber, jsonstring, i)
							end if
													
						case else
							child->setErrorMessage(invalidValue, jsonstring, i)
					end select
					
					if (currentItem->_datatype = jsonNull ) then
						if (i = parseEnd andAlso child->_datatype <> malformed) then
							FastCopy(this._value, child->_value)
							this._datatype = child->_datatype
							this._error = child->_error
							delete child
							child = 0
							return 
						else
							currentItem->setErrorMessage(0, jsonstring, i+1)
							goto cleanup
						end if
					elseif (state = valueTokenClosed) then
						if (child->_datatype = malformed) then
							currentItem->SetMalformed()
						end if
						currentItem->AppendChild(child, true)
					else
						if (child->_parent = 0) then 
							delete child
						end if
					end if
					valueEnd = 0
					child = 0
				end if
				
				if state <> nestEnd then
					goto resetStateJump
				end if
		
			case resetState:
				resetStateJump:
				child = 0
				if ( currentItem->_datatype = jsonArray ) then
					state = valueToken
					currentType = -2
					valueStart = i+1
					valueEnd = 0
					trimLeftActive = true
				else
					currentType = -2
					state = none
				end if
		end select
	next
	
	if (isStringOpen) then
		currentItem->setErrorMessage(stringNotClosed, jsonstring, endIndex)
		goto cleanup 
	end if
	if (state = keyTokenClosed) then
		currentItem->setErrorMessage(expectedKey, jsonstring, endIndex)
		goto cleanup
	end if
	
	return
	cleanup:
		if (child <> 0) then
			delete child
		end if
end sub

sub JsonBase.Parse(byref inputString as string)
	this.destructor()
	this.constructor()
	fbJsonInternal.FastTrimWhitespace(inputString)
	this.Parse( cast (byte ptr, strptr(inputstring)), len(inputString)-1)
end sub

function JsonBase.getError() as string
	return this._error
end function

constructor JsonItem()
	base()
	' Nothing to do
end constructor

constructor JsonItem(byref jsonString as string)
	base(jsonstring)
end constructor

destructor JsonItem()
	base.destructor()
end destructor

operator JsonItem.[](newKey as string) byref as JsonItem	
	if ( this._datatype = jsonObject and this._count > -1 ) then
		for i as integer = 0 to this._count
			if ( this._children[i]->_key = newkey ) then
				return *cast(JsonItem ptr,this._children[i])
			end if
		next
	end if
	
	#ifdef fbJSON_debug
		print "fbJSON Error: Key '"& key & "' not found in object "& this.key
		end -1
	#endif
	return *new JsonItem()
end operator

operator JsonItem.[](index as integer) byref as JsonItem
	if ( index <= this._count ) then
		return *cast(JsonItem ptr,this._children[index])
	end if
	
	#ifdef fbJSON_debug
		print "fbJSON Error: "& index & " out of bounds in "& this.key &". Actual size is "& this.count
		end -1
	#else
		return *new JsonItem()
	#endif
end operator

property JsonItem.Key() as string
	return this._key
end property

property JsonItem.Key(newkey as string)
	if ( this._key = newKey ) then return

	if ( this.Parent.ContainsKey(newKey) ) then return
	
	this._key = newKey
end property

property JsonItem.datatype() as jsonDataType
	return this._dataType
end property

property JsonItem.Count() as integer
	return this._count + 1
end property

property JsonItem.Value( byref newValue as string)
	using fbJsonInternal
	
	' TODO: Optimize this, according to the parser optimizations
	
	' First, handle strings in quotes:
	select case as const newValue[0]
	case jsonToken.Quote
		if ( newValue[len(newValue)-1] = jsonToken.Quote ) then
			this._dataType = jsonString
			this._value = mid(newValue,2, len(newValue)-2)
			if ( DeEscapeString(this._value) = false ) then
				this.setMalformed()
			end if
		else
			this.setMalformed()
		end if
	case 45, 48,49,50,51,52,53,54,55,56,57 '-, 0 - 9
		if (isValidDouble(newValue) ) then
			this._datatype = jsonNumber
			this._value = str(cdbl(newValue))
		else			
			this._datatype = jsonString
			this._value = newValue
				
			if ( DeEscapeString(this._value) = false ) then
				this.setMalformed()
			end if
		end if
	case 110,78, 102,70, 116,84 ' n, f, t
		select case lcase(newValue)
		case "null"
			this._value = newValue
			this._dataType = jsonNull
		case "true", "false"
			this._value = newValue
			this._dataType = jsonBool
		case else:
			' strict vs. nonstrict mode?
			this._datatype = jsonString
			this._value = newValue
				
			if ( DeEscapeString(this._value) = false ) then
				this.setMalformed()
			end if
		end select
	case else
		this._dataType = jsonString
		this._value = newValue
		if ( DeEscapeString(this._value) = false ) then
			this.setMalformed()
		end if
	end select
end property

property JsonItem.Value() as string
	return this._value
end property

function JsonItem.AddItem(newKey as string, newValue as string) as boolean
	if ( len(newKey) = 0 orElse this.containsKey(newKey) ) then
		return false
	end if
	
	if ( this._datatype = jsonNull ) then
		this._datatype = jsonObject
	end if
	
	if ( this._datatype = jsonObject ) then
		dim child as JsonItem ptr = new JsonItem
		child->Value = newValue
		child->_key = newKey
		return this.AppendChild(child)
	end if
	return false
end function

function JsonItem.AddItem(newKey as string, item as JsonItem) as boolean
	if ( len(newKey) = 0 orElse this.containsKey(newKey) ) then
		return false
	end if
	
	if ( this._datatype = jsonNull ) then
		this._datatype = jsonObject
	end if
	
	if ( this._datatype = jsonObject ) then
		dim child as JsonItem ptr = callocate(sizeof(JsonItem))
		*child = item
		child->_key = newKey
		return this.AppendChild(child)
	end if
	return false
end function

function JsonItem.AddItem(newValue as string) as boolean
	if ( this._datatype = jsonArray or this._datatype = jsonNull ) then
		this._datatype = jsonArray
		dim child as JsonItem ptr = new JsonItem
		child->value = newValue
		return this.AppendChild(child)
	end if
	return false
end function

function JsonItem.AddItem(item as JsonItem) as boolean
	if ( this._datatype = jsonArray or this._datatype = jsonNull ) then
		this._datatype = jsonArray
		dim child as JsonItem ptr = callocate(sizeof(JsonItem))
		*child = item
		return this.AppendChild(child) 		
	end if
	return false
end function

function JsonItem.RemoveItem(newKey as string) as boolean
	dim as integer index = -1
	
	if ( this._datatype = jsonObject ) then
		for i as integer = 0 to this._count
			if ( this._children[i]->_key = newkey ) then
				index = i
				exit for
			end if
		next
	end if
	
	return this.RemoveItem(index)
end function

function JsonItem.RemoveItem(index as integer) as boolean
	if ( index <= this._count andAlso index > -1 ) then
		delete this._children[index]
		if ( index < this.Count -1 ) then
			for i as integer = index to this._count 
				this._children[i] = this._children[i+1]
			next
		end if
		
		this._count -= 1
		this._children = reallocate(this._children, sizeof(jsonItem ptr) * this._count)
		
		return true
	end if
	return false
end function

function JsonItem.ToString(level as integer = 0) as string
	dim as string result
	
	' TODO: Clean up this mess.
	if ( this.datatype = jsonObject ) then
		result = "{"
	elseif ( this.datatype = jsonArray ) then
		result = "["
	elseif ( level = 0 ) then
		return this._value
	end if
		
	for i as integer = 0 to this._count 
		result += chr(10) + string((level +1) * 2, " ") 
		if ( this.datatype = jsonObject ) then
			result += """" & this[i]._key & """ : " 
		end if
		
		if ( this[i].Count >= 1 ) then
			result += this[i].toString(level+1)
		else
			if ( this[i].datatype = jsonString) then
				result += """" & this[i]._value & """"
			else
				result += this[i]._value
			end if
		end if
		if ( i < this.Count - 1 ) then
			result += ","
		else
			level -= 1
			result += chr(10)
		end if
		
	next
	
	if this.datatype = jsonObject  then
		result += string((level +1) * 2, " ")  + "}"
	elseif ( this.datatype = jsonArray ) then
		result +=  string((level +1) * 2, " ") +"]"
	end if
	
	return result
end function
