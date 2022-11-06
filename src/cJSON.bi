'   Copyright (c) 2009-2017 Dave Gamble and cJSON contributors

'   Permission is hereby granted, free of charge, to any person obtaining a copy
'   of this software and associated documentation files (the "Software"), to deal
'   in the Software without restriction, including without limitation the rights
'   to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'   copies of the Software, and to permit persons to whom the Software is
'   furnished to do so, subject to the following conditions:

'   The above copyright notice and this permission notice shall be included in
'   all copies or substantial portions of the Software.

'   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'   AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'   OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'   THE SOFTWARE.

#pragma once

#include once "crt/stddef.bi"

#ifdef __FB_WIN32__
	extern "Windows"
#else
	extern "C"
#endif

#define cJSON__h

#ifdef __FB_WIN32__
	#define __WINDOWS__
	#define CJSON_CDECL __cdecl
	#define CJSON_STDCALL __stdcall
	#define CJSON_EXPORT_SYMBOLS
	'' TODO: #define CJSON_PUBLIC(type) __declspec(dllexport) type CJSON_STDCALL
#else
	#define CJSON_CDECL
	#define CJSON_STDCALL
	#define CJSON_PUBLIC(type) type
#endif

const CJSON_VERSION_MAJOR = 1
const CJSON_VERSION_MINOR = 7
const CJSON_VERSION_PATCH = 15
const cJSON_Invalid = 0
const cJSON_False = 1 shl 0
const cJSON_True = 1 shl 1
const cJSON_NULL = 1 shl 2
const cJSON_Number = 1 shl 3
const cJSON_String = 1 shl 4
const cJSON_Array = 1 shl 5
const cJSON_Object = 1 shl 6
const cJSON_Raw = 1 shl 7
const cJSON_IsReference = 256
const cJSON_StringIsConst = 512

type cJSON
	next as cJSON ptr
	prev as cJSON ptr
	child as cJSON ptr
	as long type
	valuestring as zstring ptr
	valueint as long
	valuedouble as double
	string as zstring ptr
end type

type cJSON_Hooks
	malloc_fn as function cdecl(byval sz as uinteger) as any ptr
	free_fn as sub cdecl(byval ptr as any ptr)
end type

type cJSON_bool as boolean
const CJSON_NESTING_LIMIT = 1000
declare function cJSON_Version() as const zstring ptr
declare sub cJSON_InitHooks(byval hooks as cJSON_Hooks ptr)
declare function cJSON_Parse(byval value as const zstring ptr) as cJSON ptr
declare function cJSON_ParseWithLength(byval value as const zstring ptr, byval buffer_length as uinteger) as cJSON ptr
declare function cJSON_ParseWithOpts(byval value as const zstring ptr, byval return_parse_end as const zstring ptr ptr, byval require_null_terminated as cJSON_bool) as cJSON ptr
declare function cJSON_ParseWithLengthOpts(byval value as const zstring ptr, byval buffer_length as uinteger, byval return_parse_end as const zstring ptr ptr, byval require_null_terminated as cJSON_bool) as cJSON ptr
declare function cJSON_Print(byval item as const cJSON ptr) as zstring ptr
declare function cJSON_PrintUnformatted(byval item as const cJSON ptr) as zstring ptr
declare function cJSON_PrintBuffered(byval item as const cJSON ptr, byval prebuffer as long, byval fmt as cJSON_bool) as zstring ptr
declare function cJSON_PrintPreallocated(byval item as cJSON ptr, byval buffer as zstring ptr, byval length as const long, byval format as const cJSON_bool) as cJSON_bool
declare sub cJSON_Delete(byval item as cJSON ptr)
declare function cJSON_GetArraySize(byval array as const cJSON ptr) as long
declare function cJSON_GetArrayItem(byval array as const cJSON ptr, byval index as long) as cJSON ptr
declare function cJSON_GetObjectItem(byval object as const cJSON const ptr, byval string as const zstring const ptr) as cJSON ptr
declare function cJSON_GetObjectItemCaseSensitive(byval object as const cJSON const ptr, byval string as const zstring const ptr) as cJSON ptr
declare function cJSON_HasObjectItem(byval object as const cJSON ptr, byval string as const zstring ptr) as cJSON_bool
declare function cJSON_GetErrorPtr() as const zstring ptr
declare function cJSON_GetStringValue(byval item as const cJSON const ptr) as zstring ptr
declare function cJSON_GetNumberValue(byval item as const cJSON const ptr) as double
declare function cJSON_IsInvalid(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsFalse(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsTrue(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsBool(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsNull(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsNumber(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsString(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsArray(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsObject(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_IsRaw(byval item as const cJSON const ptr) as cJSON_bool
declare function cJSON_CreateNull() as cJSON ptr
declare function cJSON_CreateTrue() as cJSON ptr
declare function cJSON_CreateFalse() as cJSON ptr
declare function cJSON_CreateBool(byval boolean as cJSON_bool) as cJSON ptr
declare function cJSON_CreateNumber(byval num as double) as cJSON ptr
declare function cJSON_CreateString(byval string as const zstring ptr) as cJSON ptr
declare function cJSON_CreateRaw(byval raw as const zstring ptr) as cJSON ptr
declare function cJSON_CreateArray() as cJSON ptr
declare function cJSON_CreateObject() as cJSON ptr
declare function cJSON_CreateStringReference(byval string as const zstring ptr) as cJSON ptr
declare function cJSON_CreateObjectReference(byval child as const cJSON ptr) as cJSON ptr
declare function cJSON_CreateArrayReference(byval child as const cJSON ptr) as cJSON ptr
declare function cJSON_CreateIntArray(byval numbers as const long ptr, byval count as long) as cJSON ptr
declare function cJSON_CreateFloatArray(byval numbers as const single ptr, byval count as long) as cJSON ptr
declare function cJSON_CreateDoubleArray(byval numbers as const double ptr, byval count as long) as cJSON ptr
declare function cJSON_CreateStringArray(byval strings as const zstring const ptr ptr, byval count as long) as cJSON ptr
declare function cJSON_AddItemToArray(byval array as cJSON ptr, byval item as cJSON ptr) as cJSON_bool
declare function cJSON_AddItemToObject(byval object as cJSON ptr, byval string as const zstring ptr, byval item as cJSON ptr) as cJSON_bool
declare function cJSON_AddItemToObjectCS(byval object as cJSON ptr, byval string as const zstring ptr, byval item as cJSON ptr) as cJSON_bool
declare function cJSON_AddItemReferenceToArray(byval array as cJSON ptr, byval item as cJSON ptr) as cJSON_bool
declare function cJSON_AddItemReferenceToObject(byval object as cJSON ptr, byval string as const zstring ptr, byval item as cJSON ptr) as cJSON_bool
declare function cJSON_DetachItemViaPointer(byval parent as cJSON ptr, byval item as cJSON const ptr) as cJSON ptr
declare function cJSON_DetachItemFromArray(byval array as cJSON ptr, byval which as long) as cJSON ptr
declare sub cJSON_DeleteItemFromArray(byval array as cJSON ptr, byval which as long)
declare function cJSON_DetachItemFromObject(byval object as cJSON ptr, byval string as const zstring ptr) as cJSON ptr
declare function cJSON_DetachItemFromObjectCaseSensitive(byval object as cJSON ptr, byval string as const zstring ptr) as cJSON ptr
declare sub cJSON_DeleteItemFromObject(byval object as cJSON ptr, byval string as const zstring ptr)
declare sub cJSON_DeleteItemFromObjectCaseSensitive(byval object as cJSON ptr, byval string as const zstring ptr)
declare function cJSON_InsertItemInArray(byval array as cJSON ptr, byval which as long, byval newitem as cJSON ptr) as cJSON_bool
declare function cJSON_ReplaceItemViaPointer(byval parent as cJSON const ptr, byval item as cJSON const ptr, byval replacement as cJSON ptr) as cJSON_bool
declare function cJSON_ReplaceItemInArray(byval array as cJSON ptr, byval which as long, byval newitem as cJSON ptr) as cJSON_bool
declare function cJSON_ReplaceItemInObject(byval object as cJSON ptr, byval string as const zstring ptr, byval newitem as cJSON ptr) as cJSON_bool
declare function cJSON_ReplaceItemInObjectCaseSensitive(byval object as cJSON ptr, byval string as const zstring ptr, byval newitem as cJSON ptr) as cJSON_bool
declare function cJSON_Duplicate(byval item as const cJSON ptr, byval recurse as cJSON_bool) as cJSON ptr
declare function cJSON_Compare(byval a as const cJSON const ptr, byval b as const cJSON const ptr, byval case_sensitive as const cJSON_bool) as cJSON_bool
declare sub cJSON_Minify(byval json as zstring ptr)
declare function cJSON_AddNullToObject(byval object as cJSON const ptr, byval name as const zstring const ptr) as cJSON ptr
declare function cJSON_AddTrueToObject(byval object as cJSON const ptr, byval name as const zstring const ptr) as cJSON ptr
declare function cJSON_AddFalseToObject(byval object as cJSON const ptr, byval name as const zstring const ptr) as cJSON ptr
declare function cJSON_AddBoolToObject(byval object as cJSON const ptr, byval name as const zstring const ptr, byval boolean as const cJSON_bool) as cJSON ptr
declare function cJSON_AddNumberToObject(byval object as cJSON const ptr, byval name as const zstring const ptr, byval number as const double) as cJSON ptr
declare function cJSON_AddStringToObject(byval object as cJSON const ptr, byval name as const zstring const ptr, byval string as const zstring const ptr) as cJSON ptr
declare function cJSON_AddRawToObject(byval object as cJSON const ptr, byval name as const zstring const ptr, byval raw as const zstring const ptr) as cJSON ptr
declare function cJSON_AddObjectToObject(byval object as cJSON const ptr, byval name as const zstring const ptr) as cJSON ptr
declare function cJSON_AddArrayToObject(byval object as cJSON const ptr, byval name as const zstring const ptr) as cJSON ptr
'' TODO: #define cJSON_SetIntValue(object, number) ((object) ? (object)->valueint = (object)->valuedouble = (number) : (number))
declare function cJSON_SetNumberHelper(byval object as cJSON ptr, byval number as double) as double
#define cJSON_SetNumberValue(object, number) iif(object <> NULL, cJSON_SetNumberHelper(object, cdbl(number)), (number))
declare function cJSON_SetValuestring(byval object as cJSON ptr, byval valuestring as const zstring ptr) as zstring ptr
'' TODO: #define cJSON_SetBoolValue(object, boolValue) ( (object != NULL && ((object)->type & (cJSON_False|cJSON_True))) ? (object)->type=((object)->type &(~(cJSON_False|cJSON_True)))|((boolValue)?cJSON_True:cJSON_False) : cJSON_Invalid )
'' TODO: #define cJSON_ArrayForEach(element, array) for(element = (array != NULL) ? (array)->child : NULL; element != NULL; element = element->next)
declare function cJSON_malloc(byval size as uinteger) as any ptr
declare sub cJSON_free(byval object as any ptr)

end extern

type JSON_VERSIONPROC as function cdecl () as zstring ptr
type JSON_PARSEPROC as function cdecl (byval value as const zstring ptr) as cJSON ptr
type JSON_DELETEPROC as sub cdecl (byval item as cJSON ptr)
type JSON_GETOBJECTITEMPROC as function cdecl (byval object as const cJSON ptr, byval string as const zstring ptr) as cJSON ptr
type JSON_ISSTRINGPROC as function cdecl (byval item as const cJSON ptr) as cJSON_bool
type JSON_ISNUMBERPROC as function cdecl (byval item as const cJSON ptr) as cJSON_bool
type JSON_ISBOOLPROC as function cdecl (byval item as const cJSON ptr) as cJSON_bool
type JSON_ISTRUEPROC as function cdecl (byval item as const cJSON ptr) as cJSON_bool
type JSON_ISFALSEPROC as function cdecl (byval item as const cJSON ptr) as cJSON_bool
type JSON_GETARRAYSIZEPROC as function cdecl (byval item as const cJSON ptr) as long
type JSON_GETARRAYITEMPROC as function cdecl (byval array as const cJSON ptr, byval index as long) as cJSON ptr

#ifdef __FB_64BIT__
   #define DECL_CJSON_DLL_NAME        "./cJSON64.dll"
   #define DECL_CJSON_VERSION         "cJSON_Version"
   #define DECL_CJSON_PARSE           "cJSON_Parse"
   #define DECL_CJSON_DELETE          "cJSON_Delete"
   #define DECL_CJSON_ISSTRING        "cJSON_IsString"
   #define DECL_CJSON_ISNUMBER        "cJSON_IsNumber"
   #define DECL_CJSON_ISBOOL          "cJSON_IsBool"
   #define DECL_CJSON_ISTRUE          "cJSON_IsTrue"
   #define DECL_CJSON_ISFALSE         "cJSON_IsFalse"
   #define DECL_CJSON_GETOBJECTITEM   "cJSON_GetObjectItem"
   #define DECL_CJSON_GETARRAYSIZE    "cJSON_GetArraySize"
   #define DECL_CJSON_GETARRAYITEM    "cJSON_GetArrayItem"
#else
   #define DECL_CJSON_DLL_NAME        "./cJSON32.dll"
   #define DECL_CJSON_VERSION         "_cJSON_Version@0"
   #define DECL_CJSON_PARSE           "_cJSON_Parse@4"
   #define DECL_CJSON_DELETE          "_cJSON_Delete@4"
   #define DECL_CJSON_ISSTRING        "_cJSON_IsString@4"
   #define DECL_CJSON_ISNUMBER        "_cJSON_IsNumber@4"
   #define DECL_CJSON_ISBOOL          "_cJSON_IsBool@4"
   #define DECL_CJSON_ISTRUE          "_cJSON_IsTrue@4"
   #define DECL_CJSON_ISFALSE         "_cJSON_IsFalse@4"
   #define DECL_CJSON_GETOBJECTITEM   "_cJSON_GetObjectItem@8"
   #define DECL_CJSON_GETARRAYSIZE    "_cJSON_GetArraySize@4"
   #define DECL_CJSON_GETARRAYITEM    "_cJSON_GetArrayItem@8"
#endif

type CJSON_TYPE
   private:
      _hLib as HMODULE = NULL

   public:
      declare constructor()
      declare destructor()
      declare property hLib( byval libHandle as HMODULE )
      declare property hLib() as HMODULE
      declare function version() as string
      declare function parse(byval jsonString as const zstring ptr) as cJSON ptr
      declare sub deleteptr(byval jsonPtr as cJSON ptr)
      declare function isString(byval jsonPtr as const cJSON ptr) as boolean
      declare function isNumber(byval jsonPtr as const cJSON ptr) as boolean
      declare function isBool(byval jsonPtr as const cJSON ptr) as boolean
      declare function isTrue(byval jsonPtr as const cJSON ptr) as boolean
      declare function isFalse(byval jsonPtr as const cJSON ptr) as boolean
      declare function getptr(byval jsonPtr as const cJSON ptr, byval itemKey as const zstring ptr) as cJSON ptr
      declare function gettext(byval jsonPtr as const cJSON ptr, byval itemKey as const zstring ptr) as string
      declare function getnumber(byval jsonPtr as const cJSON ptr, byval itemKey as const zstring ptr) as double
      declare function getbool(byval jsonPtr as const cJSON ptr, byval itemKey as const zstring ptr) as boolean
      declare function arraycount(byval jsonPtr as const cJSON ptr) as long
      declare function arrayitem(byval jsonPtr as const cJSON ptr, byval index as long) as cJSON ptr
end type

constructor CJSON_TYPE
   if _hLib = NULL then
      _hLib = LoadLibrary(DECL_CJSON_DLL_NAME)
      if _hLib = NULL then
         ? "Unable to load " & DECL_CJSON_DLL_NAME
      end if
   end if   
end constructor

destructor CJSON_TYPE
   if _hLib then
      FreeLibrary(_hLib)
      _hLib = NULL
   end if   
end destructor

property CJSON_TYPE.hLib( byval libHandle as HMODULE )
   _hLib = libHandle
end property

property CJSON_TYPE.hLib() as HMODULE
   property = _hLib
end property

function CJSON_TYPE.version() as string
   if this.hLib = NULL then exit function
   dim pProc as JSON_VERSIONPROC = _
   cast(JSON_VERSIONPROC, GetProcAddress(this.hLib, DECL_CJSON_VERSION))
   if pProc then
      dim psz as const zstring ptr = pProc()
      if psz then function = *psz
   end if
end function

function CJSON_TYPE.parse(byval jsonString as const zstring ptr) as cJSON ptr
   if this.hLib = NULL then exit function
   dim pProc as JSON_PARSEPROC = _
   cast(JSON_PARSEPROC, GetProcAddress(this.hLib, DECL_CJSON_PARSE))
   if pProc then
      function = pProc(jsonString)
   end if
end function

sub CJSON_TYPE.deleteptr(byval jsonPtr as cJSON ptr)
   if this.hLib = NULL then exit sub
   dim pProc as JSON_DELETEPROC = _
   cast(JSON_DELETEPROC, GetProcAddress(this.hLib, DECL_CJSON_DELETE))
   if pProc then pProc(jsonPtr)
end sub

function CJSON_TYPE.isString(byval jsonPtr as const cJSON ptr) as boolean
   if this.hLib = NULL then exit function
   dim pProc as JSON_ISSTRINGPROC = _
   cast(JSON_ISSTRINGPROC, GetProcAddress(this.hLib, DECL_CJSON_ISSTRING))
   if pProc then
      function = cbool(pProc(jsonPtr))
   end if
end function

function CJSON_TYPE.isNumber(byval jsonPtr as const cJSON ptr) as boolean
   if this.hLib = NULL then exit function
   dim pProc as JSON_ISNUMBERPROC = _
   cast(JSON_ISNUMBERPROC, GetProcAddress(this.hLib, DECL_CJSON_ISNUMBER))
   if pProc then
      function = cbool(pProc(jsonPtr))
   end if
end function

function CJSON_TYPE.isBool(byval jsonPtr as const cJSON ptr) as boolean
   if this.hLib = NULL then exit function
   dim pProc as JSON_ISBOOLPROC = _
   cast(JSON_ISBOOLPROC, GetProcAddress(this.hLib, DECL_CJSON_ISBOOL))
   if pProc then
      function = cbool(pProc(jsonPtr))
   end if
end function

function CJSON_TYPE.isTrue(byval jsonPtr as const cJSON ptr) as boolean
   if this.hLib = NULL then exit function
   dim pProc as JSON_ISTRUEPROC = _
   cast(JSON_ISTRUEPROC, GetProcAddress(this.hLib, DECL_CJSON_ISTRUE))
   if pProc then
      function = cbool(pProc(jsonPtr))
   end if
end function

function CJSON_TYPE.isFalse(byval jsonPtr as const cJSON ptr) as boolean
   if this.hLib = NULL then exit function
   dim pProc as JSON_ISFALSEPROC = _
   cast(JSON_ISFALSEPROC, GetProcAddress(this.hLib, DECL_CJSON_ISFALSE))
   if pProc then
      function = cbool(pProc(jsonPtr))
   end if
end function

function CJSON_TYPE.getptr(byval jsonPtr as const cJSON ptr, byval itemKey as const zstring ptr) as cJSON ptr
   if this.hLib = NULL then exit function
   dim pProc as JSON_GETOBJECTITEMPROC = _
   cast(JSON_GETOBJECTITEMPROC, GetProcAddress(this.hLib, DECL_CJSON_GETOBJECTITEM))
   if pProc then
      function = pProc(jsonPtr, itemKey)
   end if
end function

function CJSON_TYPE.gettext(byval jsonPtr as const cJSON ptr, byval itemKey as const zstring ptr) as string
   if this.hLib = NULL then exit function
   dim as cJSON ptr jsonPtr2 = this.getptr(jsonPtr, itemKey)
   if jsonPtr2 then 
      if (this.isString(jsonPtr2) andalso (jsonPtr2->valuestring <> NULL)) then
         function = *jsonPtr2->valuestring
      end if   
   end if
end function

function CJSON_TYPE.getnumber(byval jsonPtr as const cJSON ptr, byval itemKey as const zstring ptr) as double
   if this.hLib = NULL then exit function
   dim as cJSON ptr jsonPtr2 = this.getptr(jsonPtr, itemKey)
   if jsonPtr2 then 
      if (this.isNumber(jsonPtr2)) then
         function = jsonPtr2->valuedouble
      end if   
   end if
end function

function CJSON_TYPE.getbool(byval jsonPtr as const cJSON ptr, byval itemKey as const zstring ptr) as boolean
   if this.hLib = NULL then exit function
   dim as cJSON ptr jsonPtr2 = this.getptr(jsonPtr, itemKey)
   if jsonPtr2 then 
      if (this.isBool(jsonPtr2)) then
         function = cbool(this.isTrue(jsonPtr2))
      end if   
   end if
end function

function CJSON_TYPE.arraycount(byval jsonPtr as const cJSON ptr) as long
   if this.hLib = NULL then exit function
   dim pProc as JSON_GETARRAYSIZEPROC = _
   cast(JSON_GETARRAYSIZEPROC, GetProcAddress(this.hLib, DECL_CJSON_GETARRAYSIZE))
   if pProc then
      function = pProc(jsonPtr)
   end if
end function

function CJSON_TYPE.arrayitem(byval jsonPtr as const cJSON ptr, byval index as long) as cJSON ptr
   if this.hLib = NULL then exit function
   dim pProc as JSON_GETARRAYITEMPROC = _
   cast(JSON_GETARRAYITEMPROC, GetProcAddress(this.hLib, DECL_CJSON_GETARRAYITEM))
   if pProc then
      function = pProc(jsonPtr, index)
   end if
end function



