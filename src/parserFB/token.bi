enum TOK
	Tassign				' =
	Tplus				' +
	Tminus				' -	
	Tdiv				' /
	Ttimes				' *	
	Tmod				' mod
	Tintegerdiv			' \
	Tshl				' shl
	Tshr				' dhr
	Tcaret				' ^

	Tgreater			' >
	Tless				' <
	Tis					' is
 
	Tdotdotdot			' ...
	Tdot				' .
	Tunderline			' _
	Texclam				' !
	Tquest				' ?
	Tdollar				' $
	Tquote				' "
	Tcolon				' :
	Tcomma				' ,
	Tpound				' #
	Tandsign			' &
	Tat					' @
	Topenparen			' (
	Tcloseparen			' )
	Topenbracket		' [
	Tclosebracket		' ]
	Topencurly			' {
	Tclosecurly			' }


	Tand
	Tor
	Txor
	Tnot
	Teqv
	Timp

	Tandalso
	Torelse

	Tlet
	Tfor
	Tto
	Tnext
	Tstep
	Texit
	Tnew
	Tdelete
	Tdo
	Tloop
	Twhile
	Twend
	Tif
	Tthen
	Telse
	Tend
	Tdim
	Tas
	Tenum
	Ttype
	Tunion
	Tfield
	Treturn
	Twith
	Tselect
	Tcase
	Tcast	

	Toption
	Texplicit
	

	
	' function
	Tdeclare
	Tsub
	Tfunction
	Tbyval
	Tbyref
	Tstdcall
	Tcdecl
	Tpascal
	Toverload
	Talias
	Texport

	' data
	Tbyte
	Tubyte
	Tshort
	Tushort
	Tinteger
	Tuinteger
	Tlongint
	Tulongint
	Tsingle
	Tdouble
	Tstring
	Tzstring
	Twstring
	Tany
	
	Tconst
	Tpointer
	Tptr
	Tunsigned
	Textern
	Tcommon
	Tshared
	Tstatic
	Tscope

	' Array
	Tredim
	Tvar
	Tpreserve
	Terase
	
	' class
	Tclass
	Tprivate
	Tprotected
	Tpublic
	Textends
	Tobject
	Tvirtual
	Tabstract
	Tconstructor
	Tdestructor
	Tproperty
	Toperator
	Tthis
	Tbase


	'
	Tinclude
	Tlibpath
	Tonce
	Telseif
	Tendif
	Tifdef
	Tifndef
	Tdefine
	Tmacro
	Tendmacro
	
	'Tinclib
	Tnamespace
	'
	Teol ' \n
	Tidentifier
	Tstrings
	Tnumbers
end enum

type TokenUnit
	as TOK			tok
	as string		identifier
	as integer		lineNumber
end type

#define B_VARIABLE 1
#define B_FUNCTION 2
#define B_SUB 4
#define B_PROPERTY 8
#define B_CTOR 16
#define B_DTOR 32
#define B_PARAM 64
#define B_TYPE 128
#define B_ENUM 256
#define B_UNION 512
#define B_CLASS 1024
#define B_INCLUDE 2048
#define B_ENUMMEMBER 4096
#define B_ALIAS 8192
#define B_BAS 16384
#define B_BI 32768
#define B_NAMESPACE 65536
#define B_MACRO 131072
#define B_SCOPE 262144
#define B_DEFINE 524288
#define B_OPERATOR 1048576

declare function identToTOK( ident as string ) as TOK

