' ########################################################################################
' Microsoft Windows
' File: CCLRHost.inc
' Contents: Class to host the .NET Common Language Runtime (CLR)
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "Afx/CWSTR.inc"
#INCLUDE ONCE "Afx/CDispInvoke.inc"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' ########################################################################################
#ifndef __Afx_IUnknown_INTERFACE_DEFINED__
#define __Afx_IUnknown_INTERFACE_DEFINED__
TYPE Afx_IUnknown AS Afx_IUnknown_
TYPE Afx_IUnknown_ EXTENDS OBJECT
	DECLARE ABSTRACT FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION AddRef () AS ULONG
	DECLARE ABSTRACT FUNCTION Release () AS ULONG
END TYPE
TYPE AFX_LPUNKNOWN AS Afx_IUnknown PTR
#endif
' ########################################################################################

' ########################################################################################
#ifndef __Afx_IDispatch_INTERFACE_DEFINED__
#define __Afx_IDispatch_INTERFACE_DEFINED__
TYPE Afx_IDispatch AS Afx_IDispatch_
TYPE Afx_IDispatch_  EXTENDS Afx_Iunknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) as HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif
' ########################################################################################

' ########################################################################################
#ifndef __Afx_ICLRMetaHost_INTERFACE_DEFINED__
#define __Afx_ICLRMetaHost_INTERFACE_DEFINED__

CONST AFX_CLSID_CLRMetaHost = "{9280188d-0e8e-4867-b30c-7fa83884e8de}"
CONST AFX_IID_ICLRMetaHost = "{D332DB9E-B9B3-4125-8207-A14884F53216}"

TYPE Afx_ICLRMetaHost AS Afx_ICLRMetaHost_
TYPE Afx_ICLRMetaHost_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetRuntime (BYVAL pwzVersion AS LPCWSTR, BYVAL riid AS REFIID, BYVAL ppRuntime AS LPVOID) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetVersionFromFile (BYVAL pwzFilePath AS LPCWSTR, BYVAL pwzBuffer AS LPWSTR, BYVAL pcchBuffer AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnumerateInstalledRuntimes (BYVAL ppEnumerator AS IEnumUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnumerateLoadedRuntimes (BYVAL hndProcess AS HANDLE, BYVAL ppEnumerator AS IEnumUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RequestRuntimeLoadedNotification (BYVAL pCallbackFunction AS ANY PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION QueryLegacyV2RuntimeBinding (BYVAL rood AS REFIID, BYVAL ppUnk AS LPVOID) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExitProcess (BYVAL oExitCode AS INT32) AS HRESULT
END TYPE

#endif
' ########################################################################################

' ########################################################################################
#ifndef __Afx_ICLRRuntimeInfo_INTERFACE_DEFINED__
#define __Afx_ICLRRuntimeInfo_INTERFACE_DEFINED__

CONST IID_ICLRRuntimeInfo = "{BD39D1D2-BA2F-486a-89B0-B4B0CB466891}"

TYPE Afx_ICLRRuntimeInfo AS Afx_ICLRRuntimeInfo_
TYPE Afx_ICLRRuntimeInfo_ EXTENDS Afx_IUnknown

   DECLARE ABSTRACT FUNCTION GetVersionString (BYVAL pwzBuffer AS LPWSTR, BYVAL pcchBuffer AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetRuntimeDirectory (BYVAL pwzBuffer AS LPWSTR, BYVAL pcchBuffer AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsLoaded (BYVAL hndProcess AS HANDLE, BYVAL pbLoaded AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION LoadErrorString (BYVAL iResourceID AS UINT, BYVAL pwzBuffer AS LPWSTR, BYVAL pcchBuffer AS DWORD PTR, BYVAL iLocaleID AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION LoadLibrary (BYVAL pwzDllName AS LPCWSTR, BYVAL phndModule AS HMODULE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetProcAddress (BYVAL pszProcName AS LPCSTR, BYVAL ppProc AS LPVOID) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetInterface (BYVAL rclsid AS REFCLSID, BYVAL riid AS REFIID, BYVAL ppUnk AS LPVOID) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsLoadable (BYVAL pbLoadable AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetDefaultStartupFlags (BYVAL dwStartupFlags AS DWORD, BYVAL pwzHostConfigFile AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDefaultStartupFlags (BYVAL pdwStartupFlags AS DWORD PTR, BYVAL pwzHostConfigFile AS LPWSTR, BYVAL pcchHostConfigFile AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT SUB BindAsLegacyV2Runtime ()
   DECLARE ABSTRACT FUNCTION IsStarted (BYVAL pbStarted AS WINBOOL PTR, BYVAL pdwStartupFlags AS DWORD PTR) AS HRESULT

END TYPE

#endif
' ########################################################################################

TYPE Afx_ICorConfiguration AS Afx_ICorConfiguration_
TYPE Afx__Type AS Afx__Type_
TYPE Afx_ICorConfiguration AS Afx_ICorConfiguration_
TYPE Afx__Evidence AS Afx__Evidence_
TYPE Afx__EventHandler AS Afx__EventHandler_
TYPE Afx__AssemblyLoadEventHandler AS Afx__AssemblyLoadEventHandler_
TYPE Afx__ResolveEventHandler AS Afx__ResolveEventHandler_
TYPE Afx__UnhandledExceptionEventHandler AS Afx__UnhandledExceptionEventHandler_
TYPE Afx__AssemblyName AS Afx__AssemblyName_
TYPE Afx__AssemblyBuilder AS Afx__AssemblyBuilder_
TYPE Afx__PermissionSet AS Afx__PermissionSet_
TYPE Afx__ObjectHandle AS Afx__ObjectHandle_
TYPE Afx__Binder AS Afx__Binder_
TYPE Afx__CultureInfo AS Afx__CultureInfo_
TYPE Afx__Assembly aS Afx__Assembly_
TYPE Afx__PolicyLevel AS Afx__PolicyLevel_
TYPE Afx_IPrincipal AS Afx_IPrincipal_
TYPE Afx__CrossAppDomainDelegate AS Afx__CrossAppDomainDelegate_
TYPE Afx__ObjRef AS Afx__ObjRef_

ENUM AssemblyBuilderAccess
   ' // IID: {F0778630-AC34-3D71-9FAB-617F61243065}
   ' // Number of constants: 5
   AssemblyBuilderAccess_Run = 1   ' (&h00000001)
   AssemblyBuilderAccess_Save = 2   ' (&h00000002)
   AssemblyBuilderAccess_RunAndSave = 3   ' (&h00000003)
   AssemblyBuilderAccess_ReflectionOnly = 6   ' (&h00000006)
   AssemblyBuilderAccess_RunAndCollect = 9   ' (&h00000009)
END ENUM

ENUM BindingFlags
   ' // IID: {3223E024-5D70-3236-A92A-6B4114B2632F}
   ' // Number of constants: 20
   BindingFlags_Default = 0   ' (&h00000000)
   BindingFlags_IgnoreCase = 1   ' (&h00000001)
   BindingFlags_DeclaredOnly = 2   ' (&h00000002)
   BindingFlags_Instance = 4   ' (&h00000004)
   BindingFlags_Static = 8   ' (&h00000008)
   BindingFlags_Public = 16   ' (&h00000010)
   BindingFlags_NonPublic = 32   ' (&h00000020)
   BindingFlags_FlattenHierarchy = 64   ' (&h00000040)
   BindingFlags_InvokeMethod = 256   ' (&h00000100)
   BindingFlags_CreateInstance = 512   ' (&h00000200)
   BindingFlags_GetField = 1024   ' (&h00000400)
   BindingFlags_SetField = 2048   ' (&h00000800)
   BindingFlags_GetProperty = 4096   ' (&h00001000)
   BindingFlags_SetProperty = 8192   ' (&h00002000)
   BindingFlags_PutDispProperty = 16384   ' (&h00004000)
   BindingFlags_PutRefDispProperty = 32768   ' (&h00008000)
   BindingFlags_ExactBinding = 65536   ' (&h00010000)
   BindingFlags_SuppressChangeType = 131072   ' (&h00020000)
   BindingFlags_OptionalParamBinding = 262144   ' (&h00040000)
   BindingFlags_IgnoreReturn = 16777216   ' (&h01000000)
END ENUM

ENUM PrincipalPolicy
   ' // IID: {7D29BC4B-8FBC-38AA-8B35-ED4539A1CF8E}
   ' // Number of constants: 3
   PrincipalPolicy_UnauthenticatedPrincipal = 0   ' (&h00000000)
   PrincipalPolicy_NoPrincipal = 1   ' (&h00000001)
   PrincipalPolicy_WindowsPrincipal = 2   ' (&h00000002)
END ENUM

' ########################################################################################
' Interface name: ICorRuntimeHost
' IID: {CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}
' Documentation string: Common Language Runtime Hosting Interface
' Inherited interface = IUnknown
' Number of methods = 19
' ########################################################################################

#ifndef __Afx_ICorRuntimeHost_INTERFACE_DEFINED__
#define __Afx_ICorRuntimeHost_INTERFACE_DEFINED__

CONST AFX_CLSID_CorRuntimeHost = "{CB2F6723-AB3A-11D2-9C40-00C04FA30A3E}"
CONST AFX_IID_ICorRuntimeHost = "{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}"

TYPE Afx_ICorRuntimeHost AS Afx_ICorRuntimeHost_
TYPE Afx_ICorRuntimeHost_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION CreateLogicalThreadState () AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteLogicalThreadState () AS HRESULT
   DECLARE ABSTRACT FUNCTION SwitchInLogicalThreadState (BYVAL pFiberCookie AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SwitchOutLogicalThreadState (BYVAL pFiberCookie AS ULONG PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION LocksHeldByLogicalThread (BYVAL pCount AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MapFile (BYVAL hFile AS ANY PTR, BYVAL hMapAddress AS ANY PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetConfiguration (BYVAL pConfiguration AS Afx_ICorConfiguration PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Start () AS HRESULT
   DECLARE ABSTRACT FUNCTION Stop () AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateDomain (BYVAL pwzFriendlyName AS WSTRING PTR, BYVAL pIdentityArray AS IUnknown PTR, BYVAL pAppDomain AS IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDefaultDomain (BYVAL pAppDomain AS IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnumDomains (BYVAL hEnum AS ANY PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION NextDomain (BYVAL hEnum AS ANY PTR, BYVAL pAppDomain AS IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CloseEnum (BYVAL hEnum AS ANY PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateDomainEx (BYVAL pwzFriendlyName AS WSTRING PTR, BYVAL pSetup AS IUnknown PTR, BYVAL pEvidence AS IUnknown PTR, BYVAL pAppDomain AS IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateDomainSetup (BYVAL pAppDomainSetup AS IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateEvidence (BYVAL pEvidence AS IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION UnloadDomain (BYVAL pAppDomain AS IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CurrentDomain (BYVAL pAppDomain AS IUnknown PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################
' Interface name: _AppDomain
' IID: {05F696DC-2B29-3663-AD8B-C4389CF2A713}
' Attributes =  256 [&h00000100] [Oleautomation]
' Inherited interface = IUnknown
' Number of methods = 67
' ########################################################################################

#ifndef __Afx__AppDomain_INTERFACE_DEFINED__
#define __Afx__AppDomain_INTERFACE_DEFINED__

CONST AFX_CLSID_AppDomain = "{5FE0A145-A82B-3D96-94E3-FD214C9D6EB9}"
CONST AFX_IID__AppDomain = "{05F696DC-2B29-3663-AD8B-C4389CF2A713}"

TYPE Afx__AppDomain AS Afx__AppDomain_
TYPE Afx__AppDomain_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pcTInfo AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS ULONG, BYVAL lcid AS ULONG, BYVAL ppTInfo AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS GUID PTR, BYVAL rgszNames AS LONG, BYVAL cNames AS ULONG, BYVAL lcid AS ULONG, BYVAL rgDispId AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS ULONG, BYVAL riid AS GUID PTR, BYVAL lcid AS ULONG, BYVAL wFlags AS SHORT, BYVAL pDispParams AS LONG, BYVAL pVarResult AS LONG, BYVAL pExcepInfo AS LONG, BYVAL puArgErr AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ToString (BYVAL pRetVal AS Afx_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Equals (BYVAL other AS VARIANT, BYVAL pRetVal AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHashCode (BYVAL pRetVal AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetType (BYVAL pRetVal AS Afx__Type PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION InitializeLifetimeService (BYVAL pRetVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetLifetimeService (BYVAL pRetVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Evidence (BYVAL pRetVal AS Afx__Evidence PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_DomainUnload (BYVAL value AS Afx__EventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_DomainUnload (BYVAL value AS Afx__EventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_AssemblyLoad (BYVAL value AS Afx__AssemblyLoadEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_AssemblyLoad (BYVAL value AS Afx__AssemblyLoadEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ProcessExit (BYVAL value AS Afx__EventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ProcessExit (BYVAL value AS Afx__EventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_TypeResolve (BYVAL value AS Afx__ResolveEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_TypeResolve (BYVAL value AS Afx__ResolveEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ResourceResolve (BYVAL value AS Afx__ResolveEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ResourceResolve (BYVAL value AS Afx__ResolveEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_AssemblyResolve (BYVAL value AS Afx__ResolveEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_AssemblyResolve (BYVAL value AS Afx__ResolveEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_UnhandledException (BYVAL value AS Afx__UnhandledExceptionEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_UnhandledException (BYVAL value AS Afx__UnhandledExceptionEventHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly_2 (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL dir AS Afx_BSTR, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly_3 (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL Evidence AS Afx__Evidence PTR, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly_4 (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL requiredPermissions AS Afx__PermissionSet PTR, BYVAL optionalPermissions AS Afx__PermissionSet PTR, BYVAL refusedPermissions AS Afx__PermissionSet PTR, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly_5 (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL dir AS Afx_BSTR, BYVAL Evidence AS Afx__Evidence PTR, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly_6 (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL dir AS Afx_BSTR, BYVAL requiredPermissions AS Afx__PermissionSet PTR, BYVAL optionalPermissions AS Afx__PermissionSet PTR, BYVAL refusedPermissions AS Afx__PermissionSet PTR, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly_7 (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL Evidence AS Afx__Evidence PTR, BYVAL requiredPermissions AS Afx__PermissionSet PTR, BYVAL optionalPermissions AS Afx__PermissionSet PTR, BYVAL refusedPermissions AS Afx__PermissionSet PTR, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly_8 (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL dir AS Afx_BSTR, BYVAL Evidence AS Afx__Evidence PTR, BYVAL requiredPermissions AS Afx__PermissionSet PTR, BYVAL optionalPermissions AS Afx__PermissionSet PTR, BYVAL refusedPermissions AS Afx__PermissionSet PTR, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefineDynamicAssembly_9 (BYVAL name AS Afx__AssemblyName PTR, BYVAL access AS AssemblyBuilderAccess, BYVAL dir AS Afx_BSTR, BYVAL Evidence AS Afx__Evidence PTR, BYVAL requiredPermissions AS Afx__PermissionSet PTR, BYVAL optionalPermissions AS Afx__PermissionSet PTR, BYVAL refusedPermissions AS Afx__PermissionSet PTR, BYVAL IsSynchronized AS VARIANT_BOOL, BYVAL pRetVal AS Afx__AssemblyBuilder PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateInstance (BYVAL AssemblyName AS Afx_BSTR, BYVAL typeName AS Afx_BSTR, BYVAL pRetVal AS Afx__ObjectHandle PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateInstanceFrom (BYVAL assemblyFile AS Afx_BSTR, BYVAL typeName AS Afx_BSTR, BYVAL pRetVal AS Afx__ObjectHandle PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateInstance_2 (BYVAL AssemblyName AS Afx_BSTR, BYVAL typeName AS Afx_BSTR, BYVAL activationAttributes AS SAFEARRAY PTR, BYVAL pRetVal AS Afx__ObjectHandle PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateInstanceFrom_2 (BYVAL assemblyFile AS Afx_BSTR, BYVAL typeName AS Afx_BSTR, BYVAL activationAttributes AS SAFEARRAY PTR, BYVAL pRetVal AS Afx__ObjectHandle PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateInstance_3 (BYVAL AssemblyName AS Afx_BSTR, BYVAL typeName AS Afx_BSTR, BYVAL ignoreCase AS VARIANT_BOOL, BYVAL bindingAttr AS BindingFlags, BYVAL Binder AS Afx__Binder PTR, BYVAL args AS SAFEARRAY PTR, BYVAL culture AS Afx__CultureInfo PTR, BYVAL activationAttributes AS SAFEARRAY PTR, BYVAL securityAttributes AS Afx__Evidence PTR, BYVAL pRetVal AS Afx__ObjectHandle PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateInstanceFrom_3 (BYVAL assemblyFile AS Afx_BSTR, BYVAL typeName AS Afx_BSTR, BYVAL ignoreCase AS VARIANT_BOOL, BYVAL bindingAttr AS BindingFlags, BYVAL Binder AS Afx__Binder PTR, BYVAL args AS SAFEARRAY PTR, BYVAL culture AS Afx__CultureInfo PTR, BYVAL activationAttributes AS SAFEARRAY PTR, BYVAL securityAttributes AS Afx__Evidence PTR, BYVAL pRetVal AS Afx__ObjectHandle PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load (BYVAL assemblyRef AS Afx__AssemblyName PTR, BYVAL pRetVal AS Afx__Assembly PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load_2 (BYVAL assemblyString AS Afx_BSTR, BYVAL pRetVal AS Afx__Assembly PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load_3 (BYVAL rawAssembly AS SAFEARRAY PTR, BYVAL pRetVal AS Afx__Assembly PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load_4 (BYVAL rawAssembly AS SAFEARRAY PTR, BYVAL rawSymbolStore AS SAFEARRAY PTR, BYVAL pRetVal AS Afx__Assembly PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load_5 (BYVAL rawAssembly AS SAFEARRAY PTR, BYVAL rawSymbolStore AS SAFEARRAY PTR, BYVAL securityEvidence AS Afx__Evidence PTR, BYVAL pRetVal AS Afx__Assembly PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load_6 (BYVAL assemblyRef AS Afx__AssemblyName PTR, BYVAL assemblySecurity AS Afx__Evidence PTR, BYVAL pRetVal AS Afx__Assembly PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Load_7 (BYVAL assemblyString AS Afx_BSTR, BYVAL assemblySecurity AS Afx__Evidence PTR, BYVAL pRetVal AS Afx__Assembly PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecuteAssembly (BYVAL assemblyFile AS Afx_BSTR, BYVAL assemblySecurity AS Afx__Evidence PTR, BYVAL pRetVal AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecuteAssembly_2 (BYVAL assemblyFile AS Afx_BSTR, BYVAL pRetVal AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecuteAssembly_3 (BYVAL assemblyFile AS Afx_BSTR, BYVAL assemblySecurity AS Afx__Evidence PTR, BYVAL args AS SAFEARRAY PTR, BYVAL pRetVal AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FriendlyName (BYVAL pRetVal AS Afx_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BaseDirectory (BYVAL pRetVal AS Afx_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RelativeSearchPath (BYVAL pRetVal AS Afx_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShadowCopyFiles (BYVAL pRetVal AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAssemblies (BYVAL pRetVal AS SAFEARRAY PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AppendPrivatePath (BYVAL Path AS Afx_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ClearPrivatePath () AS HRESULT
   DECLARE ABSTRACT FUNCTION SetShadowCopyPath (BYVAL s AS Afx_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ClearShadowCopyPath () AS HRESULT
   DECLARE ABSTRACT FUNCTION SetCachePath (BYVAL s AS Afx_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetData (BYVAL name AS Afx_BSTR, BYVAL data AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetData (BYVAL name AS Afx_BSTR, BYVAL pRetVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetAppDomainPolicy (BYVAL domainPolicy AS Afx__PolicyLevel PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetThreadPrincipal (BYVAL principal AS Afx_IPrincipal PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetPrincipalPolicy (BYVAL policy AS PrincipalPolicy) AS HRESULT
   DECLARE ABSTRACT FUNCTION DoCallBack (BYVAL theDelegate AS Afx__CrossAppDomainDelegate PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DynamicDirectory (BYVAL pRetVal AS Afx_BSTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: _ObjectHandle
' IID: {EA675B47-64E0-3B5F-9BE7-F7DC2990730D}
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 8
' ########################################################################################

#ifndef __Afx__ObjectHandle_INTERFACE_DEFINED__
#define __Afx__ObjectHandle_INTERFACE_DEFINED__

CONST AFX_CLSID_ObjectHandle = "{ABEB0459-03B9-35AF-96E1-66BB7BC923F7}"
CONST AFX_IID__ObjectHandle = "{EA675B47-64E0-3B5F-9BE7-F7DC2990730D}"

TYPE Afx__ObjectHandle AS Afx__ObjectHandle_
TYPE Afx__ObjectHandle_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_ToString (BYVAL pRetVal AS Afx_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Equals (BYVAL obj AS VARIANT, BYVAL pRetVal AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHashCode (BYVAL pRetVal AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetType (BYVAL pRetVal AS Afx__Type PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetLifetimeService (BYVAL pRetVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION InitializeLifetimeService (BYVAL pRetVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateObjRef (BYVAL requestedType AS Afx__Type PTR, BYVAL pRetVal AS Afx__ObjRef PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Unwrap (BYVAL pRetVal AS VARIANT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' CCLRHost - Class to host the .NET Common Language Runtime (CLR)
' ########################################################################################
TYPE CCLRHost

Private:
   m_bUninitCOM AS BOOLEAN
   m_pCoreeLib AS ANY PTR
   m_pCorRuntimeHost AS Afx_ICorRuntimeHost PTR
   m_pAppDomain AS Afx__AppDomain PTR

Public:
   DECLARE CONSTRUCTOR (BYREF wszVersion AS WSTRING = "v4.0.30319")
   DECLARE DESTRUCTOR
   DECLARE PROPERTY RuntimeHost () AS Afx_ICorRuntimeHost PTR
   DECLARE PROPERTY DefaultDomain () AS Afx__AppDomain PTR
   DECLARE PROPERTY CurrentDomain () AS Afx__AppDomain PTR
   DECLARE FUNCTION CreateDomain (BYREF wszFriendlyName AS WSTRING, BYVAL pIdentityArray AS IUnknown PTR) AS Afx__AppDomain PTR
   DECLARE FUNCTION UnloadDomain (BYVAL pDomain AS Afx__AppDomain PTR) AS HRESULT
   DECLARE FUNCTION CreateInstance (BYREF AssemblyName AS CBSTR, BYREF typeName AS CBSTR) AS IDispatch PTR
   DECLARE FUNCTION CreateInstance (BYVAL pDomain AS Afx__AppDomain PTR, BYREF AssemblyName AS CBSTR, BYREF typeName AS CBSTR) AS IDispatch PTR
   DECLARE FUNCTION CreateInstanceFrom (BYREF AssemblyName AS CBSTR, BYREF typeName AS CBSTR) AS IDispatch PTR
   DECLARE FUNCTION CreateInstanceFrom (BYVAL pDomain AS Afx__AppDomain PTR, BYREF AssemblyName AS CBSTR, BYREF typeName AS CBSTR) AS IDispatch PTR

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CCLRHost (BYREF wszVersion AS WSTRING = "v4.0.30319")

   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE

   ' // Create an instance of the ICLRMetaHost interface
   DIM pCLRMetaHost AS Afx_ICLRMetaHost PTR
   DIM CLSID_CLRMetaHost_ AS GUID = TYPE(&h9280188d, &h0e8e, &h4867, {&hb3, &h0c, &h7f, &ha8, &h38, &h84, &he8, &hde})
   DIM IID_ICLRMetaHost_ AS GUID = TYPE(&hD332DB9E, &hB9B3, &h4125, {&h82, &h07, &hA1, &h48, &h84, &hF5, &h32, &h16})
   m_pCoreeLib = DyLibLoad("mscoree.dll")
   IF m_pCoreeLib = 0 THEN EXIT CONSTRUCTOR
   DIM pProc AS FUNCTION (BYVAL clsid AS REFCLSID, BYVAL riid AS REFIID, BYVAL ppInterface AS LPVOID) AS HRESULT
   pProc = DyLibSymbol(m_pCoreeLib, "CLRCreateInstance")
   IF pProc = 0 THEN EXIT CONSTRUCTOR
   hr = pProc(@CLSID_CLRMetaHost_, @IID_ICLRMetaHost_, @pCLRMetaHost)
   IF hr <> S_OK THEN SetLastError(E_FAIL) : EXIT CONSTRUCTOR

   ' // Gets the ICLRRuntimeInfo interface that corresponds to a particular version of
   ' // the common language runtime (CLR)
   DIM pRunTimeInfo AS Afx_ICLRRuntimeInfo PTR
   DIM IID_ICLRRuntimeInfo_ AS GUID = TYPE(&hBD39D1D2, &hBA2F, &h486a, {&h89, &hB0, &hB4, &hB0, &hCB, &h46, &h68, &h91})
   hr = pCLRMetaHost->GetRuntime(@wszVersion, @IID_ICLRRuntimeInfo_, @pRuntimeInfo)
   pCLRMetaHost->Release
   IF hr <> S_OK THEN SetLastError(hr) : EXIT CONSTRUCTOR

   ' // See if the runtime associated with this interface can be loaded into
   ' // the current process, taking into account other runtimes that might
   ' // already be loaded into the process.
   DIM bLoadable AS WINBOOL
   hr = pRunTimeInfo->IsLoadable(@bLoadable)
   IF bLoadable = FALSE THEN pRunTimeInfo->Release : SetLastError(E_FAIL) : EXIT CONSTRUCTOR
   ' // Loads the CLR into the current process and returns a reference to
   ' // the ICorRuntimeHost interface. ICorRuntimeHost works will all
   ' // versions of the .NET runtime.
   DIM CLSID_CorRuntimeHost_ AS GUID = TYPE(&hCB2F6723, &hAB3A, &h11D2, {&h9C, &h40, &h00, &hC0, &h4F, &hA3, &h0A, &h3E})
   DIM IID_ICorRuntimeHost_ AS GUID = TYPE(&hCB2F6722, &hAB3A, &h11D2, {&h9C, &h40, &h00, &hC0, &h4F, &hA3, &h0A, &h3E})
   hr = pRuntimeInfo->GetInterface(@CLSID_CorRuntimeHost_, @IID_ICorRuntimeHost_, @m_pCorRuntimeHost)
   pRuntimeInfo->Release
   IF hr <> S_OK THEN SetLastError(hr) : EXIT CONSTRUCTOR

   ' // Start the runtime (this also creates a default AppDomain)
   hr = m_pCorRuntimeHost->Start
   ' // Get the default AppDomain
   DIM pUnk AS Afx_IUnknown PTR
   IF hr = S_OK THEN hr = m_pCorRuntimeHost->GetDefaultDomain(cast(ANY PTR, @pUnk))
   IF hr <> S_OK THEN SetLastError(hr) : EXIT CONSTRUCTOR
   DIM IID__AppDomain AS GUID = TYPE(&h05F696DC, &h2B29, &h3663, {&hAD, &h8B, &hC4, &h38, &h9C, &hF2, &hA7, &h13})
   pUnk->QueryInterface(@IID__AppDomain, @m_pAppDomain)

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CCLRHost
   ' // Stops execution of code in the runtime
   IF m_pCorRuntimeHost THEN m_pCorRuntimeHost->Stop
   ' // Cleanup
   IF m_pAppDomain THEN m_pAppDomain->Release
   IF m_pCorRuntimeHost THEN m_pCorRuntimeHost->Release
   IF m_pCoreeLib THEN DyLibFree(m_pCoreeLib)
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a reference to the ICorRuntimeHost interface.
' ========================================================================================
PRIVATE PROPERTY CCLRHost.RuntimeHost () AS Afx_ICorRuntimeHost PTR
   PROPERTY = m_pCorRuntimeHost
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns a reference to the default domain.
' ========================================================================================
PRIVATE PROPERTY CCLRHost.DefaultDomain () AS Afx__AppDomain PTR
   PROPERTY = m_pAppDomain
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns a reference to the domain loaded on the current thread.
' ========================================================================================
PRIVATE PROPERTY CCLRHost.CurrentDomain () AS Afx__AppDomain PTR
   DIM pDomain AS Afx__AppDomain PTR
   DIM hr AS HRESULT = m_pCorRuntimeHost->CurrentDomain(cast(ANY PTR, @pDomain))
   IF hr <> S_OK THEN SetLastError(hr)
   PROPERTY = pDomain
END PROPERTY
' ========================================================================================

' ========================================================================================
' Creates an application domain.
' ========================================================================================
PRIVATE FUNCTION CCLRHost.CreateDomain (BYREF wszFriendlyName AS WSTRING, BYVAL pIdentityArray AS IUnknown PTR) AS Afx__AppDomain PTR
   DIM pDomain AS Afx__AppDomain PTR
   DIM hr AS HRESULT = m_pCorRuntimeHost->CreateDomain(@wszFriendlyName, pIdentityArray, cast(ANY PTR, pDomain))
   IF hr <> S_OK THEN SetLastError(hr)
   FUNCTION = pDomain
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unloads the specified application domain from the current process.
' ========================================================================================
PRIVATE FUNCTION CCLRHost.UnloadDomain (BYVAL pDomain AS Afx__AppDomain PTR) AS HRESULT
   IF m_pCorRuntimeHost = NULL THEN SetLastError(E_POINTER) : RETURN E_POINTER
   DIM hr AS HRESULT = m_pCorRuntimeHost->UnloadDomain(cast(ANY PTR, pDomain))
   IF hr <> S_OK THEN SetLastError(hr)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new instance of the specified type defined in the specified assembly.
' ========================================================================================
PRIVATE FUNCTION CCLRHost.CreateInstance (BYREF AssemblyName AS CBSTR, BYREF typeName AS CBSTR) AS IDispatch PTR
   ' // The default domain must exist
   IF m_pAppDomain = NULL THEN SetLastError(E_FAIL) : RETURN NULL
   ' // Create the instance
   DIM pObjectHandle AS Afx__ObjectHandle PTR
   DIM hr AS HRESULT = m_pAppDomain->CreateInstance(AssemblyName, typeName, @pObjectHandle)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN NULL
   ' // Unwrap the object
   DIM vObj AS VARIANT
   hr = pObjectHandle->Unwrap(@vObj)
   pObjectHandle->Release
   IF hr <> S_OK THEN RETURN NULL
   RETURN vObj.pdispVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new instance of the specified type defined in the specified assembly.
' ========================================================================================
PRIVATE FUNCTION CCLRHost.CreateInstance (BYVAL pDomain AS Afx__AppDomain PTR, BYREF AssemblyName AS CBSTR, BYREF typeName AS CBSTR) AS IDispatch PTR
   IF pDomain = NULL THEN SetLastError(E_FAIL) : RETURN NULL
   ' // Create the instance
   DIM pObjectHandle AS Afx__ObjectHandle PTR
   DIM hr AS HRESULT = pDomain->CreateInstance(AssemblyName, typeName, @pObjectHandle)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN NULL
   ' // Unwrap the object
   DIM vObj AS VARIANT
   hr = pObjectHandle->Unwrap(@vObj)
   pObjectHandle->Release
   IF hr <> S_OK THEN RETURN NULL
   RETURN vObj.pdispVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new instance of the specified type defined in the specified assembly file.
' ========================================================================================
PRIVATE FUNCTION CCLRHost.CreateInstanceFrom (BYREF AssemblyName AS CBSTR, BYREF typeName AS CBSTR) AS IDispatch PTR
   ' // The default domain must exist
   IF m_pAppDomain = NULL THEN SetLastError(E_FAIL) : RETURN NULL
   ' // Create the instance
   DIM pObjectHandle AS Afx__ObjectHandle PTR
   DIM hr AS HRESULT = m_pAppDomain->CreateInstanceFrom(AssemblyName, typeName, @pObjectHandle)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN NULL
   ' // Unwrap the object
   DIM vObj AS VARIANT
   hr = pObjectHandle->Unwrap(@vObj)
   pObjectHandle->Release
   IF hr <> S_OK THEN RETURN NULL
   RETURN vObj.pdispVal
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new instance of the specified type defined in the specified assembly file.
' ========================================================================================
PRIVATE FUNCTION CCLRHost.CreateInstanceFrom (BYVAL pDomain AS Afx__AppDomain PTR, BYREF AssemblyName AS CBSTR, BYREF typeName AS CBSTR) AS IDispatch PTR
   IF pDomain = NULL THEN SetLastError(E_FAIL) : RETURN NULL
   ' // Create the instance
   DIM pObjectHandle AS Afx__ObjectHandle PTR
   DIM hr AS HRESULT = pDomain->CreateInstanceFrom(AssemblyName, typeName, @pObjectHandle)
   IF hr <> S_OK THEN SetLastError(hr) : RETURN NULL
   ' // Unwrap the object
   DIM vObj AS VARIANT
   hr = pObjectHandle->Unwrap(@vObj)
   pObjectHandle->Release
   IF hr <> S_OK THEN RETURN NULL
   RETURN vObj.pdispVal
END FUNCTION
' ========================================================================================

END NAMESPACE
