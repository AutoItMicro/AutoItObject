/**
* @package	AutoItObject
* @copyright	Copyright (C) The AutoItObject-Team. All rights reserved.
* @license	Artistic License 2.0, see Artistic.txt
*
* This file is part of AutoItObject.
*
* AutoItObject is free software; you can redistribute it and/or modify
* it under the terms of the Artistic License as published by Larry Wall,
* either version 2.0, or (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
* See the Artistic License for more details.
*
* You should have received a copy of the Artistic License with this Kit,
* in the file named "Artistic.txt".  If not, you can get a copy from
* <http://www.perlfoundation.org/artistic_license_2_0> OR
* <http://www.opensource.org/licenses/artistic-license-2.0.php>
*
* A complete list of contributors is available in dll.h
*
*/
#include <windows.h>
#include "dll.h"
#include "stdio.h"
#include "AutoItObjectClass.h"
#include "AutoItObject.h"
#include "AutoItWrapObject.h"
#include "AutoItEnumObject.h"


BOOL WINAPI DllMain(HINSTANCE hinstDLL,DWORD fdwReason,LPVOID lpvReserved)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		// Do something...
		break;
	case DLL_PROCESS_DETACH:
		// Clean up?
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	}

	return TRUE;
}


#ifdef _RELEASE
EXTERN_C BOOL WINAPI _DllMainCRTStartup(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	return TRUE;
}
#endif


EXPORT(AutoItObjectClass*) CreateAutoItObjectClass()
{

	return new AutoItObjectClass;
}


EXPORT(AutoItObject*) CreateAutoItObject()
{

	return new AutoItObject;
}


EXPORT(AutoItObject*) CloneAutoItObject(AutoItObject* old)
{

	return new AutoItObject(old);
}


EXPORT(void) Initialize(AUTOITFUNCTIONPROXY autoitfunctionproxy, AUTOITENUMFUNCTIONPROXY autoitenumfunctionproxy)
{
	AutoItObject::Init(autoitfunctionproxy);
	AutoItEnumObject::Init(autoitenumfunctionproxy);
}


EXPORT(void) AddMethod(AutoItObject* object,wchar_t* method,wchar_t* value, AutoItElement::SCOPE new_scope)
{
	object->AddMethod(method, value, new_scope);
}


EXPORT(void) RemoveMember(AutoItObject* object,wchar_t* member)
{
	object->RemoveMember(member);
}


EXPORT(void) AddEnum(AutoItObject* object, wchar_t* function_next, wchar_t* function_reset, wchar_t* function_skip)
{
	object->AddEnum(function_next, function_reset, function_skip);
}


EXPORT(void) AddProperty(AutoItObject* object,wchar_t* property_name,  AutoItElement::SCOPE new_scope, VARIANT* property_value)
{
	object->AddProperty(property_name, new_scope, property_value);
}


EXPORT(AutoItWrapObject*) CreateWrapperObject(IUnknown* object, bool fNoUnknown)
{
	return new AutoItWrapObject(object, fNoUnknown);
}


EXPORT(AutoItWrapObject*) CreateWrapperObjectEx(IUnknown* object, wchar_t* sMethods, bool fNoUnknown, bool fCallFree)
{
	return new AutoItWrapObject(object, sMethods, fNoUnknown, fCallFree);
}


EXPORT(void) WrapperAddMethod(AutoItWrapObject* object,wchar_t* method,wchar_t* types, unsigned long index)
{
	AutoItWrapElement *elem = new AutoItWrapElement;
	elem->SetName(method);
	elem->SetIndex(index);
	elem->SetTypes(types);
	object->AddMember(elem);
}


EXPORT(AutoItWrapObject*) AutoItObjectCreateObject(wchar_t* sCLSID, wchar_t* sIID, wchar_t* sMethods)
{
	CLSID clsid;
	HRESULT success = CLSIDFromString(sCLSID, &clsid); // first try string representation of the CLSID
	if (success != NOERROR)
	{
		success = CLSIDFromProgID(sCLSID, &clsid); // if that didn't work try ProgID
		if (success != NOERROR) return NULL; // and if that failed return failure
	}

	IID iid;
	if (sIID==NULL)
	{
		iid = IID_IDispatch; // default is IDispatch
	}
	else
	{
		success = IIDFromString(sIID, &iid); // converting from string representation of IID to IID
		if (success != NOERROR) return NULL;
	}

	if (CoInitialize(0)==S_FALSE) CoUninitialize(); // just checking

	PVOID ppv;
	success = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, iid, &ppv);
	if (success != S_OK) return NULL;

	AutoItWrapObject* pRetObj = reinterpret_cast<AutoItWrapObject*>(ppv); // reinterpreting returned
	if (sMethods==NULL)
	{
		if (sIID==NULL) return pRetObj; // no sIID and no sMethods defined means 'classic' IDispatch, similar to AutoIt's ObjCreate function
		else return new AutoItWrapObject(pRetObj, FALSE);
	}
	return new AutoItWrapObject(pRetObj, sMethods, FALSE);
}


EXPORT(AutoItWrapObject*) AutoItObjectCreateObjectEx(wchar_t* sModule, wchar_t* sCLSID, wchar_t* sIID, wchar_t* sMethods, bool fWrapp)
{
	CoFreeUnusedLibrariesEx(0, 0);

	CLSID clsid;
	HRESULT success = CLSIDFromString(sCLSID, &clsid); // converting from string representation of CLSID to CLSID
	if (success != NOERROR) return NULL; // return failure

	IID iid;
	if (sIID==NULL)
	{
		iid = IID_IDispatch; // default is IDispatch
	}
	else
	{
		success = IIDFromString(sIID, &iid); // converting from string representation of IID to IID
		if (success != NOERROR) return NULL;
	}

	// Load the module
	HMODULE hDll = NULL;
	if (!fWrapp)
	{
		hDll = CoLoadLibrary(sModule, TRUE);
	}
	else
	{
		hDll = LoadLibraryW(sModule);
	}
	if (hDll==NULL) return NULL; // if it can't be loaded return failure

	// define DllGetClassObject function ( http://msdn.microsoft.com/en-us/library/ms680760(VS.85).aspx )
	typedef HRESULT (__stdcall *pDllGetClassObject)(IN REFCLSID, IN REFIID, OUT PVOID);
	// see if this module does export DllGetClassObject
	pDllGetClassObject GetClassObject = reinterpret_cast<pDllGetClassObject>(GetProcAddress(hDll, "DllGetClassObject"));
	if (GetClassObject==NULL) // if not
	{
		FreeLibrary(hDll); //  free the module
		return NULL; // and return failure
	}

	// If yes then call it
	IClassFactory *pIFactory;
	success = GetClassObject(clsid, IID_IClassFactory, &pIFactory);
	// check for failure
	if (success != S_OK)
	{
		FreeLibrary(hDll); //  free the module and return failure
		return NULL;
	}

	ITypeLib *pTypeLib;
	INT iRegFlag = 0;
	if (!fWrapp)
	{	
		if (LoadTypeLibEx(sModule, REGKIND_NONE, &pTypeLib) == S_OK)
		{
			// Register TypeLib.
			if (RegisterTypeLib(pTypeLib, sModule, NULL) == S_OK) 
			{
				iRegFlag = 2; // successfully registered
			}
			else
			{
				// In case of some error try registering on per-user level.
				if (RegisterTypeLibForUser(pTypeLib, sModule, NULL) == S_OK) 
				{
					iRegFlag = 3; // per-user registration
				}
				else 
				{
					iRegFlag = 1; // neither worked
				}
			}
		}
	}

	// Now call CreateInstance with IClassFactory object with specified iid
	PVOID ppv;
	success = pIFactory->CreateInstance(NULL, iid, &ppv);

	// Check for failure
	if (success != S_OK)
	{
		if (iRegFlag != 0)
		{
			TLIBATTR *pTLibAttrr;
			if (pTypeLib->GetLibAttr(&pTLibAttrr) == S_OK)	
			{
				switch (iRegFlag)
				{
				case 2:
					UnRegisterTypeLib(pTLibAttrr->guid, pTLibAttrr->wMajorVerNum, pTLibAttrr->wMinorVerNum, pTLibAttrr->lcid, pTLibAttrr->syskind);
				case 3:
					UnRegisterTypeLibForUser(pTLibAttrr->guid, pTLibAttrr->wMajorVerNum, pTLibAttrr->wMinorVerNum, pTLibAttrr->lcid, pTLibAttrr->syskind);
				}	
				// Free allocated memory
				pTypeLib->ReleaseTLibAttr(pTLibAttrr);
			}
			// Release ITypeLib:
			pTypeLib->Release();
		}
		// Release IClassFactory:
		pIFactory->Release();
		FreeLibrary(hDll); //  free the module
		return NULL; // and return failure
	}

	// TypeLib object is no longer needed, release it:
	if (iRegFlag != 0) pTypeLib->Release();

	// ClassFactory object is no longer needed, release it:
	pIFactory->Release();

	// Wrapp returned pointer as AutoItWrapObject object
	if (sMethods == NULL)
	{
		if (!fWrapp)
		{	
			// FIXME: One DLL handle is left open here!
			return reinterpret_cast<AutoItWrapObject*>(ppv); // no sIID and no sMethods defined means 'classic' IDispatch
		}
		else
		{
			AutoItWrapObject* pRetObject = new AutoItWrapObject(reinterpret_cast<IUnknown*>(ppv), FALSE); // Methods will be added later
			pRetObject->SetParentDllHandle(hDll);
			return pRetObject;
		}
	}

	// Methods are specified
	AutoItWrapObject* pRetObject = new AutoItWrapObject(reinterpret_cast<IUnknown*>(ppv), sMethods, FALSE);
	pRetObject->SetParentDllHandle(hDll);
	return pRetObject;
}


EXPORT(AutoItWrapObject*) CreateDllCallObject(wchar_t* sModule, wchar_t* sMethods, DWORD iFlags = 0)
{
	// Load the module
	HMODULE hDll = LoadLibraryExW(sModule, NULL, iFlags);
	if (hDll==NULL) return NULL; // if it can't be loaded return failure

	AutoItWrapObject* pRetObject;
	if (sMethods == NULL)
	{
		pRetObject = new AutoItWrapObject(NULL, TRUE, TRUE);
	}
	else
	{
		pRetObject = new AutoItWrapObject(NULL, sMethods, TRUE, FALSE, TRUE);
	}
	pRetObject->SetParentDllHandle(hDll);
	return pRetObject;
}


EXPORT(void) MemoryCallEntry(DWORD e, DWORD f)
{
	if (e == 0xDEAD && f == 0xBEEF) {
#ifdef _RELEASE
		HANDLE hStdout = GetStdHandle(STD_OUTPUT_HANDLE);
		DWORD dwWritten;
		WriteFile(hStdout, "Lol. You found the easter-egg. \r\n", sizeof "Lol. You found the easter-egg. \r\n", &dwWritten, NULL);
		FlushFileBuffers(hStdout);
#else
		printf("Lol. You found the easter-egg. \r\n"); fflush(stdout);
#endif
	}
}


EXPORT(ULONG) IUnknownAddRef(IUnknown* object)
{
	return object->AddRef();
}


EXPORT(ULONG) IUnknownRelease(IUnknown* object)
{
	return object->Release();
}


EXPORT(PVOID) ReturnThis(PVOID anything)
{
	return anything;
}



INT debugprintW(const wchar_t *format, ...)
{
#ifndef _RELEASE
	va_list args;
	va_start(args, format);
	int ret = vwprintf(format, args);
	va_end(args);
	fflush(stdout);
	return ret;
#endif
	return 0;
}


INT Compare(const wchar_t *s1, const wchar_t *s2)
{
	return CompareStringW(LOCALE_SYSTEM_DEFAULT, NORM_IGNORECASE, s1, -1, s2, -1) - CSTR_EQUAL;
}


VARTYPE VarType(VARTYPE vtype, LPCWSTR wtype)
{
	if (vtype != VT_BSTR) return VT_ERROR;
	if (Compare(L"none", wtype)==0) {
		return VT_EMPTY;
	}else if (Compare(L"byte", wtype)==0) {
		return VT_UI1;
	}else if ((Compare(L"boolean", wtype)==0)||(Compare(L"bool", wtype)==0)) {
		return VT_BOOL;
	}else if (Compare(L"short", wtype)==0) {
		return VT_I2;
	}else if ((Compare(L"ushort", wtype)==0)||(Compare(L"word", wtype)==0)) {
		return VT_UI2;
	}else if ((Compare(L"dword", wtype)==0)||(Compare(L"ulong", wtype)==0)||(Compare(L"uint", wtype)==0)) {
		return VT_UI4;
	}else if ((Compare(L"long", wtype)==0)||(Compare(L"int", wtype)==0)) {
		return VT_I4;
	}else if (Compare(L"variant", wtype)==0) {
		return VT_VARIANT;
	}else if (Compare(L"int64", wtype)==0) {
		return VT_I8;
	}else if (Compare(L"uint64", wtype)==0) {
		return VT_UI8;
	}else if (Compare(L"float", wtype)==0) {
		return VT_R4;
	}else if (Compare(L"double", wtype)==0) {
		return VT_R8;
	}else if (Compare(L"hresult", wtype)==0) {
		//return VT_HRESULT; // not working properly
		return VT_UI4;
	}else if (Compare(L"str", wtype)==0){
		return VT_LPSTR;
	}else if (Compare(L"wstr", wtype)==0){
		return VT_LPWSTR;
	}else if (Compare(L"bstr", wtype)==0){
		return VT_BSTR;
	}else if ((Compare(L"ptr", wtype)==0)||(Compare(L"handle", wtype)==0)||(Compare(L"hwnd", wtype)==0)) {
		//return VT_PTR; // not working properly
#ifdef _M_X64
		return VT_UI8;
#else
		return VT_UI4;
#endif
	}else if ((Compare(L"int_ptr", wtype)==0)||(Compare(L"long_ptr", wtype)==0)||(Compare(L"lresult", wtype)==0)||(Compare(L"lparam", wtype)==0)) {
		//return VT_INT_PTR; // not working properly
#ifdef _M_X64
		return VT_I8;
#else
		return VT_I4;
#endif
	}else if ((Compare(L"uint_ptr", wtype)==0)||(Compare(L"ulong_ptr", wtype)==0)||(Compare(L"dword_ptr", wtype)==0)||(Compare(L"wparam", wtype)==0)) {
		//return VT_UINT_PTR; // not working properly
#ifdef _M_X64
		return VT_UI8;
#else
		return VT_UI4;
#endif
	}else if ((Compare(L"idispatch", wtype)==0)||(Compare(L"object", wtype)==0)) {
		return VT_DISPATCH;
	}
	return VT_ILLEGAL;
}