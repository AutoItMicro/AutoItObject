; #INDEX# =======================================================================================================================
; Title .........: AutoItObject v1.2.2.0
; AutoIt Version : 3.3
; Language ......: English (language independent)
; Description ...: Brings Objects to AutoIt.
; Author(s) .....: monoceres, trancexx, Kip, Prog@ndy
; Copyright .....: Copyright (C) The AutoItObject-Team. All rights reserved.
; License .......: Artistic License 2.0, see Artistic.txt
;
; This file is part of AutoItObject.
;
; AutoItObject is free software; you can redistribute it and/or modify
; it under the terms of the Artistic License as published by Larry Wall,
; either version 2.0, or (at your option) any later version.
;
; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
; See the Artistic License for more details.
;
; You should have received a copy of the Artistic License with this Kit,
; in the file named "Artistic.txt".  If not, you can get a copy from
; <http://www.perlfoundation.org/artistic_license_2_0> OR
; <http://www.opensource.org/licenses/artistic-license-2.0.php>
;
; ------------------------ AutoItObject CREDITS: ------------------------
; Copyright (C) by:
; The AutoItObject-Team:
; 	Andreas Karlsson (monoceres)
; 	Dragana R. (trancexx)
; 	Dave Bakker (Kip)
; 	Andreas Bosch (progandy, Prog@ndy)
;
; ===============================================================================================================================
#include-once
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6


; #CURRENT# =====================================================================================================================
;_AutoItObject_AddDestructor
;_AutoItObject_AddEnum
;_AutoItObject_AddMethod
;_AutoItObject_AddProperty
;_AutoItObject_Class
;_AutoItObject_CLSIDFromString
;_AutoItObject_CoCreateInstance
;_AutoItObject_Create
;_AutoItObject_DllOpen
;_AutoItObject_DllStructCreate
;_AutoItObject_IDispatchToPtr
;_AutoItObject_IUnknownAddRef
;_AutoItObject_IUnknownRelease
;_AutoItObject_ObjCreate
;_AutoItObject_ObjCreateEx
;_AutoItObject_ObjectFromDtag
;_AutoItObject_PtrToIDispatch
;_AutoItObject_RemoveMember
;_AutoItObject_Shutdown
;_AutoItObject_Startup
;_AutoItObject_VariantClear
;_AutoItObject_VariantCopy
;_AutoItObject_VariantFree
;_AutoItObject_VariantInit
;_AutoItObject_VariantRead
;_AutoItObject_VariantSet
;_AutoItObject_WrapperAddMethod
;_AutoItObject_WrapperCreate
; ===============================================================================================================================

; #INTERNAL_NO_DOC# =============================================================================================================
;__Au3Obj_OleUninitialize
;__Au3Obj_IUnknown_AddRef
;__Au3Obj_IUnknown_Release
;__Au3Obj_GetMethods
;__Au3Obj_SafeArrayCreate
;__Au3Obj_SafeArrayDestroy
;__Au3Obj_SafeArrayAccessData
;__Au3Obj_SafeArrayUnaccessData
;__Au3Obj_SafeArrayGetUBound
;__Au3Obj_SafeArrayGetLBound
;__Au3Obj_SafeArrayGetDim
;__Au3Obj_CreateSafeArrayVariant
;__Au3Obj_ReadSafeArrayVariant
;__Au3Obj_CoTaskMemAlloc
;__Au3Obj_CoTaskMemFree
;__Au3Obj_CoTaskMemRealloc
;__Au3Obj_GlobalAlloc
;__Au3Obj_GlobalFree
;__Au3Obj_SysAllocString
;__Au3Obj_SysCopyString
;__Au3Obj_SysReAllocString
;__Au3Obj_SysFreeString
;__Au3Obj_SysStringLen
;__Au3Obj_SysReadString
;__Au3Obj_PtrStringLen
;__Au3Obj_PtrStringRead
;__Au3Obj_FunctionProxy
;__Au3Obj_EnumFunctionProxy
;__Au3Obj_ObjStructGetElements
;__Au3Obj_ObjStructMethod
;__Au3Obj_ObjStructDestructor
;__Au3Obj_ObjStructPointer
;__Au3Obj_PointerCall
;__Au3Obj_Mem_DllOpen
;__Au3Obj_Mem_FixReloc
;__Au3Obj_Mem_FixImports
;__Au3Obj_Mem_LoadLibraryEx
;__Au3Obj_Mem_FreeLibrary
;__Au3Obj_Mem_GetAddress
;__Au3Obj_Mem_VirtualProtect
;__Au3Obj_Mem_Base64Decode
;__Au3Obj_Mem_BinDll
;__Au3Obj_Mem_BinDll_X64
; ===============================================================================================================================

; #DATATYPES# =====================================================================================================================
; none - no value (only valid for return type, equivalent to void in C)
; byte - an unsigned 8 bit integer
; boolean - an unsigned 8 bit integer
; short - a 16 bit integer
; word, ushort - an unsigned 16 bit integer
; int, long - a 32 bit integer
; bool - a 32 bit integer
; dword, ulong, uint - an unsigned 32 bit integer
; hresult - an unsigned 32 bit integer
; int64 - a 64 bit integer
; uint64 - an unsigned 64 bit integer
; ptr - a general pointer (void *)
; hwnd - a window handle (pointer wide)
; handle - an handle (pointer wide)
; float - a single precision floating point number
; double - a double precision floating point number
; int_ptr, long_ptr, lresult, lparam - an integer big enough to hold a pointer when running on x86 or x64 versions of AutoIt
; uint_ptr, ulong_ptr, dword_ptr, wparam - an unsigned integer big enough to hold a pointer when running on x86 or x64 versions of AutoIt
; str - an ANSI string (a minimum of 65536 chars is allocated)
; wstr - a UNICODE wide character string (a minimum of 65536 chars is allocated)
; bstr - a composite data type that consists of a length prefix, a data string and a terminator
; variant - a tagged union that can be used to represent any other data type
; idispatch, object - a composite data type that represents object with IDispatch interface
; ===============================================================================================================================

;--------------------------------------------------------------------------------------------------------------------------------------
#region Variable definitions

Global Const $gh_AU3Obj_kernel32dll = DllOpen("kernel32.dll")
Global Const $gh_AU3Obj_oleautdll = DllOpen("oleaut32.dll")
Global Const $gh_AU3Obj_ole32dll = DllOpen("ole32.dll")

Global Const $__Au3Obj_X64 = @AutoItX64

Global Const $__Au3Obj_VT_EMPTY = 0
Global Const $__Au3Obj_VT_NULL = 1
Global Const $__Au3Obj_VT_I2 = 2
Global Const $__Au3Obj_VT_I4 = 3
Global Const $__Au3Obj_VT_R4 = 4
Global Const $__Au3Obj_VT_R8 = 5
Global Const $__Au3Obj_VT_CY = 6
Global Const $__Au3Obj_VT_DATE = 7
Global Const $__Au3Obj_VT_BSTR = 8
Global Const $__Au3Obj_VT_DISPATCH = 9
Global Const $__Au3Obj_VT_ERROR = 10
Global Const $__Au3Obj_VT_BOOL = 11
Global Const $__Au3Obj_VT_VARIANT = 12
Global Const $__Au3Obj_VT_UNKNOWN = 13
Global Const $__Au3Obj_VT_DECIMAL = 14
Global Const $__Au3Obj_VT_I1 = 16
Global Const $__Au3Obj_VT_UI1 = 17
Global Const $__Au3Obj_VT_UI2 = 18
Global Const $__Au3Obj_VT_UI4 = 19
Global Const $__Au3Obj_VT_I8 = 20
Global Const $__Au3Obj_VT_UI8 = 21
Global Const $__Au3Obj_VT_INT = 22
Global Const $__Au3Obj_VT_UINT = 23
Global Const $__Au3Obj_VT_VOID = 24
Global Const $__Au3Obj_VT_HRESULT = 25
Global Const $__Au3Obj_VT_PTR = 26
Global Const $__Au3Obj_VT_SAFEARRAY = 27
Global Const $__Au3Obj_VT_CARRAY = 28
Global Const $__Au3Obj_VT_USERDEFINED = 29
Global Const $__Au3Obj_VT_LPSTR = 30
Global Const $__Au3Obj_VT_LPWSTR = 31
Global Const $__Au3Obj_VT_RECORD = 36
Global Const $__Au3Obj_VT_INT_PTR = 37
Global Const $__Au3Obj_VT_UINT_PTR = 38
Global Const $__Au3Obj_VT_FILETIME = 64
Global Const $__Au3Obj_VT_BLOB = 65
Global Const $__Au3Obj_VT_STREAM = 66
Global Const $__Au3Obj_VT_STORAGE = 67
Global Const $__Au3Obj_VT_STREAMED_OBJECT = 68
Global Const $__Au3Obj_VT_STORED_OBJECT = 69
Global Const $__Au3Obj_VT_BLOB_OBJECT = 70
Global Const $__Au3Obj_VT_CF = 71
Global Const $__Au3Obj_VT_CLSID = 72
Global Const $__Au3Obj_VT_VERSIONED_STREAM = 73
Global Const $__Au3Obj_VT_BSTR_BLOB = 0xfff
Global Const $__Au3Obj_VT_VECTOR = 0x1000
Global Const $__Au3Obj_VT_ARRAY = 0x2000
Global Const $__Au3Obj_VT_BYREF = 0x4000
Global Const $__Au3Obj_VT_RESERVED = 0x8000
Global Const $__Au3Obj_VT_ILLEGAL = 0xffff
Global Const $__Au3Obj_VT_ILLEGALMASKED = 0xfff
Global Const $__Au3Obj_VT_TYPEMASK = 0xfff

Global Const $__Au3Obj_tagVARIANT = "word vt;word r1;word r2;word r3;ptr data; ptr"

Global Const $__Au3Obj_VARIANT_SIZE = DllStructGetSize(DllStructCreate($__Au3Obj_tagVARIANT, 1))
Global Const $__Au3Obj_PTR_SIZE = DllStructGetSize(DllStructCreate('ptr', 1))
Global Const $__Au3Obj_tagSAFEARRAYBOUND = "ulong cElements; long lLbound;"

Global $ghAutoItObjectDLL = -1, $giAutoItObjectDLLRef = 0

;===============================================================================
#interface "IUnknown"
Global Const $sIID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
; Definition
Global $dtagIUnknown = "QueryInterface hresult(ptr;ptr*);" & _
		"AddRef dword();" & _
		"Release dword();"
; List
Global $ltagIUnknown = "QueryInterface;" & _
		"AddRef;" & _
		"Release;"
;===============================================================================
;===============================================================================
#interface "IDispatch"
Global Const $sIID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
; Definition
Global $dtagIDispatch = $dtagIUnknown & _
		"GetTypeInfoCount hresult(dword*);" & _
		"GetTypeInfo hresult(dword;dword;ptr*);" & _
		"GetIDsOfNames hresult(ptr;ptr;dword;dword;ptr);" & _
		"Invoke hresult(dword;ptr;dword;word;ptr;ptr;ptr;ptr);"
; List
Global $ltagIDispatch = $ltagIUnknown & _
		"GetTypeInfoCount;" & _
		"GetTypeInfo;" & _
		"GetIDsOfNames;" & _
		"Invoke;"
;===============================================================================

#endregion Variable definitions
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region Misc

DllCall($gh_AU3Obj_ole32dll, 'long', 'OleInitialize', 'ptr', 0)
OnAutoItExitRegister("__Au3Obj_OleUninitialize")
Func __Au3Obj_OleUninitialize()
	; Author: Prog@ndy
	DllCall($gh_AU3Obj_ole32dll, 'long', 'OleUninitialize')
	_AutoItObject_Shutdown(True)
EndFunc   ;==>__Au3Obj_OleUninitialize

Func __Au3Obj_IUnknown_AddRef($vObj)
	Local $sType = "ptr"
	If IsObj($vObj) Then $sType = "idispatch"
	Local $tVARIANT = DllStructCreate($__Au3Obj_tagVARIANT)
	; Actual call
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "long", "DispCallFunc", _
			$sType, $vObj, _
			"dword", $__Au3Obj_PTR_SIZE, _ ; offset (4 for x86, 8 for x64)
			"dword", 4, _ ; CC_STDCALL
			"dword", $__Au3Obj_VT_UINT, _
			"dword", 0, _ ; number of function parameters
			"ptr", 0, _ ; parameters related
			"ptr", 0, _ ; parameters related
			"ptr", DllStructGetPtr($tVARIANT))
	If @error Or $aCall[0] Then Return SetError(1, 0, 0)
	; Collect returned
	Return DllStructGetData(DllStructCreate("dword", DllStructGetPtr($tVARIANT, "data")), 1)
EndFunc   ;==>__Au3Obj_IUnknown_AddRef

Func __Au3Obj_IUnknown_Release($vObj)
	Local $sType = "ptr"
	If IsObj($vObj) Then $sType = "idispatch"
	Local $tVARIANT = DllStructCreate($__Au3Obj_tagVARIANT)
	; Actual call
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "long", "DispCallFunc", _
			$sType, $vObj, _
			"dword", 2 * $__Au3Obj_PTR_SIZE, _ ; offset (8 for x86, 16 for x64)
			"dword", 4, _ ; CC_STDCALL
			"dword", $__Au3Obj_VT_UINT, _
			"dword", 0, _ ; number of function parameters
			"ptr", 0, _ ; parameters related
			"ptr", 0, _ ; parameters related
			"ptr", DllStructGetPtr($tVARIANT))
	If @error Or $aCall[0] Then Return SetError(1, 0, 0)
	; Collect returned
	Return DllStructGetData(DllStructCreate("dword", DllStructGetPtr($tVARIANT, "data")), 1)
EndFunc   ;==>__Au3Obj_IUnknown_Release

Func __Au3Obj_GetMethods($tagInterface)
	Local $sMethods = StringReplace(StringRegExpReplace($tagInterface, "\h*(\w+)\h*(\w+\*?)\h*(\((.*?)\))\h*(;|;*\z)", "$1\|$2;$4" & @LF), ";" & @LF, @LF)
	If $sMethods = $tagInterface Then $sMethods = StringReplace(StringRegExpReplace($tagInterface, "\h*(\w+)\h*(;|;*\z)", "$1\|" & @LF), ";" & @LF, @LF)
	Return StringTrimRight($sMethods, 1)
EndFunc   ;==>__Au3Obj_GetMethods

Func __Au3Obj_ObjStructGetElements($sTag, ByRef $sAlign)
	Local $sAlignment = StringRegExpReplace($sTag, "\h*(align\h+\d+)\h*;.*", "$1")
	If $sAlignment <> $sTag Then
		$sAlign = $sAlignment
		$sTag = StringRegExpReplace($sTag, "\h*(align\h+\d+)\h*;", "")
	EndIf
	; Return StringRegExp($sTag, "\h*\w+\h*(\w+)\h*", 3) ; DO NOT REMOVE THIS LINE
	Return StringTrimRight(StringRegExpReplace($sTag, "\h*\w+\h*(\w+)\h*(\[\d+\])*\h*(;|;*\z)\h*", "$1;"), 1)
EndFunc   ;==>__Au3Obj_ObjStructGetElements

#endregion Misc
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region SafeArray
Func __Au3Obj_SafeArrayCreate($vType, $cDims, $rgsabound)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "ptr", "SafeArrayCreate", "dword", $vType, "uint", $cDims, 'ptr', $rgsabound)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SafeArrayCreate

Func __Au3Obj_SafeArrayDestroy($pSafeArray)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "int", "SafeArrayDestroy", "ptr", $pSafeArray)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SafeArrayDestroy

Func __Au3Obj_SafeArrayAccessData($pSafeArray, ByRef $pArrayData)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "int", "SafeArrayAccessData", "ptr", $pSafeArray, 'ptr*', 0)
	If @error Then Return SetError(1, 0, 1)
	$pArrayData = $aCall[2]
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SafeArrayAccessData

Func __Au3Obj_SafeArrayUnaccessData($pSafeArray)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "int", "SafeArrayUnaccessData", "ptr", $pSafeArray)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SafeArrayUnaccessData

Func __Au3Obj_SafeArrayGetUBound($pSafeArray, $iDim, ByRef $iBound)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "int", "SafeArrayGetUBound", "ptr", $pSafeArray, 'uint', $iDim, 'long*', 0)
	If @error Then Return SetError(1, 0, 1)
	$iBound = $aCall[3]
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SafeArrayGetUBound

Func __Au3Obj_SafeArrayGetLBound($pSafeArray, $iDim, ByRef $iBound)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "int", "SafeArrayGetLBound", "ptr", $pSafeArray, 'uint', $iDim, 'long*', 0)
	If @error Then Return SetError(1, 0, 1)
	$iBound = $aCall[3]
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SafeArrayGetLBound

Func __Au3Obj_SafeArrayGetDim($pSafeArray)
	Local $aResult = DllCall($gh_AU3Obj_oleautdll, "uint", "SafeArrayGetDim", "ptr", $pSafeArray)
	If @error Then Return SetError(1, 0, 0)
	Return $aResult[0]
EndFunc   ;==>__Au3Obj_SafeArrayGetDim

Func __Au3Obj_CreateSafeArrayVariant(ByRef Const $aArray)
	; Author: Prog@ndy
	Local $iDim = UBound($aArray, 0), $pData, $pSafeArray, $bound, $subBound, $tBound
	Switch $iDim
		Case 1
			$bound = UBound($aArray) - 1
			$tBound = DllStructCreate($__Au3Obj_tagSAFEARRAYBOUND)
			DllStructSetData($tBound, 1, $bound + 1)
			$pSafeArray = __Au3Obj_SafeArrayCreate($__Au3Obj_VT_VARIANT, 1, DllStructGetPtr($tBound))
			If 0 = __Au3Obj_SafeArrayAccessData($pSafeArray, $pData) Then
				For $i = 0 To $bound
					_AutoItObject_VariantInit($pData + $i * $__Au3Obj_VARIANT_SIZE)
					_AutoItObject_VariantSet($pData + $i * $__Au3Obj_VARIANT_SIZE, $aArray[$i])
				Next
				__Au3Obj_SafeArrayUnaccessData($pSafeArray)
			EndIf
			Return $pSafeArray
		Case 2
			$bound = UBound($aArray, 1) - 1
			$subBound = UBound($aArray, 2) - 1
			$tBound = DllStructCreate($__Au3Obj_tagSAFEARRAYBOUND & $__Au3Obj_tagSAFEARRAYBOUND)
			DllStructSetData($tBound, 3, $bound + 1)
			DllStructSetData($tBound, 1, $subBound + 1)
			$pSafeArray = __Au3Obj_SafeArrayCreate($__Au3Obj_VT_VARIANT, 2, DllStructGetPtr($tBound))
			If 0 = __Au3Obj_SafeArrayAccessData($pSafeArray, $pData) Then
				For $i = 0 To $bound
					For $j = 0 To $subBound
						_AutoItObject_VariantInit($pData + ($j + $i * ($subBound + 1)) * $__Au3Obj_VARIANT_SIZE)
						_AutoItObject_VariantSet($pData + ($j + $i * ($subBound + 1)) * $__Au3Obj_VARIANT_SIZE, $aArray[$i][$j])
					Next
				Next
				__Au3Obj_SafeArrayUnaccessData($pSafeArray)
			EndIf
			Return $pSafeArray
		Case Else
			Return 0
	EndSwitch
EndFunc   ;==>__Au3Obj_CreateSafeArrayVariant

Func __Au3Obj_ReadSafeArrayVariant($pSafeArray)
	; Author: Prog@ndy
	Local $iDim = __Au3Obj_SafeArrayGetDim($pSafeArray), $pData, $lbound, $bound, $subBound
	Switch $iDim
		Case 1
			__Au3Obj_SafeArrayGetLBound($pSafeArray, 1, $lbound)
			__Au3Obj_SafeArrayGetUBound($pSafeArray, 1, $bound)
			$bound -= $lbound
			Local $array[$bound + 1]
			If 0 = __Au3Obj_SafeArrayAccessData($pSafeArray, $pData) Then
				For $i = 0 To $bound
					$array[$i] = _AutoItObject_VariantRead($pData + $i * $__Au3Obj_VARIANT_SIZE)
				Next
				__Au3Obj_SafeArrayUnaccessData($pSafeArray)
			EndIf
			Return $array
		Case 2
			__Au3Obj_SafeArrayGetLBound($pSafeArray, 2, $lbound)
			__Au3Obj_SafeArrayGetUBound($pSafeArray, 2, $bound)
			$bound -= $lbound
			__Au3Obj_SafeArrayGetLBound($pSafeArray, 1, $lbound)
			__Au3Obj_SafeArrayGetUBound($pSafeArray, 1, $subBound)
			$subBound -= $lbound
			Local $array[$bound + 1][$subBound + 1]
			If 0 = __Au3Obj_SafeArrayAccessData($pSafeArray, $pData) Then
				For $i = 0 To $bound
					For $j = 0 To $subBound
						$array[$i][$j] = _AutoItObject_VariantRead($pData + ($j + $i * ($subBound + 1)) * $__Au3Obj_VARIANT_SIZE)
					Next
				Next
				__Au3Obj_SafeArrayUnaccessData($pSafeArray)
			EndIf
			Return $array
		Case Else
			Return 0
	EndSwitch
EndFunc   ;==>__Au3Obj_ReadSafeArrayVariant

#endregion SafeArray
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region Memory

Func __Au3Obj_CoTaskMemAlloc($iSize)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_ole32dll, "ptr", "CoTaskMemAlloc", "uint_ptr", $iSize)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_CoTaskMemAlloc

Func __Au3Obj_CoTaskMemFree($pCoMem)
	; Author: Prog@ndy
	DllCall($gh_AU3Obj_ole32dll, "none", "CoTaskMemFree", "ptr", $pCoMem)
	If @error Then Return SetError(1, 0, 0)
EndFunc   ;==>__Au3Obj_CoTaskMemFree

Func __Au3Obj_CoTaskMemRealloc($pCoMem, $iSize)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_ole32dll, "ptr", "CoTaskMemRealloc", 'ptr', $pCoMem, "uint_ptr", $iSize)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_CoTaskMemRealloc

Func __Au3Obj_GlobalAlloc($iSize, $iFlag)
	Local $aCall = DllCall($gh_AU3Obj_kernel32dll, "ptr", "GlobalAlloc", "dword", $iFlag, "dword_ptr", $iSize)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_GlobalAlloc

Func __Au3Obj_GlobalFree($pPointer)
	Local $aCall = DllCall($gh_AU3Obj_kernel32dll, "ptr", "GlobalFree", "ptr", $pPointer)
	If @error Or $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc   ;==>__Au3Obj_GlobalFree

#endregion Memory
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region SysString

Func __Au3Obj_SysAllocString($str)
	; Author: monoceres
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "ptr", "SysAllocString", "wstr", $str)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SysAllocString
Func __Au3Obj_SysCopyString($pBSTR)
	; Author: Prog@ndy
	If Not $pBSTR Then Return SetError(2, 0, 0)
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "ptr", "SysAllocStringLen", "ptr", $pBSTR, "uint", __Au3Obj_SysStringLen($pBSTR))
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SysCopyString

Func __Au3Obj_SysReAllocString(ByRef $pBSTR, $str)
	; Author: Prog@ndy
	If Not $pBSTR Then Return SetError(2, 0, 0)
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "int", "SysReAllocString", 'ptr*', $pBSTR, "wstr", $str)
	If @error Then Return SetError(1, 0, 0)
	$pBSTR = $aCall[1]
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SysReAllocString

Func __Au3Obj_SysFreeString($pBSTR)
	; Author: Prog@ndy
	If Not $pBSTR Then Return SetError(2, 0, 0)
	DllCall($gh_AU3Obj_oleautdll, "none", "SysFreeString", "ptr", $pBSTR)
	If @error Then Return SetError(1, 0, 0)
EndFunc   ;==>__Au3Obj_SysFreeString

Func __Au3Obj_SysStringLen($pBSTR)
	; Author: Prog@ndy
	If Not $pBSTR Then Return SetError(2, 0, 0)
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "uint", "SysStringLen", "ptr", $pBSTR)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_SysStringLen

Func __Au3Obj_SysReadString($pBSTR, $iLen = -1)
	; Author: Prog@ndy
	If Not $pBSTR Then Return SetError(2, 0, '')
	If $iLen < 1 Then $iLen = __Au3Obj_SysStringLen($pBSTR)
	If $iLen < 1 Then Return SetError(1, 0, '')
	Return DllStructGetData(DllStructCreate("wchar[" & $iLen & "]", $pBSTR), 1)
EndFunc   ;==>__Au3Obj_SysReadString

Func __Au3Obj_PtrStringLen($pStr)
	; Author: Prog@ndy
	Local $aResult = DllCall($gh_AU3Obj_kernel32dll, 'int', 'lstrlenW', 'ptr', $pStr)
	If @error Then Return SetError(1, 0, 0)
	Return $aResult[0]
EndFunc   ;==>__Au3Obj_PtrStringLen

Func __Au3Obj_PtrStringRead($pStr, $iLen = -1)
	; Author: Prog@ndy
	If $iLen < 1 Then $iLen = __Au3Obj_PtrStringLen($pStr)
	If $iLen < 1 Then Return SetError(1, 0, '')
	Return DllStructGetData(DllStructCreate("wchar[" & $iLen & "]", $pStr), 1)
EndFunc   ;==>__Au3Obj_PtrStringRead

#endregion SysString
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region Proxy Functions

Func __Au3Obj_FunctionProxy($FuncName, $oSelf) ; allows binary code to call autoit functions
	Local $arg = $oSelf.__params__ ; fetch params
	If IsArray($arg) Then
		Local $ret = Call($FuncName, $arg) ; Call
		If @error = 0xDEAD And @extended = 0xBEEF Then Return 0
		$oSelf.__error__ = @error ; set error
		$oSelf.__result__ = $ret ; set result
		Return 1
	EndIf
	; return error when params-array could not be created
EndFunc   ;==>__Au3Obj_FunctionProxy

Func __Au3Obj_EnumFunctionProxy($iAction, $FuncName, $oSelf, $pVarCurrent, $pVarResult)
	Local $Current, $ret
	Switch $iAction
		Case 0 ; Next
			$Current = $oSelf.__bridge__(Number($pVarCurrent))
			$ret = Execute($FuncName & "($oSelf, $Current)")
			If @error Then Return False
			$oSelf.__bridge__(Number($pVarCurrent)) = $Current
			$oSelf.__bridge__(Number($pVarResult)) = $ret
			Return 1
		Case 1 ;Skip
			Return False
		Case 2 ; Reset
			$Current = $oSelf.__bridge__(Number($pVarCurrent))
			$ret = Execute($FuncName & "($oSelf, $Current)")
			If @error Or Not $ret Then Return False
			$oSelf.__bridge__(Number($pVarCurrent)) = $Current
			Return True
	EndSwitch
EndFunc   ;==>__Au3Obj_EnumFunctionProxy

#endregion Proxy Functions
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region Call Pointer

Func __Au3Obj_PointerCall($sRetType, $pAddress, $sType1 = "", $vParam1 = 0, $sType2 = "", $vParam2 = 0, $sType3 = "", $vParam3 = 0, $sType4 = "", $vParam4 = 0, $sType5 = "", $vParam5 = 0, $sType6 = "", $vParam6 = 0, $sType7 = "", $vParam7 = 0, $sType8 = "", $vParam8 = 0, $sType9 = "", $vParam9 = 0, $sType10 = "", $vParam10 = 0, $sType11 = "", $vParam11 = 0, $sType12 = "", $vParam12 = 0, $sType13 = "", $vParam13 = 0, $sType14 = "", $vParam14 = 0, $sType15 = "", $vParam15 = 0, $sType16 = "", $vParam16 = 0, $sType17 = "", $vParam17 = 0, $sType18 = "", $vParam18 = 0, $sType19 = "", $vParam19 = 0, $sType20 = "", $vParam20 = 0)
	; Author: Ward, Prog@ndy, trancexx
	Local Static $pHook, $hPseudo, $tPtr, $sFuncName = "MemoryCallEntry"
	If $pAddress Then
		If Not $pHook Then
			Local $sDll = "AutoItObject.dll"
			If $__Au3Obj_X64 Then $sDll = "AutoItObject_X64.dll"
			$hPseudo = DllOpen($sDll)
			If $hPseudo = -1 Then
				$sDll = "kernel32.dll"
				$sFuncName = "GlobalFix"
				$hPseudo = DllOpen($sDll)
			EndIf
			Local $aCall = DllCall($gh_AU3Obj_kernel32dll, "ptr", "GetModuleHandleW", "wstr", $sDll)
			If @error Or Not $aCall[0] Then Return SetError(7, @error, 0) ; Couldn't get dll handle
			Local $hModuleHandle = $aCall[0]
			$aCall = DllCall($gh_AU3Obj_kernel32dll, "ptr", "GetProcAddress", "ptr", $hModuleHandle, "str", $sFuncName)
			If @error Then Return SetError(8, @error, 0) ; Wanted function not found
			$pHook = $aCall[0]
			$aCall = DllCall($gh_AU3Obj_kernel32dll, "bool", "VirtualProtect", "ptr", $pHook, "dword", 7 + 5 * $__Au3Obj_X64, "dword", 64, "dword*", 0)
			If @error Or Not $aCall[0] Then Return SetError(9, @error, 0) ; Unable to set MEM_EXECUTE_READWRITE
			If $__Au3Obj_X64 Then
				DllStructSetData(DllStructCreate("word", $pHook), 1, 0xB848)
				DllStructSetData(DllStructCreate("word", $pHook + 10), 1, 0xE0FF)
			Else
				DllStructSetData(DllStructCreate("byte", $pHook), 1, 0xB8)
				DllStructSetData(DllStructCreate("word", $pHook + 5), 1, 0xE0FF)
			EndIf
			$tPtr = DllStructCreate("ptr", $pHook + 1 + $__Au3Obj_X64)
		EndIf
		DllStructSetData($tPtr, 1, $pAddress)
		Local $aRet
		Switch @NumParams
			Case 2
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName)
			Case 4
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1)
			Case 6
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2)
			Case 8
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3)
			Case 10
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4)
			Case 12
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5)
			Case 14
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6)
			Case 16
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7)
			Case 18
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8)
			Case 20
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9)
			Case 22
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10)
			Case 24
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11)
			Case 26
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12)
			Case 28
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12, $sType13, $vParam13)
			Case 30
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12, $sType13, $vParam13, $sType14, $vParam14)
			Case 32
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12, $sType13, $vParam13, $sType14, $vParam14, $sType15, $vParam15)
			Case 34
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12, $sType13, $vParam13, $sType14, $vParam14, $sType15, $vParam15, $sType16, $vParam16)
			Case 36
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12, $sType13, $vParam13, $sType14, $vParam14, $sType15, $vParam15, $sType16, $vParam16, $sType17, $vParam17)
			Case 38
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12, $sType13, $vParam13, $sType14, $vParam14, $sType15, $vParam15, $sType16, $vParam16, $sType17, $vParam17, $sType18, $vParam18)
			Case 40
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12, $sType13, $vParam13, $sType14, $vParam14, $sType15, $vParam15, $sType16, $vParam16, $sType17, $vParam17, $sType18, $vParam18, $sType19, $vParam19)
			Case 42
				$aRet = DllCall($hPseudo, $sRetType, $sFuncName, $sType1, $vParam1, $sType2, $vParam2, $sType3, $vParam3, $sType4, $vParam4, $sType5, $vParam5, $sType6, $vParam6, $sType7, $vParam7, $sType8, $vParam8, $sType9, $vParam9, $sType10, $vParam10, $sType11, $vParam11, $sType12, $vParam12, $sType13, $vParam13, $sType14, $vParam14, $sType15, $vParam15, $sType16, $vParam16, $sType17, $vParam17, $sType18, $vParam18, $sType19, $vParam19, $sType20, $vParam20)
			Case Else
				If Mod(@NumParams, 2) Then Return SetError(4, 0, 0) ; Bad number of parameters
				Return SetError(5, 0, 0) ; Max number of parameters exceeded
		EndSwitch
		Return SetError(@error, @extended, $aRet) ; All went well. Error description and return values like with DllCall()
	EndIf
	Return SetError(6, 0, 0) ; Null address specified
EndFunc   ;==>__Au3Obj_PointerCall

#endregion Call Pointer
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region Embedded DLL

Func __Au3Obj_Mem_DllOpen($bBinaryImage = 0, $sSubrogor = "cmd.exe")
	If Not $bBinaryImage Then
		If $__Au3Obj_X64 Then
			$bBinaryImage = __Au3Obj_Mem_BinDll_X64()
		Else
			$bBinaryImage = __Au3Obj_Mem_BinDll()
		EndIf
	EndIf
	; Make structure out of binary data that was passed
	Local $tBinary = DllStructCreate("byte[" & BinaryLen($bBinaryImage) & "]")
	DllStructSetData($tBinary, 1, $bBinaryImage) ; fill the structure
	; Get pointer to it
	Local $pPointer = DllStructGetPtr($tBinary)
	; Start processing passed binary data. 'Reading' PE format follows.
	Local $tIMAGE_DOS_HEADER = DllStructCreate("char Magic[2];" & _
			"word BytesOnLastPage;" & _
			"word Pages;" & _
			"word Relocations;" & _
			"word SizeofHeader;" & _
			"word MinimumExtra;" & _
			"word MaximumExtra;" & _
			"word SS;" & _
			"word SP;" & _
			"word Checksum;" & _
			"word IP;" & _
			"word CS;" & _
			"word Relocation;" & _
			"word Overlay;" & _
			"char Reserved[8];" & _
			"word OEMIdentifier;" & _
			"word OEMInformation;" & _
			"char Reserved2[20];" & _
			"dword AddressOfNewExeHeader", _
			$pPointer)
	; Move pointer
	$pPointer += DllStructGetData($tIMAGE_DOS_HEADER, "AddressOfNewExeHeader") ; move to PE file header
	$pPointer += 4 ; size of skipped $tIMAGE_NT_SIGNATURE structure
	; In place of IMAGE_FILE_HEADER structure
	Local $tIMAGE_FILE_HEADER = DllStructCreate("word Machine;" & _
			"word NumberOfSections;" & _
			"dword TimeDateStamp;" & _
			"dword PointerToSymbolTable;" & _
			"dword NumberOfSymbols;" & _
			"word SizeOfOptionalHeader;" & _
			"word Characteristics", _
			$pPointer)
	; Get number of sections
	Local $iNumberOfSections = DllStructGetData($tIMAGE_FILE_HEADER, "NumberOfSections")
	; Move pointer
	$pPointer += 20 ; size of $tIMAGE_FILE_HEADER structure
	; Determine the type
	Local $tMagic = DllStructCreate("word Magic;", $pPointer)
	Local $iMagic = DllStructGetData($tMagic, 1)
	Local $tIMAGE_OPTIONAL_HEADER
	If $iMagic = 267 Then ; x86 version
		If $__Au3Obj_X64 Then Return SetError(1, 0, -1) ; incompatible versions
		$tIMAGE_OPTIONAL_HEADER = DllStructCreate("word Magic;" & _
				"byte MajorLinkerVersion;" & _
				"byte MinorLinkerVersion;" & _
				"dword SizeOfCode;" & _
				"dword SizeOfInitializedData;" & _
				"dword SizeOfUninitializedData;" & _
				"dword AddressOfEntryPoint;" & _
				"dword BaseOfCode;" & _
				"dword BaseOfData;" & _
				"dword ImageBase;" & _
				"dword SectionAlignment;" & _
				"dword FileAlignment;" & _
				"word MajorOperatingSystemVersion;" & _
				"word MinorOperatingSystemVersion;" & _
				"word MajorImageVersion;" & _
				"word MinorImageVersion;" & _
				"word MajorSubsystemVersion;" & _
				"word MinorSubsystemVersion;" & _
				"dword Win32VersionValue;" & _
				"dword SizeOfImage;" & _
				"dword SizeOfHeaders;" & _
				"dword CheckSum;" & _
				"word Subsystem;" & _
				"word DllCharacteristics;" & _
				"dword SizeOfStackReserve;" & _
				"dword SizeOfStackCommit;" & _
				"dword SizeOfHeapReserve;" & _
				"dword SizeOfHeapCommit;" & _
				"dword LoaderFlags;" & _
				"dword NumberOfRvaAndSizes", _
				$pPointer)
		; Move pointer
		$pPointer += 96 ; size of $tIMAGE_OPTIONAL_HEADER
	ElseIf $iMagic = 523 Then ; x64 version
		If Not $__Au3Obj_X64 Then Return SetError(1, 0, -1) ; incompatible versions
		$tIMAGE_OPTIONAL_HEADER = DllStructCreate("word Magic;" & _
				"byte MajorLinkerVersion;" & _
				"byte MinorLinkerVersion;" & _
				"dword SizeOfCode;" & _
				"dword SizeOfInitializedData;" & _
				"dword SizeOfUninitializedData;" & _
				"dword AddressOfEntryPoint;" & _
				"dword BaseOfCode;" & _
				"uint64 ImageBase;" & _
				"dword SectionAlignment;" & _
				"dword FileAlignment;" & _
				"word MajorOperatingSystemVersion;" & _
				"word MinorOperatingSystemVersion;" & _
				"word MajorImageVersion;" & _
				"word MinorImageVersion;" & _
				"word MajorSubsystemVersion;" & _
				"word MinorSubsystemVersion;" & _
				"dword Win32VersionValue;" & _
				"dword SizeOfImage;" & _
				"dword SizeOfHeaders;" & _
				"dword CheckSum;" & _
				"word Subsystem;" & _
				"word DllCharacteristics;" & _
				"uint64 SizeOfStackReserve;" & _
				"uint64 SizeOfStackCommit;" & _
				"uint64 SizeOfHeapReserve;" & _
				"uint64 SizeOfHeapCommit;" & _
				"dword LoaderFlags;" & _
				"dword NumberOfRvaAndSizes", _
				$pPointer)
		; Move pointer
		$pPointer += 112 ; size of $tIMAGE_OPTIONAL_HEADER
	Else
		Return SetError(1, 0, -1) ; incompatible versions
	EndIf
	; Extract data
	Local $iEntryPoint = DllStructGetData($tIMAGE_OPTIONAL_HEADER, "AddressOfEntryPoint") ; if loaded binary image would start executing at this address
	Local $pOptionalHeaderImageBase = DllStructGetData($tIMAGE_OPTIONAL_HEADER, "ImageBase") ; address of the first byte of the image when it's loaded in memory
	$pPointer += 8 ; skipping IMAGE_DIRECTORY_ENTRY_EXPORT
	; Import Directory
	Local $tIMAGE_DIRECTORY_ENTRY_IMPORT = DllStructCreate("dword VirtualAddress; dword Size", $pPointer)
	; Collect data
	Local $pAddressImport = DllStructGetData($tIMAGE_DIRECTORY_ENTRY_IMPORT, "VirtualAddress")
;~ 	Local $iSizeImport = DllStructGetData($tIMAGE_DIRECTORY_ENTRY_IMPORT, "Size")
	$pPointer += 8 ; size of $tIMAGE_DIRECTORY_ENTRY_IMPORT
	$pPointer += 24 ; skipping IMAGE_DIRECTORY_ENTRY_RESOURCE, IMAGE_DIRECTORY_ENTRY_EXCEPTION, IMAGE_DIRECTORY_ENTRY_SECURITY
	; Base Relocation Directory
	Local $tIMAGE_DIRECTORY_ENTRY_BASERELOC = DllStructCreate("dword VirtualAddress; dword Size", $pPointer)
	; Collect data
	Local $pAddressNewBaseReloc = DllStructGetData($tIMAGE_DIRECTORY_ENTRY_BASERELOC, "VirtualAddress")
	Local $iSizeBaseReloc = DllStructGetData($tIMAGE_DIRECTORY_ENTRY_BASERELOC, "Size")
	$pPointer += 8 ; size of IMAGE_DIRECTORY_ENTRY_BASERELOC
	$pPointer += 40 ; skipping IMAGE_DIRECTORY_ENTRY_DEBUG, IMAGE_DIRECTORY_ENTRY_COPYRIGHT, IMAGE_DIRECTORY_ENTRY_GLOBALPTR, IMAGE_DIRECTORY_ENTRY_TLS, IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG
	$pPointer += 40 ; five more generally unused data directories
	; Load the victim
	Local $pBaseAddress = __Au3Obj_Mem_LoadLibraryEx($sSubrogor, 1) ; "lighter" loading, DONT_RESOLVE_DLL_REFERENCES
	If @error Then Return SetError(2, 0, -1) ; Couldn't load subrogor
	Local $pHeadersNew = DllStructGetPtr($tIMAGE_DOS_HEADER) ; starting address of binary image headers
	Local $iOptionalHeaderSizeOfHeaders = DllStructGetData($tIMAGE_OPTIONAL_HEADER, "SizeOfHeaders") ; the size of the MS-DOS stub, the PE header, and the section headers
	; Set proper memory protection for writting headers (PAGE_READWRITE)
	If Not __Au3Obj_Mem_VirtualProtect($pBaseAddress, $iOptionalHeaderSizeOfHeaders, 4) Then Return SetError(3, 0, -1) ; Couldn't set proper protection for headers
	; Write NEW headers
	DllStructSetData(DllStructCreate("byte[" & $iOptionalHeaderSizeOfHeaders & "]", $pBaseAddress), 1, DllStructGetData(DllStructCreate("byte[" & $iOptionalHeaderSizeOfHeaders & "]", $pHeadersNew), 1))
	; Dealing with sections. Will write them.
	Local $tIMAGE_SECTION_HEADER
	Local $iSizeOfRawData, $pPointerToRawData
	Local $iVirtualSize, $iVirtualAddress
	Local $pRelocRaw
	For $i = 1 To $iNumberOfSections
		$tIMAGE_SECTION_HEADER = DllStructCreate("char Name[8];" & _
				"dword UnionOfVirtualSizeAndPhysicalAddress;" & _
				"dword VirtualAddress;" & _
				"dword SizeOfRawData;" & _
				"dword PointerToRawData;" & _
				"dword PointerToRelocations;" & _
				"dword PointerToLinenumbers;" & _
				"word NumberOfRelocations;" & _
				"word NumberOfLinenumbers;" & _
				"dword Characteristics", _
				$pPointer)
		; Collect data
		$iSizeOfRawData = DllStructGetData($tIMAGE_SECTION_HEADER, "SizeOfRawData")
		$pPointerToRawData = $pHeadersNew + DllStructGetData($tIMAGE_SECTION_HEADER, "PointerToRawData")
		$iVirtualAddress = DllStructGetData($tIMAGE_SECTION_HEADER, "VirtualAddress")
		$iVirtualSize = DllStructGetData($tIMAGE_SECTION_HEADER, "UnionOfVirtualSizeAndPhysicalAddress")
		If $iVirtualSize And $iVirtualSize < $iSizeOfRawData Then $iSizeOfRawData = $iVirtualSize
		; Set MEM_EXECUTE_READWRITE for sections (PAGE_EXECUTE_READWRITE for all for simplicity)
		If Not __Au3Obj_Mem_VirtualProtect($pBaseAddress + $iVirtualAddress, $iVirtualSize, 64) Then
			$pPointer += 40 ; size of $tIMAGE_SECTION_HEADER structure
			ContinueLoop
		EndIf
		; Clean the space
		DllStructSetData(DllStructCreate("byte[" & $iVirtualSize & "]", $pBaseAddress + $iVirtualAddress), 1, DllStructGetData(DllStructCreate("byte[" & $iVirtualSize & "]"), 1))
		; If there is data to write, write it
		If $iSizeOfRawData Then DllStructSetData(DllStructCreate("byte[" & $iSizeOfRawData & "]", $pBaseAddress + $iVirtualAddress), 1, DllStructGetData(DllStructCreate("byte[" & $iSizeOfRawData & "]", $pPointerToRawData), 1))
		; Relocations
		If $iVirtualAddress <= $pAddressNewBaseReloc And $iVirtualAddress + $iSizeOfRawData > $pAddressNewBaseReloc Then $pRelocRaw = $pPointerToRawData + ($pAddressNewBaseReloc - $iVirtualAddress)
		; Imports
		If $iVirtualAddress <= $pAddressImport And $iVirtualAddress + $iSizeOfRawData > $pAddressImport Then __Au3Obj_Mem_FixImports($pPointerToRawData + ($pAddressImport - $iVirtualAddress), $pBaseAddress) ; fix imports in place
		; Move pointer
		$pPointer += 40 ; size of $tIMAGE_SECTION_HEADER structure
	Next
	; Fix relocations
	If $pAddressNewBaseReloc And $iSizeBaseReloc Then __Au3Obj_Mem_FixReloc($pRelocRaw, $iSizeBaseReloc, $pBaseAddress, $pOptionalHeaderImageBase, $iMagic = 523)
	; Entry point address
	Local $pEntryFunc = $pBaseAddress + $iEntryPoint
	; DllMain simulation
	__Au3Obj_PointerCall("bool", $pEntryFunc, "ptr", $pBaseAddress, "dword", 1, "ptr", 0) ; DLL_PROCESS_ATTACH
	; Get pseudo-handle
	Local $hPseudo = DllOpen($sSubrogor)
	__Au3Obj_Mem_FreeLibrary($pBaseAddress) ; decrement reference count
	Return $hPseudo
EndFunc   ;==>__Au3Obj_Mem_DllOpen

Func __Au3Obj_Mem_FixReloc($pData, $iSize, $pAddressNew, $pAddressOld, $fImageX64)
	Local $iDelta = $pAddressNew - $pAddressOld ; dislocation value
	Local $tIMAGE_BASE_RELOCATION, $iRelativeMove
	Local $iVirtualAddress, $iSizeofBlock, $iNumberOfEntries
	Local $tEnries, $iData, $tAddress
	Local $iFlag = 3 + 7 * $fImageX64 ; IMAGE_REL_BASED_HIGHLOW = 3 or IMAGE_REL_BASED_DIR64 = 10
	While $iRelativeMove < $iSize ; for all data available
		$tIMAGE_BASE_RELOCATION = DllStructCreate("dword VirtualAddress; dword SizeOfBlock", $pData + $iRelativeMove)
		$iVirtualAddress = DllStructGetData($tIMAGE_BASE_RELOCATION, "VirtualAddress")
		$iSizeofBlock = DllStructGetData($tIMAGE_BASE_RELOCATION, "SizeOfBlock")
		$iNumberOfEntries = ($iSizeofBlock - 8) / 2
		$tEnries = DllStructCreate("word[" & $iNumberOfEntries & "]", DllStructGetPtr($tIMAGE_BASE_RELOCATION) + 8)
		; Go through all entries
		For $i = 1 To $iNumberOfEntries
			$iData = DllStructGetData($tEnries, 1, $i)
			If BitShift($iData, 12) = $iFlag Then ; check type
				$tAddress = DllStructCreate("ptr", $pAddressNew + $iVirtualAddress + BitAND($iData, 0xFFF)) ; the rest of $iData is offset
				DllStructSetData($tAddress, 1, DllStructGetData($tAddress, 1) + $iDelta) ; this is what's this all about
			EndIf
		Next
		$iRelativeMove += $iSizeofBlock
	WEnd
	Return 1 ; all OK!
EndFunc   ;==>__Au3Obj_Mem_FixReloc

Func __Au3Obj_Mem_FixImports($pImportDirectory, $hInstance)
	Local $hModule, $tFuncName, $sFuncName, $pFuncAddress
	Local $tIMAGE_IMPORT_MODULE_DIRECTORY, $tModuleName
	Local $tBufferOffset2, $iBufferOffset2
	Local $iInitialOffset, $iInitialOffset2, $iOffset
	While 1
		$tIMAGE_IMPORT_MODULE_DIRECTORY = DllStructCreate("dword RVAOriginalFirstThunk;" & _
				"dword TimeDateStamp;" & _
				"dword ForwarderChain;" & _
				"dword RVAModuleName;" & _
				"dword RVAFirstThunk", _
				$pImportDirectory)
		If Not DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY, "RVAFirstThunk") Then ExitLoop ; the end
		$tModuleName = DllStructCreate("char Name[64]", $hInstance + DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY, "RVAModuleName"))
		$hModule = __Au3Obj_Mem_LoadLibraryEx(DllStructGetData($tModuleName, "Name")) ; load the module, full load
		$iInitialOffset = $hInstance + DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY, "RVAFirstThunk")
		$iInitialOffset2 = $hInstance + DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY, "RVAOriginalFirstThunk")
		If $iInitialOffset2 = $hInstance Then $iInitialOffset2 = $iInitialOffset
		$iOffset = 0 ; back to 0
		While 1
			$tBufferOffset2 = DllStructCreate("ptr", $iInitialOffset2 + $iOffset)
			$iBufferOffset2 = DllStructGetData($tBufferOffset2, 1) ; value at that address
			If Not $iBufferOffset2 Then ExitLoop ; zero value is the end
			If BitShift(BinaryMid($iBufferOffset2, $__Au3Obj_PTR_SIZE, 1), 7) Then ; MSB is set for imports by ordinal, otherwise not
				$pFuncAddress = __Au3Obj_Mem_GetAddress($hModule, BitAND($iBufferOffset2, 0xFFFFFF)) ; the rest is ordinal value
			Else
				$tFuncName = DllStructCreate("word Ordinal; char Name[64]", $hInstance + $iBufferOffset2)
				$sFuncName = DllStructGetData($tFuncName, "Name")
				$pFuncAddress = __Au3Obj_Mem_GetAddress($hModule, $sFuncName)
			EndIf
			DllStructSetData(DllStructCreate("ptr", $iInitialOffset + $iOffset), 1, $pFuncAddress) ; and this is what's this all about
			$iOffset += $__Au3Obj_PTR_SIZE ; size of $tBufferOffset2
		WEnd
		$pImportDirectory += 20 ; size of $tIMAGE_IMPORT_MODULE_DIRECTORY
	WEnd
	Return 1 ; all OK!
EndFunc   ;==>__Au3Obj_Mem_FixImports

Func __Au3Obj_Mem_Base64Decode($sData) ; Ward
	Local $bOpcode
	If $__Au3Obj_X64 Then
		$bOpcode = Binary("0x4156415541544D89CC555756534C89C34883EC20410FB64104418800418B3183FE010F84AB00000073434863D24D89C54889CE488D3C114839FE0F84A50100000FB62E4883C601E8B501000083ED2B4080FD5077E2480FBEED0FB6042884C00FBED078D3C1E20241885500EB7383FE020F841C01000031C083FE03740F4883C4205B5E5F5D415C415D415EC34863D24D89C54889CE488D3C114839FE0F84CA0000000FB62E4883C601E85301000083ED2B4080FD5077E2480FBEED0FB6042884C078D683E03F410845004983C501E964FFFFFF4863D24D89C54889CE488D3C114839FE0F84E00000000FB62E4883C601E80C01000083ED2B4080FD5077E2480FBEED0FB6042884C00FBED078D389D04D8D7501C1E20483E03041885501C1F804410845004839FE747B0FB62E4883C601E8CC00000083ED2B4080FD5077E6480FBEED0FB6042884C00FBED078D789D0C1E2064D8D6E0183E03C41885601C1F8024108064839FE0F8536FFFFFF41C7042403000000410FB6450041884424044489E84883C42029D85B5E5F5D415C415D415EC34863D24889CE4D89C6488D3C114839FE758541C7042402000000410FB60641884424044489F04883C42029D85B5E5F5D415C415D415EC341C7042401000000410FB6450041884424044489E829D8E998FEFFFF41C7042400000000410FB6450041884424044489E829D8E97CFEFFFFE8500000003EFFFFFF3F3435363738393A3B3C3DFFFFFFFEFFFFFF000102030405060708090A0B0C0D0E0F10111213141516171819FFFFFFFFFFFF1A1B1C1D1E1F202122232425262728292A2B2C2D2E2F3031323358C3")
	Else
		$bOpcode = Binary("0x5557565383EC1C8B6C243C8B5424388B5C24308B7424340FB6450488028B550083FA010F84A1000000733F8B5424388D34338954240C39F30F848B0100000FB63B83C301E8890100008D57D580FA5077E50FBED20FB6041084C00FBED078D78B44240CC1E2028810EB6B83FA020F841201000031C083FA03740A83C41C5B5E5F5DC210008B4C24388D3433894C240C39F30F84CD0000000FB63B83C301E8300100008D57D580FA5077E50FBED20FB6041084C078DA8B54240C83E03F080283C2018954240CE96CFFFFFF8B4424388D34338944240C39F30F84D00000000FB63B83C301E8EA0000008D57D580FA5077E50FBED20FB6141084D20FBEC278D78B4C240C89C283E230C1FA04C1E004081189CF83C70188410139F374750FB60383C3018844240CE8A80000000FB654240C83EA2B80FA5077E00FBED20FB6141084D20FBEC278D289C283E23CC1FA02C1E006081739F38D57018954240C8847010F8533FFFFFFC74500030000008B4C240C0FB60188450489C82B44243883C41C5B5E5F5DC210008D34338B7C243839F3758BC74500020000000FB60788450489F82B44243883C41C5B5E5F5DC210008B54240CC74500010000000FB60288450489D02B442438E9B1FEFFFFC7450000000000EB99E8500000003EFFFFFF3F3435363738393A3B3C3DFFFFFFFEFFFFFF000102030405060708090A0B0C0D0E0F10111213141516171819FFFFFFFFFFFF1A1B1C1D1E1F202122232425262728292A2B2C2D2E2F3031323358C3")
	EndIf
	Local $tCodeBuffer = DllStructCreate("byte[" & BinaryLen($bOpcode) & "]")
	DllStructSetData($tCodeBuffer, 1, $bOpcode)
	__Au3Obj_Mem_VirtualProtect(DllStructGetPtr($tCodeBuffer), DllStructGetSize($tCodeBuffer), 64)
	If @error Then Return SetError(1, 0, "")
	Local $iLen = StringLen($sData)
	Local $tOut = DllStructCreate("byte[" & $iLen & "]")
	Local $tState = DllStructCreate("byte[16]")
	Local $Call = __Au3Obj_PointerCall("int", DllStructGetPtr($tCodeBuffer), "str", $sData, "dword", $iLen, "ptr", DllStructGetPtr($tOut), "ptr", DllStructGetPtr($tState))
	If @error Then Return SetError(2, 0, "")
	Return BinaryMid(DllStructGetData($tOut, 1), 1, $Call[0])
EndFunc   ;==>__Au3Obj_Mem_Base64Decode

Func __Au3Obj_Mem_LoadLibraryEx($sModule, $iFlag = 0)
	Local $aCall = DllCall($gh_AU3Obj_kernel32dll, "handle", "LoadLibraryExW", "wstr", $sModule, "handle", 0, "dword", $iFlag)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_Mem_LoadLibraryEx

Func __Au3Obj_Mem_FreeLibrary($hModule)
	Local $aCall = DllCall($gh_AU3Obj_kernel32dll, "bool", "FreeLibrary", "handle", $hModule)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc   ;==>__Au3Obj_Mem_FreeLibrary

Func __Au3Obj_Mem_GetAddress($hModule, $vFuncName)
	Local $sType = "str"
	If IsNumber($vFuncName) Then $sType = "int" ; if ordinal value passed
	Local $aCall = DllCall($gh_AU3Obj_kernel32dll, "ptr", "GetProcAddress", "handle", $hModule, $sType, $vFuncName)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>__Au3Obj_Mem_GetAddress

Func __Au3Obj_Mem_VirtualProtect($pAddress, $iSize, $iProtection)
	Local $aCall = DllCall($gh_AU3Obj_kernel32dll, "bool", "VirtualProtect", "ptr", $pAddress, "dword_ptr", $iSize, "dword", $iProtection, "dword*", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc   ;==>__Au3Obj_Mem_VirtualProtect

Func __Au3Obj_Mem_BinDll()
    Local $sData = "TVpAAAEAAAACAAAA//8AALgAAAAAAAAACgAAAAAAAAAOH7oOALQJzSG4AUzNIVdpbjMyIC5ETEwuDQokQAAAAFBFAABMAQMAtyn9TAAAAAAAAAAA4AACIwsBCQAAOAAAABYAAAAAAAAnkwAAABAAAABQAAAAAAAQABAAAAACAAAFAAAAAAAAAAUAAAAAAAAAALAAAAACAAAAAAAAAgAABQAAEAAAEAAAAAAQAAAQAAAAAAAAEAAAAACQAAAgAgAAIJIAAAgBAAAAoAAAcAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAISSAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALk1QUkVTUzEAgAAAABAAAAAoAAAAAgAAAAAAAAAAAAAAAAAA4AAA4C5NUFJFU1My4gUAAACQAAAABgAAACoAAAAAAAAAAAAAAAAAAOAAAOAucnNyYwAAAHADAAAAoAAAAAQAAAAwAAAAAAAAAAAAAAAAAABAAADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdjIuMTcIAHcmAAD/AHQkBGoI/xUAAVAAEFD/FTQMQDVsIgAmRmASVYsA7IPsIIPk8NkAwNlUJBjffCQCEN9sJBCLFrAIQEQCUQhMx+MNkF4oneeRzUECsMhAEhgPAAAAABgY/P///zcIAg3ABBSD0gDraCw65DLYLqMNsI5EARH3wipRh5sNwUWCYRDJw1aLAPGDJgCDZhgAFI1GCBMAXBMAi8ZIXi6xaEAOB1DoQBDOkDVojGD1D1IBpgNew4sBwwCNQQjDi0EYw4AxAASJQRjCBAAUi0EcGsShAVZXAIv5jXcIZoM+EQB0B3CkhoDVDhDAYJZo8F9B5gEGiUcQX15WgAIE" & _
            "AYPBCFH/FWgqQI8X9gwEfQEIODCTDKAmAKR1L/4ACRB8n72AHIUOCuD//5+1yESCAJBosAjt9XAbABBkligQJCRkxlAIXBe/AjECMAoABrYbgAqNRhhQxwQGHFEAECEwi0YgCIvnAFEIg2YIIABe/0EzwFaLdQAMM8lXZolN9KgGYG8A5AZgrkYAkNXYBz8jnVgEAG9chA+MWJQPgFikj1i0j1gEwI9Y1I9Y5G8MUPRvNH96XAQeQEAgAGBchE4EAOmIReqIReuIAEXsiEXtiEXuA8ZF70Z0G76AKQDgM9Lzp3QMiwBNEIkBuAJAAAOA6xCLRQgc8gICQDAD/OWVLMwAOhTBMAmA0IiEETXwD4fwD8dwN9IKAFwPoAUzyYXAD5QEwYvBwhDjIGoAiFJGoRYgNffYGwjAQMIIPg2hJmDokwQAuAERAE5gBbS4HuFEsGhEUAcAseiMLhgAAOBk9SEzA9wmARl7E2YqBACpMYmdQQQywECTaDQGYQIQDECRmGNEkR/RO/YENEowI0EpPhAA/wFT9RJGGEJX+ZJ1CIs99xL/ANdqAjPJW0CLUNP/kvz/Ek0IiUYYGIvQ/QL7AgPLA0TT/SKLRhyOyvBvcT2TzLdhfLMFwLHL9wEgfBoAwNcifLMMMATC519eW111AqEgAASNSAE7Tgh2IEZXgQQEjUQAAoCHlNz7//+L+DMAwFk5RgR2DosADosMgYkMh0ABO0YEcvL/NvwCSAhZjUwAAok+ColOCF9EQDVOQPEQ+G9EMCFhDE6ECASJBolGBBUCXsMC/0kEi0EEVEAAsAM9h7CYsEigEJhIEOkFUzsG8QQz/8cGOGsCOX4IEHYei1sCHLiFAdt0DovL6LGQADS1DnO04wchJy7R6MXAAl0EDwA9gG8VjUYoUP/XjS9GSAqCowBnAleB8bEGi3YghfYJBxrzDzfjMD0B4RWLVRiLAEUMg+wYU4vKAIPhAVZXD4XCgTxgLyzwUJgLIQAMD4QjAgAAPUB81wF1bYt1HIMk" & _
            "fgg4kHKwEjKzDTCI0/9AOTwDDFC4/VAJvIi/CGAQfE4wcPxwC4AwmF9BlzGYD0BBRzGYr0G3ATCYP0FnMZg/AEAXkT5iAACgpjChBgAVQoAYQmEDyCeUwMHgBAPGAVD/dD4I6UFGANOjZx5BAAiDeAgIAA+EtSYwCIwDsNjE8R8Dtd420pPHAw+FhTEFRHYkSSRI69tfMJUHXxAJAQ+FeTCwaNACNObRAnEDoEEH0gJxAwAwQGeBWwAgAEiRPshwA2zQZFEHIpmTMwMMJD17uNGlsViHgAvqwIELjQBFHFD/djD/FWyEThBAkNMEMJ6ARsV5LGCWmHBQnmECiUGAwwccaglYZolBgN0HHIlwGIsGVhT/UAR4AogEg8YKKFbrjPkwDLMQwAo4UOk+HxF4JCEBmQ1WWpANgbUe3gN0BxMui30Ig3+jMMEIA4t1IO965xoydyQHMDcE6UdNEfgV/HVMhSAYZJBFFhB9IFdk1IClhjJglniADow/XJA5ADDxbweyiPxvF8Dxb4dhhT4dMjCwLjADnHg0fgQqATPbO8MPjBD2AoCw6AdxtHP8ANIYwBF+DIs8hwCJfQhmO8sPhW1UePA7sgQhvwPdsxQYz+ji3wPUkJDjpYDwQJhsMtQmgATxQMh2LID0UEgriRJ1HG8Lg/iAMAiNA0YAzQC5DCAAAACJRRxmOQgPMoUmWrAmUcG3GYFPuy9AIQJbFMlbBGoDglUKM8BbhclhBIkITSCLy2tkZjsIEXQLUzG0OUUgM2QGjUQ3CFBnAriClwgE6ZYKXQMDW/EAK4WwG+AgGyCYIFATqNsdYGgEMDUh2geRLwQAx4j+BjEILhAKEby4bFgSbjAIDIH1AeUb/8MdAtiwSASD4LYAQAQERET33xv/g+cQ/kdHgQKJRfQDBP6NRfBXDQIg6TROCA0HbSsvIQX/NrTtARyZAykm/yoI8Iw+1twZYQwgiUYk0gCZMACwiIUwWMYO0AgwJJBYhN5YhA4ApRagxvBfQQeoFxBNDFFQ"
    $sData &= "iUX4gClrU4XbfkaLw4hFAcdF9GUDAIldAvyLXfSJRe4AMQA2rJVQUQAAi00YjRxECPCBAWATzTHYBoABMTgM8d/Ez1D3/F+XBRC5segVALZYhH9sBJYcoAGQfhLVUHtnV3tjuJ4lkIWUTVDXhRMHFQlIoukCiNF1i/i42RVmBTkHD4UYBRh3BZhcBgUYBqkGEfUC6hUwiBPxXzfSzGeW6GFcEdTyDlABAae/EAZlNHUgWVBMwsYEzRYlCdoAhdt24tmkX9B8UYFZh7eoYbDYkw/RFHbQFHON3XqIZDbBAnEHi3YYAIl19IP+/3UnkK1xM8C1MYN98AJnVukXiUrpfHEKvQHFB7BEBUkEGQR0CrgO5tVR8g5QG5L+a0gQFc4D1R0ULfgx/C34Cvgt+CL4LRjCWQAtaPyDZmAxEokgRjAdJ2oDZolGFyhYVhSAFShSTyI8EQGJXmCNXkhTZQPgcQVZMEEyU//X/05gCO0AXSb9//1/BUYDmQPrBbgYYcwAyTvCJGYd4e4A1g4RYu0AZhgSwoCbNwKFzj1j1gChugThIO4g34PI/4YdIrB4vH4/ZcUq8SDoitEU2IP7/34AJIPGDIsGV4sKPJiF/+YP8Gz+IFeufg4zpXwBdLoDUA0QAFCFsMhGgqAZAVeLzeCpMHTQAYI/+P9PByRDRWaDOH50JYBVAcHnAlOLHAfAPhEPwAaJLAdb6wkBVY1ODOj5A30QS11GEjGzXQfSE0/hJTEAiV4EiV4IiV4YDIleLKCPAShQiQBeGIleHIleIBqJXiTmEWUhUVgUoIYAgg4l74sAO8N0BguLyOiZUS2eCPAvNdUmgE0DagIgYK4JAmhwhQAs0NKHcMXaIvDfIvDlpQFOEhIVBZ4TT6PDAKIBkPMFYddHWRZgG18Ww1AWlFgEgLA+kNiFsHgEKAAC/IsMiIlN+NcCAdKAAFUbsggEYgACFP9qBAEUeRtAkiEBFGErQCHeIiHLETUy/0X8TQk7AUcQcoOLRxjdAAAO/3cgi87/" & _
            "dw0cUOgk3RS1EslpJuLBAp4S4PpBD4WdvgLRkx2aoeUAgOEZdgFQxt85IOAH8XRchC8wAAARx0YE3gNB5zRljhHA/Uj5FnU4jV4oU4xdP41F+CWNXVtmiVADzGNVtftUR7tVr5dRSTZJGz0LX0UCjSUJ4PIhlx0UHQDxb4NGgSEFAFGupioCRRzHUQA6EvAPCl3CGGIJYIPG319AeAuEomcR61nccEAO9wF6FbH+Cwdl0HAf4EUR66JwgAv3AZVKGLFeCAeccOoAJ4n6FJE++yC7EGiIfB+Ab1ERfGQ1EIEswhYQMIhfqpkBLUkCuAbWA1FLCABWV2ogi/noYWjvehSTqwafq2CvFuCePFCAUexAgQKpF/EuzqKRN3RSMeCMDk8FVtoNC+0qXqoi0QTT0ARPIADYBDTIR0IBQLeAnwaUldEDVAmcBjqYx5AAcByAUlpQl+AkEsDWLaILI/9ACItACLRBpshBpmkznPUC6ThVEukjuHwfsMmHpHxa2lWjyQfrcnDACPcBVZlwUQUHcHAfgAkXVThwAAb3AZdwsQEHV0Rw2gAnlhEp9gVQXQHckiI8Ek0Mg/mcdQAQ9kUYAnUQuJDyDbJKIaIBRRgB62Tu/itQ4GDEoLIRdA8Vg/mbpgsAgRipiCB1BR4uaQA5WEIO8Mwe9CHMcT4fcdDfoMQBx0ZZDFIS8skMCQB1QREIKTlZyMCJBTvDBQVBRgVFamjoH+1aAeOUUGvw1aDVQXxkLAAMgBVAmnVki30ciwB3CIP+AnQJgwT+Aw+FTgERB4sAzsHhBDPSZoMAfAHwCI1O/g8glcJAsW02yBMAgPBQObww/VAINkMKshngP1CXYAaQg4RAN6AmkBWwaB0sTjAgLCEDCOj/cPjaM4FEsC7QN5ifWeckzxBISAiqHkC43ouRP/BAdGgtIBCLwXYaIWxGMYgHXyPQDScBdgIDi1joM/+YUAcgsKi/6ERwNYUZBIBOHvvQHoRZ9+KfBKAsAhoVEKRiGhA5EyzhVESACwho" & _
            "3P0VkMxHddkrWZAeIdYqL1G3lJOLXQKQKhMJULdTIgCcKi/gUCDrSIP5lnVJR3FBdAdSE7LO0xcDigMIONGyiIdwJe2bIH14B64osi5RK4Hunsfg/QKSD9cl+RDyNpOCgKApbyM/EfI2LG4jCN4/wEAAoIZ2bFDWgU4Btx7uk3I/4AQUi8GDSAAQ/zPJiQiJSBQEiUgUwOAJJEYEwqJAafwDEv92DDTyn2KHQDVZ9BBEQPsAn2o/n1CorvZTDQFuPz/gIATCOFEws32V49EML7BoUAQAM8lmOR8PhLqirRbHXQQ7dQ/9BwICAHQDQ+sFLQaJBBBBjQRPYABQJxBemEwXMCQCA4RKjLTQGKTVA2qiNpTI5NEUE5/eFKC2E9cPALcKQmaJCEJAAkBmhcl18V43oZWyOJwfaIwai05qOQCBkIjAATPAM/8EZjkBdDag4cQAMIBsNpizU1cBWX+QmMEDjUQBAgCLTgiJBAqDwkIEPHDUSPDThYBgAFAnvT5gaAQeFbOIEwQhWUSeROE/MASLQHsEMCAQJC0HeENjRDCmNwHTh9D4UEjRSHfkDyhkBCp1OjPAV2pICE4ugI5fhLLYRBFpCANqKpougfv/7yASMgF0ogCgEQBDiV4bAAFyDODPAACLw+sbiDkUD3QSKSQ6dBPg7WPaJ5COAgHQSPNk0GgkgAZehCCUMTMU21lD0RkFIQCJGISyHzDo56/SD1CXUqq2Af4J8QBG/g1hqg0BC4kY64ap4TqpYQkGIDkBIQ2LwukjF0EAFggyY59Th9AtIMDg9/BPc9tyYjSwU4eQxRfvpjEqP6DKQRQchMFBHBwyOENMYZURugUSFWUhIWIUAfFfQWz2IfiLVfyD4hPVWMQPhabhFIREBikJMpoCUMRvMgLrEaBUwpinsoO8AXwPgX2PzV7dRQiNAUX/UFFR3RxmNABEST4FikX/6wiWaIBOkgWQ7GEEqMZvilqYqGNgJmgBrO/AWcv6GaAJQagfQYpK2agAlIAG2V0M2UUM"
    $sData &= "uKxPgJoFb9DmYEMQORF0CCuLwVdeEMAAQAekUIPAUgeh5rIO6A4FOC51BmosXwFmiThCjQRRnSgg2V+NQDPAOEUQwvUetkNzdWzAXozQKODHkHiQeGSxEIkzRjTeNGCNAY1GxitgZUIACItdDINlCACAukjwYFtEMVhGAYoCOCiAMZVoxIOT2OWaAS4BIvkE/IlVCBiNBFPGPmA2mA/BV+dQrpCI0EhENOWIAgjrZEygQIdVkO6Rle7QlICAXk1Wrj8wgh5RUCHoERDrDASDZRBlB4pPgX6kzgR+HYECmChCgRKLcBZBKbITYDMSi0UY/wpFFI1EJirB4EYEQoERATtV/A+Gci4QsKvcQRQKJ1QsI3ISIBYd4m9kbxl1AP5GupU1EAggcBQxM4fGDzHoBwOgagDeTKDyIBGLRjRaBfJfAQEnYjhw4NoD/xXoOPAPdSNtVMNiV7F4B7E+tuEdAvIoA4s+FAwgj5JRToIv8CePcsbr9/4nRS56ArCilAAJG8Yx8V83tKByYusnU1cqPKB/Iub2J1cjaH8yX1vKJxJxYaWczeKkAlQZAkTx39LK+lrEqCA2MvWFD/YfBd8hvmuw3Zv0/0FL/yu4wYvw/7HGC/8N9M+A1eU4QX8yPSCQWMX/3YafAX+AL5oBSpUB1F2EnwGPMDMErMkCOFwkDF0HpQ1wV5EJSTqqLPSR+EWQk/jlQxMiWVPcQERizAJgOKUNwhnwoUaBLnYnLp9xyRzGEQGVwpTt173yrGIlRlvbqCE1wfIH4X0AuQoVaAhTsh4PwO1mEljGSrLAA0DH08yBkKNiVBEAiweJRkDrOv95N2lK4iXlWgH6HrEu4uQAyWJXwCvr1JhOly3AEMGOHyVFBU0MgexMwQ0Cg+IBVnUGth9BJ4ESmJ8QUQei1oQTzCIL4SUQQBgWA6AHBQUvwDZSkkaQg8VTl7ADsMwHsYME0bdAYbYzLQ1CCqIXkk4gCBpQDIsUig0BiwBJCAMKiVXgixTRgeKyHQCoPAZ5BQdK" & _
            "g8r+Qrig2iO4AehaAo1B/5krwosJ8FfR/nYTsmjsNzEgddDWEgMN/v+fuFPErjMxbAM2t2DiVQaqYO0JJvRgAAG2hahggg32AGxgsFhHLpBYxN1o5G/uAcREFAEEiX3M8j7w4EgSlNAvgZ5tCgOAeQAFSIPI/kAPhEAfUNEEnpNTVxaxWIcu5hEHweZCTdDr8TEtdKARI54NISkUBypLgyhAA9iL86DQyGNgrEABBgjDPQDfTsAy4MFGCRA/UYROB4OUNQASje8XEgHrLOiUgJUGi8iLxitF6Pxcot4FRkqg+wN2CJEIsE7Ks5jACnpAkgOg8BOLTWLwzQ1OQNBIhKRZAThFKpjwUMhgHBVNdhQA6lsRARITslgFn5UloKaS9XALLBYBV0ZXXQBmiQoqSBQd0KTA3gAA5EkTBIhggKidAIP4CJZPYDbYB8APkNgE/0DYAy7eBMo4oKAAEXU3agEE6Bzf///uRQCQpVkCQNegvwHyMCEZl7A+8KAWoQZgFRjQBgGk6FEETfyIAemAyAIOAA8LdSoagP7tTQqPEPA0gSFCAg4ip8I8sMB/D7/iVwAV7BAgUCeiJoA+yp6Q4VANHOEAOGoC60UmnCBRl8MJe5wvUPQcGAGhJhEIZ+bzFqhdAomJYROZAAToPKL4L7CUCXGjNpEJMqDYMUCATsH5Aoic8aCjNpAJVwkTCccQCQDhgX7drRIAjTZxFx4uwI8YZpSOUMIpFA+EYattDCRQ8UDYSzJbBOEgm+wQDl4cDuPqBqHhUNnhAEX82RjhUEQF8ICA7gX/AgwdqvBSAH/d8NINbx44dWFhF7YwkaqxrqJGhgB7QgIwsVgCohaAYJVYxy9pIOsHxzBF/HYqYWoAAQDo5mjcfiJRlgClbhABAFZUamoBwO8kFRjRgh91UlmZIAU5AOsljcAQyI1ggTAFvnlAAgDoYId5AF0BKXX8WYsIzg+3ET0EZokUAA5BQWaF0nXvIOmTeTAIdUpmO3DBdfAafCAhJkXUjUVg1P5g" & _
            "YQgCPJjp1CiGPSAaD4TtxQytGyUsdRV1BAAZAXVUdRwIahTrL2xhUieEyQbk22yDBll0FeDlRJ4KkXeSGEWwnmSjVtF9qOYU/C4nor8C60MtIRAJD4WvvQANB2uXMQByJZAQChLRiJUgQS5nJuTQLIIGI5aQD4XvfkWBDR32KR7hAB32CDYAoG8mDARZah6GOJI7cIcAWDywWESfSIOZXrbdPYALaCS+C4qXFELZbP8xqd1sLxDTbLRYBZgSkGWWyKCl9gGLAQ/adqHpATyZi00K9IkEmSlhGCmhRfEp8R7FcrkCrQLp2SNGCBHpaiBF7H4FwYNF53MIsa5+gW5vECIG8XBLYNAE5IopQg1mBZCV66AiO8EYD4T1wkugHAMPhBbEVhFQE/FQKG2OIwiJhBIFgb9Ql2LjQA+3fAdGBlAYlQrgQlBSK4CQPuScmIIfUYeC+wreBXGESHsFdR8Vn5CU9/HS7QVpBJDgCYCQJBAQbtmQ58kGxdIBUAwPUKyRviULaaCQGy7RiAQqCfBnBQjeP2CfNrEZGfICNBApQoKgWKzQalIFIkEnZ0ABBmXgwxHpflAzEWpVEfxhgwR1CNt4UHWPNwU0I00zFWIoQYBP8VAoIGsn/JmGwVEFD4XbPQF2A9Cd1mV2Bj3xCBUuGhAV1SvQ6aQC6E9+ChEXhbZ6aKGaGCSAtqZoOIG20ptogXnSIYSm/WgRJaQ1gLYp1QnUVdE4QoFTRwIIhizijRNWDOtoFFHyAhUeNAkegySYgDZqQD408//fhA4w2If+//D4aw9oRwFF4Il95MdFIej/lQM5OHRmfhxgqSQhGFPhCmtCKqHpAXVA7D5YUMf9X4f9b1BHP4WO2nGgDAS4KyeAuhcwoHsBviQhryLqrEHKjhnwrzYExqRDAKu4EdTWZCJWEfCLXvgHD7dG8FNNPLEASQzIsmNgbCNZWXY0wFFHJCy8F7Jm0VhEDmUbFrT4JixQ5SCpFaqHEAz4XwcNnAzSE7JumNNHTgepMAAcFZIT"
    $sData &= "MdjH4QFEV2jQx/FR5/DAdKvDB0XMbCHwE0W8gAIglOPHQ1d1dVWEZrETjYW0uoYAJTkALnZAjQNXYhfgORBY9E9gR/NfwcADO8c6dTTpEi0BIfEBL4kTUWApHRumHgCeWESODnRYHxAMLrDYBc3Q2MQbVSCjeAYc/wB16FD/deT/FUBQlRBDAYlFtI0iRbReXJLYhyupVhiiwnOBLqlWD4UQpQE7CN+LHRSqCtEH/gDiWCIHE33YM/aDbscWCeClBgbamaGhMU5RAN4pIABIqdoXoXcBvgFQgyIvt8ioEMUaFm9hGwCSGMbrIyDzWk5StyG0F1pLgA8B8EApvKBMB7EI8aCYICECiEwWGBTpCy19KQsehXAwg5XsywGVnpKgxgZmiUQKDhjpEXQjIEASZgqLAOvmuS7zhSsQ4iIdINEWoDlAp04BDAh1NosPbQGJTdQaD7cJhWF5UdTSTFFlgpGYD7AIlI9QAMERWwBF0QXCdwTr3vkryF0QXidxkG1RE9lcBqrNYAVU0U0V3VTo4VK3RgRqSjcQD9ooEHUTpablE4ICMS05IrIQoB+CI63TwgzwX0ddbZCBsVDHrwagvQLTkSaOhrOHwQLREm4pEwz1kaYQDA8g8n9DPel5K5SC0CyPBdlz62JIICUBCvo3MEgCJXUYagP5RglwIunN6GNSR+CzAeDQJKAVASSiA1GFoZbgKaQATBZFBv6QkDwTCUH3oNB4hC+CQ04jIc0ILBKS/twa8HB7hB84QgZWEeDCIEQOYfMSlNKXYfMc6zlVIxyVgDVT+VJjF6YCUKfwT4MYrefUERD/ReBKDmEMMXgcALFTBP3AiC7GAtiaAOAxAHXOOtC1cV/EQRmxQWeSB1GRBxCwyFJEzAKZAo2RjgDFMJEO66LOJumtJQ4KAIDA7FGn1RJzdaVW9Q88UH4lQZ0rxpErWFManStiDAAIuSLULgDQHivgkQKxYgitAK1CLoEQEB91GrFSAf91xIetckEI6w2+AVIfIYQn+/IBklQh" & _
            "AACWeeS2B36+X6CNBytGKCptdJAz5OT4FLRT7RzwBXoxsKVdAyQgAORCA5DGIlAO7Ss0B4vI6UXQ0lswAzzsAwbPVGgtgRVUk0uF6VTifUroYhJiAQACDOh4wQhZf1l2RqCDJtpWwBEModUlYlDwywIIXCGsJZwvoMQVnMO1wypfwQVqSOgyGUF+lfDqHiTI5SBTsiNhfgNNAIxQ5cBIG4xCwSgUDEWD3ih6EawBIboBFOg/4c+uayeQE65rY7FGwkaiuTbxlkaiuTbuRpEXYYX+o1oBwyak4Z0qJI1F3PkF9QZK0A4GQocBZNRAc4BpZwMpMTJDIJcHOV0MAnUQV740VaUGfSDspQDwtR7RWKTGDgwKSsDupzFwHnXAreqPOHXKrQCedcLvuwDspfoHMJUHUkGOOkWqHGCQ0wVRFxIG4eoF61k2JQMEcRJ2PIAy5ckEUsiZMx5cwO7MNRAhU1OGCQCxiGyVNcbqlxQMQTIsnQF8UIGtmw1CXRbTJpQiMSHmAFZUV9EQDh1C5B0S6xXwPg6RBZAhpVM5OF0YdWkNNovQM8HMBesJKALE0aWBv7P/QBjACQloOOoeEKIg0CDDD0CERnHR2IQfhUYk2AfRRB31D90D8XyJBTP2RRBCPRMCrJAC0cVTMvOSJKAbAaoKAAQ0AgVqAl7rGVRGBFgMAX/vvWE/aA7hP2hsQwD4iwiNAFX0Uo1V5FJToCZMwNAnsLczTzfG7FITZIAhBQU+pFCXBODkRFfgRMexDkRj6RRwFA+3SBqx/q0QxQH6DfBfgdTYskv2AkzSDKMiGq5NAVMGIt0k+vEQIBEtoN0UGmdwlg6QwAn4ADldaheQLCF2BPTrSkIVBibNFTYsERb0q31WHGTAQEYSfSYFGyJ19IUmi9hXQk2gytmisDhskygqUVGoMWMvLYcYqhgA0lPjX6cuC+tKQS0AwMwtADnlCnUFEjvGdCXaERDQgoZVQQQTPJBlQEfCYZ+B5QYElQuRIMbud1KDEAjQh9DqDQBQ" & _
            "B0MQ2Mfw7gsCJ1YGavX/FSwGQwEMjY0GBCGCxlRvIKgBKFTwiwEkGCBEBMKz46A2ymNUNIMPQxggTwDBDRQCoRYgggagtwP/FTD4gCSCNOzqBiQECL4moDCGJZ8GmGCFhhKD4IdySAJKU4DGgaQHahES6fc2dQpoDFz5IELIakYBVmgAUHtZCSX0VAWwAgVA5FxZ8UBoBivYUBsFJWLMUGs3F1ZowFCLrAoltFA7CyWoUJtfaeAXaKBQSw4lkKnBVQypQYRcTMFFdFxcVcFFaFxMwEVYXHxVwEVIXDzBRUYzygVjHlykFLMHah9c5PeyFAdqCFwE1AdJB6odBgCkj0UAObN6UhRgpYJGAbXANSQAULtaCiXw6jVKaAkl4FBbFQglzFAbEQqCC7VF/ObDYIVGCrXnUEJFCbXSUAIIlXQYFFZocEB64EYiXsMAagnrCmoD6wYAahLrAmoLWF4gw8Z2SzCFxHQVBACV5ELGxgYwBUAnR/WW5EZnw0BThHclVgBtAUtFAlJORUwzMnxyBFBGByX3NlY2JzCHVBYGBwARRgByZWUAbHN0cghsZW5XgBVERgYlBwkAcMGUJiYHECaXB9BUx0YHkCaUR1dG9XYFkEZWNoQWJgcACVJUb3QGwPQWhkXGTFcwWRRhxFYHMIdmlMZWJlQHYGZWJjcHcCWHkUZXBhTlEFN0ZAhIYW5kMDD01oYAFyZXliOQ5nYGdNUZEsTG9jYGIAJA2VDwxlbWIzQE9BYICDDENJVE5GAk99bGPTyXJHAGk0RECz0SVW51c0FlzJRWNleExwUASW5pdGlhbGlEeukQVW5pODkkB1AWRleW5DZHBxHmNgYUVGFzaxRNZW3hI0nBoABMgcEET0xFQVVUgUEAAbcAAaMAAboAAQG6AAG7AQEAkgABVAABSgAAAQgAAQkAAQIAAAEKAAEZAAEAEwABDwABGgAAAREAARgAARcAAAEMAAEHAAEAfQABhQABBQEAATMAARIBAUcCAAE9AAGI+BAgEGAD"
    $sData &= "EFARYGcMAIDNAP+LeAQL/1AAdDWLUAiLMAMA8Cvyi96LSBAIK8t0IxpYMCA/AOC/ArwivQDNChAsfgCNLme/AAC9IE2HMKCdArCwcy9Hjl6UwwAAgFUgnSnwEAAEYBAAxIA1AKCAAFz31mKKCQCAOEBNclWAxzOAvwIMYLZ4RDGAP3iM8YNOSA11WADoLwA6AAALwHQ+6EChWGCVJkdXFybE1k1AVzZG548GNYh0AwJFJaMDeFeLENj/0DgAAQuGCHCAeISCBUUFFUgF/9NYi/6twlyB0DOAb1UEsojNGqEADCu9A/919iQAQBfOA2JH4GQFsp4QEtDqZGwAMAI11QayKgPMCIQgUGe/fo1eQQRfEIHH7m0MsOmquEIeagOwGpauICMDz4nymFFShVQ9UJJFAYAA/89cgBoBAJFpAiAAYSsBACHEAG6rDODIAK4MYMoFbAHi7cEzIBMAEJUQHwAQRoYAMQQQBgLARgBBAHIAZ6oUIkcEedIK8kUAZAoAZQBmRFDHCHTXRMI1bkzQxgZEwzRWcOzwRgFjTUBsKVfLCnMVoFxTVgzyxgC2bXBw/FAR0UYMrCmmSgVpRQFn3aBONHDnAlBUFFBHDADRK5FVEAIAMUvSK2fGAwVKIgAQUrTQ1gtgZccBTSwi1hACEEG1bQBkJYA8Q1QeQscIW3VlAnStIWwFVSxUABFHVz3CNWUQdABobFBFxgRD9SJhRFDWCdL11AygxgI1EADlxwBdY7UkY40DDTLNFchSFXDaIJXOAm4uHoTQOFIUYVsxI9YXBBB3LEhSRCBsbMYcMMQWNjeH8SSmJlkR9hnwxuYCAJL1Vgdi9lYH4EYGQodWBlIGEDZHVybXUnb2cOYC0uD3DPUBnbHtFG1pVQB9FXRUYIoYd7WFtwBBRgSFE2SNYZYvUCRw9VZFcMakaUTQckFUbwCU0ALLVn3KExXQTRGtFbBo/QBu1SQsEcYidWzdRIzXBUvWBVDXKoDLWWalII0SnXE2ADTXvYA0Y1cbklYJ" & _
            "whps898ZBxD8BRAdc52TDRQk1csYc90nPRgAQZNiJFD3VnfCcl0YbjyS14f25lYJwjIACvwAEDorAQB8RkK5ARA8pbXx/84FAP/PKxIEAeCg4TBAwSMHAsCiQ8fBMGBIBQBECOUBVeBqBjAAOoDDAeUhAaAE6oE0YIknBCAG5wKtxSElA6BiDYRCggKnASLoTSNeJgUaDADETgwD6H4cQAgUZkQ6AAA0gA7hBAREEvgFiJQArL7+BcTiA6gARCxIVlhYVlgA0DxSA1ZyOh4AA3J6fHx6fI4AdKBwA0YFpDwAGHIsiriAws4AShaSelo8SEQAMooyDuaUUjwASFhEMqBEMvYAHEIcSkpQSkQAHAMEHAOOHAUAXkYgrC4cID4AZAM++oZcEHYAfhi6VkBG/lQAKhQSNhQHDiIANC4WFirEGi4ALDAWHg4iSCIALLQ0SgMK3B4ADg5+NCYwKirSCAAgMAAlIkJEKgQqIimKCAA5APvLQVgcIyP4/7gN6BMYAABmiUYI6c4oAABmg/gRdShW/xVgUAAQD7cHUGoAVlb/FYhQABD/dfyLTQjokhsAAIhGCOnOKAAAZoP4BXUfagVqAFZW/xWIUAAQ/3X8i00I6AwdAADdXgjpzigAAGaD+AR1H2oEagBWVv8ViFAAEP91/ItNCOjFHAAA2V4I6c4oAABmg/gMD4TFAQAAUGoAVlb/FYhQABD/dfyLTQjo9RcAAOmyKAAAZoP4DA+EoAEAAFBqAFZW/xWIUAAQhcAPhI0BAAAPtwZWZoP4Aw+FhgAAAItGCIlF/P8VYFAAEA+3B1BqAFZW/xWIUAAQD7cHZoP4EXUIikX86a0mAABmg/gSdAZmg/gCdQlmi0X86X4mAABmg/gTdQuLRfyJRgjpzigAAGaD+AR1CNtF/On3JgAAZoP4BXUI20X86dImAABmg/gVdApmg/gUD4UCAQAAi0X8memyKAAAZoP4BQ+F2wAAAN1GCN1d+P8VYFAAEA+3B1BqAFZW/xWIUAAQD7cH" & _
            "AAAAALMp/UwAAAAAcJAAAAEAAAASAAAAEgAAACiQAACBkAAA+ZEAAJc/AABvPwAArz8AAGBAAAAxQQAAQD8AACo/AAAUPwAAHEMAAMc/AADrPwAAv0MAAMxDAABYPwAAf0MAAIc/AADZQwAAF0AAAEF1dG9JdE9iamVjdC5kbGwAyZAAANGQAADbkAAA55AAAACRAAAbkQAALZEAAECRAABYkQAAbJEAAICRAACWkQAApZEAALWRAADAkQAA0JEAAN2RAADokQAAQWRkRW51bQBBZGRNZXRob2QAQWRkUHJvcGVydHkAQXV0b0l0T2JqZWN0Q3JlYXRlT2JqZWN0AEF1dG9JdE9iamVjdENyZWF0ZU9iamVjdEV4AENsb25lQXV0b0l0T2JqZWN0AENyZWF0ZUF1dG9JdE9iamVjdABDcmVhdGVBdXRvSXRPYmplY3RDbGFzcwBDcmVhdGVEbGxDYWxsT2JqZWN0AENyZWF0ZVdyYXBwZXJPYmplY3QAQ3JlYXRlV3JhcHBlck9iamVjdEV4AElVbmtub3duQWRkUmVmAElVbmtub3duUmVsZWFzZQBJbml0aWFsaXplAE1lbW9yeUNhbGxFbnRyeQBSZW1vdmVNZW1iZXIAUmV0dXJuVGhpcwBXcmFwcGVyQWRkTWV0aG9kAAAAAQACAAMABAAFAAYABwAIAAkACgALAAwADQAOAA8AEAARAAAAAISSAAAAAAAAAAAAANiSAACEkgAAkJIAAAAAAAAAAAAA5ZIAAJCSAACYkgAAAAAAAAAAAAABkwAAmJIAAKCSAAAAAAAAAAAAABqTAACgkgAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0kgAAx5IAAAAAAADxkgAAAAAAAAuTAAAAAAAAtwAAgAAAAAAAAAAAAAAAAAAAAAAAAEdldE1vZHVsZUhhbmRsZUEAAABHZXRQcm9jQWRkcmVzcwBLRVJORUwzMi5ETEwAU0hMV0FQSS5kbGwAAABTdHJUb0ludDY0RXhX"
    $sData &= "AG9sZTMyLmRsbAAAAENvSW5pdGlhbGl6ZQBPTEVBVVQzMi5kbGwAYOgAAAAAWAWfAgAAizAD8CvAi/5mrcHgDIvIUK0ryAPxi8hXUUmKRDkGiAQxdfaL1ovP6FwAAABeWivAiQQytBAr0CvJO8pzJovZrEEk/jzodfJDg8EErQvAeAY7wnPl6wYDw3jfA8Irw4lG/OvW6AAAAABfgceM////sOmquJsCAACr6AAAAABYBRwCAADpDAIAAFWL7IPsFIoCVjP2Rjl1CIlN8IgBiXX4xkX/AA+G4wEAAFNXgH3/AIoMMnQMikQyAcDpBMDgBArIRoNl9ACITf4PtkX/i30IK/g79w+DoAEAAITJD4kXAQAAgH3/AIscMnQDwesEgeP//w8ARoF9+IEIAACL+3Mg0e/2wwF0FIHn/wcAAAPwgceBAAAAgHX/AetLg+d/60WD4wPB7wKD6wB0N0t0J0t0FUt1MoHn//8DAI10MAGBx0FEAADrz4Hn/z8AAIHHQQQAAEbrEYHn/wMAAAPwg8dB67OD5z9HgH3/AHQJD7ccMsHrBOsMM9tmixwygeP/DwAAD7ZF/4B1/wED8IvDg+APg/gPdAWNWAPrOEaB+/8PAAB0CMHrBIPDEusngH3/AHQNiwQywegEJf//AADrBA+3BDJGjZgRAQAARoH7EAEBAHRfi0X4K8eF23RCi33wA8eJXeyLXfiKCP9F+ED/TeyIDB9174pN/uskgH3/AA+2HDJ0DQ+2RDIBwesEweAEC9iLffiLRfD/RfiIHDhG/0X00OGDffQIiE3+D4ya/v//60kzwDhF/3QTikQy/MZF/wAl/AAAAMHgBUbrDGaLRDL7JcAPAADR4IPhfwPIjUQJCIXAdBaLDDKLXfiLffCDRfgEg8YESIkMH3XqD7ZF/4tNCCvIO/EPgiH+//9fW4tF+F7JwgQA6Tuz//80ev//NwEAAAAQAAAAgAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
            "AAAAAAAAAAAAAAAAAAABABAAAAAYAACAAAAAAAAAAAAAAAAAAAABAAEAAAAwAACAAAAAAAAAAAAAAAAAAAABAAkEAABIAAAAWKAAABgDAAAAAAAAAAAAABgDNAAAAFYAUwBfAFYARQBSAFMASQBPAE4AXwBJAE4ARgBPAAAAAAC9BO/+AAABAAIAAQAAAAIAAgABAAAAAgAAAAAAAAAAAAQAAAACAAAAAAAAAAAAAAAAAAAAdgIAAAEAUwB0AHIAaQBuAGcARgBpAGwAZQBJAG4AZgBvAAAAUgIAAAEAMAA0ADAAOQAwADQAQgAwAAAAMAAIAAEARgBpAGwAZQBWAGUAcgBzAGkAbwBuAAAAAAAxAC4AMgAuADIALgAwAAAANAAIAAEAUAByAG8AZAB1AGMAdABWAGUAcgBzAGkAbwBuAAAAMQAuADIALgAyAC4AMAAAAHoAKQABAEYAaQBsAGUARABlAHMAYwByAGkAcAB0AGkAbwBuAAAAAABQAHIAbwB2AGkAZABlAHMAIABvAGIAagBlAGMAdAAgAGYAdQBuAGMAdABpAG8AbgBhAGwAaQB0AHkAIABmAG8AcgAgAEEAdQB0AG8ASQB0AAAAAAA6AA0AAQBQAHIAbwBkAHUAYwB0AE4AYQBtAGUAAAAAAEEAdQB0AG8ASQB0AE8AYgBqAGUAYwB0AAAAAABYABoAAQBMAGUAZwBhAGwAQwBvAHAAeQByAGkAZwBoAHQAAAAoAEMAKQAgAFQAaABlACAAQQB1AHQAbwBJAHQATwBiAGoAZQBjAHQALQBUAGUAYQBtAAAASgARAAEATwByAGkAZwBpAG4AYQBsAEYAaQBsAGUAbgBhAG0AZQAAAEEAdQB0AG8ASQB0AE8AYgBqAGUAYwB0AC4AZABsAGwAAAAAAHoAIwABAFQAaABlACAAQQB1AHQAbwBJAHQATwBiAGoAZQBjAHQALQBUAGUAYQBtAAAAAABtAG8AbgBvAGMAZQByAGUAcwAsACAAdAByAGEA" & _
            "bgBjAGUAeAB4ACwAIABLAGkAcAAsACAAUAByAG8AZwBBAG4AZAB5AAAAAABEAAAAAQBWAGEAcgBGAGkAbABlAEkAbgBmAG8AAAAAACQABAAAAFQAcgBhAG4AcwBsAGEAdABpAG8AbgAAAAAACQSwBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=="
    Return __Au3Obj_Mem_Base64Decode($sData)
EndFunc   ;==>__Au3Obj_Mem_BinDll

Func __Au3Obj_Mem_BinDll_X64()
    Local $sData = "TVpAAAEAAAACAAAA//8AALgAAAAAAAAACgAAAAAAAAAOH7oOALQJzSG4AUzNIVdpbjY0IC5ETEwuDQokQAAAAFBFAABkhgMAxyn9TAAAAAAAAAAA8AAiIgsCCQAASgAAABwAAAAAAABXswAAABAAAAAAAIABAAAAABAAAAACAAAFAAIAAAAAAAUAAgAAAAAAANAAAAACAAAAAAAAAgAAAQAAEAAAAAAAABAAAAAAAAAAABAAAAAAAAAQAAAAAAAAAAAAABAAAAAAsAAAJAIAACSyAAA4AQAAAMAAAHADAAAAgAAAcAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACIsgAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC5NUFJFU1MxAKAAAAAQAAAAMAAAAAIAAAAAAAAAAAAAAAAAAOAAAOAuTVBSRVNTMtIGAAAAsAAAAAgAAAAyAAAAAAAAAAAAAAAAAADgAADgLnJzcmMAAABwAwAAAMAAAAAEAAAAOgAAAAAAAAAAAAAAAAAAQAAAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdjIuMTcKADsvAABAAFNIg+wgSIvZAP8V/E8AAEyLAMNIi8i6CAAAAABIg8QgW0j/ASVCUAAAzMxOP8AyI41UvSZgVIMhBACDYSAAVoA0WBCM8F9BCwhIXgBKFcOOhrSYgFRIkExXgE5iHY1LUAiUpeItSIsBwwBIjUEIw8zMzACLQSDDiVEgwxCLQSQOQDKMlAjARYKAlEhHAiF0hRlmg3kIwCBfgLSYT6dAPryGEISbCLLobJZ4JAJAxAkgXCQwSIt0AiQ4SIlHEN7xdbkCEIyADsV/EK7sCYBJADP2SIvaSDtBzlNQi8v/FQyYEPAPjDSGjCtAEgL34UjHwf8AgBTxABTc" & _
            "EoB+6m8BAIkHD7cLSIPDEAJmibcAwAJmOxzOde0FgPxiDg15EF4AAC9RjQUMUoXIkY2UGFADAFMJk4C0uAQREvAPlQYAMwYxIxMT8B+EsIg4FIG0SIwHSEyLABJFM9tMi8lIAI1I2MZA4MDGAEDnRkSJWNhmoAjAjRDeRIhY4aoGIG4A4wZAbgDlgAZgfgyETkAgAAVgFOwI4W4MBF8FDAPxBiBvAPMGAEVvAPUGYG8MdA9gxLQTUffAtAgghcC0E4VQVwAQtDi8XrABPCiA/R+0M0wXDKAAgdRIFA2wg1D3EIC0qISAtIOk5AAKTYkYuAJAAACA6w5NiQhJiwABSYvJ/1AIM1jAvQBIHzGQEIXBBJBIRALCtBgEAZC0mM3UmAQzA5RcQPAFADCTvAMQnDjwQBm8GNwrBjAhAWZDFW6AUhGgqAWNSwGtEiBms9Mw/EA5XEHi5SBe5STg9Qe4AVsA7wFXMoMQQQj/UQJ1EeiEYQEvIz8jM8DrAykCyEVSp9ODYVUEiVEQwNFSywKLAkiL8UgAi8pJi/lJi9iAAwFIi0QkUEiNAE4wSIlGKEiJB14YSIl+HwCNNX0jY8aDk8bR0MN1OPAwAZKIpQ1oEEiJcIBsgAcSRBVUFRRktUmWBFI00/4YEJG0iB5WkLTTvEqG5LxKE8Q7SvEP3KTQSEfSH0yQRAr3COFJD0DkgIz+DLT/SnEE8nDrkAQyQP9KkDRAbBYEsNNc5460+IRCy5nNSYWGu2g4KA8Dt00ASQPsiJhToPMEg5jTlYhrhDgwiMMFSQPcEbBcJEAESItsJEhHEVBICIt8JFhREUFeQR5dQVznQv9HzFJWkViwQZS4hWAADG/SImwklQF0JBghN40sBURpAjPtEnMCOXkUGHYqPRdDfwMsBghIhe10ENCMzgY8EKMr/8dIg8YIBDt7GHLYxwgQ6CjsBNkoQK9GjUtwqBKGdW0GEpO4uISyq6EwGuoNISNsGwPdAF1AJSPpOPIHHDET9AgRAEX0cEtIAooE" & _
            "IEG5RznxRY1p+QBFD7fQRY1x+gFmRSPVD4XsMCBQRGj8UDjuFPYAwAR1CUWEwQ8ghJo/A4H6fPz/CP8PhYR4wkmCCugARDlzEA+FWsCe8BGAMPM/iNMPgLR4vHgP9EAJYLxz/1AJzNQI+CIxoNTIAP5wCxAwiD9Bt9H4BjCwU0xHMYivAUAXMohfQccxGIBPQXeRHnQwAyFUcSMCjLQYvTAQBckaCzuacYzUSADwh9RIEay0yIQQjpDOhgWB+nqAHwB1HzP/OXkMGA+E1SRwl4CFtFhEWQOCtCjBRHltIGjXXQZUpEo1cOvTlacwdW2fcGufAIo+MbK4EAYk0jsgA2UAWBllECZlEG8AFHQcALgFAAKA6a0OkG/wBBNITwlIi4wkRLA08F/BnCwwA6yoJYH6e+ghtwHEEagtSUhIjVQk8wpCBHqwc1y39HEBA0DU+ISA1NjAt/4/EKQQysa0L5DYfUAUAY1PCWaJSBgWkh5qUHDwTGAQeuKE9IkgCBaNVkDpZGVxCQpIjVFYIIQHEheTFUCRBlgRiTgGwENXdzxxw0NyPDEVMApR/2QyR1EEcQOwwA04P0AEiUMI6XJ8MKgPwV839laQkwfyAEJYg/cEXw2NT0iE/LCBjnZf57BzDEP3UQKgVwBOKEyLAkYgSIlMJD0C1qD5HlwdBUiL+EiJQ3ttAgdIi8+RAp0gADP/O9cPjNQFAAAAi0EYQQPFBDvQD43GNIC0CGAEsajMtEiCbGJGRAiFtRUKdSqs4SBQp8kKSYvM6AUXiwEBCTl+sQKFUFS8QrBS/EAYGoAQtNNQ/FC4lgYSsPlIAqKsgckBuQwgeAEvYJaD8FA45XggSLbhBfSaBEEfQEgiLgGLAcdEi+eDOf15KAFBD5TERDvnxStALMVL6GY7KXQPHESLzUHL5IEKlMAISo1U6XkHBEBMKo0EBhHglAnAngmQ3hDIAABQBVKj8hwXb1BnW1ISFxc2MCoIMAhlMYJREUGLxPfYPQFIABvJSSPOSI0MUUkM"
    $sData &= "gFwWbzAYjIH2BHgTxQHdAX0CvRGLTMJoCGEgdQYwZQBI99kCSQPNQffccJAUCBdIG8BI99B0oECTNGJcKCQcuJLu8pyQ2eVXYVAG0CbBL1yckXJBFODGABSUnXoY4gQkwwBVCa4E42RNQBRMi7TtJUyyEYAThLDoBpHbOwAQtAhQ3VgkkMhHwkNtDgHMQM4sVPoBgcwksIj+XwFdFNYSBbDz7kfU2PS/CwSTCDEWjdTII4WAFHw+UNUBhDRAHKag3BBNix5AotQIQ/EBJf0eSIPvGMFqF8AA7QF1z0m6C4GY17HoCZlWxZToFoBklugEdGyIFI2lBwDpScUPHfkE4/0C1tlZ9QV13TnNVSM6Akm6GnBvgFOXD1CCJAv1JyX2BQBIg8IIgzr/dWIuNe4JBGaJKyEZTCmNQ1UA3IUeAWuRPRpMi4S01JlBjQYrAfj33xvAJQkWBJXewpOQURs/oJibXbdXGs9gm5OfsCBbAZkGMI4ggAE0BIP4/3UkcgEPkEMHQAID9FB5HNRoQNUDkF5VeZCQ8wZIAjl7EHQKuA7G7VHFUUiL0OYDogPgNtVhBHG5XmUIy335BLz1JYRZKUWLbxChGUCNF1H1QYE5RMkBKhAQmCXt8ACBCfDmB4ZFVJigdBWUmMqSmFCemFCmkN8gNqCFUTiQtHhBFEgrA9WNWe2NKUmNGb0OHV5ASZYX0ZlRmGRXmOC00JyT6EsRCSVdAmY8ia7mB9ExpFwArRUEYOEwYbEwHgkAt4ifiC7itUCJDfnS/04MXRf3EdAlA3UE6wW4GAHA1MhFApW0uKUAk7S41hwwVxUwjhP0pWcxPiTiXfEAM9tIIIvqlhKQk4Vhl0JqXxFG3hPAcFBPs1gNaoQQQEkOhcB0I/+Arh5wjLDjhSGXnTOITAp2FA+g8wHr534XX8CwmI3OnHAQgbAIT+BHg7Q41QnAI1+BVPgvdhHPYhf0LHYxAksYO3MYc3YjMAUBsTgk3QAMwkiJewwOKlAMH9AKT6DjAcpqKVAh" & _
            "VPbcixAPsgifD0CHD0DLZGI2iONHp4KCtOirAotyG0ADj7RSaF8T4VwT5KyCQXwAkMiBvs6A1Pj0SAfT6CS9LDHxAW4bVUYF7LQxORGJAYlZCIkAWQxIiVkQiVkNGIlZHCTg3QBZEiJglaUZAVnGIRAcaVCM0rj35KsBJJRrQDKP1RcwI9C4hIIO9u7xEMNFdJ4KgIyOlBmCrQwD2mUCIgKh684iMMAahFJRxJkRsoyOIB4g06EUPqAPEWIeHeDLQl4bDyBWUoM5YQwhAD0y+iwRBoEEWBgMwBGzYCQTg2EorxABAxE4zSJBAiIk0CRfEehOAkQ5bhgPhpYPQVA0Q56LErmh5w5dKdCO0FAph74uMHNqMwJGyirAQdAkoWwQSiBwn2hQLkEUaTzN1RbKCoEU/64aAxlFIwFHkFQ106pQNQMQ9F+cNEiMQJS0kxMgSKceE4tW2iDgIU0H4UsBzhQgjgKkgOzvSSFCJFxCUsvkWRPOKBFDgIbpiwJwIJIqY7Yi+Q+FRL0aONLz2cQBD4RAmxSwGIfxH8QAMAjGADDoHnBsAoTQYHEcFBehFwDuMnRkcgETm8GA0p9Wt0cURFD3pDIgGgMBozEA6gL0kHiEhGsUEKY2UBoAT6CBM9fuADjYjjBozPFfB1pXowowWVD/T18MwRIUQhZS2GPiUv0C/QpcbBETdBEDEUsNe0mLpcEA+W4dwDQlLxDYLgBUVCfSIVjHAN4b8K+MEOMZGTk+5gvQ2EBDC6jVnm6AKgohsV6ewBccfB9gafsR67N8Q8D3AU4esf41yTfsFQp8FRjAJ/4iEZEuNiHfMI0NzIgf8G3xEYhIi/gBJh+BaLED1gChqSgRWImWI+BHAg/BArheBj4LQQYJjWdeMZw+waSyGN6wEqQuE/CSBYC+0rYeO7Ad7a3bd9LtlN9kEe3UoxHsZYkAy+jWDYM3rZAA7gHt8QLtJUFc6en9Gk3xEDo1ENQUk60e2P6FHdiwLRBgTS2SBpAc0BZCH9TRE1MZnhYv" & _
            "wZwRgDLIkkfhEhBklQYuAVCnIO51AikJCj9DjDLs7gMQ4d454CNTDRmNDSxUAamivVWcPQXpqRZJVhzWlIDYm4hLgIhtWJqaiLDu14bUjh+QyRde1WFIfB+AyRc+fMPL9wF3l3yhugF8wcmHhgvAJ32WGSwSD1CGoVkzbjzsW3kDkE/U6BUwqM+JUBdBRKj6AHURuIJeGJMBAABEaCAA6wDtM9s5WRR0DxeD+pviFQABV8jRF9AZC5CDZWRxYNBbAaamcWy72iDg0BIbqgBQDzCEgJF4R+HjEgkgm3VHmglQdb2VAOGVuRYGgU4uAAJSClKDAgORB4G8Du8sAAIXYOgCoIGQ+EUBsN46qK/5UMhILSsQaRxFsQZF0kMRNEiAP/BQ+BUz0LQIELE4vbgc1Agk5O81EmZBgzwuJNDIAPT/UCk4AkD1UGkRvBBdI4BQIxUAUQeQYBaUo4UQ9DBVZDYBT9DhX6R+AdAAENQY9E+0uPxEskikTgKNFHIUQCWdgIAORqEPAkIXgRIQmXVbgSKLUBA7CNMPhNamSjCoPwDxcNgMAkyLEI0AQv9MjRxAZkMAgzzaCHWJRIsBwzvWdgyNQokhcEVVID0DxLC0SKUtgJD0QKTN1LZAr9gHh1k3E0ADZ/ABeK1RIAAlAjiSIREtAlQAQlI0Y6kEFUCdCekgQhjccflQ6Bg7QoUaOyEUOylTAdC0qKGyOJ06OZw6PZw6MEywMP1QGB0N0digxtMG0DryQAc7+QC51Bkgui2RVDwEXTwj9G+aOzCcOzWNoL0gOUKWJHVKOXKp/Z4MgWA2NSghEmqOT7EYJ2UT5hEjZF2ssjNcd6CDI+tvFmkVWEET9ZAD4q8hBlD+6pN0ED5JBOmQAz5JKyslXkKnbyNUdgAhtxHuFSNdARSNOedh5eIU2JiDjZT4VRMwHpphhCNSFDmYBJLWYDwBpSY5kRgLkBgkrwDBrlYZ3LDgHWSDOwB0Em4w4IdhIWQg0+RjJCXeS26fNAi/MlUvohEj4XAF"
    $sData &= "NlUv4ZoQVgThnPAE0qIXIlYR+k0COSkREkmQEcKL9boUUW2W8wLxQKgOJsdMjUdEAhYSsFPXEM2AQgdA8G+8PmCWiDKA9C/8H5xUEhCQIHBlloNSt22uBAIC/8Y79Q+EGgrRSBAUkHRcHCKROCMdJ2iwBMVVEkTkVRJDDhgPtw/GIVAlp8EUYyvOv15GcwyME8ARTcBFpJMAicZJAAzeIMC0uNTcDkBHZHkBTLhABABmgzwIO3UbZiKJLICSNKashgJIgI5RMASRlEgBwAQ0QClJ/8FB/8IAS40MCWY5LAENdcbrArou4aXwBV1BBMymVuI9IVFIffQGrlGo3ASyGC4ZJdr/FWkzkCUNvSqKFwI/828WQJTDNudf5yM0FACL02aJfHP+6GBs5QsKEQFmlhiAC/P/71cgkGCWE0AH0XNcpD8jQccEJJDyTACACxPrKYv3gnkDZjk7dB0RCEEgvjpOBpITQ/eCBPEP/C+KAwxDZjkYOXXs4hUjBPUJjUYBuiIBQBT7EFmADUsVPsc1dQjBEF4aY6/JRNPpQVEu0fEPiNFB6e0cWhphwjMH+wM0c6ygFROLDMfpHR3OTg9gWEJj8gBJi+lIO/NJixb4fhFuG/CdnKOqAjsd3nzvuhplHFWREFAkhodFYgYLoAr1BU4AuajPxNQoqwNBuFygGgDPMwnS/xUcNjEwsy0mABkKZjktBJoEsO6ABhj/FRRgYMgFl8P3ZNs6aCG4cYPscjFARQIAZfaAEvyQQhdCAgIC8f8VJLwAafQABop+W7BeII9k8MBi/IBCB9SAM92Ewv0FLNxzYCbGFeAf4Mr9BjTcX8LTGYT9Bdx/QETIPQjzDxDkQSD/8qDVCz/B1AvfILsAbmQQjwCiQCfHAMJFjVksAEWNUS5FO8F0RAxUgFEHEYMAsU6kwAIQdQQsgIGE8h+sFQBKQIBQh00171PwA9leEiKvGQA7QwAMdlxEjUQJAsTmBGTOJkn34AYFhfkN8v8v3BCL+EQ5SwgIdhtNxmWw" & _
            "uBDkqC8AixQIUgChMQAqKbA0giAX7LCAHokLngxBkrjTBJE4ZFYCA0gHiTTI/0PeCCVcgF0RgvUPTCQIVVaSWgniqHoAjQVcXR3xCm4J0BgQBJEYh5AYpxcDAjCJcAiJcCBAJECCtmhMpKM6BXEOSE2L6CJMIOQC5iqgHSygAMot82BLSALopxM4RVAo4eIStkWAlNgF6PkB9kSL/olFVJDiY9RoJ6AG3UmL7QAzyUSNYAFmgwM7fHUMi8aGO+BCBghFAOtk5gEwuKNAQlcgDQB1VYogYpbIsIPO5+0qNaoMAqRh7SU1ni1DSQNJHiDhlQIb0ZIjIQcCzwIvxESCJgsCgJY1smgc9H+c1Agd+GC2BsaKCpE0yB5Q8VCoZ+UiROE9cMoB0qm5U15dTjhPEDzjJ/YOWJUhoX5iv4DiujGDe0AASnVQbCQW6iGzuITk8RAGUQ++eAKFRw4/YFMkQmKf4FT2BYtgWaoQIOcT/8tIY/sj6yFyPoSPzlgZIuezAMtI/89Ig///dX9yPs/g5vMFvjwiHwa0LEUTJ6gisj5S59AuoVoRPRUz6yfeCdE8wYOu7mPNM96mBKXPEDSQEBAf4DwxjkminYEO/uEqEz4EMCNtRFCiwhJApTFwQOsULjOB9Gi1I/8VhGRABdYM4idXoy3wDFyKAgi0XhEDjRXrNVCeGNUDv0DWA9/DpiZwrfIB1jWALfARdC3wOaz5CaoFBgSyvhEq/vAOPI9U86ASBdYEH0HaBCTfBJ/AL2cATXBOBSxTJMYCI+QRi8OWOkVNpIMEUJYCYUMIiQBBQEEPtsGJQVNUoh5AoksU8dnWHqgSMM9RDjpH2PjU5aCRhxoF4fBRmk57TR/itPcN7juXrTAClKjqMYIv9dLcQjphNri3QAVAl9TEg9DPwwhACepmkTiEtW6ExG8eBf5CUdPgkXMaJGD6EusqJYvI5MOAjC7r5Q2cuB9AF+3Yo34ZrYZwgexAAJI/wLSY/pACJ4UsQhdQOBCUq9cRDxC3yEX+" & _
            "K2AWNIIcKCsIhMF0LoH6BRItdSZGWAHmBiQWMISr+RYDZokDSYtFhktguQYgoj6SExGU00VVB0CxM81nELRTJYHBp2BsIumkbGDGsrPcBMOODC5qgLRYRNAlwEiC5aoWaY0BAEQkeAMoi8WJQKy+CZIZNAI8I4yxI1znIUMJEEUBRQEMjUX/ScfGkg0BkLkiHI3PNAYOCA5wSYvBTImkiEp3knRPLslX4df/A/9BjX4JTKZsEY3xsnhMh8N0INQFZCF1FLxkmEoGhQaCCcAtWEEXjUYZeKgIHqIPI5sFdLj2LRAN4fpQbHQwIw3AtAjeWORvlohBRcLWVmDGEYvqiYCpNESNUgRIiYxCJOpX8eA41tAqMDgCHtAq8UD4QgPlFxN4M9ImEsCDJAoGfRPAYmVQf4DQKIRgljNFoJUEWpjClC92BAoZgUZkzBCktIgAfPgIKOsxLREM1gzAtIjdzQZJFCv9MN0jQoHCU4KAbqEB9g6RADHmGwkcQEkajQTe7BAuARa7k4TRSMPtRyjlW3gzwDA5A1EANRB0Cf/IhkUg6wL/yJ0iIXGNZBTVAeJcQAGtuDVB/yDPuBYjsnhtJpRDQPD3UGhNCjPAAGZCiQR/jUgIx44fYggAD7fYnMIlBg8CL+wHvgnRFaAtNjPSuQBmQL0ZyQF8bUkEZQMLgPkETCRIRIv6ZgCJBGlBD7cE/KBGhABs7gbHZjvaDyGEVOJMENQIlQADAHVIjUrw6MLVgAEX2DPARDv4dFAU5jLAL+EIzeiMIwEAAIgD6W0vVRt4ETJ7ZSEBhojRDlCkyEXAj0CEuEEouguAzmWiXUfUqJSPLgVXBClRBAbkITICsXIrPRC40QP98BsEkE5DzGJ0RTUQKaoNgC6SWgP5UAOGYhQA6zPBPQTrOUG5EswRtCOU7ckXyujt1NBY4Qtt2FECtQ1VQshREX6LXxFg9RNdMuJKUBLC0qikEY+OGQVZEodXEkFjEpTLGumw/REGaqMdUBfTqBSALkadv310xNRQ"
    $sData &= "EgYc6KM1FegxUNKoxIB+gi4xCH/Y1Vg2SVGB0xcjzAO4YkAGCMjo+NOlcA+EypEycMLEr94GhCTI+gWR7WAC/gjBJCEAPJFeUwBRLSAgBVC2BaDHBZDABcDMPFEAEqSr0xLpBNp1KUpJ6QRyDYB0Fr1iARAlAADzDxHFUt7M0U69WIRRTpW00BUys9oV4dEET0DWBCDfBIccOukh8FGKAC9/6Rp1AAdJi3T8COs1/25D0AsTDhDApUYAPQB4EL8ALe4B0jZgreUg6wdIjVQ17j/yq10XSc4WkCn96UEAg8m2AuAvAC0BSg4ghhfdECGjCUQebiA4BsqYwtJXJl8CEBJhcDQxIQlJmRIzX/Y58QP2DnOQE9AxIRQDOSFQ0zlhHTkRuQAAAhwA6A4pAdpRZPoG2EgIK8sPtzI5QLCA5DGgbdYA4JEJjjLS/NCEkcjl70sCtjKRSOA+wJTI8Z6eWBbQEpCBXZeWN9VDsE6TN1+wW78T5Esj0RQ5URU5Ea0MtoqHw2R5I24GgG1hJhyTFJBI5I+Q3hqj7xU4+xrGh+BBByRQUpdTGWWRFB1l+BGKBh2KTMgT61lBuaka6z2R+GBS5/0GyOgK+E9gBo7PPLGu4eOIwXhNxKEU0pewPuPr8iQ73jMPhYkRdRCs0HUgqj4wQcNfZxQIRPwOEKDpAO2VAHJOU0iXVEiqzwbvMSM06cbdAkoTQCxRVUUzhr6jcPxQyFs44lsAQRKNRwiBAAT8hVk2GUSLyI3BxFCHQSfgBWWJ6U0HAQDo+s/UIg0QkYADETChSTPSBTPJx0QklikQIG80+A6l0JcTXBMhoekiZkGJLQRvJSK48Ss44G+NQtdhgk1StOkm11kiTSIyd/PfJBADFCXURxAln0AHUYKh+gNtJVECofoCGaEaGRDwHfICvc0quO0sHfIACSsdQqQd0phKgdLI1VZQA1OUQzDEHbjCjbCeE7gPUBCR27fB3bdvkAHTGGJWVBYwCBHeSYlU3qqWEhUgrhJkZQGRA4kE" & _
            "7z4Ek65yXpMWQqQmBAfeQLLOK5oFhLoTohcEJkIgCGEPt/ZKG+D8EJkgjg+AHDFiB0Q7DMgPhHxFAMJEsfPM8UCYZuYihQLxSNwPIIUt6QFJi3zcCH64Gg6CDPJV0VpgqBFGC8CA99Zf3BUg5FCpAETcCC3pvrUBuJYPQS8y9C8hKQgRQYjwFbQrbCBFhnobkFK0uIz8BL4GEhTzEAFt6ghTtKMMHXbK0E8gjiDQiKvIIF0jj//6lwG9CgBhdxO+JSF1ARYAhRTZVQgJ8ABGCfAAHQiEL1hg92vlEFRQScBTSWBdKWPklQRl8whMwPGWQYVd56EQhIgLxTB6EgIVC/HFBFQmOiFUQKeOaw8RDUSNUPHkxCjEjABYN6BdN2H24Hbs8bAF3DEFFnZZkAREcUgX8w/mwG1UeCL2IJhtEA+4lhCCAg+FXfapF21gCTAROCNkwA8QVXSB8hkR3kLg3FgBqXYRNbIynZJshswWdGzhXSvFdgCtDGwjWiFsYaHLVjhsEJeXeqZNf3Cx9KFhFQXg8F3ngFVQqAEexf1idTUzyQRmDy818h0Bdh8G8g9cNejdGkQA7sgBcw1I5i4BEPYugCwoXSRIA8HrENFjdX8sUOKyAA4AIE8AbfMGhTHSt1NZ9EODXvSzbocwTO5bkd/7oA4CTiVgPjLyTvBfLFkSjx4hhPSf7A8CWiSkTAKyJNTw8Pgp71MSljmBZ1MCImhEEACWU2XjB5Lhzk8kORJ+UoSRzm7DAY5CMQCzOG2DEOKl4G8CTUCLeRNMiWQkIOhAyJpREPTfxIB7YgdolULhIdZlgOYzISZGgbKwjrxLIOwou5QBKPghIJRWBFUIkQWW3tQDL0sVsiWjKsDaJIAM/3CrxuRfB74HgCecB2bKJWB2XClkLFD4dVYGEYEUdjXBByZQE8UDznVgbAQCC4BO921TBdrAUScAA64jUCqivRXRESjpNPkAOV5FQIcQVPJbYUcC6wmNALyBMN5qC+ILEQt2CNMKoc4EAQJs2LDQ" & _
            "G8Joaim+HTMDHASQU0RF19e0WDSeY+B6UTo1QkgCjwmWDSGu2GBiBeHXcEmLTQKwQFmOKX4qUAMhjaRTAULJIb4JBHDZIc0hHQT5AXGlu2IvYRUpFEmLXdZC0BzojiiLwEjB4AMxCD9wRPoIUhsUAaFwAylTTSD6iQkukeELAnUCTQGcBRHbGI1HAYYDIVUYjUsMAY1TAYmcJJww4F9QwBSSpCuwCEYXIb44LpoyPPFQWFpikcMpGSO6G4GA9+D4Ui7QOBdQ0LiGkUhboTNm3gV2PAkA/jtQcIBvp3mmDkBFiAy3HFC45iphVpiA/UCo3fVgNqOkCvm00MQBDKIrI/dJM9LiKgLEhwmNQhGLymYNO9iLwr7IoIIBEBI8ELKALHgSRC7wik4NAgCEyISCkO65Q1xrIKUy2HUbeLKoLIMDYVEHjGKWCBmo8JDhmjLKBAApD7cI627dEgJjPgDsWgFDB4Uwi+OFMPAnejLYdOFGCAMK5sIxQAhuE+AC3yLQIbBA3CPRIsa0uNUAw5S4xRUBFGQKMD4ecVEJojgD4OUznWCSJkX33Sj0jlgJ8EDYZl0DjULwje0Qy41CMi/gZgCZ+gDEBSThAvMhROFgChngAl8s9wahCUiDz2o2YiwtAtC+mZCcQxJYMYbdeMQwgHQ/LnEDdmiAjP535q0GIhlYayBVBTI4QlyODPFrsDvPehZRJibRNjU3qTMmJJOHLZkakGAsPJHxAEwu2qNw6BMy2OnnOfbXxB1t8wJtvGEAZiEzbGEHZiFSImkh4CV4bWCyH+ACn7S4AJOu4IgGFiFSQLSoDLE+nav+MdkPhZjMZodSd8KAUgsPUBJFRwHTaKSyI2azkNTIYBxtdDBMSsw06bk8zgAFEJRD5IL+ULj4n1hlWjAhBa0wbWVBBWlFXfUKJ8/DXfUy60iKHNQWgSg9dSNdIU0xKZUBcOteh1BfRWGpIEY7hRN4zlpRbBwGWMARXztyyqM0JA+M6J+qz7HIKlcxrnpgzQRZPyE/B0SF"
    $sData &= "+HRXaTNKRVNO8KkXO+kQdgkljyAKBaQCEMET8iIEIabQuglipxCLw2oK8kBwS0TClusEM8EPfYVeCdIP4Tc2ESX61kBbC3IxI12UUJBcRlBF4vMy+RFFviZQUSFABBUlScI6ZIDfRxFRDyGSG3rHYo4160O5vzk7wUpJkXPVLGKXXNwaTfCPIjoIIJcwpq2jmlC60eOeS4oQA294AC0Qrkqh1ClJjVUoC0mJbTDyGFAk/wQaQnPQxJDU2KT4RUUkgUT6CgjSRtBFoNAhZXQzMsPAGUYizhTmrhn0gDKXxiCv75qO3ZAW4Lo2KLluSIEepqGKBK7pAYy0g+z9B3nsRo6hQCmMIywo9sCMaq8sKIzjIGig0wya/niBXsfOOt5XJ2gVwOKjsswlJul8SQHM6Q6M4S4F6fw2hsERkofBEe6K4VW4i/lBubqhAKCojf5VC4nkoMAGRIqec3CNjiuvFgK+c+kb9gClCem5KjEAQBCkmB2k2L0gj74kEQNpgRjtoRiBITkGT4i1COi8xq4kq5fe4uAy9QCf/p2v0Prr36kWWKWwOUJqYbg59xJqo7gpamrhtfkPEM4m4PznLtqy6O0ejVDorn+SbwsBrnhEQ5chA0EpC6QGEEJ3AClZCbBBfR/zdRDzCA9vBRm+mjD/8OeZ5bAuQRvS0+ETDLRSF17GkPnAAoP4AXW+adDGK/QgVuqB53YRPtUg5AQH1J0r4SzE2LNZHYGHpLCzXsfSCVFgNgXrRVA5JuiSvnpc49NhIK3gISUDyCkm6yiEA3dKSByEwLRoXQiE0Vagh9JX2EKguicKL8AmlxF0tWpuAuXSin0DcH0DeOLrQAfmnDhJDZqUEF770bR4lK7JBf8VtWaz0SzgpTnK1WONEypEjRP7jTP133UJjSNQjTO9BEUEjVPRBUg4RhVT1+BYKv8VlKposN47QNPmAp8jsdrrJAPMwjOiDwVCGQEFdXkHzLEEFm+xvmglLJBAcoUcFSHOAf/QhOC9uP8aB+VEGDA+NpIM" & _
            "YGIG2ABLCDf+2tJCkLRI/X9SwQc1B7+YIS0TbAhEyAb32Bv/g+cb/oPHgs1TYYHUBZEkkBQg7gNQGJYYQNjIDg47+3R59OJETNY9gAMYBqA1+C5AdzDI8B9AF+IhGuYasBhEQEG0mAQhSwFBGgsPt1EYJTeMBgSE9wV/lN7dpoAAdTARBTyiODdNGI4okl4nFBOSOoD3Aky5kQg0hVMS2+BBAOtfITgxJ7whOM4eQEWC1IvlGoukwjgDjKRMCozT9F6MkYsfUNnCZi4C5QjyxGH7BVBbYrkwJ9IDeyBNi11jyg0wrugojn9cHKAvRBRlPAIf0nJTR+ByBlUBMSyCuwER+y4kUBEwEyQWG1HQD1CkGJwZ41FfsoIFhiFAJsbGARCQFo8Arh5AirVjHRDMzIH5rVDebQNKmvYSqP/umwgDOLn15mBRQWWoLOl+CCDSNvUI5Ik5QbheZ3CfB+FFCKklfQtEsOJwL+0ECkj/YAgcBMEBwcW+iaI1McG5AEYMQOQxEkFIpisAEtQYJSBJB4/4MIguoFMozqltvjLyTADgTIfQaCSQvt8vCAsGrZGyYAckRuLgAglmho6wSUb0Vr4DoglZClYzAldE5HD6QIjhkRCNDVfUWByANcQp0OUbKVBatMxsDAdYQIrFr6pYQ4nFkVhDiMV6qlhDh8VjWEOGxUWrWMOFxS5Yw1QfrTISu/VRPHDtKhJwxQLX5haxC1cccG1FJHDEANfqV3IJR/xVnbGNXh7h4TTM+i90EZVkcO1XEXBliTRwVnHgZlxx1gqcI2sY/QZYLKKKNdRYvIA1xFh8raiUGo0NrFgMjDWcqlicijWMWCyJNXSqWIyJNVxYHIg1RPVYLLcqjsOAxVNYw4GlNHQaYRVIe6BkI+sh0uq2gAEaviuyPmHFIut7DJosIoILmi0hBNl5PR0AADCFxHQVBJXkAkDGxgYwRSdHBfCW5EZnQ1OEB5vVABcRABBLRVJOCEVMMzKMclRGBwAl9zZWNjeHhFAWBgcAEUZy" & _
            "ZQBlAGxzdHJsZSFuV4AVREYmBwlAAHDBlCYmFyYHkAfQVMdGlyYEkEdXRvV2lUYmUDaEFiYHAFJUQW90BsD0FkbGTGBXMFkVYcRWN4cGYJTGViZUZ2YGUCY3B3All0aHVgYU5RBTYswQ5kYGCANDb21wYXJlhEkCaW5nV50hQWwDbG9jAAAo/g0BABJvbGVNQkNvkYAAAENMU0lERnJvcW3sw3NZAmdJRLQg2CRR5VY3V8ZMaQhlc0V4XJDkliZAlxbGlqaXD1EF4paGk0NyZWF0AGVJbnN0YW5jgEBBFTa31lTWxho/khQMCpCAARBPTAhFQVVUkUEBtwAAAaMAAboBAboAAAG7AQGSAAEAVAABSgABCAAAAQkAAQIAAQoAAAEZAAETAAEADwABGgABEQAAARgAARcAAQwAAAEHAAF9AAEAhQABBQEBMwAAARIBAUcAAT0QAAGI+BAQYAPwEFARECB5AtEQaRO+HwGANACgAPxQWAiJEy0A7h4AiNNUB6CnBvB1tYjHA4A0gI9ltnhEAcwBUq3wg262BAAgmHkFkIVuxAQNCwTAdEvohWBglUYgR1cXxpY/QFcGMEYHoAWVhc4EDAh8sAJFFZSlBkAQhKWGp3WVBaqSENBUQAELhnhggHiEEoSFJ2jxjzKNpWwIPlkQizBIRQMGHeCP1CojGEsFEcBu3lAMsojNqgAMKwCAaPRfZ08CdAjcPCB2YriB9K8MsM5g1or072xsYACA9HDr3BMJIUgYqzLAtMAKdfbrBM1IjT0ouQGw6RSquAyNB6tiIxGEFaCVteX1lc5sqQJkzYLSARAlOHYD8F8SpkMBAP+fSkAcAQBo8hBJvBxEM8ExmBIcQ8XMQfwcxMJHaCyZ3AskwrcchTPCMSDYFRwzBBAGwCZIAEEAcgBnFCKHSgR5Jgf2RQBkAGUrAGZUUMcJdETCNV5uTNDGBkQDMCwB19oA8EYBY21AfC3HDGtzNaBcU1YP8sYAxNfaAxNWFdJGDbwpRgVRaXUBZw2B"
    $sData &= "XwBONHBxB1BUF1BHDADVN6XspgIcQpzSN08AzEe8VSYcI1UD0NYUYMcBWk0sItYZAlBB/QBk1bWAPENUKELHCXUVA3Z0TSIAwzdQhUMlEHTitSRcUwhBB4DGBmSqvDRUOxJGBWU88qTbD6DGAlUQAOXHAGOylSVjTQQACT7UdY00ZdMtT0DHRRg3HNNNs10gUUYj1iMEUHc8TFMQRGxswhowxBY2xzD3JKYmRhEAw/QGwOYCkvVWB2IG8FbnRgZCh1YGAFIWNkdXJtcSUHZ25gLSoMBJbW99wn0WaXUADRd0VGCH1g9WdQcQZERQSUGm3h0GMJUYb3UFZ1ye1k4F/RhUCUALXbBstan1PFEDL9AZUQOL1hPgVmHmyhNhPMLWYcR5fbBkt30AdQ0DnGVWDNIuATCP3VE2ADTdgDRlKgf1IX9ptSDMwTfdcQDVAQVQ/Z1zANNDBzA0yxtznSn3/RkAQbO9F28EUJPHU60tGm48ktekBhBu9SCdPAdAIDMTPQ00Y2Q3QsA8pprxTg9kBwAPNAAGAA8yC3ABBgQCAAYyAg8WEAAHaIJKjcATUjzDQSYBAMBBFQHAQQMBAMAhiQFvAU5BAC0BDAHHFhhkAAoAGFQJABg0CAgAGDJcR0GWAABAQYUAQEFzAAFAIcM0GXQJAAAZZAgAGVQHAEUZ3QAZMhWpIgQOQQAOEHwfIEgZAB10CwBUHZ0AHZ0AHZ0AHTIIGeAX0OwjgSYgAUAgxk80DswBDzABHgH9ANwA58EAtmC1Ay0RHQFNABTRTQAtUQosoCBjwCIAJmhPAB9kqAAAH1SnAB80pgANHwGgAJWyDRAKDQDKTQL9ABRSjVIOTNBAyAQMABSSjULEE6gdAhIdAhE9ARAAGVTSDTIPPQMP7PAgdbTAKQ1SCW0k7VEA/89coaogCCYMQCpOBSgcsNTKQkwcAMdCHGHJQlUMupjwxwCULIDIAXK6DMAVfmcAsMwAtMws4HgGAPss5HkGACvm3ozAEi5oANDFALwBVssB" & _
            "wQzQAQEvVwkVDpWwyhXQdpNgIzEJfKJcYNIdbqQsxBEx3p7AJTavEKbqowuQLIDPAZK+n0CQzwLuY+A9Fh0gHBGc0gej6UoGAO0MwBXwHEDkpcI+XBzBAP0hHNCAtMI7eCwNAi0AAGU8jMARQy7dMEQcoEbJAOzBzAECPvuApcoCBBygw0I8HBCnykJ0HKDKQqwcEN7KQih2UODLAK0Q+ELwsJYqEg59EKwcMKEFDiyxRcEBkCzEEdsM0Bmhxc0BKbJcwBIsHJCnykJ8HHDMQsgcMGHtHg4sQcEBaCzEEbz7LMQRLmXQEMIRcQzAFYgeZQBwcdMog+0EoAKgzwAcLMDPAR5XTywDwgFLDNB5wgFXZyyEyAHEDMAeIsGRpsoAXRVwcpXwkNUKQwjw/38cA4kgEAA5AP7LQZAcI072/88OAABmO94PhTUHAABJi8jorND//0iL2GZBOTT8dRNJi0T8CEiJA0mLTPwISIsB/1AISIt0JFBIi0QkYEmJXP4ISIt8JFhIiQTuSIkc70yLfCRI6cYxAABIjQ3EVAAA6EBFAABFM/9BO8cPhbwAAABIjXxtAEGNRwhmQTkE/HUHSYt0/AjrNkSLyEUzwEiL1kiLzv8VDFEAAEE7x3UYSYt0/AhBuAEAAABJi81Ii9bo4B8AAOsHSI01PFQAALkAAAEAAAAAxin9TAAAAABwsAAAAQAAABIAAAASAAAAKLAAAIWwAAD9sQAAeE8AAGhPAACATwAArFAAAMxRAAAgTwAA/E4AANhOAABMVAAAiE8AAMRPAAAwVQAAOFUAAExPAADcVAAAcE8AAEBVAAAkUAAAQXV0b0l0T2JqZWN0X1g2NC5kbGwAzbAAANWwAADfsAAA67AAAASxAAAfsQAAMbEAAESxAABcsQAAcLEAAISxAACasQAAqbEAALmxAADEsQAA1LEAAOGxAADssQAAQWRkRW51bQBBZGRNZXRob2QAQWRkUHJvcGVydHkAQXV0b0l0T2JqZWN0Q3JlYXRlT2Jq" & _
            "ZWN0AEF1dG9JdE9iamVjdENyZWF0ZU9iamVjdEV4AENsb25lQXV0b0l0T2JqZWN0AENyZWF0ZUF1dG9JdE9iamVjdABDcmVhdGVBdXRvSXRPYmplY3RDbGFzcwBDcmVhdGVEbGxDYWxsT2JqZWN0AENyZWF0ZVdyYXBwZXJPYmplY3QAQ3JlYXRlV3JhcHBlck9iamVjdEV4AElVbmtub3duQWRkUmVmAElVbmtub3duUmVsZWFzZQBJbml0aWFsaXplAE1lbW9yeUNhbGxFbnRyeQBSZW1vdmVNZW1iZXIAUmV0dXJuVGhpcwBXcmFwcGVyQWRkTWV0aG9kAAAAAQACAAMABAAFAAYABwAIAAkACgALAAwADQAOAA8AEAARAAAAAIiyAAAAAAAAAAAAAAyzAACIsgAAoLIAAAAAAAAAAAAAFbMAAKCyAACwsgAAAAAAAAAAAAAxswAAsLIAAMCyAAAAAAAAAAAAAEqzAADAsgAAAAAAAAAAAAAAAAAAAAAAAAAAAADosgAAAAAAAPuyAAAAAAAAAAAAAAAAAAAhswAAAAAAAAAAAAAAAAAAO7MAAAAAAAAAAAAAAAAAALcAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEdldE1vZHVsZUhhbmRsZUEAAABHZXRQcm9jQWRkcmVzcwBLRVJORUwzMgBTSExXQVBJLmRsbAAAAFN0clRvSW50NjRFeFcAb2xlMzIuZGxsAAAAQ29Jbml0aWFsaXplAE9MRUFVVDMyLmRsbABXVlNRUkFQSI0FTwMAAEiLMEgD8EgrwEiL/matweAMSIvIUK0ryEgD8YvIV0SLwf/JikQ5BogEMXX1SIvPSIvW6MIAAABeWoHqABAAACvJO8pzSovZrP/BPP91DYoGJP08FXXrrP/B6xc8jXUNigYkxzwFddqs/8HrBiT+POh1z1Fbg8EErQvAeAY7wnPB6wYDw3i7A8Irw4lG/OuySI09XP///7DpqrhT"
    $sData &= "AwAAq0iNBakCAACLUAyLeAgL/3Q9SIswSAPwSCvySIveSItIFEgry3Qoi1AQSAPySAP+K8Ar0gvQrMHiB9DocvYL0AvSdAtIA9pIKQtIO/dy4UiNBVsCAADpUQIAAFNVVldBVEFVQVaKAkG+AQAAAEUy0kU7xkWL6EyL2ogBTIvhRYvOQYveSYv2D4YSAgAARYTSQYvBdBVBinwDAUGKBAPA6ARAwOcEQAr46wRBijwDRQPOM+1FD7bCQYvFQSvARDvID4PHAQAAQIT/D4kcAQAARYTSQ4sMC3QDwekEgeH//w8ARQPOgfuBCAAAi9FzINHqQYTOdBSB4v8HAABFA8iBwoEAAABFMtbrUYPif+tJweoCg+EDdD7/yXQs/8l0F//JdTiB4v//AwBHjUwIAYHCQUQAAOvPgeL/PwAAgcJBBAAARQPO6xSB4v8DAABFA8iDwkHrsIPiP0ED1kWE0nQKQw+3DAvB6QTrC2ZDiwwLgeH/DwAAQQ+2wkUy1kQDyIvBg+APg/gPdAWNSAPrO0UDzoH5/w8AAHQIwekEg8ES6yhFhNJ0DEOLDAvB6QQPt8nrBUMPtwwLgcERAQAAQYPBAoH5EAEBAHRgi8MrwoXJSGPQdERJA9QD2YoCSQPWQYgENEkD9oPB/3Xv6yxFhNJ0FUMPtkwLAUMPtgQLweEEwegEC8jrBUMPtgwLQYgMNEED3kkD9kUDzkED7kAC/4P9CA+Mjf7//+tlRYTSdBtBjUH8RQPOi8hBigwDgeH8AAAAweEFRTLS6xFmQ4tMC/tBi8GB4cAPAAADyUCKx4PgfwPBjUwACEljwYXJdCFJjRQDi8HB4AJEA8gD2IsCSIPCBEGJBDRIg8YEg8H/de1BD7bKQYvVK9FEO8oPgu79//+Lw0FeQV1BXF9eXVvD6Xik//9MWf///////yAAAAAAEAAAAKAAAAAAAIABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" & _
            "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAQAAAAGAAAgAAAAAAAAAAAAAAAAAAAAQABAAAAMAAAgAAAAAAAAAAAAAAAAAAAAQAJBAAASAAAAFjAAAAYAwAAAAAAAAAAAAAYAzQAAABWAFMAXwBWAEUAUgBTAEkATwBOAF8ASQBOAEYATwAAAAAAvQTv/gAAAQACAAEAAAACAAIAAQAAAAIAAAAAAAAAAAAEAAAAAgAAAAAAAAAAAAAAAAAAAHYCAAABAFMAdAByAGkAbgBnAEYAaQBsAGUASQBuAGYAbwAAAFICAAABADAANAAwADkAMAA0AEIAMAAAADAACAABAEYAaQBsAGUAVgBlAHIAcwBpAG8AbgAAAAAAMQAuADIALgAyAC4AMAAAADQACAABAFAAcgBvAGQAdQBjAHQAVgBlAHIAcwBpAG8AbgAAADEALgAyAC4AMgAuADAAAAB6ACkAAQBGAGkAbABlAEQAZQBzAGMAcgBpAHAAdABpAG8AbgAAAAAAUAByAG8AdgBpAGQAZQBzACAAbwBiAGoAZQBjAHQAIABmAHUAbgBjAHQAaQBvAG4AYQBsAGkAdAB5ACAAZgBvAHIAIABBAHUAdABvAEkAdAAAAAAAOgANAAEAUAByAG8AZAB1AGMAdABOAGEAbQBlAAAAAABBAHUAdABvAEkAdABPAGIA" & _
            "agBlAGMAdAAAAAAAWAAaAAEATABlAGcAYQBsAEMAbwBwAHkAcgBpAGcAaAB0AAAAKABDACkAIABUAGgAZQAgAEEAdQB0AG8ASQB0AE8AYgBqAGUAYwB0AC0AVABlAGEAbQAAAEoAEQABAE8AcgBpAGcAaQBuAGEAbABGAGkAbABlAG4AYQBtAGUAAABBAHUAdABvAEkAdABPAGIAagBlAGMAdAAuAGQAbABsAAAAAAB6ACMAAQBUAGgAZQAgAEEAdQB0AG8ASQB0AE8AYgBqAGUAYwB0AC0AVABlAGEAbQAAAAAAbQBvAG4AbwBjAGUAcgBlAHMALAAgAHQAcgBhAG4AYwBlAHgAeAAsACAASwBpAHAALAAgAFAAcgBvAGcAQQBuAGQAeQAAAAAARAAAAAEAVgBhAHIARgBpAGwAZQBJAG4AZgBvAAAAAAAkAAQAAABUAHIAYQBuAHMAbABhAHQAaQBvAG4AAAAAAAkEsAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
    Return __Au3Obj_Mem_Base64Decode($sData)
EndFunc   ;==>__Au3Obj_Mem_BinDll_X64

#endregion Embedded DLL
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region DllStructCreate Wrapper

Func __Au3Obj_ObjStructMethod(ByRef $oSelf, $vParam1 = 0, $vParam2 = 0)
	Local $sMethod = $oSelf.__name__
	Local $tStructure = DllStructCreate($oSelf.__tag__, $oSelf.__pointer__)
	Local $vOut
	Switch @NumParams
		Case 1
			$vOut = DllStructGetData($tStructure, $sMethod)
		Case 2
			If $oSelf.__propcall__ Then
				$vOut = DllStructSetData($tStructure, $sMethod, $vParam1)
			Else
				$vOut = DllStructGetData($tStructure, $sMethod, $vParam1)
			EndIf
		Case 3
			$vOut = DllStructSetData($tStructure, $sMethod, $vParam2, $vParam1)
	EndSwitch
	If IsPtr($vOut) Then Return Number($vOut)
	Return $vOut
EndFunc   ;==>__Au3Obj_ObjStructMethod

Func __Au3Obj_ObjStructDestructor(ByRef $oSelf)
	If $oSelf.__new__ Then __Au3Obj_GlobalFree($oSelf.__pointer__)
EndFunc   ;==>__Au3Obj_ObjStructDestructor

Func __Au3Obj_ObjStructPointer(ByRef $oSelf, $vParam = Default)
	If $oSelf.__propcall__ Then Return SetError(1, 0, 0)
	If @NumParams = 1 Or IsKeyword($vParam) Then Return $oSelf.__pointer__
	Return Number(DllStructGetPtr(DllStructCreate($oSelf.__tag__, $oSelf.__pointer__), $vParam))
EndFunc   ;==>__Au3Obj_ObjStructPointer

#endregion DllStructCreate Wrapper
;--------------------------------------------------------------------------------------------------------------------------------------


;--------------------------------------------------------------------------------------------------------------------------------------
#region Public UDFs

Global Enum $ELTYPE_NOTHING, $ELTYPE_METHOD, $ELTYPE_PROPERTY
Global Enum $ELSCOPE_PUBLIC, $ELSCOPE_READONLY, $ELSCOPE_PRIVATE

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_AddDestructor
; Description ...: Adds a destructor to an AutoIt-object
; Syntax.........: _AutoItObject_AddDestructor(ByRef $oObject, $sAutoItFunc)
; Parameters ....: $oObject     - the object to modify
;                  $sAutoItFunc - the AutoIt-function wich represents this destructor.
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: monoceres (Andreas Karlsson)
; Modified.......:
; Remarks .......: Adding a method that will be called on object destruction. Can be called multiple times.
; Related .......: _AutoItObject_AddProperty, _AutoItObject_AddEnum, _AutoItObject_RemoveMember, _AutoItObject_AddMethod
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_AddDestructor(ByRef $oObject, $sAutoItFunc)
	Return _AutoItObject_AddMethod($oObject, "~", $sAutoItFunc, True)
EndFunc   ;==>_AutoItObject_AddDestructor

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_AddEnum
; Description ...: Adds an Enum to an AutoIt-object
; Syntax.........: _AutoItObject_AddEnum(ByRef $oObject, $sNextFunc, $sResetFunc [, $sSkipFunc = ''])
; Parameters ....: $oObject     - the object to modify
;                  $sNextFunc   - The function to be called to get the next entry
;                  $sResetFunc  - The function to be called to reset the enum
;                  $sSkipFunc   - [optional] The function to be called to skip elements (not supported by AutoIt)
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_AddMethod, _AutoItObject_AddProperty, _AutoItObject_RemoveMember
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_AddEnum(ByRef $oObject, $sNextFunc, $sResetFunc, $sSkipFunc = '')
	; Author: Prog@ndy
	If Not IsObj($oObject) Then Return SetError(2, 0, 0)
	DllCall($ghAutoItObjectDLL, "none", "AddEnum", "idispatch", $oObject, "wstr", $sNextFunc, "wstr", $sResetFunc, "wstr", $sSkipFunc)
	If @error Then Return SetError(1, @error, 0)
	Return True
EndFunc   ;==>_AutoItObject_AddEnum

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_AddMethod
; Description ...: Adds a method to an AutoIt-object
; Syntax.........: _AutoItObject_AddMethod(ByRef $oObject, $sName, $sAutoItFunc [, $fPrivate = False])
; Parameters ....: $oObject     - the object to modify
;                  $sName       - the name of the method to add
;                  $sAutoItFunc - the AutoIt-function wich represents this method.
;                  $fPrivate    - [optional] Specifies whether the function can only be called from within the object. (default: False)
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......: The first parameter of the AutoIt-function is always a reference to the object. ($oSelf)
;                  This parameter will automatically be added and must not be given in the call.
;                  The function called '__default__' is accesible without a name using brackets ($return = $oObject())
; Related .......: _AutoItObject_AddProperty, _AutoItObject_AddEnum, _AutoItObject_RemoveMember
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_AddMethod(ByRef $oObject, $sName, $sAutoItFunc, $fPrivate = False)
	; Author: Prog@ndy
	If Not IsObj($oObject) Then Return SetError(2, 0, 0)
	Local $iFlags = 0
	If $fPrivate Then $iFlags = $ELSCOPE_PRIVATE
	DllCall($ghAutoItObjectDLL, "none", "AddMethod", "idispatch", $oObject, "wstr", $sName, "wstr", $sAutoItFunc, 'dword', $iFlags)
	If @error Then Return SetError(1, @error, 0)
	Return True
EndFunc   ;==>_AutoItObject_AddMethod

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_AddProperty
; Description ...: Adds a property to an AutoIt-object
; Syntax.........: _AutoItObject_AddProperty(ByRef $oObject, $sName [, $iFlags = $ELSCOPE_PUBLIC [, $vData = ""]])
; Parameters ....: $oObject     - the object to modify
;                  $sName       - the name of the property to add
;                  $iFlags      - [optional] Specifies the access to the property
;                  $vData       - [optional] Initial data for the property
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......: The property called '__default__' is accesible without a name using brackets ($value = $oObject())
;                  + $iFlags can be:
;                  |$ELSCOPE_PUBLIC   - The Property has public access.
;                  |$ELSCOPE_READONLY - The property is read-only and can only be changed from within the object.
;                  |$ELSCOPE_PRIVATE  - The property is private and can only be accessed from within the object.
;                  +
;                  + Initial default value for every new property is nothing (no value).
; Related .......: _AutoItObject_AddMethod, _AutoItObject_AddEnum, _AutoItObject_RemoveMember
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_AddProperty(ByRef $oObject, $sName, $iFlags = $ELSCOPE_PUBLIC, $vData = "")
	; Author: Prog@ndy
	Local Static $tStruct = DllStructCreate($__Au3Obj_tagVARIANT)
	If Not IsObj($oObject) Then Return SetError(2, 0, 0)
	Local $pData = 0
	If @NumParams = 4 Then
		$pData = DllStructGetPtr($tStruct)
		_AutoItObject_VariantInit($pData)
		$oObject.__bridge__(Number($pData)) = $vData
	EndIf
	DllCall($ghAutoItObjectDLL, "none", "AddProperty", "idispatch", $oObject, "wstr", $sName, 'dword', $iFlags, 'ptr', $pData)
	Local $error = @error
	If $pData Then _AutoItObject_VariantClear($pData)
	If $error Then Return SetError(1, $error, 0)
	Return True
EndFunc   ;==>_AutoItObject_AddProperty

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_Class
; Description ...: AutoItObject COM wrapper function
; Syntax.........: _AutoItObject_Class()
; Parameters ....:
; Return values .: Success      - object with defined:
;                   -methods:
;                  |	Create([$oParent = 0]) - creates AutoItObject object
;                  |	AddMethod($sName, $sAutoItFunc [, $fPrivate = False]) - adds new method
;                  |	AddProperty($sName, $iFlags = $ELSCOPE_PUBLIC, $vData = 0) - adds new property
;                  |	AddDestructor($sAutoItFunc) - adds destructor
;                  |	AddEnum($sNextFunc, $sResetFunc [, $sSkipFunc = '']) - adds enum
;                  |	RemoveMember($sMember) - removes member
;                   -properties:
;                  |	Object - readonly property representing the last created AutoItObject object
; Author ........: trancexx
; Modified.......:
; Remarks .......: "Object" propery can be accessed only once for one object. After that new AutoItObject object is created.
;                  +Method "Create" will discharge previous AutoItObject object and create a new one.
; Related .......: _AutoItObject_Create
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_Class()
	Local $aCall = DllCall($ghAutoItObjectDLL, "idispatch", "CreateAutoItObjectClass")
	If @error Then Return SetError(1, @error, 0)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_Class

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_CLSIDFromString
; Description ...: Converts a string to a CLSID-Struct (GUID-Struct)
; Syntax.........: _AutoItObject_CLSIDFromString($sString)
; Parameters ....: $sString     - The string to convert
; Return values .: Success      - DLLStruct in format $tagGUID
;                  Failure      - 0
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_CoCreateInstance
; Link ..........: http://msdn.microsoft.com/en-us/library/ms680589(VS.85).aspx
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_CLSIDFromString($sString)
	Local $tCLSID = DllStructCreate("dword;word;word;byte[8]")
	Local $aResult = DllCall($gh_AU3Obj_ole32dll, 'long', 'CLSIDFromString', 'wstr', $sString, 'ptr', DllStructGetPtr($tCLSID))
	If @error Then Return SetError(1, @error, 0)
	If $aResult[0] <> 0 Then Return SetError(2, $aResult[0], 0)
	Return $tCLSID
EndFunc   ;==>_AutoItObject_CLSIDFromString

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_CoCreateInstance
; Description ...: Creates a single uninitialized object of the class associated with a specified CLSID.
; Syntax.........: _AutoItObject_CoCreateInstance($rclsid, $pUnkOuter, $dwClsContext, $riid, ByRef $ppv)
; Parameters ....: $rclsid       - The CLSID associated with the data and code that will be used to create the object.
;                  $pUnkOuter    - If NULL, indicates that the object is not being created as part of an aggregate.
;                  +If non-NULL, pointer to the aggregate object's IUnknown interface (the controlling IUnknown).
;                  $dwClsContext - Context in which the code that manages the newly created object will run.
;                  +The values are taken from the enumeration CLSCTX.
;                  $riid         - A reference to the identifier of the interface to be used to communicate with the object.
;                  $ppv          - [out byref] Variable that receives the interface pointer requested in riid.
;                  +Upon successful return, *ppv contains the requested interface pointer. Upon failure, *ppv contains NULL.
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_ObjCreate, _AutoItObject_CLSIDFromString
; Link ..........: http://msdn.microsoft.com/en-us/library/ms686615(VS.85).aspx
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_CoCreateInstance($rclsid, $pUnkOuter, $dwClsContext, $riid, ByRef $ppv)
	$ppv = 0
	Local $aResult = DllCall($gh_AU3Obj_ole32dll, 'long', 'CoCreateInstance', 'ptr', $rclsid, 'ptr', $pUnkOuter, 'dword', $dwClsContext, 'ptr', $riid, 'ptr*', 0)
	If @error Then Return SetError(1, @error, 0)
	$ppv = $aResult[5]
	Return SetError($aResult[0], 0, $aResult[0] = 0)
EndFunc   ;==>_AutoItObject_CoCreateInstance

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_Create
; Description ...: Creates an AutoIt-object
; Syntax.........: _AutoItObject_Create( [$oParent = 0] )
; Parameters ....: $oParent     - [optional] an AutoItObject whose methods & properties are copied. (default: 0)
; Return values .: Success      - AutoIt-Object
;                  Failure      - 0
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_Class
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_Create($oParent = 0)
	; Author: Prog@ndy
	Local $aResult
	Switch IsObj($oParent)
		Case True
			$aResult = DllCall($ghAutoItObjectDLL, "idispatch", "CloneAutoItObject", 'idispatch', $oParent)
		Case Else
			$aResult = DllCall($ghAutoItObjectDLL, "idispatch", "CreateAutoItObject")
	EndSwitch
	If @error Then Return SetError(1, @error, 0)
	Return $aResult[0]
EndFunc   ;==>_AutoItObject_Create

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_DllOpen
; Description ...: Creates an object associated with specified dll
; Syntax.........: _AutoItObject_DllOpen($sDll [, $sTag = "" [, $iFlag = 0]])
; Parameters ....: $sDll - Dll for which to create an object
;                  $sTag - [optional] String representing function return value and parameters.
;                  $iFlag - [optional] Flag specifying the level of loading. See MSDN about LoadLibraryEx function for details. Default is 0.
; Return values .: Success      - Dispatch-Object
;                  Failure      - 0
; Author ........: trancexx
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_WrapperCreate
; Link ..........: http://msdn.microsoft.com/en-us/library/ms684179(VS.85).aspx
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_DllOpen($sDll, $sTag = "", $iFlag = 0)
	Local $sTypeTag = "wstr"
	If $sTag = Default Or Not $sTag Then $sTypeTag = "ptr"
	Local $aCall = DllCall($ghAutoItObjectDLL, "idispatch", "CreateDllCallObject", "wstr", $sDll, $sTypeTag, __Au3Obj_GetMethods($sTag), "dword", $iFlag)
	If @error Or Not IsObj($aCall[0]) Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_DllOpen

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_DllStructCreate
; Description ...: Object wrapper for DllStructCreate and related functions
; Syntax.........: _AutoItObject_DllStructCreate($sTag [, $vParam = 0])
; Parameters ....: $sTag     - A string representing the structure to create (same as with DllStructCreate)
;                  $vParam   - [optional] If this parameter is DLLStruct type then it will be copied to newly allocated space and maintained during lifetime of the object. If this parameter is not suplied needed memory allocation is done but content is initialized to zero. In all other cases function will not allocate memory but use parameter supplied as the pointer (same as DllStructCreate)
; Return values .: Success      - Object-structure
;                  Failure      - 0, @error is set to error value of DllStructCreate() function.
; Author ........: trancexx
; Modified.......:
; Remarks .......: AutoIt can't handle pointers properly when passed to or returned from object methods. Use Number() function on pointers before using them with this function.
;                  +Every element of structure must be named. Values are accessed through their names.
;                  +Created object exposes:
;                  +  - set of dynamic methods in names of elements of the structure
;                  +  - readonly properties:
;                  |	__tag__ - a string representing the object-structure
;                  |	__size__ - the size of the struct in bytes
;                  |	__alignment__ - alignment string (e.g. "align 2")
;                  |	__count__ - number of elements of structure
;                  |	__elements__ - string made of element names separated by semicolon (;)
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_DllStructCreate($sTag, $vParam = 0)
	Local $fNew = False
	Local $tSubStructure = DllStructCreate($sTag)
	If @error Then Return SetError(@error, 0, 0)
	Local $iSize = DllStructGetSize($tSubStructure)
	Local $pPointer = $vParam
	Select
		Case @NumParams = 1
			; Will allocate fixed 128 extra bytes due to possible misalignment and other issues
			$pPointer = __Au3Obj_GlobalAlloc($iSize + 128, 64) ; GPTR
			If @error Then Return SetError(3, 0, 0)
			$fNew = True
		Case IsDllStruct($vParam)
			$pPointer = __Au3Obj_GlobalAlloc($iSize, 64) ; GPTR
			If @error Then Return SetError(3, 0, 0)
			$fNew = True
			DllStructSetData(DllStructCreate("byte[" & $iSize & "]", $pPointer), 1, DllStructGetData(DllStructCreate("byte[" & $iSize & "]", DllStructGetPtr($vParam)), 1))
		Case @NumParams = 2 And $vParam = 0
			Return SetError(3, 0, 0)
	EndSelect
	Local $sAlignment
	Local $sNamesString = __Au3Obj_ObjStructGetElements($sTag, $sAlignment)
	Local $aElements = StringSplit($sNamesString, ";", 2)
	Local $oObj = _AutoItObject_Class()
	For $i = 0 To UBound($aElements) - 1
		$oObj.AddMethod($aElements[$i], "__Au3Obj_ObjStructMethod")
	Next
	$oObj.AddProperty("__tag__", $ELSCOPE_READONLY, $sTag)
	$oObj.AddProperty("__size__", $ELSCOPE_READONLY, $iSize)
	$oObj.AddProperty("__alignment__", $ELSCOPE_READONLY, $sAlignment)
	$oObj.AddProperty("__count__", $ELSCOPE_READONLY, UBound($aElements))
	$oObj.AddProperty("__elements__", $ELSCOPE_READONLY, $sNamesString)
	$oObj.AddProperty("__new__", $ELSCOPE_PRIVATE, $fNew)
	$oObj.AddProperty("__pointer__", $ELSCOPE_READONLY, Number($pPointer))
	$oObj.AddMethod("__default__", "__Au3Obj_ObjStructPointer")
	$oObj.AddDestructor("__Au3Obj_ObjStructDestructor")
	Return $oObj.Object
EndFunc   ;==>_AutoItObject_DllStructCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_IDispatchToPtr
; Description ...: Returns pointer to AutoIt's object type
; Syntax.........: _AutoItObject_IDispatchToPtr(ByRef $oIDispatch)
; Parameters ....: $oIDispatch  - Object
; Return values .: Success      - Pointer to object
;                  Failure      - 0
; Author ........: monoceres, trancexx
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_PtrToIDispatch, _AutoItObject_CoCreateInstance, _AutoItObject_ObjCreate
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_IDispatchToPtr($oIDispatch)
	Local $aCall = DllCall($ghAutoItObjectDLL, "ptr", "ReturnThis", "idispatch", $oIDispatch)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_IDispatchToPtr

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_IUnknownAddRef
; Description ...: Increments the refrence count of an IUnknown-Object
; Syntax.........: _AutoItObject_IUnknownAddRef($vUnknown)
; Parameters ....: $vUnknown    - IUnkown-pointer or object itself
; Return values .: Success      - New reference count.
;                  Failure      - 0, @error is set.
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_IUnknownAddRef(Const $vUnknown)
	; Author: Prog@ndy
	Local $sType = "ptr"
	If IsObj($vUnknown) Then $sType = "idispatch"
	Local $aCall = DllCall($ghAutoItObjectDLL, "dword", "IUnknownAddRef", $sType, $vUnknown)
	If @error Then Return SetError(1, @error, 0)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_IUnknownAddRef

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_IUnknownRelease
; Description ...: Decrements the refrence count of an IUnknown-Object
; Syntax.........: _AutoItObject_IUnknownRelease($vUnknown)
; Parameters ....: $vUnknown    - IUnkown-pointer or object itself
; Return values .: Success      - New reference count.
;                  Failure      - 0, @error is set.
; Author ........: trancexx
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_IUnknownRelease(Const $vUnknown)
	Local $sType = "ptr"
	If IsObj($vUnknown) Then $sType = "idispatch"
	Local $aCall = DllCall($ghAutoItObjectDLL, "dword", "IUnknownRelease", $sType, $vUnknown)
	If @error Then Return SetError(1, @error, 0)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_IUnknownRelease

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_ObjCreate
; Description ...: Creates a reference to a COM object
; Syntax.........: _AutoItObject_ObjCreate($sID [, $sRefId = Default [, $tagInterface = Default ]] )
; Parameters ....: $sID - Object identifier. Either string representation of CLSID or ProgID
;                  $sRefId - [optional] String representation of the identifier of the interface to be used to communicate with the object. Default is the value of IDispatch
;                  $tagInterface - [optional] String defining the methods of the Interface, see Remarks for _AutoItObject_WrapperCreate function for details
; Return values .: Success      - Dispatch-Object
;                  Failure      - 0
; Author ........: trancexx
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_ObjCreateEx, _AutoItObject_WrapperCreate
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_ObjCreate($sID, $sRefId = Default, $tagInterface = Default)
	Local $sTypeRef = "wstr"
	If $sRefId = Default Or Not $sRefId Then $sTypeRef = "ptr"
	Local $sTypeTag = "wstr"
	If $tagInterface = Default Or Not $tagInterface Then $sTypeTag = "ptr"
	Local $aCall = DllCall($ghAutoItObjectDLL, "idispatch", "AutoItObjectCreateObject", "wstr", $sID, $sTypeRef, $sRefId, $sTypeTag, __Au3Obj_GetMethods($tagInterface))
	If @error Or Not IsObj($aCall[0]) Then Return SetError(1, 0, 0)
	If $sTypeRef = "ptr" And $sTypeTag = "ptr" Then _AutoItObject_IUnknownRelease($aCall[0])
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_ObjCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_ObjCreateEx
; Description ...: Creates a reference to a COM object
; Syntax.........: _AutoItObject_ObjCreateEx($sModule, $sCLSID [, $sRefId = Default [, $tagInterface = Default [, $fWrapp = False]]] )
; Parameters ....: $sModule - Full path to the module with class (object)
;                  $sCLSID - Object identifier. String representation of CLSID.
;                  $sRefId - [optional] String representation of the identifier of the interface to be used to communicate with the object. Default is the value of IDispatch
;                  $tagInterface - [optional] String defining the methods of the Interface, see Remarks for _AutoItObject_WrapperCreate function for details
;                  $fWrapped - [optional] Specifies whether to wrapp created object.
; Return values .: Success      - Dispatch-Object
;                  Failure      - 0
; Author ........: trancexx
; Modified.......:
; Remarks .......: This function doesn't require any additional registration of the classes and interaces supported in the server module.
;                 +In case $tagInterface is specified $fWrapp parameter is ignored.
;                 +If $sRefId is left default then first supported interface by the coclass is returned (the default dispatch).
; Related .......: _AutoItObject_ObjCreate, _AutoItObject_WrapperCreate
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_ObjCreateEx($sModule, $sID, $sRefId = Default, $tagInterface = Default, $fWrapp = False)
	Local $sTypeRef = "wstr"
	If $sRefId = Default Or Not $sRefId Then $sTypeRef = "ptr"
	Local $sTypeTag = "wstr"
	If $tagInterface = Default Or Not $tagInterface Then
		$sTypeTag = "ptr"
	Else
		$fWrapp = True
	EndIf
	Local $aCall = DllCall($ghAutoItObjectDLL, "idispatch", "AutoItObjectCreateObjectEx", "wstr", $sModule, "wstr", $sID, $sTypeRef, $sRefId, $sTypeTag, __Au3Obj_GetMethods($tagInterface), "bool", $fWrapp)
	If @error Or Not IsObj($aCall[0]) Then Return SetError(1, 0, 0)
	If Not $fWrapp Then _AutoItObject_IUnknownRelease($aCall[0])
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_ObjCreateEx

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_ObjectFromDtag
; Description ...: Creates custom object defined with "dtag" interface description string
; Syntax.........: _AutoItObject_ObjectFromDtag($sFunctionPrefix, $dtagInterface [, $fNoUnknown = False])
; Parameters ....: $sFunctionPrefix  - The prefix of the functions you define as object methods
;                  $dtagInterface - string describing the interface (dtag)
;                  $fNoUnknown - [optional] NOT an IUnkown-Interface. Do not call "Release" method when out of scope (Default: False, meaining to call Release method)
; Return values .: Success      - object type
;                  Failure      - 0
; Author ........: trancexx
; Modified.......:
; Remarks .......: Main purpose of this function is to create custom objects that serve as event handlers for other objects.
;                  +Registered callback functions (defined methods) are left for AutoIt to free at its convenience on exit.
; Related .......: _AutoItObject_ObjCreate, _AutoItObject_ObjCreateEx, _AutoItObject_WrapperCreate
; Link ..........: http://msdn.microsoft.com/en-us/library/ms692727(VS.85).aspx
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_ObjectFromDtag($sFunctionPrefix, $dtagInterface, $fNoUnknown = False)
	Local $sMethods = __Au3Obj_GetMethods($dtagInterface)
	$sMethods = StringReplace(StringReplace(StringReplace(StringReplace($sMethods, "object", "idispatch"), "variant*", "ptr"), "hresult", "long"), "bstr", "ptr")
	Local $aMethods = StringSplit($sMethods, @LF, 3)
	Local $iUbound = UBound($aMethods)
	Local $sMethod, $aSplit, $sNamePart, $aTagPart, $sTagPart, $sRet, $sParams
	; Allocation. Read http://msdn.microsoft.com/en-us/library/ms810466.aspx to see why like this (object + methods):
	Local $tInterface = DllStructCreate("ptr[" & $iUbound + 1 & "]", __Au3Obj_CoTaskMemAlloc($__Au3Obj_PTR_SIZE * ($iUbound + 1)))
	If @error Then Return SetError(1, 0, 0)
	For $i = 0 To $iUbound - 1
		$aSplit = StringSplit($aMethods[$i], "|", 2)
		If UBound($aSplit) <> 2 Then ReDim $aSplit[2]
		$sNamePart = $aSplit[0]
		$sTagPart = $aSplit[1]
		$sMethod = $sFunctionPrefix & $sNamePart
		$aTagPart = StringSplit($sTagPart, ";", 2)
		$sRet = $aTagPart[0]
		$sParams = StringReplace($sTagPart, $sRet, "", 1)
		$sParams = "ptr" & $sParams
		DllStructSetData($tInterface, 1, DllCallbackGetPtr(DllCallbackRegister($sMethod, $sRet, $sParams)), $i + 2) ; Freeing is left to AutoIt.
	Next
	DllStructSetData($tInterface, 1, DllStructGetPtr($tInterface) + $__Au3Obj_PTR_SIZE) ; Interface method pointers are actually pointer size away
	Return _AutoItObject_WrapperCreate(DllStructGetPtr($tInterface), $dtagInterface, $fNoUnknown, True) ; and first pointer is object pointer that's wrapped
EndFunc   ;==>_AutoItObject_ObjectFromDtag

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_PtrToIDispatch
; Description ...: Converts IDispatch pointer to AutoIt's object type
; Syntax.........: _AutoItObject_PtrToIDispatch($pIDispatch)
; Parameters ....: $pIDispatch  - IDispatch pointer
; Return values .: Success      - object type
;                  Failure      - 0
; Author ........: monoceres, trancexx
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_IDispatchToPtr, _AutoItObject_WrapperCreate
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_PtrToIDispatch($pIDispatch)
	Local $aCall = DllCall($ghAutoItObjectDLL, "idispatch", "ReturnThis", "ptr", $pIDispatch)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_PtrToIDispatch

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_RemoveMember
; Description ...: Removes a property or a function from an AutoIt-object
; Syntax.........: _AutoItObject_RemoveMember(ByRef $oObject, $sMember)
; Parameters ....: $oObject     - the object to modify
;                  $sMember     - the name of the member to remove
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_AddMethod, _AutoItObject_AddProperty, _AutoItObject_AddEnum
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_RemoveMember(ByRef $oObject, $sMember)
	; Author: Prog@ndy
	If Not IsObj($oObject) Then Return SetError(2, 0, 0)
	If $sMember = '__default__' Then Return SetError(3, 0, 0)
	DllCall($ghAutoItObjectDLL, "none", "RemoveMember", "idispatch", $oObject, "wstr", $sMember)
	If @error Then Return SetError(1, @error, 0)
	Return True
EndFunc   ;==>_AutoItObject_RemoveMember

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_Shutdown
; Description ...: frees the AutoItObject DLL
; Syntax.........: _AutoItObject_Shutdown()
; Parameters ....: $fFinal    - [optional] Force shutdown of the library? (Default: False)
; Return values .: Remaining reference count (one for each call to _AutoItObject_Startup)
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......: Usage of this function is optonal. The World wouldn't end without it.
; Related .......: _AutoItObject_Startup
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_Shutdown($fFinal = False)
	; Author: Prog@ndy
	If $giAutoItObjectDLLRef <= 0 Then Return 0
	$giAutoItObjectDLLRef -= 1
	If $fFinal Then $giAutoItObjectDLLRef = 0
	If $giAutoItObjectDLLRef = 0 Then DllCall($ghAutoItObjectDLL, "ptr", "Initialize", "ptr", 0, "ptr", 0)
	Return $giAutoItObjectDLLRef
EndFunc   ;==>_AutoItObject_Shutdown

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_Startup
; Description ...: Initializes AutoItObject
; Syntax.........: _AutoItObject_Startup( [$fLoadDLL = False [, $sDll = "AutoitObject.dll"]] )
; Parameters ....: $fLoadDLL    - [optional] specifies whether an external DLL-file should be used (default: False)
;                  $sDLL        - [optional] the path to the external DLL (default: AutoitObject.dll or AutoitObject_X64.dll)
; Return values .: Success      - True
;                  Failure      - False
; Author ........: trancexx, Prog@ndy
; Modified.......:
; Remarks .......: Automatically switches between 32bit and 64bit mode if no special DLL is specified.
; Related .......: _AutoItObject_Shutdown
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_Startup($fLoadDLL = False, $sDll = "AutoitObject.dll")
	Local Static $__Au3Obj_FunctionProxy = DllCallbackGetPtr(DllCallbackRegister("__Au3Obj_FunctionProxy", "int", "wstr;idispatch"))
	Local Static $__Au3Obj_EnumFunctionProxy = DllCallbackGetPtr(DllCallbackRegister("__Au3Obj_EnumFunctionProxy", "int", "dword;wstr;idispatch;ptr;ptr"))
	If $ghAutoItObjectDLL = -1 Then
		If $fLoadDLL Then
			If $__Au3Obj_X64 And @NumParams = 1 Then $sDll = "AutoItObject_X64.dll"
			$ghAutoItObjectDLL = DllOpen($sDll)
		Else
			$ghAutoItObjectDLL = __Au3Obj_Mem_DllOpen()
		EndIf
		If $ghAutoItObjectDLL = -1 Then Return SetError(1, 0, False)
	EndIf
	If $giAutoItObjectDLLRef <= 0 Then
		$giAutoItObjectDLLRef = 0
		DllCall($ghAutoItObjectDLL, "ptr", "Initialize", "ptr", $__Au3Obj_FunctionProxy, "ptr", $__Au3Obj_EnumFunctionProxy)
		If @error Then
			DllClose($ghAutoItObjectDLL)
			$ghAutoItObjectDLL = -1
			Return SetError(2, 0, False)
		EndIf
	EndIf
	$giAutoItObjectDLLRef += 1
	Return True
EndFunc   ;==>_AutoItObject_Startup

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_VariantClear
; Description ...: Clears the value of a variant
; Syntax.........: _AutoItObject_VariantClear($pvarg)
; Parameters ....: $pvarg       - the VARIANT to clear
; Return values .: Success      - 0
;                  Failure      - nonzero
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_VariantFree
; Link ..........: http://msdn.microsoft.com/en-us/library/ms221165.aspx
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_VariantClear($pvarg)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "long", "VariantClear", "ptr", $pvarg)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_VariantClear

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_VariantCopy
; Description ...: Copies a VARIANT to another
; Syntax.........: _AutoItObject_VariantCopy($pvargDest, $pvargSrc)
; Parameters ....: $pvargDest   - Destionation variant
;                  $pvargSrc    - Source variant
; Return values .: Success      - 0
;                  Failure      - nonzero
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_VariantRead
; Link ..........: http://msdn.microsoft.com/en-us/library/ms221697.aspx
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_VariantCopy($pvargDest, $pvargSrc)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "long", "VariantCopy", "ptr", $pvargDest, 'ptr', $pvargSrc)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_VariantCopy

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_VariantFree
; Description ...: Frees a variant created by _AutoItObject_VariantSet
; Syntax.........: _AutoItObject_VariantFree($pvarg)
; Parameters ....: $pvarg       - the VARIANT to free
; Return values .: Success      - 0
;                  Failure      - nonzero
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......: Use this function on variants created with _AutoItObject_VariantSet function (when first parameter for that function is 0).
; Related .......: _AutoItObject_VariantClear
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_VariantFree($pvarg)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "long", "VariantClear", "ptr", $pvarg)
	If @error Then Return SetError(1, 0, 1)
	If $aCall[0] = 0 Then __Au3Obj_CoTaskMemFree($pvarg)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_VariantFree

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_VariantInit
; Description ...: Initializes a variant.
; Syntax.........: _AutoItObject_VariantInit($pvarg)
; Parameters ....: $pvarg       - the VARIANT to initialize
; Return values .: Success      - 0
;                  Failure      - nonzero
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_VariantClear
; Link ..........: http://msdn.microsoft.com/en-us/library/ms221402.aspx
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_VariantInit($pvarg)
	; Author: Prog@ndy
	Local $aCall = DllCall($gh_AU3Obj_oleautdll, "long", "VariantInit", "ptr", $pvarg)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc   ;==>_AutoItObject_VariantInit

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_VariantRead
; Description ...: Reads the value of a VARIANT
; Syntax.........: _AutoItObject_VariantRead($pVariant)
; Parameters ....: $pVariant    - Pointer to VARaINT-structure
; Return values .: Success      - value of the VARIANT
;                  Failure      - 0
; Author ........: monoceres, Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_VariantSet
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_VariantRead($pVariant)
	; Author: monoceres, Prog@ndy
	Local $var = DllStructCreate($__Au3Obj_tagVARIANT, $pVariant), $data
	; Translate the vt id to a autoit dllcall type
	Local $VT = DllStructGetData($var, "vt"), $type
	Switch $VT
		Case $__Au3Obj_VT_I1, $__Au3Obj_VT_UI1
			$type = "byte"
		Case $__Au3Obj_VT_I2
			$type = "short"
		Case $__Au3Obj_VT_I4
			$type = "int"
		Case $__Au3Obj_VT_I8
			$type = "int64"
		Case $__Au3Obj_VT_R4
			$type = "float"
		Case $__Au3Obj_VT_R8
			$type = "double"
		Case $__Au3Obj_VT_UI2
			$type = 'word'
		Case $__Au3Obj_VT_UI4
			$type = 'uint'
		Case $__Au3Obj_VT_UI8
			$type = 'uint64'
		Case $__Au3Obj_VT_BSTR
			Return __Au3Obj_SysReadString(DllStructGetData($var, "data"))
		Case $__Au3Obj_VT_BOOL
			$type = 'short'
		Case BitOR($__Au3Obj_VT_ARRAY, $__Au3Obj_VT_UI1)
			Local $pSafeArray = DllStructGetData($var, "data")
			Local $bound, $pData, $lbound
			If 0 = __Au3Obj_SafeArrayGetUBound($pSafeArray, 1, $bound) Then
				__Au3Obj_SafeArrayGetLBound($pSafeArray, 1, $lbound)
				$bound += 1 - $lbound
				If 0 = __Au3Obj_SafeArrayAccessData($pSafeArray, $pData) Then
					Local $tData = DllStructCreate("byte[" & $bound & "]", $pData)
					$data = DllStructGetData($tData, 1)
					__Au3Obj_SafeArrayUnaccessData($pSafeArray)
				EndIf
			EndIf
			Return $data
		Case BitOR($__Au3Obj_VT_ARRAY, $__Au3Obj_VT_VARIANT)
			Return __Au3Obj_ReadSafeArrayVariant(DllStructGetData($var, "data"))
		Case $__Au3Obj_VT_DISPATCH
			Return _AutoItObject_PtrToIDispatch(DllStructGetData($var, "data"))
		Case $__Au3Obj_VT_PTR
			Return DllStructGetData($var, "data")
		Case $__Au3Obj_VT_ERROR
			Return Default
		Case Else
			_AutoItObject_VariantClear($pVariant)
			Return SetError(1, 0, '')
	EndSwitch

	$data = DllStructCreate($type, DllStructGetPtr($var, "data"))

	Switch $VT
		Case $__Au3Obj_VT_BOOL
			Return DllStructGetData($data, 1) <> 0
	EndSwitch
	Return DllStructGetData($data, 1)

EndFunc   ;==>_AutoItObject_VariantRead

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_VariantSet
; Description ...: sets the value of a varaint or creates a new one.
; Syntax.........: _AutoItObject_VariantSet($pVar, $vVal, $iSpecialType = 0)
; Parameters ....: $pVar        - Pointer to the VARIANT to modify (0 if you want to create it new)
;                  $vVal        - Value of the VARIANT
;                  $iSpecialType - [optional] Modify the automatic type. NOT FOR GENERAL USE!
; Return values .: Success      - Pointer to the VARIANT
;                  Failure      - 0
; Author ........: monoceres, Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_VariantRead
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_VariantSet($pVar, $vVal, $iSpecialType = 0)
	; Author: monoceres, Prog@ndy
	If Not $pVar Then
		$pVar = __Au3Obj_CoTaskMemAlloc($__Au3Obj_VARIANT_SIZE)
		_AutoItObject_VariantInit($pVar)
	Else
		_AutoItObject_VariantClear($pVar)
	EndIf
	Local $tVar = DllStructCreate($__Au3Obj_tagVARIANT, $pVar)
	Local $iType = $__Au3Obj_VT_EMPTY, $vDataType = ''

	Switch VarGetType($vVal)
		Case "Int32"
			$iType = $__Au3Obj_VT_I4
			$vDataType = 'int'
		Case "Int64"
			$iType = $__Au3Obj_VT_I8
			$vDataType = 'int64'
		Case "String", 'Text'
			$iType = $__Au3Obj_VT_BSTR
			$vDataType = 'ptr'
			$vVal = __Au3Obj_SysAllocString($vVal)
		Case "Double"
			$vDataType = 'double'
			$iType = $__Au3Obj_VT_R8
		Case "Float"
			$vDataType = 'float'
			$iType = $__Au3Obj_VT_R4
		Case "Bool"
			$vDataType = 'short'
			$iType = $__Au3Obj_VT_BOOL
			If $vVal Then
				$vVal = 0xffff
			Else
				$vVal = 0
			EndIf
		Case 'Ptr'
			If $__Au3Obj_X64 Then
				$iType = $__Au3Obj_VT_UI8
			Else
				$iType = $__Au3Obj_VT_UI4
			EndIf
			$vDataType = 'ptr'
		Case 'Object'
			_AutoItObject_IUnknownAddRef($vVal)
			$vDataType = 'ptr'
			$iType = $__Au3Obj_VT_DISPATCH
		Case "Binary"
			; ARRAY OF BYTES !
			Local $tSafeArrayBound = DllStructCreate($__Au3Obj_tagSAFEARRAYBOUND)
			DllStructSetData($tSafeArrayBound, 1, BinaryLen($vVal))
			Local $pSafeArray = __Au3Obj_SafeArrayCreate($__Au3Obj_VT_UI1, 1, DllStructGetPtr($tSafeArrayBound))
			Local $pData
			If 0 = __Au3Obj_SafeArrayAccessData($pSafeArray, $pData) Then
				Local $tData = DllStructCreate("byte[" & BinaryLen($vVal) & "]", $pData)
				DllStructSetData($tData, 1, $vVal)
				__Au3Obj_SafeArrayUnaccessData($pSafeArray)
				$vVal = $pSafeArray
				$vDataType = 'ptr'
				$iType = BitOR($__Au3Obj_VT_ARRAY, $__Au3Obj_VT_UI1)
			EndIf
		Case "Array"
			$vDataType = 'ptr'
			$vVal = __Au3Obj_CreateSafeArrayVariant($vVal)
			$iType = BitOR($__Au3Obj_VT_ARRAY, $__Au3Obj_VT_VARIANT)
		Case Else ;"Keyword" ; all keywords and unknown Vartypes will be handled as "default"
			$iType = $__Au3Obj_VT_ERROR
			$vDataType = 'int'
	EndSwitch
	If $vDataType Then
		DllStructSetData(DllStructCreate($vDataType, DllStructGetPtr($tVar, 'data')), 1, $vVal)

		If @NumParams = 3 Then $iType = $iSpecialType
		DllStructSetData($tVar, 'vt', $iType)
	EndIf
	Return $pVar
EndFunc   ;==>_AutoItObject_VariantSet

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_WrapperAddMethod
; Description ...: Adds additional methods to the Wrapper-Object, e.g if you want alternative parameter types
; Syntax.........: _AutoItObject_WrapperAddMethod(ByRef $oWrapper, $sReturnType, $sName, $sParamTypes, $ivtableIndex)
; Parameters ....: $oWrapper     - The Object you want to modify
;                  $sReturnType  - the return type of the function
;                  $sName        - The name of the function
;                  $sParamTypes  - the parameter types
;                  $ivTableIndex - Index of the function in the object's vTable
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: _AutoItObject_WrapperCreate
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_WrapperAddMethod(ByRef $oWrapper, $sReturnType, $sName, $sParamTypes, $ivtableIndex)
	; Author: Prog@ndy
	If Not IsObj($oWrapper) Then Return SetError(2, 0, 0)
	DllCall($ghAutoItObjectDLL, "none", "WrapperAddMethod", 'idispatch', $oWrapper, 'wstr', $sName, "wstr", StringRegExpReplace($sReturnType & ';' & $sParamTypes, "\s|(;+\Z)", ''), 'dword', $ivtableIndex)
	If @error Then Return SetError(1, @error, 0)
	Return True
EndFunc   ;==>_AutoItObject_WrapperAddMethod

; #FUNCTION# ====================================================================================================================
; Name...........: _AutoItObject_WrapperCreate
; Description ...: Creates an IDispatch-Object for COM-Interfaces normally not supporting it.
; Syntax.........: _AutoItObject_WrapperCreate($pUnknown, $tagInterface [, $fNoUnknown = False [, $fCallFree = False]])
; Parameters ....: $pUnknown     - Pointer to an IUnknown-Interface not supporting IDispatch
;                  $tagInterface - String defining the methods of the Interface, see Remarks for details
;                  $fNoUnknown   - [optional] $pUnknown is NOT an IUnkown-Interface. Do not release when out of scope (Default: False)
;                  $fCallFree   - [optional] Internal parameter. Do not use.
; Return values .: Success      - Dispatch-Object
;                  Failure      - 0, @error set
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......: $tagInterface can be a string in the following format (dtag):
;                  +  "FunctionName ReturnType(ParamType1;ParamType2);FunctionName2 ..."
;                  +    - FunctionName is the name of the function you want to call later
;                  +    - ReturnType is the return type (like DLLCall)
;                  +    - ParamType is the type of the parameter (like DLLCall) [do not include the THIS-param]
;                  +
;                  +Alternative Format where only method names are listed (ltag) results in different format for calling the functions/methods later. You must specify the datatypes in the call then:
;                  +  $oObject.function("returntype", "1stparamtype", $1stparam, "2ndparamtype", $2ndparam, ...)
;                  +
;                  +The reuturn value of a call is always an array (except an error occured, then it's 0):
;                  +  - $array[0] - containts the return value
;                  +  - $array[n] - containts the n-th parameter
; Related .......: _AutoItObject_WrapperAddMethod
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AutoItObject_WrapperCreate($pUnknown, $tagInterface, $fNoUnknown = False, $fCallFree = False)
	If Not $pUnknown Then Return SetError(1, 0, 0)
	Local $sMethods = __Au3Obj_GetMethods($tagInterface)
	Local $aResult
	If $sMethods Then
		$aResult = DllCall($ghAutoItObjectDLL, "idispatch", "CreateWrapperObjectEx", 'ptr', $pUnknown, 'wstr', $sMethods, "bool", $fNoUnknown, "bool", $fCallFree)
	Else
		$aResult = DllCall($ghAutoItObjectDLL, "idispatch", "CreateWrapperObject", 'ptr', $pUnknown, "bool", $fNoUnknown)
	EndIf
	If @error Then Return SetError(2, @error, 0)
	Return $aResult[0]
EndFunc   ;==>_AutoItObject_WrapperCreate

#endregion Public UDFs
;--------------------------------------------------------------------------------------------------------------------------------------