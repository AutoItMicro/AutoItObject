;#AutoIt3Wrapper_UseX64=n ; if you are on x64 and want x86
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6

#include "AutoItObject.au3"
#include <Array.au3> ; for displaying array

; Error monitoring
Global $oError = ObjEvent("AutoIt.Error", "_ErrFunc")
Func _ErrFunc()
	ConsoleWrite("! COM Error !  Number: 0x" & Hex($oError.number, 8) & "   ScriptLine: " & $oError.scriptline & " - " & $oError.windescription & @CRLF)
	Return
EndFunc   ;==>_ErrFunc

; Start AutoItObject
_AutoItObject_Startup()


; SYSTEM (kernel) is:
ConsoleWrite("!SYSTEM = " & _SystemModuleInformation() & @CRLF)

; See all modules
Global $aArrayOfSystemModuled = _SystemModuleInformation(False)

_ArrayDisplay($aArrayOfSystemModuled, "SystemModuleInformation")




; FUNCTION:

Func _SystemModuleInformation($fGetSYSTEM = True)
	; First call is to NtQuerySystemInformation is to determine required size of the buffer (SYSTEM_INFORMATION_CLASS parameter is set to SystemModuleInformation)
	Local $aCall = DllCall("ntdll.dll", "long", "NtQuerySystemInformation", _
			"dword", 11, _ ; SystemModuleInformation
			"ptr", 0, _
			"dword", 0, _
			"dword*", 0)
	; Check for possible error (on AutoIt side only)
	If @error Then Return SetError(1, 0, "")
	Local $iSize = $aCall[4]
	; Make raw buffer to collect to
	Local $oBufferRaw = _AutoItObject_DllStructCreate("byte[" & $iSize & "]")
	; Fill the buffer
	$aCall = DllCall("ntdll.dll", "long", "NtQuerySystemInformation", _
			"dword", 11, _ ; SystemModuleInformation
			"ptr", $oBufferRaw(), _
			"dword", $iSize, _
			"dword*", 0)
	If @error Then Return SetError(2, 0, "")
	; Address of SYSTEM_MODULE_INFORMATION structure
	Local $pPointer = $oBufferRaw()
	; Some definitions
	#cs
		SYSTEM_MODULE_INFORMATION structure is defined as:

		typedef struct _SYSTEM_MODULE_INFORMATION {
		DWORD_PTR            ModulesCount;
		SYSTEM_MODULE        Modules[0]; // array of SYSTEM_MODULE structures
		} SYSTEM_MODULE_INFORMATION, *PSYSTEM_MODULE_INFORMATION;

		And SYSTEM_MODULE structure is (free interptetation):

		typedef struct _SYSTEM_MODULE {
		DWORD_PTR            Reserved[2];
		PVOID                ImageBaseAddress;
		DWORD                ImageSize;
		DWORD                Flags;
		WORD                 Index;
		WORD                 Unknown;  // seems to be set for modules accesible from user mode
		WORD                 LoadCount;
		WORD                 ModuleNameOffset;
		BYTE                 ImageName[MAXIMUM_FILENAME_LENGTH];
		} SYSTEM_MODULE, *PSYSTEM_MODULE;
	#ce
	; I can write now:
	Local $oCount = _AutoItObject_DllStructCreate("dword_ptr ModulesCount", $pPointer)
	; ...and move pointer to first SYSTEM_MODULE structure
	$pPointer += $oCount.__size__
	; Collect data (overall number of system modules):
	Local $iCount = $oCount.ModulesCount
	; Dimensioning array to fill
	Local $aArray[$iCount + 1][7] = [["Index", "ImageName", "ImageBaseAddress", "ImageSize (bytes)", "LoadCount", "Location", "Flags"]]
	; Needed variables
	Local $oSYSTEM_MODULE, $iIndex, $iNameOffset, $ImageName
	; Loop through all SYSTEM_MODULE structures and fill the array
	For $i = 1 To $iCount
		$oSYSTEM_MODULE = _AutoItObject_DllStructCreate("dword_ptr Reserved[2];" & _
				"ptr ImageBaseAddress;" & _
				"dword ImageSize;" & _
				"dword Flags;" & _
				"word Index;" & _
				"word Unknown;" & _
				"word LoadCount;" & _
				"word ModuleNameOffset;" & _
				"char ImageName[256]", _
				$pPointer)
		$iIndex = $oSYSTEM_MODULE.Index
		$aArray[$i][0] = $iIndex
		$iNameOffset = $oSYSTEM_MODULE.ModuleNameOffset
		$ImageName = $oSYSTEM_MODULE.ImageName
		$aArray[$i][5] = StringLeft($ImageName, $iNameOffset)
		$ImageName = StringTrimLeft($ImageName, $iNameOffset)
		; If kernel name is wanted, return it
		If $iIndex = 0 And $fGetSYSTEM Then Return $ImageName
		$aArray[$i][1] = $ImageName
		$aArray[$i][2] = Ptr($oSYSTEM_MODULE.ImageBaseAddress)
		$aArray[$i][3] = $oSYSTEM_MODULE.ImageSize
		$aArray[$i][4] = $oSYSTEM_MODULE.LoadCount
		$aArray[$i][6] = "0x" & Hex($oSYSTEM_MODULE.Flags, 8) ; in this form just for better 'visualisation'
		; Move pointer to next structure
		$pPointer += $oSYSTEM_MODULE.__size__
	Next
	; If array is wanted return it
	Return $aArray
EndFunc   ;==>_SystemModuleInformation
