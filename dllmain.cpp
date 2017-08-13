// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"

void Too_Late_To_Destruct();

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:    break;
	case DLL_THREAD_ATTACH:     break;
	case DLL_THREAD_DETACH:     break;
	case DLL_PROCESS_DETACH:    Too_Late_To_Destruct();
		break;
	}
	return TRUE;
}

/*
#include <tchar.h>                          //
#include <OleAuto.h>                        //
TCHAR ModuleFileName[512] = _T("");         //

// DllMain “à‚Å‚±‚ê‚ð‚â‚é
::GetModuleFileName(hModule, ModuleFileName, 512);

VARIANT __stdcall modulePath()      //
{
    VARIANT ret;
    ::VariantInit(&ret);
    ret.vt = VT_BSTR;
    ret.bstrVal = SysAllocString(ModuleFileName);
    return ret;
}
*/