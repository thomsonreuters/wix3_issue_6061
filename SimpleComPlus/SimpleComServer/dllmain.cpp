// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"

HMODULE g_thisModule;

BOOL APIENTRY DllMain( HMODULE hModule,
					   DWORD  ul_reason_for_call,
					   LPVOID lpReserved
					 )
{
	OutputDebugStringW(L"In DllMain()");
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		g_thisModule = hModule;
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

