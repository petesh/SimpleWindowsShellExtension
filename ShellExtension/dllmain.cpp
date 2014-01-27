// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "ClassFactory.h"
#include <VersionHelpers.h>
#include <objbase.h>

static HINSTANCE g_hInst;
UINT g_dllRefCnt;

#pragma data_seg(".text")
#define INITGUID
#include <initguid.h>
#include <shlguid.h>
#include "comguid.h"
#pragma data_seg()

#define SHELL_EXT_NAME (TEXT("UserShellExtSizeLimit"))
#define SHELL_EXT_NAME_PERIOD (TEXT("UserShellExtSizeLimit."))

BOOL APIENTRY
DllMain( HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		g_hInst = (HINSTANCE)hModule;
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

STDAPI
DllCanUnloadNow()
{
	return (g_dllRefCnt ? S_FALSE : S_OK);
}

STDAPI
DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppReturn)
{
	*ppReturn = 0;
	if (!IsEqualCLSID(rclsid, CLSID_UserShellExtSizeLimit))
		return CLASS_E_CLASSNOTAVAILABLE;

	CClassFactory *fac = new CClassFactory();
	if (fac == 0)
		return E_OUTOFMEMORY;

	HRESULT result = fac->QueryInterface(riid, ppReturn);
	fac->Release();
	return result;
}

typedef struct {
	HKEY hRootKey;
	LPTSTR lpszSubKey;
	LPTSTR lpszValueName;
	LPTSTR lpszData;
} REGSTRUCT, *LPREGSTRUCT;

STDAPI
DllRegisterServer(void)
{
	OutputDebugString(TEXT("Registering Server"));
	int i;
	HKEY hKey;
	LRESULT lResult;
	DWORD dwDisp;
	TCHAR szSubKey[MAX_PATH];
	TCHAR szCLSID[MAX_PATH];
	TCHAR szModule[MAX_PATH];
	LPWSTR pwsz;

	StringFromIID(CLSID_UserShellExtSizeLimit, &pwsz);
	if (pwsz)
	{
		LPMALLOC pMalloc = 0;
		_tcscpy_s(szCLSID, ARRAYSIZE(szCLSID), pwsz);
		CoGetMalloc(1, &pMalloc);
		if (pMalloc)
		{
			pMalloc->Free(pwsz);
			pMalloc->Release();
		}
	}
	GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule));
	REGSTRUCT ClsIdEntries[] = {
		HKEY_CLASSES_ROOT, TEXT("CLSID\\%s"), NULL, SHELL_EXT_NAME,
		HKEY_CLASSES_ROOT, TEXT("CLSID\\%s"), TEXT("InfoTip"), SHELL_EXT_NAME_PERIOD,
		HKEY_CLASSES_ROOT, TEXT("CLSID\\%s\\InprocServer32"), NULL, TEXT("%s"),
		HKEY_CLASSES_ROOT, TEXT("CLSID\\%s\\InprocServer32"), TEXT("ThreadingModel"), TEXT("Apartment"),
		NULL, NULL, NULL, NULL };

	for (i = 0; ClsIdEntries[i].hRootKey; i++)
	{
		wsprintf(szSubKey, ClsIdEntries[i].lpszSubKey, szCLSID);
		lResult = RegCreateKeyEx(ClsIdEntries[i].hRootKey, szSubKey,
			0,
			NULL,
			REG_OPTION_NON_VOLATILE,
			KEY_WRITE,
			NULL,
			&hKey,
			&dwDisp);
		if (lResult == NOERROR) {
			TCHAR szData[MAX_PATH];
			wsprintf(szData, ClsIdEntries[i].lpszData, szModule);
			lResult = RegSetValueEx(hKey,
				ClsIdEntries[i].lpszValueName,
				0,
				REG_SZ,
				(LPBYTE)szData,
				(lstrlen(szData) + 1) * sizeof(TCHAR));
			RegCloseKey(hKey);
		} else {
			return SELFREG_E_CLASS;
		}
	}

	lstrcpy(szSubKey, TEXT("*\\ShellEx\\ContextMenuHandlers\\UserShellExtSizeLimit"));
	lResult = RegCreateKeyEx(HKEY_CLASSES_ROOT,
		szSubKey,
		0,
		NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE,
		NULL,
		&hKey,
		&dwDisp);

	if (lResult == NOERROR)
	{
		TCHAR szData[MAX_PATH];
		lstrcpy(szData, szCLSID);
		lResult = RegSetValueEx(hKey,
			NULL,
			0,
			REG_SZ,
			(LPBYTE)szData,
			(lstrlen(szData) + 1) * sizeof(TCHAR));

		RegCloseKey(hKey);
	}
	else
		return SELFREG_E_CLASS;

	lstrcpy(szSubKey, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"));
	lResult = RegCreateKeyEx(HKEY_LOCAL_MACHINE,
		szSubKey,
		0,
		NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE,
		NULL,
		&hKey,
		&dwDisp);

	if (lResult == NOERROR)
	{
		TCHAR szData[MAX_PATH];
		lstrcpy(szData, SHELL_EXT_NAME);
		lResult = RegSetValueEx(hKey,
			szCLSID,
			0,
			REG_SZ,
			(LPBYTE)szData,
			(lstrlen(szData) + 1) * sizeof(TCHAR));

		RegCloseKey(hKey);
	}
	else
		return SELFREG_E_CLASS;

	return S_OK;
}

STDAPI
DllUnregisterServer(VOID)
{
	OutputDebugString(TEXT("Unregistering server"));
	INT i;
	LRESULT lResult;
	TCHAR szSubKey[MAX_PATH];
	TCHAR szCLSID[MAX_PATH];
	LPWSTR pwsz;

	// get the CLSID in string form
	StringFromIID(CLSID_UserShellExtSizeLimit, &pwsz);

	if (pwsz)
	{
		_tcscpy_s(szCLSID, ARRAYSIZE(szCLSID), pwsz);
		LPMALLOC pMalloc;
		CoGetMalloc(1, &pMalloc);
		if (pMalloc)
		{
			pMalloc->Free(pwsz);
			pMalloc->Release();
		}
	}

	// CLSID entries
	REGSTRUCT ClsidEntries[] = {
		HKEY_CLASSES_ROOT, TEXT("CLSID\\%s\\InprocServer32"), NULL, NULL,
		HKEY_CLASSES_ROOT, TEXT("CLSID\\%s"), NULL, NULL,
		NULL, NULL, NULL, NULL };

	for (i = 0; ClsidEntries[i].hRootKey; i++)
	{
		// create the sub key string.
		wsprintf(szSubKey, ClsidEntries[i].lpszSubKey, szCLSID);
		lResult = RegDeleteKey(ClsidEntries[i].hRootKey, szSubKey);
	}

	// Delete the Context Menu & approved key
	lstrcpy(szSubKey, TEXT("*\\ShellEx\\ContextMenuHandlers\\UserShellExtSizeLimit"));
	lResult = RegDeleteKey(HKEY_CLASSES_ROOT, szSubKey);

	lstrcpy(szSubKey, TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"));
	HKEY hTmpKey;
	if (RegOpenKey(HKEY_LOCAL_MACHINE, szSubKey, &hTmpKey) == ERROR_SUCCESS) {
		RegDeleteValue(hTmpKey, szCLSID);
		RegCloseKey(hTmpKey);
	}
	return S_OK;
}
