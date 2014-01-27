/******************************************************************
*
*  Project.....:  Windows View (Namespace Extension)
*
*  Application.:  WINVIEW.dll
*  Module......:  Utility.cpp
*  Description.:  Helper functions (implementation)
*
*  Compiler....:  MS Visual C++
*  Written by..:  D. Esposito
*
*  Environment.:  Windows 9x/NT
*
*******************************/

#include "stdafx.h"

/*---------------------------------------------------------------*/
//                        INCLUDE section
/*---------------------------------------------------------------*/
#include "Utility.h"

extern HIMAGELIST g_himlSmall;

#define MAIN_KEY_STRING        (TEXT("Software\\UserShellExtSizeLimit"))
#define VALUE_STRING           (TEXT("Display Settings"))

#define INITIAL_COLUMN1_SIZE   100	// state
#define INITIAL_COLUMN2_SIZE   100	// class
#define INITIAL_COLUMN3_SIZE   60	// hwnd
#define INITIAL_COLUMN4_SIZE   100	// title



/*---------------------------------------------------------------*/
// Registry utility
/*---------------------------------------------------------------*/

// GetKeyName
DWORD GetKeyName(HKEY hKeyRoot, LPCTSTR pszSubKey, DWORD dwIndex,
	LPTSTR pszOut, DWORD dwSize)
{
	HKEY hKey;
	LONG lResult;
	FILETIME ft;

	if (!pszOut)
		return 0;

	if (!pszSubKey)
		pszSubKey = TEXT("");

	// open the specified key
	lResult = RegOpenKeyEx(hKeyRoot, pszSubKey, 0,
		KEY_ENUMERATE_SUB_KEYS, &hKey);

	if (lResult != ERROR_SUCCESS)
		return 0;

	// get the specified subkey
	lResult = RegEnumKeyEx(hKey, dwIndex, pszOut, &dwSize,
		NULL, NULL, NULL, &ft);
	RegCloseKey(hKey);

	if (lResult == ERROR_SUCCESS)
		return dwSize;

	return 0;
}


// GetValueName
BOOL GetValueName(HKEY hKeyRoot, LPCTSTR pszSubKey, DWORD dwIndex,
	LPTSTR pszOut, DWORD dwSize)
{
	HKEY hKey;
	LONG lResult;
	DWORD dwType;

	if (!pszOut)
		return 0;

	if (!pszSubKey)
		pszSubKey = TEXT("");

	// open the specified key
	lResult = RegOpenKeyEx(hKeyRoot, pszSubKey, 0,
		KEY_QUERY_VALUE, &hKey);

	if (lResult != ERROR_SUCCESS)
		return 0;

	// get the specified subkey
	lResult = RegEnumValue(hKey, dwIndex, pszOut, &dwSize, NULL,
		&dwType, NULL, NULL);

	RegCloseKey(hKey);

	return (lResult == ERROR_SUCCESS);
}



// SaveGlobalSettings
BOOL SaveGlobalSettings(VOID)
{
	HKEY hKey;
	LONG lResult;
	DWORD dwDisp;

	lResult = RegCreateKeyEx(HKEY_CURRENT_USER,
		MAIN_KEY_STRING,
		0,
		NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_ALL_ACCESS,
		NULL,
		&hKey,
		&dwDisp);

	if (lResult != ERROR_SUCCESS)
		return FALSE;

	// create an array to put our data in
	DWORD dwArray[5];
	//dwArray[0] = g_nColumn1;
	//dwArray[1] = g_nColumn2;
	//dwArray[2] = g_nColumn3;
	//dwArray[3] = g_nColumn4;
	dwArray[4] = 0;

	lResult = RegSetValueEx(hKey, VALUE_STRING, 0,
		REG_BINARY, (LPBYTE)dwArray, sizeof(dwArray));

	RegCloseKey(hKey);

	if (lResult != ERROR_SUCCESS)
		return FALSE;

	return TRUE;
}


// GetGlobalSettings
BOOL GetGlobalSettings(VOID)
{
	HKEY hKey;
	LONG lResult;

	// set up the default data
	//g_nColumn1 = INITIAL_COLUMN1_SIZE;
	//g_nColumn2 = INITIAL_COLUMN2_SIZE;
	//g_nColumn3 = INITIAL_COLUMN3_SIZE;
	//g_nColumn4 = INITIAL_COLUMN4_SIZE;

	lResult = RegOpenKeyEx(HKEY_CURRENT_USER,
		MAIN_KEY_STRING,
		0,
		KEY_ALL_ACCESS,
		&hKey);

	if (lResult != ERROR_SUCCESS)
		return FALSE;

	// create an array to put our data in
	DWORD dwArray[5];
	DWORD dwType;
	DWORD dwSize = sizeof(dwArray);

	// retrieve settings
	lResult = RegQueryValueEx(hKey,
		VALUE_STRING,
		NULL,
		&dwType,
		(LPBYTE)dwArray,
		&dwSize);

	RegCloseKey(hKey);

	if (lResult != ERROR_SUCCESS)
		return FALSE;

	//g_nColumn1 = dwArray[0];
	//g_nColumn2 = dwArray[1];
	//g_nColumn3 = dwArray[2];
	//g_nColumn4 = dwArray[3];
	//g_bShowIDW = dwArray[4];

	return TRUE;
}



// CompareItems (sort listview by PIDL/HWND)
int CALLBACK CompareItems(LPARAM lParam1, LPARAM lParam2, LPARAM lpData)
{
	//CShellFolder* pFolder = reinterpret_cast<CShellFolder*>(lpData);
	//if(!pFolder)
	//   return 0;

	//return pFolder->CompareIDs(0, reinterpret_cast<LPITEMIDLIST>(lParam1),
	//                              reinterpret_cast<LPITEMIDLIST>(lParam2));
	return 0;
}


// create imagelists for the listview
VOID CreateImageLists(VOID)
{
	//if( g_himlSmall )
	//   ImageList_Destroy( g_himlSmall );

	//// set the small image list
	//g_himlSmall = ImageList_Create( 16, 16, ILC_COLORDDB|ILC_MASK, 4, 0 );
	//if( g_himlSmall )
	//{
	//	HICON hIcon;
	//    hIcon = (HICON) LoadImage( g_hInst, 
	//		MAKEINTRESOURCE(IDI_WINDOW), IMAGE_ICON, 
	//		16, 16, LR_DEFAULTCOLOR );
	//   ImageList_AddIcon( g_himlSmall, hIcon );

	//   hIcon = (HICON) LoadImage( g_hInst, 
	//	   MAKEINTRESOURCE(IDI_PARWND), IMAGE_ICON, 
	//	   16, 16, LR_DEFAULTCOLOR );
	//   ImageList_AddIcon( g_himlSmall, hIcon );

	//   hIcon = (HICON) LoadImage( g_hInst, 
	//	   MAKEINTRESOURCE(IDI_WND), IMAGE_ICON, 
	//	   16, 16, LR_DEFAULTCOLOR );
	//   ImageList_AddIcon( g_himlSmall, hIcon );
	//
	//}
}

// DestroyImageLists
VOID DestroyImageLists(VOID)
{
	//if( g_himlSmall )
	//	ImageList_Destroy( g_himlSmall );
}


// WideCharToLocal
INT WideCharToLocal(LPTSTR pLocal, LPWSTR pWide, DWORD dwChars)
{
	*pLocal = NULL;
	size_t nWideLength = 0;
#ifndef UNICODE
	wchar_t *pwszSubstring;
#endif

	nWideLength = wcslen(pWide);

#ifdef UNICODE
	if (nWideLength < dwChars)
	{
		wcsncpy_s(pLocal, MAX_PATH, pWide, dwChars);
	}
	else
	{
		wcsncpy_s(pLocal, MAX_PATH, pWide, dwChars - 1);
		pLocal[dwChars - 1] = NULL;
	}
#else
	if (nWideLength < dwChars)
	{
		WideCharToMultiByte(CP_ACP,
			0,
			pWide,
			-1,
			pLocal,
			dwChars,
			NULL,
			NULL);
	}
	else
	{
		pwszSubstring = new WCHAR[dwChars];
		wcsncpy_s(pwszSubstring, pWide, dwChars - 1);
		pwszSubstring[dwChars - 1] = NULL;

		WideCharToMultiByte(CP_ACP,
			0,
			pwszSubstring,
			-1,
			pLocal,
			dwChars,
			NULL,
			NULL);

		delete[] pwszSubstring;
	}
#endif

	return lstrlen(pLocal);
}



// LocalToWideChar
INT LocalToWideChar(LPWSTR pWide, LPTSTR pLocal, DWORD dwChars)
{
	*pWide = 0;
#ifdef UNICODE
	size_t nLocalLength = wcslen(pLocal);
	if (nLocalLength < dwChars)
	{
		wcsncpy_s(pWide, MAX_PATH, pLocal, dwChars);
	} else {
		wcsncpy_s(pWide, MAX_PATH, pLocal, dwChars - 1);
		pWide[dwChars - 1] = 0;
	}
#else
	MultiByteToWideChar(CP_ACP, 0, pLocal, -1, pWide, dwChars);
#endif
	return lstrlenW(pWide);
}

