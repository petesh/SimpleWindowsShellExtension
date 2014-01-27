
#pragma once

DWORD GetKeyName(HKEY hKeyRoot, LPCTSTR pszSubKey, DWORD dwIndex, LPTSTR pszOut, DWORD dwSize);
BOOL GetValueName(HKEY hKeyRoot, LPCTSTR pszSubKey, DWORD dwIndex, LPTSTR pszOut, DWORD dwSize);
BOOL SaveGlobalSettings(VOID);
BOOL GetGlobalSettings(VOID);
INT WideCharToLocal(LPTSTR pLocal, LPWSTR pWide, DWORD dwChars);
INT LocalToWideChar(LPWSTR pWide, LPTSTR pLocal, DWORD dwChars);

#define SHELL_EXT_NAME (TEXT("UserShellExtSizeLimit"))
#define SHELL_EXT_NAME_PERIOD (TEXT("UserShellExtSizeLimit."))
