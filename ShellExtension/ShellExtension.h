
#pragma once
#include "stdafx.h"

class CShellExtension : public IShellExtInit, IContextMenu
{
protected:
	ULONG m_refCnt;
public:
	CShellExtension();
	virtual ~CShellExtension();

	// IUnknown
	STDMETHOD(QueryInterface) (REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef) (void);
	STDMETHOD_(ULONG, Release) (void);

	// IShellExtInit
	STDMETHODIMP Initialize(LPCITEMIDLIST, LPDATAOBJECT, HKEY);

	// IContextMenu
	STDMETHODIMP GetCommandString(UINT_PTR, UINT, UINT *, LPSTR, UINT);
	STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO);
	STDMETHODIMP QueryContextMenu(HMENU, UINT, UINT, UINT, UINT);

private:
	TCHAR m_szFile[MAX_PATH];
	int m_numFiles;
};
