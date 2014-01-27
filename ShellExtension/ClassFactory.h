
#pragma once

#include "stdafx.h"

extern HINSTANCE g_hInst;
extern UINT g_dllRefCnt;

class CClassFactory : public IClassFactory
{
protected:
	ULONG		m_objRefCnt;
public:
	CClassFactory();
	virtual ~CClassFactory();

	STDMETHODIMP QueryInterface(REFIID, LPVOID *);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP CreateInstance(LPUNKNOWN, REFIID, LPVOID *);
	STDMETHODIMP LockServer(BOOL) { return E_NOTIMPL; }
};