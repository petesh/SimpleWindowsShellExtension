
#include "stdafx.h"
#include "ClassFactory.h"
#include "ShellExtension.h"

CClassFactory::CClassFactory()
{
	m_objRefCnt = 1;
	InterlockedIncrement(&g_dllRefCnt);
}

CClassFactory::~CClassFactory()
{
	InterlockedDecrement(&g_dllRefCnt);
}

STDMETHODIMP
CClassFactory::QueryInterface(REFIID riid, LPVOID *ppReturn)
{
	*ppReturn = 0;

	if (IsEqualIID(riid, IID_IUnknown)) {
		*ppReturn = this;
	} else {
		if (IsEqualIID(riid, IID_IClassFactory))
			*ppReturn = (IClassFactory *)this;
	}

	if (*ppReturn) {
		LPUNKNOWN pUnk = (LPUNKNOWN)(*ppReturn);
		pUnk->AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

STDMETHODIMP
CClassFactory::CreateInstance(LPUNKNOWN pUnknown, REFIID riid, LPVOID *ppObject)
{
	*ppObject = 0;
	if (pUnknown != NULL)
		return CLASS_E_NOAGGREGATION;

	CShellExtension *pExt = new CShellExtension();
	if (pExt == 0)
		return E_OUTOFMEMORY;

	HRESULT hResult = pExt->QueryInterface(riid, ppObject);
	pExt->Release();
	return hResult;
}

STDMETHODIMP_(ULONG)
CClassFactory::AddRef()
{
	return InterlockedIncrement(&m_objRefCnt);
}

STDMETHODIMP_(ULONG)
CClassFactory::Release()
{
	ULONG dwRefCnt = InterlockedDecrement(&m_objRefCnt);
	if (dwRefCnt == 0)
		delete this;
	return dwRefCnt;
}

