/*
  Copyright (c) 2022 Makoto Tanabe <mtanabe.sj@outlook.com>
  Licensed under the MIT License.

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.
*/
#pragma once
#include <ocidl.h>
#include <objsafe.h>


//#define SUPPORT_FREE_THREADING_COM

#ifndef LIB_ADDREF
// no explicit reference counting is performed to keep the app alive.
// if that's needed, re-define LIB_ADDREF and LIB_RELEASE in stdafx.h.
#define LIB_ADDREF
#define LIB_RELEASE
#endif//#ifndef LIB_ADDREF

//#define NO_VTABLE __declspec(novtable)
#define NO_VTABLE

/////////////////////////////////////////////////////////////////

#define DELEGATE_IUNKNOWN_TO_IDISPATCHIMPL(iobj, piid, plibid) \
STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppvObj) { return IDispatchImpl<iobj, piid, plibid>::QueryInterface(riid, ppvObj); } \
STDMETHOD_(ULONG, AddRef)() { return IDispatchImpl<iobj, piid, plibid>::AddRef(); } \
STDMETHOD_(ULONG, Release)() { return IDispatchImpl<iobj, piid, plibid>::Release(); }

#define DELEGATE_IUNKNOWN_TO_IDISPATCHWITHOBJECTSAFETYIMPL(iobj, piid, plibid) \
STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppvObj) { return IDispatchWithObjectSafetyImpl<iobj, piid, plibid>::QueryInterface(riid, ppvObj); } \
STDMETHOD_(ULONG, AddRef)() { return IDispatchImpl<iobj, piid, plibid>::AddRef(); } \
STDMETHOD_(ULONG, Release)() { return IDispatchImpl<iobj, piid, plibid>::Release(); }

#define DELEGATE_IUNKNOWN_REF_TO_IDISPATCHWITHOBJECTSAFETYIMPL(iobj, piid, plibid) \
STDMETHOD_(ULONG, AddRef)() { return IDispatchImpl<iobj, piid, plibid>::AddRef(); } \
STDMETHOD_(ULONG, Release)() { return IDispatchImpl<iobj, piid, plibid>::Release(); }

/////////////////////////////////////////////////////////////////

template <class T>
T* ADDREFASSIGN(T *p) {
	if (p)
		dynamic_cast<IUnknown*>(p)->AddRef();
	return p;
}
template <class T>
void FINALRELEASE(T** p) {
	T* p2 = (T*)InterlockedExchangePointer((LPVOID*)p, NULL);
	if (p2)
		dynamic_cast<IUnknown*>(p2)->Release();
}


//NOTE: make sure you include <combaseapi.h> to avoid a compiler error.
template <class T>
class _ComRef : public T
{
private:
	STDMETHOD_(ULONG, AddRef)() = 0;
	STDMETHOD_(ULONG, Release)() = 0;
};


template <class T>
class AutoComRel {
public:
	AutoComRel(T* p = NULL) : _p(ADDREFASSIGN(p)) {}
	~AutoComRel() { release(); }

	void release() { FINALRELEASE<T>(&_p); }
	BOOL attach(T* p) {
		release();
		if (!p) return FALSE;
		_p = p;
		return TRUE;
	}
	BOOL attach(T** pp) {
		if (!attach(*pp)) return FALSE;
		*pp = NULL;
		return TRUE;
	}
	T* detach() {
		T* p = _p;
		_p = NULL;
		return p;
	}
	operator T*() const { return _p; }
	T** operator &() { return &_p; }
	_ComRef<T>* operator->() const throw() {
		return (_ComRef<T>*)_p;
	}
	T* _p;
};


class _FTComLock
{
public:
	_FTComLock() :
#ifdef SUPPORT_FREE_THREADING_COM
		_locks(0), _mutex(NULL),
#endif//#ifdef SUPPORT_FREE_THREADING_COM
		_ref(1)
	{
		LIB_ADDREF;
#ifdef SUPPORT_FREE_THREADING_COM
		_mutex = ::CreateMutex(0, FALSE, 0);
#endif//#ifdef SUPPORT_FREE_THREADING_COM
	}
	virtual ~_FTComLock()
	{
		ASSERT(_ref == 0);
#ifdef SUPPORT_FREE_THREADING_COM
		HANDLE h = InterlockedExchangePointer(&_mutex, NULL);
		if (h)
			::CloseHandle(h);
#endif//#ifdef SUPPORT_FREE_THREADING_COM
		LIB_RELEASE;
	}

	LONG Lock()
	{
#ifdef SUPPORT_FREE_THREADING_COM
		DWORD res = ::WaitForSingleObject(_mutex, INFINITE);
		ASSERT(res == WAIT_OBJECT_0);
		LONG c = InterlockedIncrement(&_locks);
		return c;
#else//#ifdef SUPPORT_FREE_THREADING_COM
		return 1;
#endif//#ifdef SUPPORT_FREE_THREADING_COM
	}
	LONG Unlock()
	{
#ifdef SUPPORT_FREE_THREADING_COM
		LONG c = InterlockedDecrement(&_locks);
		::ReleaseMutex(_mutex);
		return c;
#else//#ifdef SUPPORT_FREE_THREADING_COM
		return 0;
#endif//#ifdef SUPPORT_FREE_THREADING_COM
	}

protected:
	ULONG _ref;
#ifdef SUPPORT_FREE_THREADING_COM
	LONG _locks;
	HANDLE _mutex;
#endif//#ifdef SUPPORT_FREE_THREADING_COM
};

class _AutoComLock
{
public:
	_AutoComLock(_FTComLock *p) : _p(p)
	{
		_p->Lock();
	}
	~_AutoComLock()
	{
		_p->Unlock();
	}
	_FTComLock *_p;
};

template <class T, const IID* piid>
class IUnknownImpl : public _FTComLock, public T
{
public:
	IUnknownImpl() {}

	// IUnknown methods
	STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppv)
	{
		AddRef();
		if (IsEqualIID(riid, IID_IUnknown))
		{
			*ppv = (LPUNKNOWN)(T*)this;
			return S_OK;
		}
		if (IsEqualIID(riid, *piid))
		{
			*ppv = dynamic_cast<T*>(this);
			return S_OK;
		}
		Release();
		return E_NOINTERFACE;
	}
	STDMETHOD_(ULONG, AddRef)()
	{
		Lock();
		ULONG c = InterlockedIncrement(&_ref);
		Unlock();
		return c;
	}
	STDMETHOD_(ULONG, Release)()
	{
		Lock();
		ULONG c = InterlockedDecrement(&_ref);
		Unlock();
		if (c == 0)
			delete this;
		return c;
	}
};


////////////////////////////////////////////////////////////////////////
// IClassFactory implementations

class NO_VTABLE IClassFactoryImpl : public _FTComLock, public IClassFactory
{
public:
	IClassFactoryImpl()
	{
		LIB_ADDREF;
	}
	virtual ~IClassFactoryImpl()
	{
		LIB_RELEASE;
	}

	// IUnknown methods
	STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppv)
	{
		AddRef();
		if (GetInterface(riid, ppv))
			return NOERROR;
		*ppv = NULL;
		Release();
		return E_NOINTERFACE;
	}
	STDMETHOD_(ULONG, AddRef)()
	{
		Lock();
		ULONG c = InterlockedIncrement(&_ref);
		Unlock();
		return c;
	}
	STDMETHOD_(ULONG, Release)()
	{
		Lock();
		ULONG c = InterlockedDecrement(&_ref);
		Unlock();
		if (c)
			return c;
		delete this;
		return 0;
	}
	// IClassFactory methods
	STDMETHOD(CreateInstance)(LPUNKNOWN punk, REFIID riid, void** ppv)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(LockServer)(BOOL fLock)
	{
		if (fLock) {
			LIB_LOCK;
		} else {
			LIB_UNLOCK;
		}
		return S_OK;
	}

protected:
	virtual BOOL GetInterface(REFIID riid, LPVOID* ppv)
	{
		if (IsEqualIID(riid, IID_IUnknown))
		{
			*ppv = (LPUNKNOWN)this;
			return TRUE;
		}
		else if (IsEqualIID(riid, IID_IClassFactory))
		{
			*ppv = (LPCLASSFACTORY)this;
			return TRUE;
		}
		return FALSE;
	}
};

template <class T>
class NO_VTABLE IClassFactoryNoAggrImpl : public IClassFactoryImpl
{
public:
	IClassFactoryNoAggrImpl() {}

	// IClassFactory methods
	STDMETHOD(CreateInstance)(LPUNKNOWN punk, REFIID riid, void** ppv)
	{
		if (NULL == ppv) return E_POINTER;
		*ppv = NULL;
		if (NULL != punk)
			return CLASS_E_NOAGGREGATION;
		return CreateInstanceInternal(riid, ppv);
	}

protected:
	virtual HRESULT CreateInstanceInternal(REFIID riid, void** ppv)
	{
		T* pObj = new T;
		HRESULT hr = E_OUTOFMEMORY;
		if (pObj)
		{
			hr = pObj->QueryInterface(riid, ppv);
			pObj->Release();
		}
		return hr;
	}
};

/////////////////////////////////////////////////////////////////

template <class T, const IID* piid, const GUID* plibid, WORD wMajor = 1, WORD wMinor = 0>
class IDispatchImpl : public _FTComLock, public T
{
public:
	IDispatchImpl(LPTYPEINFO pTI = NULL) : _pti(pTI)
	{
		LIB_ADDREF;
	}
	~IDispatchImpl()
	{
		FINALRELEASE(&_pti);
		LIB_RELEASE;
	}

	// IUnknown methods
	STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppv)
	{
		AddRef();
		if (IsEqualIID(riid, IID_IUnknown))
		{
			*ppv = (LPUNKNOWN)(T*)this;
			return S_OK;
		}
		if (IsEqualIID(riid, IID_IDispatch))
		{
			*ppv = (LPDISPATCH)(T*)this;
			return S_OK;
		}
		if (IsEqualIID(riid, *piid))
		{
			*ppv = dynamic_cast<T*>(this);
			return S_OK;
		}
		Release();
		return E_NOINTERFACE;
	}
	STDMETHOD_(ULONG, AddRef)()
	{
		Lock();
		ULONG c = InterlockedIncrement(&_ref);
		Unlock();
		return c;
	}
	STDMETHOD_(ULONG, Release)()
	{
		Lock();
		ULONG c = InterlockedDecrement(&_ref);
		Unlock();
		if (c == 0)
			delete this;
		return c;
	}

	// IDispatch methods
	STDMETHOD(GetIDsOfNames)(REFIID riid, OLECHAR** rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
	{
		if (IID_NULL != riid)
			return DISP_E_UNKNOWNINTERFACE;
		LPTYPEINFO pTI;
		HRESULT hr = GetTypeInfo(0, lcid, &pTI);
		if (SUCCEEDED(hr)) {
			hr = DispGetIDsOfNames(pTI, rgszNames, cNames, rgDispId);
			pTI->Release();
		}
		return hr;
	}

	STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
	{
		if (0 != itinfo)
			return TYPE_E_ELEMENTNOTFOUND;
		if (!pptinfo)
			return E_POINTER;
		Lock();
		if (!_pti) {
			LPTYPELIB pTypeLib;
			HRESULT hr = LoadRegTypeLib(*plibid, wMajor, wMinor, PRIMARYLANGID(lcid), &pTypeLib);
			if (FAILED(hr)) {
				TCHAR szPath[_MAX_PATH];
				GetModuleFileName(LibInstanceHandle, szPath, ARRAYSIZE(szPath));
				hr = LoadTypeLib(szPath, &pTypeLib);
				if (FAILED(hr)) {
					Unlock();
					return hr;
				}
			}
			hr = GetTypeInfoOfGuidInternal(pTypeLib); // set _pti to the typelib
			if (SUCCEEDED(hr))
				hr = GetTypeInfo2(pTypeLib);
			// release pTypeLib. _pti has a reference on it.
			pTypeLib->Release();
			if (FAILED(hr)) {
				Unlock();
				return hr;
			}
		}
		_pti->AddRef();
		*pptinfo = _pti;
		Unlock();
		return NOERROR;
	}

	STDMETHOD(GetTypeInfoCount)(UINT* pctinfo)
	{
		*pctinfo = 1;
		return S_OK;
	}

	STDMETHOD(Invoke)(DISPID dispid, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
	{
		LPTYPEINFO pTI;
		HRESULT hr = GetTypeInfo(0, lcid, &pTI);
		if (FAILED(hr))
			return hr;
		hr = pTI->Invoke((LPDISPATCH)this, dispid, wFlags, pDispParams, pvarResult, pexcepinfo, puArgErr);
		// TODO: consider filling in EXCEPINFO.
		pTI->Release();
		return hr;
	}

protected:
	LPTYPEINFO _pti;

	virtual HRESULT GetTypeInfoOfGuidInternal(LPTYPELIB pTypeLib)
	{
		Lock();
		ASSERT(_pti == NULL);
		HRESULT hr = pTypeLib->GetTypeInfoOfGuid(*piid, &_pti);
		Unlock();
		return hr;
	}
	virtual HRESULT GetTypeInfo2(LPTYPELIB ptl) { return S_OK; }
};


template <class T, const IID* piid, const GUID* plibid, WORD wMajor = 1, WORD wMinor = 0>
class NO_VTABLE IDispatchWithObjectSafetyImpl : public IDispatchImpl<T, piid, plibid, wMajor, wMinor>, public IObjectSafety
{
public:
	IDispatchWithObjectSafetyImpl(LPTYPEINFO pTI = NULL) : IDispatchImpl<T, piid, plibid>(pTI)
	{
		_safetyOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA | INTERFACESAFE_FOR_UNTRUSTED_CALLER;
	}

	// IUnknown methods
	STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppv)
	{
		AddRef();
		if (IsEqualIID(riid, IID_IObjectSafety))
		{
			*ppv = (IObjectSafety*)this;
			return S_OK;
		}
		Release();
		return IDispatchImpl<T, piid, plibid>::QueryInterface(riid, ppv);
	}

	// IObjectSafety methods
	STDMETHOD(GetInterfaceSafetyOptions) (/* [in] */ REFIID riid, /* [out] */ DWORD *pdwSupportedOptions, /* [out] */ DWORD *pdwEnabledOptions)
	{
		if (riid != IID_IDispatch && riid != *piid)
			return E_NOINTERFACE;
		*pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA | INTERFACESAFE_FOR_UNTRUSTED_CALLER;
		*pdwEnabledOptions = _safetyOptions;
		return S_OK;
	}
	/*
	Returns one of the following values:
	S_OK The object is safe for loading.
	E_NOINTERFACE The riid parameter specifies an interface that is unknown to the object.
	E_FAIL The dwOptionSetMask parameter specifies an option that is not supported by the object.
	*/
	STDMETHOD(SetInterfaceSafetyOptions) (/* [in] */ REFIID riid, /* [in] */ DWORD dwOptionSetMask, /* [in] */ DWORD dwEnabledOptions)
	{
		if (riid != IID_IDispatch && riid != *piid)
			return E_NOINTERFACE;
		if (!(dwOptionSetMask & (INTERFACESAFE_FOR_UNTRUSTED_DATA | INTERFACESAFE_FOR_UNTRUSTED_CALLER)))
			return E_FAIL;
		_safetyOptions = (dwOptionSetMask & dwEnabledOptions);
		return S_OK;
	}

protected:
	// IObjectSafety related
	DWORD _safetyOptions;
};


/////////////////////////////////////////////////////////////////

template <class T, const IID* piid>
class NO_VTABLE IDispatchEventSinkImpl : public _FTComLock, public T {
public:
	IDispatchEventSinkImpl() { LIB_ADDREF; }
	virtual ~IDispatchEventSinkImpl() { LIB_RELEASE; }

	// IUnknown methods
	STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppv)
	{
		HRESULT hr = NOERROR;
		AddRef();
		if (GetInterface(riid, ppv))
			return S_OK;
		*ppv = NULL;
		Release();
		return E_NOINTERFACE;
	}
	STDMETHOD_(ULONG, AddRef)() {
		Lock();
		ULONG c = InterlockedIncrement(&_ref);
		Unlock();
		return c;
	}
	STDMETHOD_(ULONG, Release)() {
		Lock();
		ULONG c = InterlockedDecrement(&_ref);
		Unlock();
		if (c == 0)
			delete this;
		return c;
	}

	// IDispatch methods
	STDMETHOD(GetTypeInfoCount) (UINT FAR* pctinfo) { return E_NOTIMPL; }
	STDMETHOD(GetTypeInfo) (UINT itinfo, LCID lcid, ITypeInfo FAR* FAR* pptinfo) { return E_NOTIMPL; }
	STDMETHOD(GetIDsOfNames) (REFIID riid, OLECHAR FAR* FAR* rgszNames, UINT cNames, LCID lcid, DISPID FAR* rgdispid) { return E_NOTIMPL; }
	STDMETHOD(Invoke) (DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pdispparams, VARIANT FAR* pvarResult, EXCEPINFO FAR* pexcepinfo, UINT FAR* puArgErr) { return DISP_E_MEMBERNOTFOUND; }

protected:
	virtual BOOL GetInterface(REFIID riid, LPVOID* ppv) {
		if (IsEqualIID(riid, IID_IUnknown))
			*ppv = (LPUNKNOWN)(T*)this;
		else if (IsEqualIID(riid, IID_IDispatch))
			*ppv = (LPDISPATCH)(T*)this;
		else if (IsEqualIID(riid, *piid))
			*ppv = (T*)this;
		else
			return FALSE;
		return TRUE;
	}
};

template <class T, const IID* piid>
class NO_VTABLE IDispatchEventSinkAdviseImpl : public IDispatchEventSinkImpl<T, piid> {
public:
	IDispatchEventSinkAdviseImpl() : _cp(NULL), _cookie(0), _hrCP(S_OK) {}
	virtual ~IDispatchEventSinkAdviseImpl() { ASSERT(_cp == NULL); }

	HRESULT Advise(LPUNKNOWN pUnkSource)
	{
		LPCONNECTIONPOINT pCP = NULL;
		LPCONNECTIONPOINTCONTAINER pCPC;
		HRESULT hr = pUnkSource->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pCPC);
		if (SUCCEEDED(hr))
		{
			hr = pCPC->FindConnectionPoint(*piid, &pCP);
			if (SUCCEEDED(hr))
			{
				IDispatchEventSinkImpl<T, piid>::Lock();
				hr = pCP->Advise(this, &_cookie);
				if (FAILED(hr))
					FINALRELEASE(&pCP);
				else
					_cp = pCP;
				IDispatchEventSinkImpl<T, piid>::Unlock();
			}
			pCPC->Release();
		}
		_hrCP = hr;
		return hr;
	}
	void Unadvise() {
		IDispatchEventSinkImpl<T, piid>::Lock();
		if (_cp && _cookie)
		{
			try
			{
				_hrCP = _cp->Unadvise(_cookie);
				FINALRELEASE(&_cp);
			}
			catch (...)
			{
				ASSERT(FALSE);
				_cp = NULL;
			}
			_cookie = 0;
		}
		IDispatchEventSinkImpl<T, piid>::Unlock();
	}

public:
	HRESULT _hrCP;
protected:
	DWORD _cookie;
	LPCONNECTIONPOINT _cp;
};

