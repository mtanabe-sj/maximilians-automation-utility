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
#include <Windows.h>
#include <OleCtl.h>
#include <ocidl.h>
#include "AxObjList.h"
#include "IDispatchImpl.h"

#ifndef ASSERT
#define ASSERT(good)
#endif//#ifndef ASSERT

#define _CookieFromAddress(p) ((ULONG)(ULONG_PTR)(LPVOID)(p))

//#define ADDREF_CPC


class ConnectionPointCallback
{
public:
	virtual void GetIID(IID* pIID) = 0;
	virtual void OnAdviseConnectionPoint() = 0;
	virtual void OnUnadviseConnectionPoint() = 0;
};


class ConnectionPointImpl : public IUnknownImpl<IConnectionPoint, &IID_IConnectionPoint>
{
public:
	ConnectionPointImpl(ConnectionPointCallback* cb, IConnectionPointContainer* cpc) : _cb(cb), _cpc(cpc), _sink(NULL)
	{
#ifdef ADDREF_CPC
		_cpc->AddRef();
#endif//#ifdef ADDREF_CPC
	}
	virtual ~ConnectionPointImpl()
	{
		FINALRELEASE(&_sink);
#ifdef ADDREF_CPC
		FINALRELEASE(&_cpc);
#else//#ifdef ADDREF_CPC
		_cpc = NULL;
#endif//#ifdef ADDREF_CPC
	}

	// *** IConnectionPoint ***
	STDMETHOD(GetConnectionInterface)(IID* pIID)
	{
		_cb->GetIID(pIID);
		return S_OK;
	}
	STDMETHOD(GetConnectionPointContainer)(IConnectionPointContainer** ppCPC)
	{
		try {
			_cpc->AddRef();
		}
		catch (...) {
			ASSERT(FALSE);
			_cpc = NULL;
			*ppCPC = NULL;
			return E_FAIL;
		}
		*ppCPC = _cpc;
		return S_OK;
	}
	STDMETHOD(Advise)(LPUNKNOWN pUnkSink, DWORD* pdwCookie)
	{
		ASSERT(!IsBadReadPtr(pUnkSink, sizeof(IUnknown)));
		ASSERT((pdwCookie == NULL) || !IsBadWritePtr(pdwCookie, sizeof(DWORD)));

		if (pUnkSink == NULL)
			return E_POINTER;

		if (_sink)
			return CONNECT_E_ADVISELIMIT;

		IID iid;
		_cb->GetIID(&iid);
		if (FAILED(pUnkSink->QueryInterface(iid, (LPVOID*)&_sink)))
			return CONNECT_E_CANNOTCONNECT;

		if (pdwCookie != NULL)
			*pdwCookie = _CookieFromAddress(this);

		_cb->OnAdviseConnectionPoint();
		return S_OK;
	}
	STDMETHOD(Unadvise)(DWORD dwCookie)
	{
		if (_CookieFromAddress(this) == dwCookie)
		{
			// a client has finished listening to our event.
			FINALRELEASE(&_sink);
			// report it to the server.
			_cb->OnUnadviseConnectionPoint();
			return S_OK;
		}
		else
		{
			return CONNECT_E_NOCONNECTION;
		}
	}
	STDMETHOD(EnumConnections)(LPENUMCONNECTIONS* ppEnum)
	{
		*ppEnum = NULL;
		return E_NOTIMPL;
	}

	HRESULT FireEvent(DISPID dispId, VARIANT* pvEventArgs, UINT cEventArgs)
	{
		if (!_sink)
			return E_UNEXPECTED;

		HRESULT hr = NOERROR;
		DISPPARAMS dispparams;
		VARIANT vaResult;
		EXCEPINFO excepInfo;
		UINT nArgErr = 0;

		memset(&dispparams, 0, sizeof(dispparams));
		memset(&excepInfo, 0, sizeof(excepInfo));
		VariantInit(&vaResult);

		dispparams.cArgs = cEventArgs;
		dispparams.rgvarg = pvEventArgs;

		try
		{
			// Fire the event.
			hr = _sink->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, &vaResult, &excepInfo, &nArgErr);
			if (FAILED(hr))
			{
				if (excepInfo.bstrSource != NULL)
					SysFreeString(excepInfo.bstrSource);
				if (excepInfo.bstrDescription != NULL)
					SysFreeString(excepInfo.bstrDescription);
				if (excepInfo.bstrHelpFile != NULL)
					SysFreeString(excepInfo.bstrHelpFile);
			}
		}
		catch (...)
		{
			hr = DISP_E_EXCEPTION;
		}

		return hr;
	}

protected:
	ConnectionPointCallback* _cb;
	IConnectionPointContainer* _cpc;
	IDispatch* _sink;
};

class ConnectionPointListImpl : public AxObjList<ConnectionPointImpl>
{
public:
	ConnectionPointListImpl(ConnectionPointCallback* pCPCallback, IConnectionPointContainer* pCPContainer) : _cpCB(pCPCallback), _cpCont(pCPContainer), _cpPos(0) {}
	~ConnectionPointListImpl()
	{
		_cpCB = NULL;
		_cpCont = NULL;
		_cpPos = 0;
		AxObjList<ConnectionPointImpl>::clear();
	}

	// *** IConnectionPointContainer ***
	HRESULT EnumConnectionPoints(IEnumConnectionPoints **ppEnum)
	{
		_cpPos = 0;
		return _cpCont->QueryInterface(IID_IEnumConnectionPoints, (LPVOID*)ppEnum);
	}
	HRESULT FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP)
	{
		if (IsBadWritePtr(ppCP, sizeof(LPCONNECTIONPOINT)))
			return E_POINTER;
		if (!_cpCB || !_cpCont)
			return E_UNEXPECTED; // called from destructor?
		IID iid;
		_cpCB->GetIID(&iid);
		if (riid != iid)
			return CONNECT_E_NOCONNECTION;
		ConnectionPointImpl* pCP = new ConnectionPointImpl(_cpCB, _cpCont);
		add(pCP);
		*ppCP = pCP;
		return S_OK;
	}

	// *** IEnumConnectionPoints methods ***
	HRESULT Next(ULONG cConnections, IConnectionPoint **rgpcn, ULONG *pcFetched)
	{
		if (pcFetched)
			*pcFetched = 0;
		long i;
		for (i = 0; i < (long)cConnections; i++)
			rgpcn[i] = NULL;
		ULONG cFetched = 0;
		for (i = _cpPos; (long)cConnections && i < size(); i++)
		{
			ConnectionPointImpl* pCP = (*this)[i];
			pCP->AddRef();
			rgpcn[i] = pCP;
			_cpPos++;
			cFetched++;
			cConnections--;
		}
		if (pcFetched)
			*pcFetched = cFetched;
		return cConnections == 0 ? S_OK : S_FALSE;
	}
	HRESULT Skip(ULONG cConnections)
	{
		for (long i = _cpPos; cConnections && i < size(); i++)
		{
			//ConnectionPointImpl* pCP = (*this)[i];
			_cpPos++;
			cConnections--;
		}
		return cConnections == 0 ? S_OK : S_FALSE;
	}
	HRESULT Reset()
	{
		_cpPos = 0;
		return S_OK;
	}
	HRESULT Clone(IEnumConnectionPoints **ppEnum)
	{
		return E_NOTIMPL;
	}

protected:
	ConnectionPointCallback* _cpCB;
	IConnectionPointContainer* _cpCont;
	long _cpPos;
};

