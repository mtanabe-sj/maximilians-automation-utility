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


/* Use this class to wrap a VARIANT. It automatically calls VariantClear() to free a space allocation made by the constructor when control leaves the current scope. Assignment methods assume no data has been assigned. So, you must call clear() before you can assign another value to an already assigned VariantAutoRel. Otherwise, the second assignment causes a space leak.
*/
class VariantAutoRel
{
public:
	VariantAutoRel() { init(); }
	VariantAutoRel(short iVal) { init(); assignShort(iVal); }
	VariantAutoRel(long lVal) { init(); assignLong(lVal); }
	VariantAutoRel(DATE dtVal) { init(); assignDate(dtVal); }
	VariantAutoRel(LPDISPATCH pdispVal) { init(); assignDisp(pdispVal); }
	VariantAutoRel(LPUNKNOWN punkVal) { init(); assignPunk(punkVal); }
	VariantAutoRel(BSTR *pbstrVal) { init(); attachB(pbstrVal); }
	VariantAutoRel(LPCSTR pszVal) { init(); assignA(pszVal); }
	VariantAutoRel(LPCWSTR pszVal) { init(); assignW(pszVal); }
	VariantAutoRel(LPBYTE pData, UINT cbData) { init(); assign(pData, cbData); }
	VariantAutoRel(HWND h) { init(); assignHandle((OLE_HANDLE)(UINT)(UINT_PTR)h); }
	virtual ~VariantAutoRel() {
		clear();
	}
	void init() {
		VariantInit(&_v);
	}
	void clear() {
		VariantClear(&_v);
	}
	void attach(VARIANT var) {
		clear();
		_v = var;
	}
	VARIANT detach() {
		VARIANT var = _v;
		init();
		return var;
	}

	// tests if no data has been assigned to the wrapped VARIANT. Notice that it does not test if an assigned data has a zero length.
	BOOL isEmpty() {
		return (_v.vt == VT_EMPTY);
	}

	HRESULT copyFrom(VARIANT* src) {
		clear();
		return VariantCopy(&_v, src);
	}
	HRESULT copyTo(VARIANT* dest) {
		return VariantCopy(dest, &_v);
	}

	// useful operators.
	operator VARIANT*() { return &_v; }
	operator VARIANT&() { return _v; }

	// a VARIANT instance the class wraps.
	VARIANT _v;

	// assigns a 2-byte signed integer.
	bool assignShort(short iVal) {
		_v.vt = VT_I2;
		_v.iVal = iVal;
		return true;
	}
	// assigns a 4-byte signed integer.
	bool assignLong(long lVal) {
		_v.vt = VT_I4;
		_v.lVal = lVal;
		return true;
	}
	// assigns an OLE date.
	bool assignDate(DATE dtVal) {
		_v.vt = VT_DATE;
		_v.date = dtVal;
		return true;
	}
	// assigns a boolean.
	bool assignBool(VARIANT_BOOL bVal) {
		_v.vt = VT_BOOL;
		_v.boolVal = bVal;
		return true;
	}
	// assigns an automation (IDispatch) object.
	bool assignDisp(LPDISPATCH pdispVal) {
		_v.vt = VT_DISPATCH;
		_v.pdispVal = pdispVal;
		pdispVal->AddRef();
		return true;
	}
	// assigns a IUnknown object.
	bool assignPunk(LPUNKNOWN punkVal) {
		_v.vt = VT_UNKNOWN;
		_v.punkVal = punkVal;
		punkVal->AddRef();
		return true;
	}
	// attaches a BSTR from a source and clears the source pointer.
	bool attachB(BSTR *pbstrVal) {
		_v.vt = VT_BSTR;
		_v.bstrVal = (BSTR)InterlockedExchangePointer((LPVOID*)pbstrVal, NULL);
		return true;
	}
	// assigns text in ASCII.
	bool assignA(LPCSTR pszVal) {
		init();
		int cc = MultiByteToWideChar(CP_ACP, 0, pszVal, -1, NULL, 0);
		if (cc == 0 || !(_v.bstrVal = SysAllocStringLen(NULL, cc)))
			return false;
		MultiByteToWideChar(CP_ACP, 0, pszVal, -1, _v.bstrVal, cc + 1);
		_v.vt = VT_BSTR;
		return true;
	}
	// assigns text in unicode.
	bool assignW(LPCWSTR pszVal) {
		_v.vt = VT_BSTR;
		_v.bstrVal = SysAllocString(pszVal);
		return true;
	}
	// assigns a range of data bytes at starting address pData and of length cbData. A safe array is used to allocate the buffer space.
	bool assign(LPBYTE pData, UINT cbData)
	{
		init();
		SAFEARRAYBOUND rgsabound[1];
		rgsabound[0].lLbound = 0;
		rgsabound[0].cElements = cbData;
		SAFEARRAY* psa = SafeArrayCreate(VT_UI1, 1, rgsabound);
		if (!psa)
			return false;
		memcpy(psa->pvData, pData, cbData);
		_v.vt = (VT_ARRAY | VT_UI1);
		_v.parray = psa;
		return true;
	}
	bool assignHandle(OLE_HANDLE h)
	{
		_v.vt = VT_I4;
		_v.ulVal = (ULONG)(ULONG_PTR)h;
		return true;
	}
};

// macros for getting a pointer to and the byte size of the binary data stored as a safe array in VariantAutoRel.
#define VARIANTAUTOREL_BINDATAPTR(var) (LPBYTE)((var)._v.parray->pvData)
#define VARIANTAUTOREL_BINDATALEN(var) ((var)._v.parray->rgsabound[0].cElements)


