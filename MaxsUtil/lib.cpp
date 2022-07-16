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
// VersionInfo.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include <initguid.h>
#include "IDispatchImpl.h"
#include "MaxsUtil_h.h"
#include "MaxsUtil_i.c"
#include "RegistryHelper.h"
#include "VersionInfoImpl.h"
#include "InputBoxImpl.h"
#include "ProgressBoxImpl.h"


// This is a table of COM classes we want to expose. DllRegisterServer and DllUnregisterServer use the table for registration purposes. If you define a new COM class, make sure it's added to this table, and add the C++ implementation class to DllGetClassObject.
static CoclassRegInfo s_cri[] =
{
	{&CLSID_VersionInfo, L"MaxsUtilLib.VersionInfo", L"Max's VersionInfo", &IID_IVersionInfo, L"VersionInfo",},
	{&CLSID_InputBox, L"MaxsUtilLib.InputBox", L"Max's InputBox", &IID_IInputBox, L"InputBox",},
	{&CLSID_ProgressBox, L"MaxsUtilLib.ProgressBox", L"Max's ProgressBox", &IID_IProgressBox, L"ProgressBox",},
};


// global variables.
ULONG LibRefCount = 0;
HMODULE LibInstanceHandle = NULL;


BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		DisableThreadLibraryCalls(hModule);
		LibInstanceHandle = hModule;
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// Special entry points required for inproc servers

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	*ppv = NULL;
	IClassFactory *pClassFactory;
	// propertysheet handlers
	if (IsEqualCLSID(rclsid, CLSID_VersionInfo))
		pClassFactory = new IClassFactoryNoAggrImpl<VersionInfoImpl>;
	else if (IsEqualCLSID(rclsid, CLSID_InputBox))
		pClassFactory = new IClassFactoryNoAggrImpl<InputBoxImpl>;
	else if (IsEqualCLSID(rclsid, CLSID_ProgressBox))
		pClassFactory = new IClassFactoryNoAggrImpl<ProgressBoxImpl>;
	else
		return CLASS_E_CLASSNOTAVAILABLE;
	if (pClassFactory == NULL)
		return E_OUTOFMEMORY;
	HRESULT hr = pClassFactory->QueryInterface(riid, ppv);
	pClassFactory->Release();
	return hr;
}

STDAPI DllCanUnloadNow(void)
{
	return (LibRefCount == 0) ? S_OK : S_FALSE;
}

DWORD DllGetVersion(void)
{
	return MAKELONG(1, 0);
}


///////////////////////////////////////////////////////////
STDAPI DllRegisterServer(void)
{
	HRESULT hr;
	HRESULT hr2 = S_OK;

	for (int i = 0; i < ARRAYSIZE(s_cri); i++) {
		CoclassRegistryHelper rh(LibInstanceHandle, s_cri + i, &LIBID_MaxsUtilLib);
		hr = rh.registerCoclass();
		if (FAILED(hr))
			hr2 = hr;
		hr = rh.registerInterface(L"MaxsUtil.dll");
		if (FAILED(hr))
			hr2 = hr;
	}

	hr = COMServerLibRegistryHelper(LibInstanceHandle, &LIBID_MaxsUtilLib).registerTypelib(L"Maximilian's Utility Library", L"MaxsUtil.dll");
	if (FAILED(hr))
		hr2 = hr;
	return hr2;
}

STDAPI DllUnregisterServer(void)
{
	HRESULT hr;
	HRESULT hr2 = S_OK;

	for (int i = 0; i < ARRAYSIZE(s_cri); i++) {
		CoclassRegistryHelper rh(LibInstanceHandle, s_cri + i, &LIBID_MaxsUtilLib);
		hr = rh.unregisterCoclass();
		if (FAILED(hr))
			hr2 = hr;

		hr = rh.unregisterInterface();
		if (FAILED(hr))
			hr2 = hr;
	}

	hr = COMServerLibRegistryHelper(LibInstanceHandle, &LIBID_MaxsUtilLib).unregisterTypelib();
	if (FAILED(hr))
		hr2 = hr;
	return hr2;
}

