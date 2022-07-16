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
#include <cguid.h> // for GUID_NULL, CLSID_NULL
#include <shlwapi.h>


// Enable this directive to register the COM server for all users of the machine. If it's not enabled, the registration is per user (i.e., the settings are written to the key HKEY_CURRENT_USER\Software\Classes).
//#define _REGISTER_PER_MACHINE

#ifdef _REGISTER_PER_MACHINE
#define REGHLP_CLASSES_HKEY HKEY_LOCAL_MACHINE
#else//#ifdef _REGISTER_PER_MACHINE
#define REGHLP_CLASSES_HKEY HKEY_CURRENT_USER
#endif//#ifdef _REGISTER_PER_MACHINE
#define REGHLP_CLASSES_SUBKEY L"Software\\Classes"
#define REGHLP_CLASSES_SUBKEY2 REGHLP_CLASSES_SUBKEY L"\\"

/* This is a registry helper class. It provides methods useful in quickly reading from, writing to a registry key, and deleting registry keys and values. The methods performs the necessary steps of opening a target registry key, reading or writing the setting, and closing the key, so you don't have to.
*/
class RegistryHelper
{
public:
	RegistryHelper(HKEY root) : _rootKey(root), _keyName{ 0 } {}
	RegistryHelper(HKEY root, LPCWSTR pszKey) : _rootKey(root), _keyName{ 0 }
	{
		if (pszKey)
			lstrcpy(_keyName, pszKey);
	}
	RegistryHelper(RegistryHelper &parent, LPCWSTR subkey) : _rootKey(parent._rootKey)
	{
		lstrcat(lstrcat(lstrcpy(_keyName, parent._keyName), L"\\"), subkey);
	}
	~RegistryHelper() {}

	// formats arguments based on fmt, and stores the result string as a key path in _keyName.
	void formatKey(LPCWSTR fmt, ...)
	{
		va_list args;
		va_start(args, fmt);
#ifdef _DEBUG
		int cc = _vscwprintf(fmt, args);
		ASSERT(cc < ARRAYSIZE(_keyName));
#endif//_DEBUG
		vswprintf_s(_keyName, ARRAYSIZE(_keyName), fmt, args);
		va_end(args);
	}
	// appends a formatted string as an additional key path to the current key path.
	void formatSubkey(LPCWSTR fmt, ...)
	{
		int taken = (int)wcslen(_keyName);
		if (taken > 0 && _keyName[taken - 1] != '\\')
		{
			_keyName[taken] = '\\';
			taken++;
		}
		LPWSTR p0 = _keyName + taken;
		int len0 = ARRAYSIZE(_keyName) - taken;
		va_list args;
		va_start(args, fmt);
#ifdef _DEBUG
		int cc = _vscwprintf(fmt, args);
		ASSERT(cc < len0);
#endif//_DEBUG
		vswprintf_s(p0, len0, fmt, args);
		va_end(args);
	}

	// formats arguments based on fmt, and writes the result string as the data of a value named valueName and created in a registry key of _keyName.
	LONG formatValue(LPCWSTR valueName, LPCWSTR fmt, ...)
	{
		WCHAR val[512];
		va_list args;
		va_start(args, fmt);
#ifdef _DEBUG
		int cc = _vscwprintf(fmt, args);
		ASSERT(cc < ARRAYSIZE(val));
#endif//_DEBUG
		vswprintf_s(val, ARRAYSIZE(val), fmt, args);
		va_end(args);
		return createValue(valueName, val);
	}

	// creates a value entry of name pszName in a registry key of _keyName, and assigns a value given by pszValue.
	LONG createValue(LPCWSTR pszName, LPCWSTR pszValue)
	{
		LONG res;
		HKEY hkey;
		DWORD dwDisposition;

		res = RegCreateKeyEx(_rootKey, _keyName, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkey, &dwDisposition);
		if (res == ERROR_SUCCESS)
		{
			res = RegSetValueEx(hkey, pszName, NULL, REG_SZ, (LPBYTE)pszValue, ((DWORD)wcslen(pszValue) + 1) * sizeof(TCHAR));
			RegCloseKey(hkey);
		}
		return res;
	}

	// deletes a subkey from a registry key of _keyName. add 'shlwapi.lib' to the linker input list because the method calls Win32 SHDeleteKey().
	LONG deleteSubkey(LPCWSTR pszSubKey)
	{
		LONG res;
		HKEY hkey;

		res = RegOpenKeyEx(_rootKey, _keyName, 0, KEY_ALL_ACCESS, &hkey);
		if (res == ERROR_SUCCESS)
		{
			res = SHDeleteKey(hkey, pszSubKey);
			RegCloseKey(hkey);
		}
		return res;
	}

	// deletes a value entry in a registry key of _keyName.
	LONG deleteValue(LPCWSTR pszValueName)
	{
		LONG res;
		HKEY hkey;

		res = RegOpenKeyEx(_rootKey, _keyName, 0, KEY_WRITE, &hkey);
		if (res == ERROR_SUCCESS)
		{
			res = RegDeleteValue(hkey, pszValueName);
			RegCloseKey(hkey);
		}
		return res;
	}

	// sets a value of name pszName in registry key _keyName to a DWORD data given by dwValue.
	LONG setDwordValue(LPCWSTR pszName, DWORD dwValue)
	{
		LONG res;
		HKEY hkey;
		DWORD dwDisposition;

		res = RegCreateKeyEx(_rootKey, _keyName, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkey, &dwDisposition);
		if (res == ERROR_SUCCESS)
		{
			res = RegSetValueEx(hkey, pszName, NULL, REG_DWORD, (LPBYTE)&dwValue, sizeof(DWORD));
			RegCloseKey(hkey);
		}
		return res;
	}

	// reads the data assigned to a DWORD value entry in registry key _keyName and passes it to the caller in pdwValue.
	LONG getDwordValue(LPCWSTR pszName, DWORD* pdwValue)
	{
		LONG res;
		HKEY hkey;

		res = RegOpenKeyEx(_rootKey, _keyName, 0, KEY_QUERY_VALUE, &hkey);
		if (res == ERROR_SUCCESS)
		{
			DWORD dwType = REG_DWORD;
			DWORD dwLength = sizeof(DWORD);
			res = RegQueryValueEx(hkey, pszName, NULL, &dwType, (LPBYTE)pdwValue, &dwLength);
			RegCloseKey(hkey);
		}
		return res;
	}

protected:
	HKEY _rootKey; // a root key, e.g., HKEY_CURRENT_USER.
	TCHAR _keyName[256]; // a fully qualified pathname of a registry key.
};

// define useful macros of rudimentary registry management.
#define Registry_CreateNameValue(hkeyRoot, pszKey, pszName, pszValue) RegistryHelper(hkeyRoot,pszKey).createValue(pszName, pszValue)
#define Registry_DeleteSubkey(hkeyRoot, pszKey, pszSubKey) RegistryHelper(hkeyRoot,pszKey).deleteSubkey(pszSubKey)
#define Registry_DeleteValue(hkeyRoot, pszKey, pszValueName) RegistryHelper(hkeyRoot,pszKey).deleteValue(pszValueName)
#define Registry_SetDwordValue(hkeyRoot, pszKey, pszName, dwValue) RegistryHelper(hkeyRoot,pszKey).setDwordValue(pszName, dwValue)
#define Registry_GetDwordValue(hkeyRoot, pszKey, pszName, pdwValue) RegistryHelper(hkeyRoot,pszKey).getDwordValue(pszName, pdwValue)


/////////////////////////////////////////////////////////////////////
#define SZ_PROXYSTUB_CLSID L"{00020420-0000-0000-C000-000000000046}"

/* This structure describes a COM server class in terms of parameters needed in registering the server with the system. The structure is used to initialize CoclassRegistryHelper with necessary registration information.
*/
struct CoclassRegInfo {
	const CLSID *clsid;
	LPCWSTR progId;
	LPCWSTR friendlyName;
	const IID *iid;
	LPCWSTR interfaceName;
};

/* This is a COM server registration helper. Specifically, use it to register or unregister the type library of a COM server. It creates (or deletes) keys and settings the type library is required of under the base key of [HKCR\typelib].
*/
class COMServerLibRegistryHelper
{
public:
	COMServerLibRegistryHelper(HINSTANCE hInst, const GUID* plibid, DWORD libVersion = MAKELONG(0, 1)) : _hInst(hInst), _plibid(plibid), _libVersion(libVersion), _wszPath{0}
	{
		StringFromIID(*_plibid, &_pwszLIBID);
		GetModuleFileName(_hInst, _wszPath, ARRAYSIZE(_wszPath));
	}
	~COMServerLibRegistryHelper()
	{
		CoTaskMemFree(_pwszLIBID);
	}

	/* generates the following typelib registration.

		[HKEY_CLASSES_ROOT]
		+ [TypeLib]
		  + [{...LIBID...}]
		   + [<major.minor>]
		    + [FLAGS]
		    | |  @=<libflags>
			| + [<LCID>]
		    |  + [win32]
		    |      @=<TLB-Path>
		    + [HELPDIR]
		        @=<HelpFolderPath>

		Parameter libflag is typically set to 0 for permitting access to typlib. It could be set to one of these.
		typedef enum tagLIBFLAGS {
		  LIBFLAG_FRESTRICTED = 0x1,
		  LIBFLAG_FCONTROL = 0x2,
		  LIBFLAG_FHIDDEN = 0x4,
		  LIBFLAG_FHASDISKIMAGE = 0x8
		} LIBFLAGS;
		For description of each LIBFLAG, refer to https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ne-oaidl-libflags.

		Parameter lcid identifies a locale the typelib is intended for. It's typically set to 0 meaning the typelib works for any locale.
	*/
	HRESULT registerTypelib(LPCWSTR libName, LPCWSTR tlbName, DWORD libflags=0, DWORD lcid=0)
	{
		DWORD res;
		HRESULT hr = S_OK;
		//[HKEY_CLASSES_ROOT\TypeLib\{...LIBID...}\1.0]
		//@="fabulous library"
		RegistryHelper rkey(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		rkey.formatSubkey(L"TypeLib\\%s\\%d.%d", _pwszLIBID, HIWORD(_libVersion), LOWORD(_libVersion));
		res = rkey.formatValue(NULL, L"%s %d.%d", libName, HIWORD(_libVersion), LOWORD(_libVersion));
		if (res == ERROR_SUCCESS)
		{
			//[HKEY_CLASSES_ROOT\TypeLib\{...LIBID...}\1.0\FLAGS]
			//@="0"
			RegistryHelper subkeyFlags(rkey, L"FLAGS");
			res = subkeyFlags.formatValue(NULL, L"%X", libflags);

			//[HKEY_CLASSES_ROOT\TypeLib\{...LIBID...}\1.0\0\win32]
			//@="c:\\program files\\...\\???.tlb"
			TCHAR szVal[_MAX_PATH];
			_replacePathTitle(tlbName, szVal);
			RegistryHelper subkeyWin32(rkey);
			subkeyWin32.formatSubkey(L"%X\\win32", lcid);
			subkeyWin32.createValue(NULL, szVal);

			//[HKEY_CLASSES_ROOT\TypeLib\{...LIBID...}\1.0\HELPDIR]
			//@="c:\\program files\\???"
			_replacePathTitle(NULL, szVal);
			RegistryHelper subkeyHelpdir(rkey, L"HELPDIR");
			res = subkeyHelpdir.createValue(NULL, szVal);
		}
		else
			hr = HRESULT_FROM_WIN32(res);
		return hr;
	}
	// removes our LIBID subkey from [HKCR\TypeLib].
	HRESULT unregisterTypelib()
	{
		RegistryHelper basekey(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2 L"TypeLib");
		long res = basekey.deleteSubkey(_pwszLIBID);
		return HRESULT_FROM_WIN32(res);
	}

protected:
	HINSTANCE _hInst;
	const CLSID* _plibid;
	DWORD _libVersion;
	TCHAR _wszPath[MAX_PATH];
	LPWSTR _pwszLIBID;

	// scans the internal path buffer _wszPath for newTitle, copies the leading part of the internal path up to where a match is found to the output buffer outVal, and appends newTitle to outVal.
	// If newTitle is NULL, the method returns the parent path.
	LPWSTR _replacePathTitle(LPCWSTR newTitle, LPWSTR outVal)
	{
		LPWSTR p = wcsrchr(_wszPath, '\\');
		ASSERT(p);
		int len = (int)(INT_PTR)(p - _wszPath);
		// special case where no new title should be set. copy the parent path.
		if (!newTitle)
		{
			// lose the trailling backslash. notice it's 'len+1', and not 'len'. Win32 lstrcpyn wants a null char space included in the size parameter.
			return lstrcpyn(outVal, _wszPath, len + 1);
		}
		len++;
		lstrcpyn(outVal, _wszPath, len + 1);
		lstrcpy(outVal + len, newTitle);
		return outVal;
	}
};

/* Use this helper class to register a COM class and an automation interface it exposes. Its register* functions generate the following trees of registry keys and settings. The unregister* functions remove them. It assumes the COM class supports the apartment threading model.

	[HKEY_CLASSES_ROOT]
	+ [CLSID]
	 + [{...clsid...}]
	  |   @=<FriendlyName>
	  + [IInProcServer32]
	  |   @=<ModulePath>
	  |   ThreadingModel="Apartment"
	  + [DefaultIcon]
	  |   @=<ModulePath,IconIndex>
	  |
	  + [ProgID]
	  |   @="<ProgId.Version>"
	  + [VersionIndependentProgID]
	  |   @="<ProgId>"
	  + [Version]
	  |   @="<Major.Minor>"
	  + [TypeLib]
	      @="{...LIBID...}"

	[HKEY_CLASSES_ROOT]
	+ [<ProgId>]
	| + [CLSID]
	| |   @="{...clsid...}"
	| + [CurVer]
	|     @="<ProgId.Version>"
	+ [<ProgId.Version>]
	  + [CLSID]
	      @="{...clsid...}"

	[HKEY_CLASSES_ROOT]
	+ [Interface]
	 + [{...iid...}]
	  |  @="<IntefaceName>"
	  + [ProxyStubClsid]
	  |  @= "{00020420-0000-0000-C000-000000000046}"
	  + [ProxyStubClsid32]
	  |  @= "{00020420-0000-0000-C000-000000000046}"
	  + [TypeLib]
	     @="{...LIBID...}"
	     Version="<Major.Minor>"
*/
class CoclassRegistryHelper : public COMServerLibRegistryHelper
{
public:
	CoclassRegistryHelper(HINSTANCE hInst, const CoclassRegInfo *cri, const CLSID* plibid, DWORD libVersion = MAKELONG(0,1)) : COMServerLibRegistryHelper(hInst, plibid, libVersion), _crinf(cri)
	{
		StringFromIID(*(_crinf->clsid), &_pwszCLSID);
		StringFromIID(*(_crinf->iid), &_pwszIID);
	}
	~CoclassRegistryHelper()
	{
		CoTaskMemFree(_pwszIID);
		CoTaskMemFree(_pwszCLSID);
	}

	HRESULT registerCoclass(int nIconIndex=0)
	{
		HRESULT hr = _registerAsCOMServer(nIconIndex);
		if (hr == S_OK)
			hr = _registerForAutomation();
		return hr;
	}
	HRESULT registerInterface(LPCWSTR tlbName)
	{
		//[HKEY_CLASSES_ROOT\Interface\{...IID...}]
		//@= "MaxsUtilLib..."
		RegistryHelper basekey(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		basekey.formatSubkey(L"Interface\\%s", _pwszIID);
		long res = basekey.createValue(NULL, _crinf->interfaceName);
		if (res != ERROR_SUCCESS)
			return HRESULT_FROM_WIN32(res);
		//[HKEY_CLASSES_ROOT\Interface\{...IID...}\ProxyStubClsid]
		//@= "{00020420-0000-0000-C000-000000000046}"
		RegistryHelper subkeyProxystubclsid(basekey, L"ProxyStubClsid");
		res = subkeyProxystubclsid.createValue(NULL, SZ_PROXYSTUB_CLSID);
		//[HKEY_CLASSES_ROOT\Interface\{...IID...}\ProxyStubClsid32]
		//@= "{00020420-0000-0000-C000-000000000046}"
		RegistryHelper subkeyProxystubclsid32(basekey, L"ProxyStubClsid32");
		res = subkeyProxystubclsid32.createValue(NULL, SZ_PROXYSTUB_CLSID);
		//[HKEY_CLASSES_ROOT\Interface\{...IID...}\TypeLib]
		//@= "{...LIBID...}"
		RegistryHelper subkeyTypelib(basekey, L"TypeLib");
		res = subkeyTypelib.createValue(NULL, _pwszLIBID);
		//"Version"="1.0"
		res = subkeyTypelib.formatValue(L"Version", L"%d.%d", HIWORD(_libVersion), LOWORD(_libVersion));
		return S_OK;
	}

	HRESULT unregisterCoclass()
	{
		if (*(_crinf->clsid) != CLSID_NULL)
		{
			RegistryHelper clsidKey(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2 L"CLSID");
			clsidKey.deleteSubkey(_pwszCLSID);
		}
		if (_crinf->progId)
		{
			TCHAR szVerProgId[80];
			wsprintf(szVerProgId, L"%s.%d", _crinf->progId, HIWORD(_libVersion));
			RegistryHelper rootKey(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY);
			rootKey.deleteSubkey(szVerProgId);
			rootKey.deleteSubkey(_crinf->progId);
		}
		return S_OK;
	}
	HRESULT unregisterInterface()
	{
		long res;
		if (*(_crinf->iid) != CLSID_NULL)
		{
			RegistryHelper interfaceKey(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2 L"Interface");
			res = interfaceKey.deleteSubkey(_pwszIID);
		}
		return HRESULT_FROM_WIN32(res);
	}

protected:
	const CoclassRegInfo *_crinf;
	LPWSTR _pwszCLSID, _pwszIID;

	HRESULT _registerAsCOMServer(int nIconIndex)
	{
		//[HKEY_CLASSES_ROOT\CLSID\{...CLSID...}]
		//@= "fabulous library"
		RegistryHelper basekey(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		basekey.formatSubkey(L"CLSID\\%s", _pwszCLSID);
		long res = basekey.createValue(NULL, _crinf->friendlyName);
		if (res != ERROR_SUCCESS)
			return HRESULT_FROM_WIN32(res);
		//[HKEY_CLASSES_ROOT\CLSID\{...CLSID...}\InProcServer32]
		//@= "c:\\program files\\...\\???.dll"
		//"ThreadingModel" = "Apartment"
		RegistryHelper subkeyInprocserver32(basekey, L"InProcServer32");
		res = subkeyInprocserver32.createValue(NULL, _wszPath);
		res = subkeyInprocserver32.createValue(L"ThreadingModel", L"Apartment");
		//[HKEY_CLASSES_ROOT\CLSID\{...CLSID...}\DefaultIcon]
		//@= "c:\\program files\\...\\???.dll,0"
		RegistryHelper subkeyDefaulticon(basekey, L"DefaultIcon");
		res = subkeyDefaulticon.formatValue(NULL, L"%s,%d", _wszPath, nIconIndex);
		return S_OK;
	}
	HRESULT _registerForAutomation()
	{
		LONG res;
		//[HKEY_CLASSES_ROOT\MaxsUtilLib.VersionInfo\CLSID]
		//@= "{...CLSID...}"
		RegistryHelper key0(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		key0.formatSubkey(L"%s\\CLSID", _crinf->progId);
		res = key0.createValue(NULL, _pwszCLSID);
		if (res != ERROR_SUCCESS)
			return HRESULT_FROM_WIN32(res);
		WCHAR szVerProgId[80];
		wsprintf(szVerProgId, L"%s.%d", _crinf->progId, HIWORD(_libVersion));
		//[HKEY_CLASSES_ROOT\MaxsUtilLib.VersionInfo\CurVer]
		//@= "MaxsUtilLib.VersionInfo.1"
		RegistryHelper key1(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		key1.formatSubkey(L"%s\\CurVer", _crinf->progId);
		res = key1.formatValue(NULL, szVerProgId);
		//[HKEY_CLASSES_ROOT\MaxsUtilLib.VersionInfo.1\CLSID]
		//@= "{...CLSID...}"
		RegistryHelper key2(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		key2.formatSubkey(L"%s\\CLSID", szVerProgId);
		res = key2.createValue(NULL, _pwszCLSID);
#ifdef NEED_PROGRAMMABLE_KEY
		//[HKEY_CLASSES_ROOT\CLSID\{...CLSID...}\Programmable]
		//""
		RegistryHelper keyCLSID(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		keyCLSID.formatSubkey(L"CLSID\\%s\\Programmable", _pwszCLSID);
		res = keyCLSID.createValue(NULL, L"");
#endif//	#ifdef NEED_PROGRAMMABLE_KEY
		//[HKEY_CLASSES_ROOT\CLSID\{...CLSID...}\VersionIndependentProgID]
		//"MaxsUtilLib.VersionInfo"
		RegistryHelper keyCLSID2(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		keyCLSID2.formatSubkey(L"CLSID\\%s\\VersionIndependentProgID", _pwszCLSID);
		res = keyCLSID2.createValue(NULL, _crinf->progId);
		//[HKEY_CLASSES_ROOT\CLSID\{...CLSID...}\ProgId]
		//"MaxsUtilLib.VersionInfo.1"
		RegistryHelper keyCLSID3(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		keyCLSID3.formatSubkey(L"CLSID\\%s\\ProgId", _pwszCLSID);
		res = keyCLSID3.createValue(NULL, szVerProgId);
		//[HKEY_CLASSES_ROOT\CLSID\{...CLSID...}\Version]
		//@= "1.0"
		RegistryHelper keyCLSID4(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		keyCLSID4.formatSubkey(L"CLSID\\%s\\Version", _pwszCLSID);
		res = keyCLSID4.formatValue(NULL, L"%d.%d", HIWORD(_libVersion), LOWORD(_libVersion));
		//[HKEY_CLASSES_ROOT\CLSID\{...CLSID...}\TypeLib]
		//@= "{...LIBID...}"
		RegistryHelper keyCLSID5(REGHLP_CLASSES_HKEY, REGHLP_CLASSES_SUBKEY2);
		keyCLSID5.formatSubkey(L"CLSID\\%s\\TypeLib", _pwszCLSID);
		res = keyCLSID5.createValue(NULL, _pwszLIBID);
		return S_OK;
	}
};

