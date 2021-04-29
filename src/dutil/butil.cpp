// Copyright (c) .NET Foundation and contributors. All rights reserved. Licensed under the Microsoft Reciprocal License. See LICENSE.TXT file in the project root for full license information.

#include "precomp.h"

// Exit macros
#define ButilExitOnLastError(x, s, ...) ExitOnLastErrorSource(DUTIL_SOURCE_BUTIL, x, s, __VA_ARGS__)
#define ButilExitOnLastErrorDebugTrace(x, s, ...) ExitOnLastErrorDebugTraceSource(DUTIL_SOURCE_BUTIL, x, s, __VA_ARGS__)
#define ButilExitWithLastError(x, s, ...) ExitWithLastErrorSource(DUTIL_SOURCE_BUTIL, x, s, __VA_ARGS__)
#define ButilExitOnFailure(x, s, ...) ExitOnFailureSource(DUTIL_SOURCE_BUTIL, x, s, __VA_ARGS__)
#define ButilExitOnRootFailure(x, s, ...) ExitOnRootFailureSource(DUTIL_SOURCE_BUTIL, x, s, __VA_ARGS__)
#define ButilExitOnFailureDebugTrace(x, s, ...) ExitOnFailureDebugTraceSource(DUTIL_SOURCE_BUTIL, x, s, __VA_ARGS__)
#define ButilExitOnNull(p, x, e, s, ...) ExitOnNullSource(DUTIL_SOURCE_BUTIL, p, x, e, s, __VA_ARGS__)
#define ButilExitOnNullWithLastError(p, x, s, ...) ExitOnNullWithLastErrorSource(DUTIL_SOURCE_BUTIL, p, x, s, __VA_ARGS__)
#define ButilExitOnNullDebugTrace(p, x, e, s, ...)  ExitOnNullDebugTraceSource(DUTIL_SOURCE_BUTIL, p, x, e, s, __VA_ARGS__)
#define ButilExitOnInvalidHandleWithLastError(p, x, s, ...) ExitOnInvalidHandleWithLastErrorSource(DUTIL_SOURCE_BUTIL, p, x, s, __VA_ARGS__)
#define ButilExitOnWin32Error(e, x, s, ...) ExitOnWin32ErrorSource(DUTIL_SOURCE_BUTIL, e, x, s, __VA_ARGS__)

// constants
// From engine/registration.h
const LPCWSTR BUNDLE_REGISTRATION_REGISTRY_UNINSTALL_KEY = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall";
const LPCWSTR BUNDLE_REGISTRATION_REGISTRY_BUNDLE_UPGRADE_CODE = L"BundleUpgradeCode";
const LPCWSTR BUNDLE_REGISTRATION_REGISTRY_BUNDLE_PROVIDER_KEY = L"BundleProviderKey";
const LPCWSTR BUNDLE_REGISTRATION_REGISTRY_BUNDLE_VARIABLE_KEY = L"variables";

// Forward declarations.
/********************************************************************
OpenBundleKey - Opens the bundle uninstallation key for a given bundle

NOTE: caller is responsible for closing key
********************************************************************/
static HRESULT OpenBundleKey(
    __in_z LPCWSTR wzBundleId,
    __in BUNDLE_INSTALL_CONTEXT context,
    __in_opt LPCWSTR szSubKey,
    __inout HKEY *key);


extern "C" HRESULT DAPI BundleGetBundleInfo(
  __in_z LPCWSTR wzBundleId,
  __in_z LPCWSTR wzAttribute,
  __out_ecount_opt(*pcchValueBuf) LPWSTR lpValueBuf,
  __inout_opt LPDWORD pcchValueBuf
  )
{
    Assert(wzBundleId && wzAttribute);

    HRESULT hr = S_OK;
    BUNDLE_INSTALL_CONTEXT context = BUNDLE_INSTALL_CONTEXT_MACHINE;
    LPWSTR sczValue = NULL;
    HKEY hkBundle = NULL;
    DWORD cchSource = 0;
    DWORD dwType = 0;
    DWORD dwValue = 0;

    if ((lpValueBuf && !pcchValueBuf) || !wzBundleId || !wzAttribute)
    {
        ButilExitOnFailure(hr = E_INVALIDARG, "An invalid parameter was passed to the function.");
    }

    if (FAILED(hr = OpenBundleKey(wzBundleId, context = BUNDLE_INSTALL_CONTEXT_MACHINE, NULL, &hkBundle)) &&
        FAILED(hr = OpenBundleKey(wzBundleId, context = BUNDLE_INSTALL_CONTEXT_USER, NULL, &hkBundle)))
    {
        ButilExitOnFailure(E_FILENOTFOUND == hr ? HRESULT_FROM_WIN32(ERROR_UNKNOWN_PRODUCT) : hr, "Failed to locate bundle uninstall key path.");
    }

    // If the bundle doesn't have the property defined, return ERROR_UNKNOWN_PROPERTY
    hr = RegGetType(hkBundle, wzAttribute, &dwType);
    ButilExitOnFailure(E_FILENOTFOUND == hr ? HRESULT_FROM_WIN32(ERROR_UNKNOWN_PROPERTY) : hr, "Failed to locate bundle property.");

    switch (dwType)
    {
        case REG_SZ:
            hr = RegReadString(hkBundle, wzAttribute, &sczValue);
            ButilExitOnFailure(hr, "Failed to read string property.");
            break;
        case REG_DWORD:
            hr = RegReadNumber(hkBundle, wzAttribute, &dwValue);
            ButilExitOnFailure(hr, "Failed to read dword property.");

            hr = StrAllocFormatted(&sczValue, L"%d", dwValue);
            ButilExitOnFailure(hr, "Failed to format dword property as string.");
            break;
        default:
            ButilExitOnFailure(hr = E_NOTIMPL, "Reading bundle info of type 0x%x not implemented.", dwType);

    }

    hr = ::StringCchLengthW(sczValue, STRSAFE_MAX_CCH, reinterpret_cast<UINT_PTR*>(&cchSource));
    ButilExitOnFailure(hr, "Failed to calculate length of string");

    if (lpValueBuf)
    {
        // cchSource is the length of the string not including the terminating null character
        if (*pcchValueBuf <= cchSource)
        {
            *pcchValueBuf = ++cchSource;
            ButilExitOnFailure(hr = HRESULT_FROM_WIN32(ERROR_MORE_DATA), "A buffer is too small to hold the requested data.");
        }

        hr = ::StringCchCatNExW(lpValueBuf, *pcchValueBuf, sczValue, cchSource, NULL, NULL, STRSAFE_FILL_BEHIND_NULL);
        ButilExitOnFailure(hr, "Failed to copy the property value to the output buffer.");
        
        *pcchValueBuf = cchSource++;        
    }

LExit:
    ReleaseRegKey(hkBundle);
    ReleaseStr(sczValue);

    return hr;
}


extern "C" HRESULT DAPI BundleGetBundleNumericVariable(
    __in LPCWSTR wzBundleId,
    __in LPCWSTR wzAttribute,
    __inout_opt LONGLONG * pllValue
)
{
    HRESULT hr = S_OK;
    DWORD dwType;
    LPWSTR sczValue = NULL;
    DWORD cbData = 0;

    if (wzBundleId && *wzBundleId && wzAttribute && *wzAttribute)
    {
        hr = BundleGetBundleVariable(wzBundleId, wzAttribute, &dwType, NULL, &cbData);
        ButilExitOnFailure(hr, "Failed to read variable type.");

        switch (dwType)
        {
        case REG_QWORD:
            Assert(cbData == sizeof(LONGLONG));

            hr = BundleGetBundleVariable(wzBundleId, wzAttribute, &dwType, pllValue, &cbData);
            ButilExitOnFailure(hr, "Failed to read variable QWORD.");
            break;

        case REG_SZ:
            hr = StrAlloc(&sczValue, cbData);
            ButilExitOnFailure(hr, "Failed to allocate memory for string variable.");

            hr = BundleGetBundleVariable(wzBundleId, wzAttribute, &dwType, &sczValue, &cbData);
            ButilExitOnFailure(hr, "Failed to read variable string.");

            hr = StrStringToInt64(sczValue, 0, pllValue);
            if (FAILED(hr))
            {
                hr = DISP_E_TYPEMISMATCH;
            }
            break;
        default:
            ButilExitOnFailure(hr = E_NOTIMPL, "Reading bundle variable of type 0x%x as number is not implemented.", dwType);
        }
    }
    else
    {
        hr = E_INVALIDARG;
        ButilExitOnFailure(hr, "Invalid agruments to read variable string.");
    }
LExit:
    ReleaseStr(sczValue);

    return hr;
}

extern "C" HRESULT DAPI BundleGetBundleStringVariable(
    __in LPCWSTR wzBundleId,
    __in LPCWSTR wzAttribute,
    __out_bcount_opt(*pcchData) LPWSTR lpData,
    __inout_opt LPDWORD pcchData
)
{
    HRESULT hr = S_OK;
    DWORD dwType;
    LONGLONG llValue = 0;
    LPWSTR sczValue = NULL;
    DWORD cbData = 0;

    if (wzBundleId && *wzBundleId && wzAttribute && *wzAttribute)
    {
        hr = BundleGetBundleVariable(wzBundleId, wzAttribute, &dwType, NULL, &cbData);
        ButilExitOnFailure(hr, "Failed to read shared variable type.");

        switch (dwType)
        {
        case REG_QWORD:
            Assert(cbData == sizeof(LONGLONG));

            hr = BundleGetBundleVariable(wzBundleId, wzAttribute, &dwType, &llValue, &cbData);
            ButilExitOnFailure(hr, "Failed to read shared variable QWORD.");

            hr = StrAllocFormatted(&sczValue, L"%d", llValue);
            ButilExitOnFailure(hr, "Failed to allocate shared string variable.");

            hr = ::StringCchLengthW(sczValue, STRSAFE_MAX_CCH, reinterpret_cast<UINT_PTR*>(&cbData));
            ButilExitOnFailure(hr, "Failed to calculate length of string");

            if (lpData)
            {
                if (*pcchData <= cbData)
                {
                    *pcchData = ++cbData;
                    ButilExitOnFailure(hr = HRESULT_FROM_WIN32(ERROR_MORE_DATA), "A buffer is too small to hold the requested data.");
                }
                // Safe to copy
                hr = ::StringCchCatNExW(*reinterpret_cast<LPWSTR*>(lpData), *pcchData, sczValue, cbData, NULL, NULL, STRSAFE_FILL_BEHIND_NULL);
                ButilExitOnFailure(hr, "Failed to copy the shared variable value to the output buffer.");

                *pcchData = cbData++;
            }
            else
            {
                *pcchData = cbData++;
                ButilExitOnFailure(hr = HRESULT_FROM_WIN32(ERROR_SUCCESS), "A buffer is too small to hold the requested data.");
            }
            break;

        case REG_SZ:
            //hr = StrAlloc(&sczValue, cbData);
            //ExitOnFailure(hr, "Failed to allocate memory for shared string variable.");

            hr = BundleGetBundleVariable(wzBundleId, wzAttribute, &dwType, lpData, pcchData);
            ButilExitOnFailure(hr, "Failed to read shared variable string.");

            break;
        default:
            ButilExitOnFailure(hr = E_NOTIMPL, "Reading bundle shared variable of type 0x%x as string is not implemented.", dwType);
        }
    }
    else
    {
        hr = E_INVALIDARG;
    }
LExit:
    ReleaseStr(sczValue);

    return hr;
}

/********************************************************************
BundleGetBundleVariable - Queries the bundle installation metadata for a given variable

RETURNS:
E_INVALIDARG
An invalid parameter was passed to the function.
HRESULT_FROM_WIN32(ERROR_UNKNOWN_PRODUCT)
The bundle is not installed
HRESULT_FROM_WIN32(ERROR_UNKNOWN_PROPERTY)
The shared variable is unrecognized
HRESULT_FROM_WIN32(ERROR_MORE_DATA)
A buffer is too small to hold the requested data.
HRESULT_FROM_WIN32(ERROR_SUCCESS)
When probing for a string variable, if no data pointer is passed but the size is and the variable is found.
E_NOTIMPL:
Tried to read a bundle variable for a type which has not been implemented

All other returns are unexpected returns from other dutil methods.
********************************************************************/
extern "C" HRESULT DAPI BundleGetBundleVariable(
    __in LPCWSTR wzBundleId,
    __in LPCWSTR wzAttribute,
    __out_opt LPDWORD pdwType,
    __out_bcount_opt(*pcbData) PVOID pvData,
    __inout_opt LPDWORD pcbData
)
{
    Assert(wzBundleId && wzAttribute);

    HRESULT hr = S_OK;
    BUNDLE_INSTALL_CONTEXT context = BUNDLE_INSTALL_CONTEXT_MACHINE;
    LPWSTR sczValue = NULL;
    HKEY hkBundle = NULL;
    DWORD cbData = 0;
    DWORD dwType = 0;
    DWORD64 qwValue = 0;

    if (!wzBundleId || !wzAttribute || (pvData && !pcbData))
    {
        ButilExitOnFailure(hr = E_INVALIDARG, "An invalid parameter was passed to the function.");
    }

    if (FAILED(hr = OpenBundleKey(wzBundleId, context = BUNDLE_INSTALL_CONTEXT_MACHINE, BUNDLE_REGISTRATION_REGISTRY_BUNDLE_VARIABLE_KEY, &hkBundle)) &&
        FAILED(hr = OpenBundleKey(wzBundleId, context = BUNDLE_INSTALL_CONTEXT_USER, BUNDLE_REGISTRATION_REGISTRY_BUNDLE_VARIABLE_KEY, &hkBundle)))
    {
        ButilExitOnFailure(E_FILENOTFOUND == hr ? HRESULT_FROM_WIN32(ERROR_UNKNOWN_PRODUCT) : hr, "Failed to locate bundle uninstall key variable path.");
    }

    // If the bundle doesn't have the shared variable defined, return ERROR_UNKNOWN_PROPERTY
    hr = RegGetType(hkBundle, wzAttribute, &dwType);
    ButilExitOnFailure(E_FILENOTFOUND == hr ? HRESULT_FROM_WIN32(ERROR_UNKNOWN_PROPERTY) : hr, "Failed to locate bundle shared variable.");

    if (pdwType)
    {
        *pdwType = dwType;
    }

    switch (dwType)
    {
    case REG_SZ:
        if (pcbData)
        {
            hr = RegReadString(hkBundle, wzAttribute, &sczValue);
            ButilExitOnFailure(hr, "Failed to read string shared variable.");

            hr = ::StringCchLengthW(sczValue, STRSAFE_MAX_CCH, reinterpret_cast<UINT_PTR*>(&cbData));
            ButilExitOnFailure(hr, "Failed to calculate length of string");

            if (pvData)
            {
                if (*pcbData <= cbData)
                {
                    *pcbData = ++cbData;
                    ButilExitOnFailure(hr = HRESULT_FROM_WIN32(ERROR_MORE_DATA), "A buffer is too small to hold the requested data.");
                }
                // Safe to copy
                hr = ::StringCchCatNExW(*reinterpret_cast<LPWSTR*>(pvData), *pcbData, sczValue, cbData, NULL, NULL, STRSAFE_FILL_BEHIND_NULL);
                ButilExitOnFailure(hr, "Failed to copy the shared variable value to the output buffer.");

                *pcbData = ++cbData;
            }
            else
            {
                *pcbData = ++cbData;
                ButilExitOnFailure(hr = HRESULT_FROM_WIN32(ERROR_SUCCESS), "A buffer is too small to hold the requested data.");
            }
        }

        break;
    case REG_QWORD:
        if (pcbData)
        {
            hr = RegReadQword(hkBundle, wzAttribute, &qwValue);
            ButilExitOnFailure(hr, "Failed to read qword shared variable.");

            if (pvData)
            {
                if (*pcbData < sizeof(DWORD64))
                {
                    *pcbData = sizeof(DWORD64);
                    ButilExitOnFailure(hr = HRESULT_FROM_WIN32(ERROR_MORE_DATA), "A numeric buffer is too small to hold the requested data.");
                }
                else if (*pcbData > sizeof(DWORD64))
                {
                    *pcbData = sizeof(DWORD64);
                    ButilExitOnFailure(hr = E_INVALIDARG, "A numeric buffer is too large to hold the requested data.");
                }

                *reinterpret_cast<DWORD64*>(pvData) = qwValue;
            }
            else
            {
                if (*pcbData < sizeof(DWORD64))
                {
                    *pcbData = sizeof(DWORD64);
                    ButilExitOnFailure(hr = HRESULT_FROM_WIN32(ERROR_SUCCESS), "A numeric buffer is too small to hold the requested data.");
                }
            }
        }
        break;
    default:
        ButilExitOnFailure(hr = E_NOTIMPL, "Reading bundle shared variable of type 0x%x not implemented.", dwType);

    }

LExit:
    ReleaseRegKey(hkBundle);
    ReleaseStr(sczValue);

    return hr;
}

HRESULT DAPI BundleEnumRelatedBundle(
  __in_z LPCWSTR wzUpgradeCode,
  __in BUNDLE_INSTALL_CONTEXT context,
  __inout PDWORD pdwStartIndex,
  __out_ecount(MAX_GUID_CHARS+1) LPWSTR lpBundleIdBuf
    )
{
    HRESULT hr = S_OK;
    HKEY hkRoot = BUNDLE_INSTALL_CONTEXT_USER == context ? HKEY_CURRENT_USER : HKEY_LOCAL_MACHINE;
    HKEY hkUninstall = NULL;
    HKEY hkBundle = NULL;
    LPWSTR sczUninstallSubKey = NULL;
    DWORD cchUninstallSubKey = 0;
    LPWSTR sczUninstallSubKeyPath = NULL;
    LPWSTR sczValue = NULL;
    DWORD dwType = 0;

    LPWSTR* rgsczBundleUpgradeCodes = NULL;
    DWORD cBundleUpgradeCodes = 0;
    BOOL fUpgradeCodeFound = FALSE;

    if (!wzUpgradeCode || !lpBundleIdBuf || !pdwStartIndex)
    {
        ButilExitOnFailure(hr = E_INVALIDARG, "An invalid parameter was passed to the function.");
    }

    hr = RegOpen(hkRoot, BUNDLE_REGISTRATION_REGISTRY_UNINSTALL_KEY, KEY_READ, &hkUninstall);
    ButilExitOnFailure(hr, "Failed to open bundle uninstall key path.");

    for (DWORD dwIndex = *pdwStartIndex; !fUpgradeCodeFound; dwIndex++)
    {
        hr = RegKeyEnum(hkUninstall, dwIndex, &sczUninstallSubKey);
        ButilExitOnFailure(hr, "Failed to enumerate bundle uninstall key path.");

        hr = StrAllocFormatted(&sczUninstallSubKeyPath, L"%ls\\%ls", BUNDLE_REGISTRATION_REGISTRY_UNINSTALL_KEY, sczUninstallSubKey);
        ButilExitOnFailure(hr, "Failed to allocate bundle uninstall key path.");
        
        hr = RegOpen(hkRoot, sczUninstallSubKeyPath, KEY_READ, &hkBundle);
        ButilExitOnFailure(hr, "Failed to open uninstall key path.");

        // If it's a bundle, it should have a BundleUpgradeCode value of type REG_SZ (old) or REG_MULTI_SZ
        hr = RegGetType(hkBundle, BUNDLE_REGISTRATION_REGISTRY_BUNDLE_UPGRADE_CODE, &dwType);
        if (FAILED(hr))
        {
            ReleaseRegKey(hkBundle);
            ReleaseNullStr(sczUninstallSubKey);
            ReleaseNullStr(sczUninstallSubKeyPath);
            // Not a bundle
            continue;
        }

        switch (dwType)
        {
            case REG_SZ:
                hr = RegReadString(hkBundle, BUNDLE_REGISTRATION_REGISTRY_BUNDLE_UPGRADE_CODE, &sczValue);
                ButilExitOnFailure(hr, "Failed to read BundleUpgradeCode string property.");
                if (CSTR_EQUAL == ::CompareStringW(LOCALE_INVARIANT, NORM_IGNORECASE, sczValue, -1, wzUpgradeCode, -1))
                {
                    *pdwStartIndex = dwIndex;
                    fUpgradeCodeFound = TRUE;
                    break;
                }

                ReleaseNullStr(sczValue);

                break;
            case REG_MULTI_SZ:
                hr = RegReadStringArray(hkBundle, BUNDLE_REGISTRATION_REGISTRY_BUNDLE_UPGRADE_CODE, &rgsczBundleUpgradeCodes, &cBundleUpgradeCodes);
                ButilExitOnFailure(hr, "Failed to read BundleUpgradeCode  multi-string property.");

                for (DWORD i = 0; i < cBundleUpgradeCodes; i++)
                {
                    LPWSTR wzBundleUpgradeCode = rgsczBundleUpgradeCodes[i];
                    if (wzBundleUpgradeCode && *wzBundleUpgradeCode)
                    {
                        if (CSTR_EQUAL == ::CompareStringW(LOCALE_INVARIANT, NORM_IGNORECASE, wzBundleUpgradeCode, -1, wzUpgradeCode, -1))
                        {
                            *pdwStartIndex = dwIndex;
                            fUpgradeCodeFound = TRUE;
                            break;
                        }
                    }
                }
                ReleaseNullStrArray(rgsczBundleUpgradeCodes, cBundleUpgradeCodes);

                break;

            default:
                ButilExitOnFailure(hr = E_NOTIMPL, "BundleUpgradeCode of type 0x%x not implemented.", dwType);

        }

        if (fUpgradeCodeFound)
        {
            if (lpBundleIdBuf)
            {
                hr = ::StringCchLengthW(sczUninstallSubKey, STRSAFE_MAX_CCH, reinterpret_cast<UINT_PTR*>(&cchUninstallSubKey));
                ButilExitOnFailure(hr, "Failed to calculate length of string");

                hr = ::StringCchCopyNExW(lpBundleIdBuf, MAX_GUID_CHARS + 1, sczUninstallSubKey, cchUninstallSubKey, NULL, NULL, STRSAFE_FILL_BEHIND_NULL);
                ButilExitOnFailure(hr, "Failed to copy the property value to the output buffer.");
            }

            break;
        }

        // Cleanup before next iteration
        ReleaseRegKey(hkBundle);
        ReleaseNullStr(sczUninstallSubKey);
        ReleaseNullStr(sczUninstallSubKeyPath);
    }

LExit:
    ReleaseStr(sczValue);
    ReleaseStr(sczUninstallSubKey);
    ReleaseStr(sczUninstallSubKeyPath);
    ReleaseRegKey(hkBundle);
    ReleaseRegKey(hkUninstall);
    ReleaseStrArray(rgsczBundleUpgradeCodes, cBundleUpgradeCodes);

    return hr;
}


HRESULT OpenBundleKey(
    __in_z LPCWSTR wzBundleId,
    __in BUNDLE_INSTALL_CONTEXT context,
    __in_opt LPCWSTR szSubKey,
    __inout HKEY *key)
{
    Assert(key && wzBundleId);
    AssertSz(NULL == *key, "*key should be null");

    HRESULT hr = S_OK;
    HKEY hkRoot = BUNDLE_INSTALL_CONTEXT_USER == context ? HKEY_CURRENT_USER : HKEY_LOCAL_MACHINE;
    LPWSTR sczKeypath = NULL;

    if (szSubKey)
    {
        hr = StrAllocFormatted(&sczKeypath, L"%ls\\%ls\\%ls", BUNDLE_REGISTRATION_REGISTRY_UNINSTALL_KEY, wzBundleId, szSubKey);
    }
    else
    {
        hr = StrAllocFormatted(&sczKeypath, L"%ls\\%ls", BUNDLE_REGISTRATION_REGISTRY_UNINSTALL_KEY, wzBundleId);
    }
    ButilExitOnFailure(hr, "Failed to allocate bundle uninstall key path.");
    
    hr = RegOpen(hkRoot, sczKeypath, KEY_READ, key);
    ButilExitOnFailure(hr, "Failed to open bundle uninstall key path.");

LExit:
    ReleaseStr(sczKeypath);

    return hr;
}
