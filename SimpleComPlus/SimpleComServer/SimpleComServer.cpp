// SimpleComServer.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"

#include "DispatchCalculatorComponent_h.h"

#include <new>  // For std::nothrow

#include <strsafe.h>    // For StringCchPrintfW
#include <OleCtl.h> // For SELFREG_E_CLASS

extern "C"
{
extern const IID DIID_IDispCalculatorComponent;
extern const IID LIBID_DispatchCalculatorComponentLib;
extern const CLSID CLSID_DispatchCalculatorComponent;
}

extern HMODULE g_thisModule;

LONG ModuleRefCount = 0;

// Some credit to https://devblogs.microsoft.com/oldnewthing/20130612-00/?p=4103
class DispatchCalculatorComponent : public IDispatch
{
private:
    ~DispatchCalculatorComponent()
    {
        InterlockedDecrement(&ModuleRefCount);
    }

public:
    DispatchCalculatorComponent() : m_cRef(1)
    {
        InterlockedIncrement(&ModuleRefCount);
    }

    // *** IUnknown ***
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv)
    {
        *ppv = nullptr;

        HRESULT hr = E_NOINTERFACE;

        if (riid == IID_IUnknown || riid == IID_IDispatch
            || riid == DIID_IDispCalculatorComponent
            )
        {
            *ppv = static_cast<IDispatch*>(this);
            AddRef();
            hr = S_OK;
        }

        return hr;
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&m_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        LONG cRef = InterlockedDecrement(&m_cRef);

        if (cRef == 0) delete this;

        return cRef;
    }

    // *** IDispatch ***

    IFACEMETHODIMP GetTypeInfoCount(UINT* pctinfo)
    {
        *pctinfo = 0;

        return E_NOTIMPL;
    }

    IFACEMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
    {
        *ppTInfo = nullptr;

        return E_NOTIMPL;
    }

    IFACEMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid)
    {
        return E_NOTIMPL;
    }

    IFACEMETHODIMP Invoke(
        DISPID dispid,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS* pdispparams,
        VARIANT* pVarResult,
        EXCEPINFO* pexecpinfo,
        UINT* puArgErr)
    {
        if (pVarResult) VariantInit(pVarResult);

        return SimpleInvoke(dispid, pdispparams, pVarResult);
    }

private:
    LONG m_cRef;

    HRESULT SimpleInvoke(DISPID dispid, DISPPARAMS* pdispparams, VARIANT* pvarResult)
    {
        switch (dispid)
        {
        case 1 /* Add */:
            // NOTE: Reverse order [right-to-left] inside the Dispatch Parameters!
            return ImplementAdd(pdispparams->rgvarg[2].intVal, pdispparams->rgvarg[1].intVal, pdispparams->rgvarg[0].pintVal);
            break;
        }

        return S_OK;
    }

    HRESULT ImplementAdd(int a, int b, int* value)
    {
        *value = a + b;
        return S_OK;
    }
};

// Credit https://stackoverflow.com/q/17041907/6586652
class DispatchCalculatorComponentFactory : public IClassFactory
{
private:
    ~DispatchCalculatorComponentFactory()
    {
        InterlockedDecrement(&ModuleRefCount);
    }

public:
    DispatchCalculatorComponentFactory() : m_cRef(1)
    {
        InterlockedIncrement(&ModuleRefCount);
    }

    // ** IUnknown **
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppObj)
    {
        *ppObj = nullptr;

        HRESULT hr = E_NOINTERFACE;

        if (riid == IID_IUnknown || riid == IID_IClassFactory
            )
        {
            *ppObj = static_cast<IClassFactory*>(this);
            AddRef();
            hr = S_OK;
        }

        return hr;
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&m_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        long cRef = InterlockedDecrement(&m_cRef);

        if (cRef == 0) delete this;

        return cRef;
    }

    // *** IClassFactory ***

    IFACEMETHODIMP CreateInstance(
        IUnknown* pUnknownOuter,
        const IID& iid,
        void** ppv)
    {
        *ppv = nullptr;

        if (pUnknownOuter != nullptr)
        {
            return CLASS_E_NOAGGREGATION;
        }

        DispatchCalculatorComponent* pObject = new (std::nothrow) DispatchCalculatorComponent();
        if (pObject == nullptr)
        {
            return E_OUTOFMEMORY;
        }

        HRESULT hr = pObject->QueryInterface(iid, ppv);
        pObject->Release();

        return hr;
    }

    IFACEMETHODIMP LockServer(BOOL bLock)
    {
        if (bLock)
        {
            InterlockedIncrement(&ModuleRefCount);
        }
        else
        {
            InterlockedDecrement(&ModuleRefCount);
        }
        return S_OK;
    }

private:
    LONG m_cRef;
};

STDAPI DllGetClassObject(const CLSID& clsid,
    const IID& iid,
    void** ppv)
{
    *ppv = nullptr;
    OutputDebugStringW(L"In DllGetClassObject()");
    if (clsid == CLSID_DispatchCalculatorComponent)
    {

        DispatchCalculatorComponentFactory* pFact = new (std::nothrow) DispatchCalculatorComponentFactory();
        if (pFact == nullptr)
        {
            return E_OUTOFMEMORY;   // avoid release on null
        }
        HRESULT hr = pFact->QueryInterface(iid, ppv);

        pFact->Release();

        return hr;
    }

    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow()
{
    // Use the CompareExchange with unlikely value trick for reading
    LONG currentRefCount = InterlockedCompareExchange(&ModuleRefCount, -1L, -1L);
    return currentRefCount == 0 ? S_OK : S_FALSE;
}

bool HelpWriteKey(
    LPCWSTR pathBelowClasses,
    LPCWSTR valueName,
    DWORD dwType,
    const void* Data,
    DWORD dwDataSize)
{
    WCHAR wszCompletePath[MAX_PATH];
    HRESULT hr = StringCchPrintfW(wszCompletePath, MAX_PATH, LR"(Software\Classes\%s)", pathBelowClasses);
    if (FAILED(hr))
        return false;

    HKEY hk;
    if (ERROR_SUCCESS != RegCreateKeyExW(HKEY_LOCAL_MACHINE, wszCompletePath, 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, nullptr, &hk, nullptr))
    {
        return false;
    }

    bool setSuccess = (ERROR_SUCCESS == RegSetValueEx(hk, valueName, 0, dwType, static_cast<const BYTE*>(Data), dwDataSize));

    bool closeSuccess = (ERROR_SUCCESS == RegCloseKey(hk));

    return setSuccess && closeSuccess;
}

STDAPI DllRegisterServer()
{
    OutputDebugStringW(L"In DllRegisterServer()");

    // This implementation is inspired by the registry.cpp implementation from
    // https://www.codeguru.com/cpp/com-tech/activex/tutorials/article.php/c5567/Step-by-Step-COM-Tutorial.htm

    WCHAR wszPayload[MAX_PATH];
    LPOLESTR lposClsId;
    HRESULT hr = StringFromCLSID(CLSID_DispatchCalculatorComponent, &lposClsId);
    if (SUCCEEDED(hr))
    {

        // Write the human readable description.
        hr = StringCchPrintfW(wszPayload, MAX_PATH, L"DispatchCalculatorComponent");
        if (SUCCEEDED(hr))
        {
            WCHAR wszDescKey[MAX_PATH];
            hr = StringCchPrintfW(wszDescKey, MAX_PATH, LR"(CLSID\%s)", lposClsId);
            if (SUCCEEDED(hr))
                if (!HelpWriteKey(wszDescKey, nullptr, REG_SZ, wszPayload, lstrlen(wszPayload) * sizeof(WCHAR)))
                    hr = SELFREG_E_CLASS;
        }

        // Write the InProcServer32 value
        if (SUCCEEDED(hr))
        {
            WCHAR wszInProc[MAX_PATH];
            hr = StringCchPrintfW(wszInProc, MAX_PATH, LR"(CLSID\%s\InprocServer32)", lposClsId);

            if (SUCCEEDED(hr))
            {
                SetLastError(ERROR_SUCCESS);
                if (GetModuleFileNameW(g_thisModule, wszPayload, MAX_PATH) == 0 ||
                    GetLastError() == ERROR_INSUFFICIENT_BUFFER)
                {
                    hr = SELFREG_E_CLASS;
                }
            }

            if (SUCCEEDED(hr))
                if (!HelpWriteKey(wszInProc, nullptr, REG_SZ, wszPayload, lstrlen(wszPayload) * sizeof(WCHAR)))
                    hr = SELFREG_E_CLASS;
        }

        // We have no ProgId.

        CoTaskMemFree(lposClsId);
        lposClsId = nullptr;
    }

    // We also need to write an interface.
    LPOLESTR lposIid = nullptr;
    if (SUCCEEDED(hr))
        hr = StringFromIID(DIID_IDispCalculatorComponent, &lposIid);

    if (SUCCEEDED(hr))
    {
        // Write the human readable description.
        hr = StringCchPrintfW(wszPayload, MAX_PATH, L"IDispCalculatorComponent");
        if (SUCCEEDED(hr))
        {
            WCHAR wszInterfDescKey[MAX_PATH];
            hr = StringCchPrintfW(wszInterfDescKey, MAX_PATH, LR"(Interface\%s)", lposIid);

            if (SUCCEEDED(hr))
                if (!HelpWriteKey(wszInterfDescKey, nullptr, REG_SZ, wszPayload, lstrlen(wszPayload) * sizeof(WCHAR)))
                    hr = SELFREG_E_CLASS;
        }

        // Write the Proxy/Stub value.
        if (SUCCEEDED(hr))
        {
            WCHAR wszProxyStub[MAX_PATH];
            hr = StringCchPrintfW(wszProxyStub, MAX_PATH, LR"(Interface\%s\ProxyStubClsid32)", lposIid);

            const WCHAR* PSDispatch = L"{00020420-0000-0000-C000-000000000046}";
            if (SUCCEEDED(hr))
                if (!HelpWriteKey(wszProxyStub, nullptr, REG_SZ, PSDispatch, lstrlen(PSDispatch) * sizeof(WCHAR)))
                    hr = SELFREG_E_CLASS;
        }

        // Add the Typelib since it doesn't seem to work without it
        if (SUCCEEDED(hr))
        {
            WCHAR wszTypeLib[MAX_PATH];
            hr = StringCchPrintfW(wszTypeLib, MAX_PATH, LR"(Interface\%s\TypeLib)", lposIid);

            if (SUCCEEDED(hr))
            {
                LPOLESTR lposLibId = nullptr;
                hr = StringFromIID(LIBID_DispatchCalculatorComponentLib, &lposLibId);
                if (SUCCEEDED(hr))
                {
                    if (!HelpWriteKey(wszTypeLib, nullptr, REG_SZ, lposLibId, lstrlen(lposLibId) * sizeof(OLECHAR)))
                        hr = SELFREG_E_CLASS;

                    // TODO: If still not working, write the Version named value.

                    CoTaskMemFree(lposLibId);
                    lposLibId = nullptr;
                }
            }
        }

        CoTaskMemFree(lposIid);
        lposIid = nullptr;
    }

    // Canonicalize our result.
    return SUCCEEDED(hr) ? S_OK : SELFREG_E_CLASS;
}

STDAPI DllUnregisterServer()
{
    OutputDebugStringW(L"In DllUnregisterServer()");

    // No ProgId entry

    WCHAR wszKeyPath[MAX_PATH];

    LPOLESTR lposClsId;
    HRESULT hr = StringFromCLSID(CLSID_DispatchCalculatorComponent, &lposClsId);

    //bool foundOrphans = false;
    if (SUCCEEDED(hr))
    {
        hr = StringCchPrintfW(wszKeyPath, MAX_PATH, LR"(Software\Classes\CLSID\%s\InprocServer32)", lposClsId);
        if (SUCCEEDED(hr))
            if (ERROR_SUCCESS != RegDeleteKeyW(HKEY_LOCAL_MACHINE, wszKeyPath))
                hr = SELFREG_E_CLASS;

        if (SUCCEEDED(hr))
            hr = StringCchPrintfW(wszKeyPath, MAX_PATH, LR"(Software\Classes\CLSID\%s)", lposClsId);
        if (SUCCEEDED(hr))
            if (ERROR_SUCCESS != RegDeleteKeyW(HKEY_LOCAL_MACHINE, wszKeyPath))
                hr = SELFREG_E_CLASS;

        CoTaskMemFree(lposClsId);
        lposClsId = nullptr;
    }

    LPOLESTR lposIid = nullptr;
    if (SUCCEEDED(hr))
        hr = StringFromIID(DIID_IDispCalculatorComponent, &lposIid);

    if (SUCCEEDED(hr))
    {
        hr = StringCchPrintfW(wszKeyPath, MAX_PATH, LR"(Software\Classes\Interface\%s\ProxyStubClsid32)", lposIid);
        if (SUCCEEDED(hr))
            if (ERROR_SUCCESS != RegDeleteKeyW(HKEY_LOCAL_MACHINE, wszKeyPath))
                hr = SELFREG_E_CLASS;

        if (SUCCEEDED(hr))
            hr = StringCchPrintfW(wszKeyPath, MAX_PATH, LR"(Software\Classes\Interface\%s\TypeLib)", lposIid);
        if (SUCCEEDED(hr))
            if (ERROR_SUCCESS != RegDeleteKeyW(HKEY_LOCAL_MACHINE, wszKeyPath))
                hr = SELFREG_E_CLASS;

        if (SUCCEEDED(hr))
            hr = StringCchPrintfW(wszKeyPath, MAX_PATH, LR"(Software\Classes\Interface\%s)", lposIid);
        if (SUCCEEDED(hr))
            if (ERROR_SUCCESS != RegDeleteKeyW(HKEY_LOCAL_MACHINE, wszKeyPath))
                hr = SELFREG_E_CLASS;

        CoTaskMemFree(lposIid);
        lposIid = nullptr;
    }

    // Canonicalize
    return SUCCEEDED(hr) ? S_OK : SELFREG_E_CLASS;
}
