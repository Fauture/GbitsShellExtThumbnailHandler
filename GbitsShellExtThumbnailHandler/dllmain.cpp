#include <windows.h>
#include <Guiddef.h>
#include <shlobj.h>                 // For SHChangeNotify
#include "ClassFactory.h"           // For the class factory
#include "Reg.h"


// When you write your own handler, you must create a new CLSID by using the 
// "Create GUID" tool in the Tools menu, and specify the CLSID value here.
const CLSID CLSID_GbitsThumbnailProvider =
// {16469314-EDDD-404F-A8CC-3F9C0A63497B}
{ 0x4ECEE19A, 0x5151, 0xa301, { 0x51, 0x51, 0x3a, 0x01, 0x4e, 0xce, 0xe1, 0x9a } };
const CLSID CLSID_GbitsJdThumbnailProvider =
// {16469314-EDDD-404F-A8CC-3F9C0A63497B}
{ 0x4ECEE19A, 0x5151, 0xa302, { 0x51, 0x51, 0x3a, 0x02, 0x4e, 0xce, 0xe1, 0x9a } };
const CLSID CLSID_GbitsDkThumbnailProvider =
// {16469314-EDDD-404F-A8CC-3F9C0A63497B}
{ 0x4ECEE19A, 0x5151, 0xa303, { 0x51, 0x51, 0x3a, 0x03, 0x4e, 0xce, 0xe1, 0x9a } };


HINSTANCE   g_hInst     = NULL;
long        g_cDllRef   = 0;

BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:
        // Hold the instance of this DLL module, we will use it to get the 
        // path of the DLL to register the component.
        g_hInst = hModule;
        DisableThreadLibraryCalls(hModule);
        break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}


//
//   FUNCTION: DllGetClassObject
//
//   PURPOSE: Create the class factory and query to the specific interface.
//
//   PARAMETERS:
//   * rclsid - The CLSID that will associate the correct data and code.
//   * riid - A reference to the identifier of the interface that the caller 
//     is to use to communicate with the class object.
//   * ppv - The address of a pointer variable that receives the interface 
//     pointer requested in riid. Upon successful return, *ppv contains the 
//     requested interface pointer. If an error occurs, the interface pointer 
//     is NULL. 
//
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
    HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

    if (IsEqualCLSID(CLSID_GbitsThumbnailProvider, rclsid))
    {
        hr = E_OUTOFMEMORY;

        ClassFactory *pClassFactory = new ClassFactory();
        if (pClassFactory)
        {
            hr = pClassFactory->QueryInterface(riid, ppv);
            pClassFactory->Release();
        }
    }

    if (IsEqualCLSID(CLSID_GbitsDkThumbnailProvider, rclsid))
    {
        hr = E_OUTOFMEMORY;

        ClassFactory* pClassFactory = new ClassFactory();
        if (pClassFactory)
        {
            hr = pClassFactory->QueryInterface(riid, ppv);
            pClassFactory->Release();
        }
    }

    if (IsEqualCLSID(CLSID_GbitsJdThumbnailProvider, rclsid))
    {
        hr = E_OUTOFMEMORY;

        ClassFactory* pClassFactory = new ClassFactory();
        if (pClassFactory)
        {
            hr = pClassFactory->QueryInterface(riid, ppv);
            pClassFactory->Release();
        }
    }

    return hr;
}


//
//   FUNCTION: DllCanUnloadNow
//
//   PURPOSE: Check if we can unload the component from the memory.
//
//   NOTE: The component can be unloaded from the memory when its reference 
//   count is zero (i.e. nobody is still using the component).
// 
STDAPI DllCanUnloadNow(void)
{
    return g_cDllRef > 0 ? S_FALSE : S_OK;
}

//
//   FUNCTION: DllRegisterServer
//
//   PURPOSE: Register the COM server and the thumbnail handler.
// 
STDAPI DllRegisterServer(void)
{
    HRESULT hr;

    wchar_t szModule[MAX_PATH];
    if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        return hr;
    }

    // Register the component.
    hr = RegisterInprocServer(szModule, CLSID_GbitsThumbnailProvider,
        L"GbitsShellExtThumbnailHandler.GbitsThumbnailProvider Class", 
        L"Apartment");
    if (SUCCEEDED(hr))
    {
        // Register the thumbnail handler. 
        hr = RegisterShellExtThumbnailHandler(L".gbits", 
            CLSID_GbitsThumbnailProvider);
        if (SUCCEEDED(hr))
        {
            // This tells the shell to invalidate the thumbnail cache. It is 
            // important because any .gbits files viewed before registering 
            // this handler would otherwise show cached blank thumbnails.
            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
        }
    }

    // Register the component.
    hr = RegisterInprocServer(szModule, CLSID_GbitsJdThumbnailProvider,
        L"GbitsShellExtThumbnailHandler.GbitsThumbnailProvider Class",
        L"Apartment");
    if (SUCCEEDED(hr))
    {
        // Register the thumbnail handler. 
        hr = RegisterShellExtThumbnailHandler(L".gbits_jd",
            CLSID_GbitsJdThumbnailProvider);
        if (SUCCEEDED(hr))
        {
            // This tells the shell to invalidate the thumbnail cache. It is 
            // important because any .gbits_jd files viewed before registering 
            // this handler would otherwise show cached blank thumbnails.
            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
        }
    }

    // Register the component.
    hr = RegisterInprocServer(szModule, CLSID_GbitsDkThumbnailProvider,
        L"GbitsShellExtThumbnailHandler.GbitsThumbnailProvider Class",
        L"Apartment");
    if (SUCCEEDED(hr))
    {
        // Register the thumbnail handler. 
        hr = RegisterShellExtThumbnailHandler(L".gbits_dk",
            CLSID_GbitsDkThumbnailProvider);
        if (SUCCEEDED(hr))
        {
            // This tells the shell to invalidate the thumbnail cache. It is 
            // important because any .gbits_dk files viewed before registering 
            // this handler would otherwise show cached blank thumbnails.
            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
        }
    }

    return hr;
}


//
//   FUNCTION: DllUnregisterServer
//
//   PURPOSE: Unregister the COM server and the thumbnail handler.
// 
STDAPI DllUnregisterServer(void)
{
    HRESULT hr = S_OK;

    wchar_t szModule[MAX_PATH];
    if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        return hr;
    }

    // Unregister the component.
    hr = UnregisterInprocServer(CLSID_GbitsThumbnailProvider);
    if (SUCCEEDED(hr))
    {
        // Unregister the thumbnail handler.
        hr = UnregisterShellExtThumbnailHandler(L".gbits");
    }

    // Unregister the component.
    hr = UnregisterInprocServer(CLSID_GbitsJdThumbnailProvider);
    if (SUCCEEDED(hr))
    {
        // Unregister the thumbnail handler.
        hr = UnregisterShellExtThumbnailHandler(L".gbits_jd");
    }

    // Unregister the component.
    hr = UnregisterInprocServer(CLSID_GbitsDkThumbnailProvider);
    if (SUCCEEDED(hr))
    {
        // Unregister the thumbnail handler.
        hr = UnregisterShellExtThumbnailHandler(L".gbits_dk");
    }

    return hr;
}