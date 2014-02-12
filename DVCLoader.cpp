/**
 * DVCLoader: Sample Code for Loading a Dynamic Virtual Channel (DVC) Plugin
 *
 * Copyright 2014 Mike McDonald <mikem@nogginware.com>
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

#include <stdio.h>
#include <stdlib.h>
#include <objbase.h>
#include <comutil.h>
#include <tsvirtualchannels.h>

typedef HRESULT (__stdcall *LPDllGetClassObject)(REFCLSID rclsid, REFIID riid, LPVOID *ppv);
typedef HRESULT (__stdcall *LPVirtualChannelGetInstance)(REFIID riid, ULONG *pNumObjs, VOID **objArray);

void showUsage()
{
	printf("Usage: DVCLoader [ <dllname> | <clsid> | <dllname:clsid> ]\n");
}

BOOL parseCLSIDString(char* cstrCLSID, CLSID* clsid)
{
	_bstr_t bstrCLSID = cstrCLSID;

	return CLSIDFromString(bstrCLSID, clsid) == NOERROR ? TRUE : FALSE;
}

IWTSPlugin* loadDVCUsingCLSID(CLSID* clsid)
{
	IWTSPlugin* pWTSPlugin = NULL;
	HRESULT hr;

	hr = CoCreateInstance(*clsid, NULL, CLSCTX_ALL, IID_IWTSPlugin, (LPVOID*)&pWTSPlugin);
	if (hr != S_OK)
	{
		fprintf(stderr, "error calling CoCreateInstance (hr=%x)\n", hr);
		return NULL;
	}

	return pWTSPlugin;
}

IWTSPlugin* loadDVCUsingDllName(char* dllName)
{
	LPVirtualChannelGetInstance pVirtualChannelGetInstance;
	IWTSPlugin* pWTSPlugin = NULL;
	LPVOID *objArray = NULL;
	ULONG numObjs = 0;
	HMODULE hModule;
	HRESULT hr;

	hModule = LoadLibrary(dllName);
	if (hModule == NULL)
	{
		fprintf(stderr, "could not load '%s'\n", dllName);
		goto EXCEPTION;
	}

	pVirtualChannelGetInstance = (LPVirtualChannelGetInstance)GetProcAddress(hModule, "VirtualChannelGetInstance");
	if (!pVirtualChannelGetInstance)
	{
		fprintf(stderr, "could not retrieve VirtualChannelGetInstance entry point\n");
		goto EXCEPTION;
	}

	hr = pVirtualChannelGetInstance(IID_IWTSPlugin, &numObjs, NULL);
	if (hr != S_OK)
	{
		fprintf(stderr, "error calling VirtualChannelGetInstance entry point (hr=%x)\n", hr);
		goto EXCEPTION;
	}

	if (numObjs == 0)
	{
		fprintf(stderr, "call to VirtualChannelGetInstance returned no objects\n");
		goto EXCEPTION;
	}

	objArray = (LPVOID*)malloc(numObjs * sizeof(LPVOID));
	if (objArray == NULL)
	{
		fprintf(stderr, "could not allocate memory for object array\n");
		goto EXCEPTION;
	}

	hr = pVirtualChannelGetInstance(IID_IWTSPlugin, &numObjs, objArray);
	if (hr != S_OK)
	{
		fprintf(stderr, "error calling VirtualChannelGetInstance entry point (hr=%x)\n", hr);
		goto EXCEPTION;
	}

	// At this point we have an array of IWTSPlugin interface pointers.  For now at least,
	// we're only going to handle the first pointer in the array.
	pWTSPlugin = (IWTSPlugin*)objArray[0];
	for (ULONG i = 1; i < numObjs; i++)
	{
		IUnknown* pUnknown = (IUnknown*)objArray[i];
		pUnknown->Release();
	}

EXCEPTION:
	if (objArray)
	{
		free(objArray);
	}
	FreeLibrary(hModule);

	return pWTSPlugin;
}

IWTSPlugin* loadDVCUsingDllNameCLSID(char* dllName, CLSID* clsid)
{
	LPDllGetClassObject pDllGetClassObject;
	IClassFactory* pClassFactory = NULL;
	IWTSPlugin* pWTSPlugin = NULL;
	HMODULE hModule;
	HRESULT hr;

	hModule = LoadLibrary(dllName);
	if (hModule == NULL)
	{
		fprintf(stderr, "could not load '%s'\n", dllName);
		goto EXCEPTION;
	}

	pDllGetClassObject = (LPDllGetClassObject)GetProcAddress(hModule, "DllGetClassObject");
	if (!pDllGetClassObject)
	{
		fprintf(stderr, "could not retrieve DllGetClassObject entry point\n");
		goto EXCEPTION;
	}

	hr = pDllGetClassObject(*clsid, IID_IClassFactory, (LPVOID*)&pClassFactory);
	if (hr != S_OK)
	{
		fprintf(stderr, "error calling DllGetClassObject (hr=%x)\n", hr);
		goto EXCEPTION;
	}

	hr = pClassFactory->CreateInstance(NULL, IID_IWTSPlugin, (LPVOID*)&pWTSPlugin);
	if (hr != S_OK)
	{
		fprintf(stderr, "error calling ICLassFactory::CreateInstance (hr=%x)\n", hr);
		goto EXCEPTION;
	}

EXCEPTION:
	if (pClassFactory)
	{
		pClassFactory->Release();
	}
	FreeLibrary(hModule);

	return pWTSPlugin;
}

int main(int argc, char **argv)
{
	IWTSPlugin* pWTSPlugin = NULL;
	CLSID clsid;
	char* colon;

	CoInitialize(NULL);

	if (argc != 2)
	{
		showUsage();
		return -1;
	}

	//
	// Load a Dynamic Virtual Channel (DVC) plugin in one of 3 possible
	// ways (http://msdn.microsoft.com/en-us/library/bb540856(v=vs.85).aspx).
	//
	// 1. Plug-inDLLName:{CLSID}
	//
	//    The plug-in is not necessarily registered in the Windows registry
	//    as a Component Object Model (COM) object, but the DLL is implemented
	//    as an in-process COM object. The RDC client will load the DLL
	//    specified by Plug-inDLLName and retrieve the COM object directly
	//    using CLSID.
	//
	// 2. Plug-inDLLName
	//
	//    The DLL implements the VirtualChannelGetInstance function and exports
	//    it by name. The RDC client will use the VirtualChannelGetInstance
	//    function to obtain IWTSPlugin interface pointers for all of the
	//    plug-ins implemented by the DLL.
	//
	// 3. {CLSID}
	//
	//    The RDC client will instantiate the plug-in as a regular COM object
	//    using CoCreateInstance with the CLSID.
	//
 
	colon = strrchr(argv[1], ':');
	if (colon)
	{
		char* dllName = argv[1];
		char* cstrCLSID = colon + 1;
		if (parseCLSIDString(cstrCLSID, &clsid))
		{
			*colon = '\0';
			pWTSPlugin = loadDVCUsingDllNameCLSID(dllName, &clsid);
		}
		else
		{
			pWTSPlugin = loadDVCUsingDllName(dllName);
		}
	}
	else
	{
		if (parseCLSIDString(argv[1], &clsid))
		{
			pWTSPlugin = loadDVCUsingCLSID(&clsid);
		}
		else
		{
			pWTSPlugin = loadDVCUsingDllName(argv[1]);
		}
	}

	printf("pWTSPlugin=%x\n", pWTSPlugin);

	CoUninitialize();

	return 0;
}
