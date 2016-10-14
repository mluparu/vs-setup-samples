// Setup.Configuration.ZeroDeps.VC.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iostream>
#include "Setup.Configuration.h"

using namespace ATL;

void PrintInstance(CComPtr<ISetupInstance2>& instance2, CComPtr<ISetupHelper>& setupHelper);
void PrintPackageReference(CComPtr<ISetupPackageReference>& package);
void PrintWorkloads(CComSafeArray<IUnknown*>& packages);

int main()
{
	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	CComPtr<ISetupConfiguration> setupConfig = nullptr;
	auto hresult = setupConfig.CoCreateInstance(__uuidof(SetupConfiguration), nullptr, CLSCTX_INPROC_SERVER);

	if (!SUCCEEDED(hresult) || setupConfig == nullptr)
	{
		std::wcout << L"Visual Studio '15' may not be installed (Component creation failed)" << std::endl;
		return 0;
	}

	CComQIPtr<ISetupConfiguration2> setupConfig2(setupConfig);
	CComQIPtr<ISetupHelper> setupHelper(setupConfig);
	if (!setupConfig2 || !setupHelper)
	{
		std::wcout << L"Unsupported version of Visual Studio '15' may be installed (ISetupConfiguration2 or ISetupHelper unavailable)" << std::endl;
		return 1;
	}
		
	CComPtr<IEnumSetupInstances> enumInstances;
	setupConfig2->EnumAllInstances(&enumInstances);
	if (!enumInstances)
	{
		std::wcout << L"No VS '15' version is installed (EnumAllInstances returned null)" << std::endl;
		return 1;
	}

	CComPtr<ISetupInstance> instance;

	std::wcout << L"Listing VS '15' instances:" << std::endl;
	std::wcout << L"-------------------------------------------" << std::endl;
	while (SUCCEEDED(enumInstances->Next(1, &instance, nullptr)) && instance)
	{
		CComQIPtr<ISetupInstance2> instance2(instance);
		if (!instance2)
		{
			std::wcout << L"Error querying instance (ISetupInstance2 unavailable)" << std::endl;
			continue;
		}
		//Print instance
		PrintInstance(instance2, setupHelper);
		
		instance = nullptr;
		std::wcout << L"*" << std::endl;
	}
    return 0;
}

void PrintInstance(CComPtr<ISetupInstance2>& instance2, CComPtr<ISetupHelper>& setupHelper)
{
	USES_CONVERSION;

	CComBSTR	bstrId;
	if (FAILED(instance2->GetInstanceId(&bstrId)))
	{
		std::wcout << L"Error reading instance id" << std::endl;
	}
	InstanceState state;
	if (FAILED(instance2->GetState(&state)))
	{
		std::wcout << L"Error reading instance state" << std::endl;
	}

	std::wcout << L"InstanceId: " << OLE2T(bstrId)
		<< L" (" << (eComplete == state ? L"Complete" : L"Incomplete") << L")"
		<< std::endl;

	CComBSTR	bstrVersion;
	if (FAILED(instance2->GetInstallationVersion(&bstrVersion)))
	{
		std::wcout << L"Error reading version" << std::endl;
	}
	else
	{
		ULONGLONG ullVersion;
		if (FAILED(setupHelper->ParseVersion(bstrVersion, &ullVersion)))
		{
			std::wcout << L"Error parsing version " << OLE2T(bstrVersion) << std::endl;
		}

		std::wcout << L"InstallationVersion: " << OLE2T(bstrVersion) << L" (" << ullVersion << L")" << std::endl;
	}

	// Reboot may have been required before the installation path was created.
	if ((eLocal & state) == eLocal)
	{
		CComBSTR bstrInstallationPath;
		if (FAILED(instance2->GetInstallationPath(&bstrInstallationPath)))
		{
			std::wcout << L"Error getting installation path" << std::endl;
		}

		std::wcout << L"InstallationPath: " << OLE2T(bstrInstallationPath) << std::endl;
	}

	// Reboot may have been required before the product package was registered (last).
	if ((eRegistered & state) == eRegistered)
	{
		CComPtr<ISetupPackageReference> product;
		instance2->GetProduct(&product);

		if (!product)
		{
			std::wcout << L"Error getting product info (ISetupPackageReference missing)" << std::endl;
			return;
		}

		PrintPackageReference(product);
		std::wcout << std::endl;

		LPSAFEARRAY lpsaPackages;
		if (FAILED(instance2->GetPackages(&lpsaPackages)))
		{
			std::wcout << L"Error enumerating packages (GetPackages failed)" << std::endl;
			return;
		}

		CComSafeArray<IUnknown*> packages;
		packages.Attach(lpsaPackages);

		PrintWorkloads(packages);
		std::wcout << std::endl;
	}
}

void PrintPackageReference(CComPtr<ISetupPackageReference>& package)
{
	USES_CONVERSION;

	CComBSTR bstrId;
	if (FAILED(package->GetId(&bstrId)))
	{
		std::wcout << L"Error getting reference id (GetId failed)" << std::endl;
	}

	CComBSTR bstrType;
	if (FAILED(package->GetType(&bstrType)))
	{
		std::wcout << L"Error getting reference type (GetType failed)" << std::endl;
	}

	std::wcout << OLE2T(bstrId) << L" (" << OLE2T(bstrType) << L")" << std::endl;
}

void PrintWorkloads(CComSafeArray<IUnknown*>& packages)
{
	ULONG cPackages = packages.GetCount();
	for (ULONG i = 0; i < cPackages; ++i)
	{
		CComQIPtr<ISetupPackageReference> package = packages.GetAt(i);
		if (!package)
		{
			std::wcout << L"Error querying package (QI for ISetupPackageReference failed)" << std::endl;
			continue;
		}
		
		std::wcout << L"	";
		PrintPackageReference(package);
	}
}
