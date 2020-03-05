#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#include <comdef.h>
#include <comdefsp.h>
#include <iostream>
#include <tchar.h>

#include "nameof.hpp"

constexpr auto check = ::_com_util::CheckError;

int wmain(int argc, WCHAR* argv[])
{
	auto libPath = argv[1];
	check(CoInitialize(nullptr));

	ITypeLibPtr pTypeLib{};

	printf("Path: %S\n", libPath);
	check(LoadTypeLib(libPath, &pTypeLib));
	{
		TLIBATTR* pTLibAttr{};
		check(pTypeLib->GetLibAttr(&pTLibAttr));
		LPOLESTR pGuidStr{};
		check(StringFromCLSID(pTLibAttr->guid, &pGuidStr));
		printf("GUID: %S\n", pGuidStr);
		CoTaskMemFree(pGuidStr);
		printf("lcid: 0x%X\n", pTLibAttr->lcid);
		printf("syskind: %s\n", nameof::nameof_enum(pTLibAttr->syskind).data());
		printf("VerNum: %u,%u\n", pTLibAttr->wMajorVerNum, pTLibAttr->wMinorVerNum);
		printf("wLibFlags: %s\n", nameof::nameof_enum((LIBFLAGS)pTLibAttr->wLibFlags).data());
		pTypeLib->ReleaseTLibAttr(pTLibAttr);

		bstr_t name, docString, helpFile;
		check(pTypeLib->GetDocumentation(-1, name.GetAddress(), docString.GetAddress(), nullptr, helpFile.GetAddress()));
		printf("Documentation Name: %S\n", name.GetBSTR());
		printf("Documentation String: %S\n", docString.GetBSTR());
		printf("Documentation HelpFile: %S\n", helpFile.GetBSTR());
	}
	{
		ITypeLib2Ptr pTypeLib2 = pTypeLib;
		ULONG cUniqueNames, cchUniqueNames;
		check(pTypeLib2->GetLibStatistics(&cUniqueNames, &cchUniqueNames));
		printf("UniqueNames: %u\n", cUniqueNames);
		printf("Count of Char UniqueNames: %u\n", cchUniqueNames);
		CUSTDATA custData{};
		check(pTypeLib2->GetAllCustData(&custData));
		for (DWORD i = 0; i < custData.cCustData; ++i)
		{
			auto& item = custData.prgCustData[i];
			LPOLESTR pGuidStr{};
			check(StringFromCLSID(item.guid, &pGuidStr));
			printf(" %u: CustData GUID: %S\n", i, pGuidStr);
			CoTaskMemFree(pGuidStr);
			variant_t varStr;
			check(VariantChangeType(&varStr, &item.varValue, VARIANT_ALPHABOOL, VT_BSTR));
			printf(" %u: CustData Value: %S\n", i, varStr.bstrVal);
		}
		ClearCustData(&custData);
	}

	printf("# of Type: %u\n", pTypeLib->GetTypeInfoCount());
	for (UINT iType = 0; iType < pTypeLib->GetTypeInfoCount(); ++iType)
	{
		printf("-------------------- %u --------------------\n", iType);
		TYPEKIND typeKind{};
		check(pTypeLib->GetTypeInfoType(iType, &typeKind));
		printf("TypeKind of Type: %s\n", nameof::nameof_enum(typeKind).data());
	
		bstr_t name, docString, helpFile;
		check(pTypeLib->GetDocumentation(iType, name.GetAddress(), docString.GetAddress(), nullptr, helpFile.GetAddress()));
		printf("Documentation Name: %S\n", name.GetBSTR());
		printf("Documentation String: %S\n", docString.GetBSTR());
		printf("Documentation HelpFile: %S\n", helpFile.GetBSTR());
	
		ITypeInfoPtr pTypeInfo{};
		check(pTypeLib->GetTypeInfo(iType, &pTypeInfo));

		TYPEATTR* pTypeAttr;
		check(pTypeInfo->GetTypeAttr(&pTypeAttr));
		LPOLESTR pGuidStr{};
		check(StringFromCLSID(pTypeAttr->guid, &pGuidStr));
		printf("GUID: %S\n", pGuidStr);
		CoTaskMemFree(pGuidStr);
		printf("lcid: 0x%X\n", pTypeAttr->lcid);
		printf("memidConstructor: %d\n", pTypeAttr->memidConstructor);
		printf("memidDestructor: %d\n", pTypeAttr->memidDestructor);
		printf("Schema: %S\n", pTypeAttr->lpstrSchema);
		printf("cbSizeInstance: %u\n", pTypeAttr->cbSizeInstance);
		printf("TypeKind: %s\n", nameof::nameof_enum(pTypeAttr->typekind).data());
		printf("# of Funcs: %u\n", pTypeAttr->cFuncs);
		printf("# of Varss: %u\n", pTypeAttr->cVars);
		printf("# of implemented interfaces: %u\n", pTypeAttr->cImplTypes);
		printf("size of this type's VTBL: %u\n", pTypeAttr->cbSizeVft);
		printf("byte alignment for an instance of this type: %u\n", pTypeAttr->cbAlignment);
		printf("Type Flags: %s\n", nameof::nameof_enum((TYPEFLAGS)pTypeAttr->wTypeFlags).data());
		printf("VerNum: %u,%u\n", pTypeAttr->wMajorVerNum, pTypeAttr->wMinorVerNum);
		//printf("tdescAlias: %u\n", pTypeAttr->tdescAlias);
		enum class IDLFLAG : USHORT {
			NONE = IDLFLAG_NONE, 
			FIN = IDLFLAG_FIN, 
			FOUT = IDLFLAG_FOUT, 
			FLCID = IDLFLAG_FLCID, 
			FRETVAL = IDLFLAG_FRETVAL,
		};
		printf("idldescType: %s\n", nameof::nameof_enum((IDLFLAG)pTypeAttr->idldescType.wIDLFlags).data());

		printf("-------------------- Vars --------------------\n");
		for (UINT iVar = 0; iVar < pTypeAttr->cVars; ++iVar)
		{
			VARDESC* pVarDesc{};
			check(pTypeInfo->GetVarDesc(iVar, &pVarDesc));
			printf(" %u: memid: %d\n", iVar, pVarDesc->memid);
			{
				BSTR names[256]{};
				UINT cNames{};
				check(pTypeInfo->GetNames(pVarDesc->memid, names, _countof(names), &cNames));
				printf(" %u: name:", iVar);
				for (UINT i = 0; i < cNames; ++i)
				{
					printf("%S, ", bstr_t(names[i], false).GetBSTR());
				}
				printf("\n");
			}

			bstr_t name, docString, helpFile;
			check(pTypeInfo->GetDocumentation(pVarDesc->memid, name.GetAddress(), docString.GetAddress(), nullptr, helpFile.GetAddress()));
			printf(" Documentation Name: %S\n", name.GetBSTR());
			printf(" Documentation String: %S\n", docString.GetBSTR());
			printf(" Documentation HelpFile: %S\n", helpFile.GetBSTR());

			printf(" %u: Schema: %S\n", iVar, pVarDesc->lpstrSchema);
			// oInst
			// lpvarValue
			// elemdescVar
			printf(" %u: wVarFlags: %s\n", iVar, nameof::nameof_enum((VARFLAGS)pVarDesc->wVarFlags).data());
			printf(" %u: varkind: 0x%08X<%s>\n", iVar, pVarDesc->varkind, nameof::nameof_enum(pVarDesc->varkind).data());

			pTypeInfo->ReleaseVarDesc(pVarDesc);
		}

		printf("-------------------- Funcs --------------------\n");
		for (UINT iFunc = 0; iFunc < pTypeAttr->cFuncs; ++iFunc)
		{
			FUNCDESC* pFuncDesc{};
			check(pTypeInfo->GetFuncDesc(iFunc, &pFuncDesc));
			printf(" %u: memid: %d\n", iFunc, pFuncDesc->memid);
			{
				BSTR names[256]{};
				UINT cNames{};
				check(pTypeInfo->GetNames(pFuncDesc->memid, names, _countof(names), &cNames));
				printf(" %u: name: ", iFunc);
				for (UINT i = 0; i < cNames; ++i)
				{
					printf("%s%S%s", i > 1 ? ", ":"" ,bstr_t(names[i], false).GetBSTR(), i == 0 ? "(" : "");
				}
				printf(")\n");
			}

			bstr_t name, docString, helpFile;
			check(pTypeInfo->GetDocumentation(pFuncDesc->memid, name.GetAddress(), docString.GetAddress(), nullptr, helpFile.GetAddress()));
			printf(" Documentation Name: %S\n", name.GetBSTR());
			printf(" Documentation String: %S\n", docString.GetBSTR());
			printf(" Documentation HelpFile: %S\n", helpFile.GetBSTR());

			// lprgscode
			// lprgelemdescParam
			printf(" %u: funckind: %s\n", iFunc, nameof::nameof_enum(pFuncDesc->funckind).data());
			printf(" %u: invkind: %s\n", iFunc, nameof::nameof_enum(pFuncDesc->invkind).data());
			printf(" %u: callconv: %s\n", iFunc, nameof::nameof_enum(pFuncDesc->callconv).data());
			printf(" %u: cParams: %u\n", iFunc, pFuncDesc->cParams);
			printf(" %u: cParamsOpt: %u\n", iFunc, pFuncDesc->cParamsOpt);
			printf(" %u: oVft: %u\n", iFunc, pFuncDesc->oVft);
			printf(" %u: cScodes: %u\n", iFunc, pFuncDesc->cScodes);
			// elemdescFunc
			printf(" %u: wFuncFlags: 0x%08X<%s>\n", iFunc, pFuncDesc->wFuncFlags, nameof::nameof_enum((FUNCFLAGS)pFuncDesc->wFuncFlags).data());

			pTypeInfo->ReleaseFuncDesc(pFuncDesc);
		}		

		pTypeInfo->ReleaseTypeAttr(pTypeAttr);

		{
			ITypeInfo2Ptr pTypeInfo2 = pTypeInfo;

		}
	}

	::CoUninitialize();

}
