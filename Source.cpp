#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "shlwapi")

#include <windows.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <comutil.h>

TCHAR szClassName[] = TEXT("Window");

VOID GetFileDetails(LPCTSTR lpszFilePath, HWND hEdit)
{
	TCHAR szFolderPath[MAX_PATH];
	TCHAR szFileName[MAX_PATH];
	lstrcpy(szFolderPath, lpszFilePath);
	PathRemoveFileSpec(szFolderPath);
	lstrcpy(szFileName, PathFindFileName(lpszFilePath));
	SetWindowText(hEdit, 0);
	CoInitialize(0);
	IShellDispatch* pUnknown = NULL;
	struct __declspec(uuid("13709620-C279-11CE-A49E-444553540000")) Cls;
	if (SUCCEEDED(CoCreateInstance(__uuidof(Cls), NULL, CLSCTX_ALL, IID_PPV_ARGS(&pUnknown))))
	{
		Folder *folder = NULL;
		VARIANT v;
		{
			VariantInit(&v);
			V_VT(&v) = VT_BSTR;
			V_BSTR(&v) = BSTR(szFolderPath);
			pUnknown->NameSpace(v, &folder);
		}
		VARIANT vNull = { 0 };
		VariantInit(&vNull);
		for (int i = 0; i < 1000; ++i)
		{
			BSTR bstrColumn = 0;
			folder->GetDetailsOf(vNull, i, &bstrColumn);
			if (lstrlen(bstrColumn) > 0)
			{
				FolderItem *item;
				folder->ParseName(BSTR(szFileName), &item);
				const _variant_t vItem((IDispatch*)item);
				BSTR bstrValue = 0;
				folder->GetDetailsOf((VARIANT)vItem, i, &bstrValue);
				if (lstrlen(bstrValue) > 0)
				{
					SendMessage(hEdit, EM_REPLACESEL, 0, (LPARAM)bstrColumn);
					SendMessage(hEdit, EM_REPLACESEL, 0, (LPARAM)TEXT(":"));
					SendMessage(hEdit, EM_REPLACESEL, 0, (LPARAM)bstrValue);
					SendMessage(hEdit, EM_REPLACESEL, 0, (LPARAM)TEXT("\r\n"));
				}
				SysFreeString(bstrValue);
				item->Release();
			}
			SysFreeString(bstrColumn);
		}
		folder->Release();
		pUnknown->Release();
		VariantClear(&vNull);
	}
	CoUninitialize();
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit;
	switch (msg)
	{
	case WM_CREATE:
		hEdit = CreateWindow(TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_VSCROLL | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_SIZE:
		MoveWindow(hEdit, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
		break;
	case WM_DROPFILES:
	{
		SetWindowText(hEdit, 0);
		const int nFileCount = DragQueryFile((HDROP)wParam, -1, NULL, 0);
		for (int i = 0; i<nFileCount; i++)
		{
			TCHAR szFilePath[MAX_PATH];
			DragQueryFile((HDROP)wParam, i, szFilePath, _countof(szFilePath));
			GetFileDetails(szFilePath, hEdit);
		}
		DragFinish((HDROP)wParam);
	}
	break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Shell を使ってファイルの詳細情報を取得する"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}