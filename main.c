// The code is based on https://github.com/little-brother/sqlite-wlx
#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <uxtheme.h>
#include <locale.h>
#include <tchar.h>
#include <shlwapi.h>
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#include <math.h>
#include "include/sqlite3.h"

#define LVS_EX_AUTOSIZECOLUMNS 0x10000000

#define WMU_UPDATE_GRID        WM_USER + 1
#define WMU_UPDATE_DATA        WM_USER + 2
#define WMU_UPDATE_ROW_COUNT   WM_USER + 3
#define WMU_UPDATE_FILTER_SIZE WM_USER + 4
#define WMU_SET_HEADER_FILTERS WM_USER + 5
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 6
#define WMU_SET_CURRENT_CELL   WM_USER + 7
#define WMU_RESET_CACHE        WM_USER + 8
#define WMU_SET_FONT           WM_USER + 9
#define WMU_SET_THEME          WM_USER + 10
#define WMU_HIDE_COLUMN        WM_USER + 11
#define WMU_SHOW_COLUMNS       WM_USER + 12
#define WMU_HOT_KEYS           WM_USER + 13  
#define WMU_HOT_CHARS          WM_USER + 14
#define WMU_EDIT_CELL          WM_USER + 20
#define WMU_UPDATE_COUNTS      WM_USER + 22

#define IDC_MAIN               100
#define IDC_TABLELIST          101
#define IDC_GRID               102
#define IDC_STATUSBAR          103
#define IDC_CELL_EDIT          104
#define IDC_BUTTON             105
#define IDC_HEADER_EDIT        1000
#define IDC_DLG_GRID           2000
#define IDC_DLG_OK             2001
#define IDC_DLG_EDIT           2200
#define IDC_DLG_LABEL          2500

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROWS          5001
#define IDM_COPY_COLUMN        5002
#define IDM_FILTER_ROW         5003
#define IDM_DARK_THEME         5004
#define IDM_HIDE_COLUMN        5010
#define IDM_DROP               5015
#define IDM_BLOB               5020
#define IDM_ADD_ROW            5021
#define IDM_DELETE_ROW         5022
#define IDM_SEPARATOR1         5030
#define IDM_SEPARATOR2         5031

#define IDM_OPEN_DB            5130
#define IDM_EXIT               5131
#define IDM_RECENT             5140
#define IDM_ABOUT              5165
#define IDM_HOMEPAGE           5166
#define IDM_WIKI               5167

#define IDI_ICON               6000

#define SB_VERSION             0
#define SB_TABLE_COUNT         1
#define SB_VIEW_COUNT          2
#define SB_TYPE                3
#define SB_ROW_COUNT           4
#define SB_CURRENT_CELL        5
#define SB_AUXILIARY           6

#define SPLITTER_WIDTH         5
#define MAX_TEXT_LENGTH        32000
#define MAX_COLUMN_LENGTH      2000
#define MAX_TABLE_LENGTH       2000
#define BLOB_VALUE             "(BLOB)"
#define MAX_RECENT_FILES       10 

#define APP_NAME               TEXT("sqlite-x")
#define APP_VERSION            TEXT("0.9.9")
#ifdef __MINGW64__
#define APP_PLATFORM               64
#else
#define APP_PLATFORM               32
#endif

static TCHAR iniPath[MAX_PATH] = {0};

struct DLGITEMTEMPLATEEX {
	DLGITEMTEMPLATE; 
	WORD menu; 
	WORD windowClass;
};

LRESULT CALLBACK cbMainWnd(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
HWND ListLoadW (HWND hParentWnd, TCHAR* fileToLoad, int showFlags);
void __stdcall ListCloseWindow(HWND hWnd);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewMain (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewCellEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK cbDlgAddRow(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK cbDlgAddTable(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
BOOL CALLBACK cbEnumSetFont (HWND hWnd, LPARAM hFont);
void updateRecentList(HWND hWnd);

void bindValue(sqlite3_stmt* stmt, int i, const char* value8);
void bindFilterValues(HWND hHeader, sqlite3_stmt* stmt);
HWND getMainWindow(HWND hWnd);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
TCHAR* getStoredString(TCHAR* name, TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords);
TCHAR* extractUrl(TCHAR* data);
void setClipboardText(const TCHAR* text);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);
void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState);
BOOL saveFile(TCHAR* path, const TCHAR* filter, const TCHAR* defExt, HWND hWnd);

int WINAPI WinMain (HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	WNDCLASS wc = {0};
	wc.lpfnWndProc = cbMainWnd;
	wc.hInstance = hInstance;
	wc.hCursor = LoadCursor (NULL, IDC_ARROW);
	wc.lpszClassName = TEXT("sqlite-x-class");	
	wc.hbrBackground = (HBRUSH) GetSysColorBrush(COLOR_APPWORKSPACE);
	wc.style = CS_DBLCLKS;
	RegisterClass(&wc);

	HMENU hFileMenu = CreatePopupMenu();
	AppendMenu(hFileMenu, MF_STRING, IDM_OPEN_DB, TEXT("Open..."));
	AppendMenu(hFileMenu, MF_STRING, IDM_SEPARATOR1, NULL);
	AppendMenu(hFileMenu, MF_STRING, IDM_EXIT, TEXT("Exit"));

	HMENU hAboutMenu = CreatePopupMenu();
	AppendMenu(hAboutMenu, MF_STRING, IDM_ABOUT, TEXT("About"));
	AppendMenu(hAboutMenu, MF_STRING, IDM_WIKI, TEXT("Wiki"));
	AppendMenu(hAboutMenu, MF_STRING, IDM_HOMEPAGE, TEXT("Homepage"));		

	HMENU hMainMenu = CreateMenu();
	AppendMenu(hMainMenu, MF_STRING | MF_POPUP, (UINT_PTR)hFileMenu, TEXT("Database"));
	AppendMenu(hMainMenu, MF_STRING | MF_POPUP, (UINT_PTR)hAboutMenu, TEXT("?"));
	ModifyMenu(hMainMenu, 1, MF_BYPOSITION | MFT_RIGHTJUSTIFY, 0, TEXT("?"));	

	HWND hWnd = CreateWindowEx(0, TEXT("sqlite-x-class"), TEXT("sqlite-x - No database selected"), WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, hMainMenu, hInstance, NULL);
	
	if (hWnd == NULL)
		return 0;
		
	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&icex);
	
	setlocale(LC_CTYPE, "");		
		
	TCHAR buf16[MAX_PATH];	
	GetModuleFileName(0, buf16, MAX_PATH);
	PathRemoveFileSpec(buf16);
	_sntprintf(iniPath, MAX_PATH, TEXT("%ls/%ls.ini"), buf16, APP_NAME);
	
	int nArgs = 0;
	TCHAR** args = CommandLineToArgvW(GetCommandLine(), &nArgs);		
		
	if (nArgs > 1) {
		TCHAR path16[MAX_PATH + 1] = {0};	
		_tcsncpy(path16, args[1], MAX_PATH);
		if (!ListLoadW (hWnd, path16, 0))
			return 0;
	}

	HICON hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON));
	SendMessage(hWnd, WM_SETICON, ICON_SMALL, (LPARAM)hIcon);

	SetWindowPos(hWnd, 0, 
		getStoredValue(TEXT("position-x"), 200),
		getStoredValue(TEXT("position-y"), 200),
		getStoredValue(TEXT("width"), 800),
		getStoredValue(TEXT("height"), 600),
		SWP_NOZORDER);	
	ShowWindow(hWnd, nCmdShow);
	SendMessage(hWnd, WM_SIZE, 0, 0);
	updateRecentList(hWnd);

	
	if (nArgs < 2 && getStoredValue(TEXT("open-last-db"), 1))
		SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(IDM_RECENT, 0), 0);
	
	MSG msg = {};
	while (GetMessage(&msg, NULL, 0, 0) > 0) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	
	return 0;
}

LRESULT CALLBACK cbMainWnd(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
	switch (uMsg) {		
		case WM_DESTROY: {
			HWND hMainWnd = GetDlgItem(hWnd, IDC_MAIN);			
			if (hMainWnd)
				ListCloseWindow(hMainWnd);
			
			RECT rc;
			GetWindowRect(hWnd, &rc);
			setStoredValue(TEXT("position-x"), rc.left);
			setStoredValue(TEXT("position-y"), rc.top);
			setStoredValue(TEXT("width"), rc.right - rc.left);
			setStoredValue(TEXT("height"), rc.bottom - rc.top);			
			
			PostQuitMessage(0);
			return 0;
		}
		break;
		
		case WM_SIZE: {
			RECT rc;
			GetClientRect(hWnd, &rc);			
			HWND hMainWnd = GetDlgItem(hWnd, IDC_MAIN);
			if (hMainWnd)
				SetWindowPos(hMainWnd, 0, 0, 0, rc.right, rc.bottom, SWP_NOMOVE | SWP_NOZORDER);
		}
		break;
		
		case WM_KEYDOWN: {
			if (wParam == VK_ESCAPE && getStoredValue(TEXT("exit-by-escape"), 1)) {
				SendMessage(hWnd, WM_DESTROY, 0, 0);
				return TRUE;
			}			
		}
		break;		
		
		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDM_ABOUT) {
				OSVERSIONINFO vi = {0};
				vi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);

				NTSTATUS (WINAPI *pRtlGetVersion)(PRTL_OSVERSIONINFOW lpVersionInformation) = NULL;
				HINSTANCE hNTdllDll = LoadLibrary(TEXT("ntdll.dll"));

				if (hNTdllDll != NULL) {
					pRtlGetVersion = (NTSTATUS (WINAPI *)(PRTL_OSVERSIONINFOW))
					GetProcAddress (hNTdllDll, "RtlGetVersion");

					if (pRtlGetVersion != NULL)
						pRtlGetVersion((PRTL_OSVERSIONINFOW)&vi);

					FreeLibrary(hNTdllDll);
				}

				if (pRtlGetVersion == NULL)
					GetVersionEx(&vi);

				BOOL isWin64 = FALSE;
				#if defined(_WIN64)
					isWin64 = TRUE;
				#else
					isWin64 = IsWow64Process(GetCurrentProcess(), &isWin64) && isWin64;
				#endif

				TCHAR buf[1024];
				_sntprintf(buf, 1023, TEXT("GUI version: %ls\nArchitecture: %ls\nSQLite version: %hs\nBuild date: %ls\n\nOS: %i.%i.%i %ls %ls"),
					APP_VERSION,
					APP_PLATFORM == 32 ? TEXT("x86") : TEXT("x86-64"),
					sqlite3_libversion(),
					TEXT(__DATE__),
					vi.dwMajorVersion, vi.dwMinorVersion, vi.dwBuildNumber, vi.szCSDVersion,
					isWin64 ? TEXT("x64") : TEXT("x32")
				);

				MessageBox(hWnd, buf, APP_NAME, MB_OK);		
			}
			
			if (cmd == IDM_HOMEPAGE)
				ShellExecute(0, 0, TEXT("https://github.com/little-brother/sqlite-x"), 0, 0 , SW_SHOW);
				
			if (cmd == IDM_WIKI)
				ShellExecute(0, 0, TEXT("https://github.com/little-brother/sqlite-x/wiki"), 0, 0 , SW_SHOW);	
			
			if (cmd == IDM_OPEN_DB) {
				TCHAR path16[MAX_PATH];
				
				OPENFILENAME ofn = {0};
				ofn.lStructSize = sizeof(ofn);
				ofn.hwndOwner = hWnd;
				ofn.lpstrFile = path16;
				ofn.lpstrFile[0] = '\0';
				ofn.nMaxFile = MAX_PATH;
				ofn.lpstrFilter = TEXT("Databases (*.sqlite, *.sqlite3, *.db, *.db3)\0*.sqlite;*.sqlite3;*.db;*.db3\0All\0*.*\0");
				ofn.nFilterIndex = 1;
				ofn.lpstrFileTitle = NULL;
				ofn.nMaxFileTitle = 0;
				ofn.lpstrInitialDir = NULL;
				ofn.Flags = OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
			
				if (!GetOpenFileName(&ofn))
					return 0;
					
				TCHAR ext16[32] = {0};
				_tsplitpath(path16, NULL, NULL, NULL, ext16);
				if (_tcslen(ext16) == 0)
					_tcscat(path16, TEXT(".sqlite"));			
					
				ListLoadW (hWnd, path16, 0);
			}
			
			if (cmd == IDM_EXIT) {
				SendMessage(hWnd, WM_CLOSE, 0, 0);
			}
			
			if (cmd >= IDM_RECENT && cmd < IDM_RECENT + MAX_RECENT_FILES) {
				TCHAR path16[MAX_PATH + 1];
				if (GetMenuString(GetSubMenu(GetMenu(hWnd), 0), cmd, path16, MAX_PATH, MF_BYCOMMAND))
					ListLoadW (hWnd, path16, 0);
			}
		}
		break;
	
	}
	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

void updateRecentList(HWND hWnd) {
	HMENU hMenu = GetSubMenu(GetMenu(hWnd), 0);
	
	for (int i = 0; i < MAX_RECENT_FILES; i++)
		DeleteMenu(hMenu, IDM_RECENT + i, MF_BYCOMMAND);
	DeleteMenu(hMenu, IDM_SEPARATOR2, MF_BYCOMMAND);

	TCHAR* buf16 = calloc(20 * MAX_PATH, sizeof(TCHAR));
	if (GetPrivateProfileString(APP_NAME, TEXT("recent"), NULL, buf16, 20 * MAX_PATH, iniPath)) {
		int i = 0;
		TCHAR* path16 = _tcstok(buf16, TEXT("?"));
		while (path16 != NULL && i < MAX_RECENT_FILES) {
			if (_taccess(path16, 0) == 0) {
				if (i == 0)	
					InsertMenu(hMenu, IDM_EXIT, MF_BYCOMMAND | MF_STRING, IDM_SEPARATOR2, NULL);

				InsertMenu(hMenu, IDM_SEPARATOR2, MF_BYCOMMAND | MF_STRING, IDM_RECENT + i, path16);
				i++;
			}
			path16 = _tcstok(NULL, TEXT("?"));
		}		
	}
	free(buf16);
}

HWND ListLoadW (HWND hParentWnd, TCHAR* fileToLoad, int showFlags) {
	int len = _tcslen(fileToLoad) + 10;
	TCHAR* journal = calloc(len + 1, sizeof(TCHAR));
	_sntprintf(journal, len, TEXT("%ls-wal"), fileToLoad);
	BOOL hasJournal = _taccess(journal, 0) == 0;
	free(journal);
	
	char* fileToLoad8 = utf16to8(fileToLoad);

	BOOL isEditable = getStoredValue(TEXT("editable"), 1);
	sqlite3 *db = 0;
	if (SQLITE_OK != sqlite3_open_v2(fileToLoad8, &db, (isEditable ? SQLITE_OPEN_READWRITE : SQLITE_OPEN_READONLY) | SQLITE_OPEN_URI, NULL)) {
		MessageBox(hParentWnd, TEXT("Error to open database"), fileToLoad, MB_OK);
		free(fileToLoad8);
		return NULL;
	}
	free(fileToLoad8);
	
	HWND hMainWnd = GetDlgItem(hParentWnd, IDC_MAIN);
	if (hMainWnd) {
		ListCloseWindow(hMainWnd);
		DestroyWindow(hMainWnd);
	}

	BOOL isStandalone = GetParent(hParentWnd) == HWND_DESKTOP;
	hMainWnd = CreateWindow(WC_STATIC, APP_NAME, WS_CHILD | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hParentWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	int cacheSize = getStoredValue(TEXT("cache-size"), 2000);
	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("FILTERROW"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("CACHE"), calloc(cacheSize, sizeof(TCHAR*)));
	SetProp(hMainWnd, TEXT("CACHESIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("DB"), (HANDLE)db);
	SetProp(hMainWnd, TEXT("HASJOURNAL"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("CACHEOFFSET"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TABLENAME8"), calloc(MAX_TABLE_LENGTH, sizeof(char)));
	SetProp(hMainWnd, TEXT("WHERE8"), calloc(MAX_TEXT_LENGTH, sizeof(char)));
	SetProp(hMainWnd, TEXT("HASROWID"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTROWNO"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTCOLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SEARCHCELLPOS"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SPLITTERPOSITION"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTFAMILY"), getStoredString(TEXT("font"), TEXT("Arial")));	
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FILTERALIGN"), calloc(1, sizeof(int)));		

	SetProp(hMainWnd, TEXT("DARKTHEME"), calloc(1, sizeof(int)));			
	SetProp(hMainWnd, TEXT("TEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR2"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERBACKCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTCELLCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SELECTIONTEXTCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SELECTIONBACKCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SPLITTERCOLOR"), calloc(1, sizeof(int)));

	*(int*)GetProp(hMainWnd, TEXT("HASJOURNAL")) = hasJournal;	
	*(int*)GetProp(hMainWnd, TEXT("SPLITTERPOSITION")) = getStoredValue(TEXT("splitter-position"), 200);
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);
	*(int*)GetProp(hMainWnd, TEXT("CACHESIZE")) = cacheSize;
	*(int*)GetProp(hMainWnd, TEXT("CACHEOFFSET")) = -1;	
	*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) = getStoredValue(TEXT("filter-row"), 0);
	*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) = getStoredValue(TEXT("dark-theme"), 0);
	*(int*)GetProp(hMainWnd, TEXT("FILTERALIGN")) = getStoredValue(TEXT("filter-align"), 0);		

	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE | (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	HDC hDC = GetDC(hMainWnd);
	float z = GetDeviceCaps(hDC, LOGPIXELSX) / 96.0; // 96 = 100%, 120 = 125%, 144 = 150%
	ReleaseDC(hMainWnd, hDC);	
	int sizes[7] = {35 * z, 110 * z, 180 * z, 225 * z, 420 * z, 500 * z, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 7, (LPARAM)&sizes);

	HWND hBtnWnd = CreateWindow(WC_BUTTON, TEXT("Add table"), WS_CHILD | WS_VISIBLE,
		0, 0, 100, 100, hMainWnd, (HMENU)IDC_BUTTON, GetModuleHandle(0), NULL);

	HWND hListWnd = CreateWindow(WC_LISTBOX, NULL, WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_TABSTOP | WS_HSCROLL,
		0, 20, 100, 100, hMainWnd, (HMENU)IDC_TABLELIST, GetModuleHandle(0), NULL);
	SetProp(hListWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hListWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hMainWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
	
	int noLines = getStoredValue(TEXT("disable-grid-lines"), 1);	
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | (noLines ? 0 : LVS_EX_GRIDLINES) | LVS_EX_LABELTIP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));
	
	HWND hHeader = ListView_GetHeader(hGridWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);
	SetWindowTheme(hHeader, TEXT(" "), TEXT(" "));
	SetProp(hHeader, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hHeader, GWLP_WNDPROC, (LONG_PTR)cbNewHeader));

	HMENU hGridMenu = CreatePopupMenu();
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_ROWS, TEXT("Copy row(s)"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_COLUMN, TEXT("Copy column"));	
	AppendMenu(hGridMenu, MF_STRING, IDM_SEPARATOR2, NULL);
	AppendMenu(hGridMenu, MF_STRING, IDM_HIDE_COLUMN, TEXT("Hide column"));		
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_FILTER_ROW, TEXT("Filters"));		
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_DARK_THEME, TEXT("Dark theme"));					
	SetProp(hMainWnd, TEXT("GRIDMENU"), hGridMenu);

	HMENU hListMenu = CreatePopupMenu();
	AppendMenu(hListMenu, MF_STRING, IDM_DROP, TEXT("Delete"));
	SetProp(hMainWnd, TEXT("LISTMENU"), hListMenu);

	HMENU hBlobMenu = CreatePopupMenu();
	AppendMenu(hBlobMenu, MF_STRING, IDM_BLOB, TEXT("Save to file"));
	SetProp(hMainWnd, TEXT("BLOBMENU"), hBlobMenu);

	sqlite3_stmt *stmt;
	sqlite3_prepare_v2(db, "select name, type from sqlite_master where type in ('table', 'view') order by type, name", -1, &stmt, 0);
	while(SQLITE_ROW == sqlite3_step(stmt)) {
		TCHAR* name16 = utf8to16((char*)sqlite3_column_text(stmt, 0));
		int pos = ListBox_AddString(hListWnd, name16);
		free(name16);

		BOOL isTable = strcmp((char*)sqlite3_column_text(stmt, 1), "table") == 0;
		SendMessage(hListWnd, LB_SETITEMDATA, pos, isTable);
	}
	sqlite3_finalize(stmt);

	TCHAR buf[255];
	_sntprintf(buf, 32, TEXT(" %ls"), APP_VERSION);	
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)buf);	

	SendMessage(hMainWnd, WMU_UPDATE_COUNTS, 0, 0);
	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hMainWnd, WMU_SET_THEME, 0, 0);
	
	ListBox_SetCurSel(hListWnd, 0);
	SendMessage(hMainWnd, WMU_UPDATE_GRID, 0, 0);
	
	ShowWindow(hMainWnd, SW_SHOW);
	SetFocus(hListWnd);	
	SetWindowText(hParentWnd, fileToLoad);
	SendMessage(hParentWnd, WM_SIZE, 0, 0);
	
	TCHAR* paths16 = calloc(20 * MAX_PATH, sizeof(TCHAR));
	_tcscat(paths16, fileToLoad);
	_tcscat(paths16, TEXT("?"));	
	
	TCHAR* buf16 = calloc(20 * MAX_PATH, sizeof(TCHAR));
	if (GetPrivateProfileString(APP_NAME, TEXT("recent"), NULL, buf16, 20 * MAX_PATH, iniPath)) {
		int i = 0;
		TCHAR* path16 = _tcstok(buf16, TEXT("?"));
		while (path16 != NULL && i < MAX_RECENT_FILES) {
			TCHAR tmp16[MAX_PATH + 1] = {0};
			_tcscat(tmp16, path16);
			_tcscat(tmp16, TEXT("?"));	

			if (_tcsstr(paths16, tmp16) == 0 && _taccess(path16, 0) == 0) {
				_tcscat(paths16, tmp16);
				i++;
			}
			
			path16 = _tcstok(NULL, TEXT("?"));
		}		
	}
	WritePrivateProfileString(APP_NAME, TEXT("recent"), paths16, iniPath);
	free(buf16);
	free(paths16);
	updateRecentList(hParentWnd);

	return hMainWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("splitter-position"), *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("filter-row"), *(int*)GetProp(hWnd, TEXT("FILTERROW")));	
	setStoredValue(TEXT("dark-theme"), *(int*)GetProp(hWnd, TEXT("DARKTHEME")));	

	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	sqlite3* db = (sqlite3*)GetProp(hWnd, TEXT("DB"));
	BOOL hasJournal = *(int*)GetProp(hWnd, TEXT("HASJOURNAL"));
	int deleteMode = getStoredValue(TEXT("delete-journal"), 0);
	TCHAR* dbpath = utf8to16(sqlite3_db_filename(db, 0));	
	sqlite3_close(db);

	if (!hasJournal && deleteMode) {
		int len = _tcslen(dbpath) + 10;
		TCHAR* journal = calloc(len + 1, sizeof(TCHAR));
		_sntprintf(journal, len, TEXT("%ls-wal"), dbpath);
		
		if (_taccess(journal, 0) == 0 && (deleteMode == 2 || deleteMode == 1 && IDYES == MessageBox(hWnd, TEXT("Delete journal files?"), TEXT("Confirmation"), MB_YESNO | MB_ICONQUESTION))) {
			DeleteFile(journal);
			_sntprintf(journal, len, TEXT("%ls-shm"), dbpath);
			DeleteFile(journal);
		}
		free(journal);
	}	
	free(dbpath);
		
	free((int*)GetProp(hWnd, TEXT("HASJOURNAL")));
	free((int*)GetProp(hWnd, TEXT("FILTERROW")));
	free((int*)GetProp(hWnd, TEXT("DARKTHEME")));	
	free((TCHAR***)GetProp(hWnd, TEXT("CACHE")));
	free((TCHAR***)GetProp(hWnd, TEXT("CACHESIZE")));	
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((int*)GetProp(hWnd, TEXT("CACHEOFFSET")));
	free((char*)GetProp(hWnd, TEXT("TABLENAME8")));
	free((char*)GetProp(hWnd, TEXT("WHERE8")));
	free((int*)GetProp(hWnd, TEXT("HASROWID")));	
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));	
	free((int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	free((int*)GetProp(hWnd, TEXT("FONTSIZE")));	
	free((int*)GetProp(hWnd, TEXT("CURRENTROWNO")));				
	free((int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));
	free((int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")));
	free((TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
	free((int*)GetProp(hWnd, TEXT("FILTERALIGN")));
	
	free((int*)GetProp(hWnd, TEXT("TEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR2")));	
	free((int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")));
	free((int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("SPLITTERCOLOR")));	

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));	
	DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
	DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("GRIDMENU")));
	DestroyMenu(GetProp(hWnd, TEXT("LISTMENU")));
	DestroyMenu(GetProp(hWnd, TEXT("BLOBMENU")));	

	RemoveProp(hWnd, TEXT("CACHESIZE"));
	RemoveProp(hWnd, TEXT("FILTERROW"));
	RemoveProp(hWnd, TEXT("DARKTHEME"));	
	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("DB"));
	RemoveProp(hWnd, TEXT("HASJOURNAL"));	
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("CACHEOFFSET"));
	RemoveProp(hWnd, TEXT("TABLENAME8"));
	RemoveProp(hWnd, TEXT("WHERE8"));
	RemoveProp(hWnd, TEXT("HASROWID"));	
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));	
	RemoveProp(hWnd, TEXT("SPLITTERPOSITION"));
	RemoveProp(hWnd, TEXT("CURRENTROWNO"));	
	RemoveProp(hWnd, TEXT("CURRENTCOLNO"));
	RemoveProp(hWnd, TEXT("SEARCHCELLPOS"));
	RemoveProp(hWnd, TEXT("FILTERALIGN"));
	RemoveProp(hWnd, TEXT("LASTFOCUS"));
	
	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("FONTFAMILY"));
	RemoveProp(hWnd, TEXT("FONTSIZE"));			
	RemoveProp(hWnd, TEXT("TEXTCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR2"));	
	RemoveProp(hWnd, TEXT("FILTERTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("FILTERBACKCOLOR"));
	RemoveProp(hWnd, TEXT("CURRENTCELLCOLOR"));
	RemoveProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("SELECTIONBACKCOLOR"));
	RemoveProp(hWnd, TEXT("SPLITTERCOLOR"));
	RemoveProp(hWnd, TEXT("BACKBRUSH"));
	RemoveProp(hWnd, TEXT("FILTERBACKBRUSH"));		
	RemoveProp(hWnd, TEXT("SPLITTERBRUSH"));	
	RemoveProp(hWnd, TEXT("GRIDMENU"));
	RemoveProp(hWnd, TEXT("LISTMENU"));	
	RemoveProp(hWnd, TEXT("BLOBMENU"));

	DestroyWindow(hWnd);
}

LRESULT CALLBACK cbNewMain(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_SIZE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int statusH = rc.bottom;

			int splitterW = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			GetClientRect(hWnd, &rc);
			HWND hBtnWnd = GetDlgItem(hWnd, IDC_BUTTON);
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);

			int h = 0;
			if (getStoredValue(TEXT("editable"), 1) == 1) {				
				HWND hHeader = ListView_GetHeader(hGridWnd);
				BOOL isFilterRow = *(int*)GetProp(hWnd, TEXT("FILTERROW")) != 0;
				BOOL hasColumns = Header_GetItemCount(hHeader) > 0;
				LONG_PTR styles = 0;
				if (!hasColumns) {
					styles = GetWindowLongPtr(hHeader, GWL_STYLE);
					SetWindowLongPtr(hHeader, GWL_STYLE, isFilterRow ? styles | HDS_FILTERBAR : styles & (~HDS_FILTERBAR));
					ListView_AddColumn(hGridWnd, TEXT("Temp"), LVCFMT_LEFT);
				}

				RECT rcHeader;
				Header_GetItemRect(hHeader, 0, &rcHeader);
				h = rcHeader.bottom - rcHeader.top;
				if (isFilterRow)
					h = round(h/2);

				if (!hasColumns) {
					SetWindowLongPtr(hHeader, GWL_STYLE, styles);
					ListView_DeleteColumn(hGridWnd, 0);
				}
			}

			SetWindowPos(hBtnWnd, 0, 0, 0, splitterW, h, SWP_NOMOVE | SWP_NOZORDER);
			SetWindowPos(hListWnd, 0, 0, h, splitterW, rc.bottom - statusH - h, SWP_NOZORDER);
			SetWindowPos(hGridWnd, 0, splitterW + SPLITTER_WIDTH, 0, rc.right - splitterW - SPLITTER_WIDTH, rc.bottom - statusH, SWP_NOZORDER);
		}
		break;
		
		case WM_PAINT: {
			PAINTSTRUCT ps = {0};
			HDC hDC = BeginPaint(hWnd, &ps);

			RECT rc;
			GetClientRect(hWnd, &rc);
			rc.left = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			rc.right = rc.left + SPLITTER_WIDTH;
			FillRect(hDC, &rc, (HBRUSH)GetProp(hWnd, TEXT("SPLITTERBRUSH")));
			EndPaint(hWnd, &ps);

			return 0;
		}
		break;

		// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
		case WM_NCHITTEST: {
			return 1;
		}
		break;
		
		case WM_SETCURSOR: {
			SetCursor(LoadCursor(0, GetProp(hWnd, TEXT("ISMOUSEHOVER")) ? IDC_SIZEWE : IDC_ARROW));
			return TRUE;
		}
		break;
		
		case WM_SETFOCUS: {
			SetFocus(GetProp(hWnd, TEXT("LASTFOCUS")));
		}
		break;						

		case WM_LBUTTONDOWN: {
			int x = GET_X_LPARAM(lParam);
			int pos = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			if (x >= pos && x <= pos + SPLITTER_WIDTH) {
				SetProp(hWnd, TEXT("ISMOUSEDOWN"), (HANDLE)1);
				SetCapture(hWnd);
			}
			return 0;
		}
		break;

		case WM_LBUTTONUP: {
			ReleaseCapture();
			RemoveProp(hWnd, TEXT("ISMOUSEDOWN"));
		}
		break;

		case WM_MOUSEMOVE: {
			DWORD x = GET_X_LPARAM(lParam);
			int* pPos = (int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			
			if (!GetProp(hWnd, TEXT("ISMOUSEHOVER")) && *pPos <= x && x <= *pPos + SPLITTER_WIDTH) {
				TRACKMOUSEEVENT tme = {sizeof(TRACKMOUSEEVENT), TME_LEAVE, hWnd, 0};
				TrackMouseEvent(&tme);	
				SetProp(hWnd, TEXT("ISMOUSEHOVER"), (HANDLE)1);
			}
			
			if (GetProp(hWnd, TEXT("ISMOUSEDOWN")) && x > 0 && x < 32000) {
				*pPos = x;
				SendMessage(hWnd, WM_SIZE, 0, 0);
			}
		}
		break;
		
		case WM_MOUSELEAVE: {
			SetProp(hWnd, TEXT("ISMOUSEHOVER"), 0);
		}
		break;

		case WM_MOUSEWHEEL: {
			if (LOWORD(wParam) == MK_CONTROL) {
				SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1: -1, 0);
				return 1;
			}
		}
		break;
		
		case WM_KEYDOWN: {
			if (SendMessage(hWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;			
		}
		break;		

		case WM_CONTEXTMENU: {
			if ((HWND)wParam == GetDlgItem(hWnd, IDC_TABLELIST)) {
				HWND hListWnd = (HWND)wParam;

				POINT p;
				GetCursorPos(&p);
				ScreenToClient(hListWnd, &p);

				int pos = SendMessage(hListWnd, LB_ITEMFROMPOINT, 0, (WPARAM)MAKELPARAM(p.x, p.y));
				int currPos = ListBox_GetCurSel(hListWnd);

				if (pos != -1 && currPos != pos) {
					ListBox_SetCurSel(hListWnd, pos);
					SendMessage(hWnd, WM_COMMAND, MAKELPARAM(IDC_TABLELIST, LBN_SELCHANGE), (LPARAM)hListWnd);
				}

				if (pos != -1 && currPos != -1 && getStoredValue(TEXT("editable"), 1) == 1) {
					ClientToScreen(hListWnd, &p);
					HMENU hMenu = GetProp(hWnd, TEXT("LISTMENU"));
					TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
				}
			}
		}
		break;
		
		case WM_CTLCOLORLISTBOX: {
			SetBkColor((HDC)wParam, *(int*)GetProp(hWnd, TEXT("BACKCOLOR")));
			SetTextColor((HDC)wParam, *(int*)GetProp(hWnd, TEXT("TEXTCOLOR")));
			return (INT_PTR)(HBRUSH)GetProp(hWnd, TEXT("BACKBRUSH"));	
		}
		break;		
		
		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDC_BUTTON) {
				struct DLGITEMTEMPLATEEX tpl = {0};
				if (IDOK == DialogBoxIndirectParam ((HINSTANCE)GetModuleHandle(NULL), (LPCDLGTEMPLATE)&tpl, hWnd, &cbDlgAddTable, (LPARAM)hWnd)) {
					HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
					ListBox_SetCurSel(hListWnd, ListBox_GetCount(hListWnd) - 1);
					SendMessage(hWnd, WMU_UPDATE_COUNTS, 0, 0);
					SetFocus(hListWnd);
					SendMessage(hWnd, WM_COMMAND, MAKELPARAM(IDC_TABLELIST, LBN_SELCHANGE), (LPARAM)hListWnd);
				}
			}

			if (cmd == IDC_TABLELIST && HIWORD(wParam) == LBN_SELCHANGE)
				SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
				
			if (cmd == IDC_TABLELIST && HIWORD(wParam) == LBN_SETFOCUS) 
				SetProp(hWnd, TEXT("LASTFOCUS"), (HWND)lParam);					

			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROWS || cmd == IDM_COPY_COLUMN) {
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				
				int colCount = Header_GetItemCount(hHeader);
				int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
				int selCount = ListView_GetSelectedCount(hGridWnd);

				if (rowNo == -1 ||
					rowNo >= rowCount ||					
					colCount == 0 ||
					cmd == IDM_COPY_CELL && colNo == -1 || 
					cmd == IDM_COPY_CELL && colNo >= colCount || 
					cmd == IDM_COPY_COLUMN && colNo == -1 || 
					cmd == IDM_COPY_COLUMN && colNo >= colCount || 					
					cmd == IDM_COPY_ROWS && selCount == 0) {
					setClipboardText(TEXT(""));
					return 0;
				}

				int cacheSize = *(int*)GetProp(hWnd, TEXT("CACHESIZE"));
				int cacheOffset = *(int*)GetProp(hWnd, TEXT("CACHEOFFSET"));
				int minRowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
				int maxRowNo = minRowNo;
				do {
					int rowNo = ListView_GetNextItem(hGridWnd, maxRowNo, LVNI_SELECTED);
					if (rowNo == -1)
						break;
					maxRowNo = rowNo;	
				} while (maxRowNo <= cacheOffset + cacheSize);

				if (rowNo < cacheOffset || 
					rowNo >= cacheOffset + cacheSize ||
					cmd == IDM_COPY_COLUMN && selCount == 1 && (0 < cacheOffset || rowCount > cacheSize) ||
					cmd == IDM_COPY_COLUMN && selCount != 1 && (minRowNo < cacheOffset || maxRowNo >= cacheOffset + cacheSize) ||					
					cmd == IDM_COPY_ROWS && selCount > 1 && (minRowNo < cacheOffset || maxRowNo >= cacheOffset + cacheSize) 
				) {
					TCHAR msg[1024];
					_sntprintf(msg, 1024, TEXT("The operation is available only for full-cached resultset.\nThis resultset has %i rows and a cache size is %i.\nCheck Wiki to learn how to increase the cache size."), rowCount, cacheSize);
					MessageBox(hWnd, msg, 0, 0);
					return 0;
				}
						
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				TCHAR* delimiter = getStoredString(TEXT("column-delimiter"), TEXT("\t"));

				int len = 0;
				if (cmd == IDM_COPY_CELL) 
					len = _tcslen(cache[rowNo - cacheOffset][colNo]);
								
				if (cmd == IDM_COPY_ROWS) {
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1) {
						for (int colNo = 0; colNo < colCount; colNo++) {
							if (ListView_GetColumnWidth(hGridWnd, colNo)) 
								len += _tcslen(cache[rowNo - cacheOffset][colNo]) + 1; /* column delimiter */
						}
													
						len++; /* \n */		
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
				}				

				if (cmd == IDM_COPY_COLUMN) {
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1 && rowNo < rowCount) {
						len += _tcslen(cache[rowNo - cacheOffset][colNo]) + 1 /* \n */;
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
					} 
				}	

				TCHAR* buf = calloc(len + 1, sizeof(TCHAR));
				if (cmd == IDM_COPY_CELL)
					_tcscat(buf, cache[rowNo - cacheOffset][colNo]);
								
				if (cmd == IDM_COPY_ROWS) {
					int pos = 0;
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1) {
						for (int colNo = 0; colNo < colCount; colNo++) {
							if (ListView_GetColumnWidth(hGridWnd, colNo)) {
								int len = _tcslen(cache[rowNo - cacheOffset][colNo]);
								_tcsncpy(buf + pos, cache[rowNo - cacheOffset][colNo], len);
								buf[pos + len] = delimiter[0];
								pos += len + 1;
							}
						}

						buf[pos - (pos > 0)] = TEXT('\n');
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
					buf[pos - 1] = 0; // remove last \n
				}				
				
				if (cmd == IDM_COPY_COLUMN) {
					int pos = 0;
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1 && rowNo < rowCount) {
						int len = _tcslen(cache[rowNo - cacheOffset][colNo]);
						_tcsncpy(buf + pos, cache[rowNo - cacheOffset][colNo], len);
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
						if (rowNo != -1 && rowNo < rowCount)
							buf[pos + len] = TEXT('\n');
						pos += len + 1;								
					} 
				}
									
				setClipboardText(buf);
				free(buf);
				free(delimiter);
			}
			
			if (cmd == IDM_ADD_ROW && getStoredValue(TEXT("editable"), 1) == 1) {
				struct DLGITEMTEMPLATEEX tpl = {0};
				if (IDOK == DialogBoxIndirectParam ((HINSTANCE)GetModuleHandle(NULL), (LPCDLGTEMPLATE)&tpl, hWnd, &cbDlgAddRow, (LPARAM)hWnd)) {
					SendMessage(hWnd, WMU_UPDATE_ROW_COUNT, 0, 0);
					SendMessage(hWnd, WMU_UPDATE_DATA, 0, 0);
					SetFocus(GetDlgItem(hWnd, IDC_GRID));
				}	
			}
			
			if (cmd == IDM_DELETE_ROW) {
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);

				if (*(int*)GetProp(hWnd, TEXT("HASROWID")) == 0 || getStoredValue(TEXT("editable"), 1) == 0)
					return 0;
				
				int selCount = ListView_GetSelectedCount(hGridWnd);
				int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
				int totalRowCount = *(int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
				sqlite3* db = (sqlite3*)GetProp(hWnd, TEXT("DB"));
				char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));
				
				char* query8 = 0;
				
				if (selCount == totalRowCount) {
					query8 = (char*)calloc (64 + strlen(tablename8), sizeof (char));
					sprintf(query8, "delete from \"%s\"", tablename8);	
				} else {		
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));		
					int cacheSize = *(int*)GetProp(hWnd, TEXT("CACHESIZE"));
					int cacheOffset = *(int*)GetProp(hWnd, TEXT("CACHEOFFSET"));				
					int minRowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					int maxRowNo = minRowNo;
					while (TRUE) {
						int rowNo = ListView_GetNextItem(hGridWnd, maxRowNo, LVNI_SELECTED);
						if (rowNo == -1)
							break;
						maxRowNo = rowNo;	
					}	
					
					if (minRowNo < cacheOffset || maxRowNo >= cacheOffset + cacheSize) {
						TCHAR msg[1024];
						_sntprintf(msg, 1024, TEXT("The operation is available only for full-cached resultset.\nThis resultset has %i rows and a cache size is %i.\nCheck Wiki to learn how to increase the cache size."), rowCount, cacheSize);
						MessageBox(hWnd, msg, 0, 0);
						return 0;
					}
					
					query8 = (char*)calloc (64 + strlen(tablename8) + selCount * 25, sizeof (char));
					sprintf(query8, "delete from \"%s\" where rowid in (0", tablename8);
					
					int colCount = Header_GetItemCount(hHeader);
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1) {
						char buf8[64];
						sprintf(buf8, ", %i", _ttoi(cache[rowNo - cacheOffset][colCount]));
						strcat(query8, buf8);
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
					strcat(query8, ")");
				}
				
				if (SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0)) {
					ListView_SetItemState(hGridWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
					SendMessage(hWnd, WMU_UPDATE_ROW_COUNT, 0, 0);
					SendMessage(hWnd, WMU_UPDATE_DATA, 0, 0);
				} else	
					MessageBoxA(hWnd, sqlite3_errmsg(db), "Error", MB_OK);
					
				free(query8);
				return 0;	
			}
			
			if (cmd == IDM_BLOB) {
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				
				if (*(int*)GetProp(hWnd, TEXT("HASROWID")) == 0)
					return MessageBox(hWnd, TEXT("The table doesn't have a rowid.\nBlob saving is not supported."), NULL, MB_OK);
				
				int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
				if (rowNo != -1) {
					int cacheSize = *(int*)GetProp(hWnd, TEXT("CACHESIZE"));
					int* pCacheOffset = (int*)GetProp(hWnd, TEXT("CACHEOFFSET"));				
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));

					if (rowNo < *pCacheOffset || rowNo >= *pCacheOffset + cacheSize)
						return 0;

					int colCount = Header_GetItemCount(hHeader);				
					sqlite3_stmt *stmt;

					sqlite3* db = (sqlite3*)GetProp(hWnd, TEXT("DB"));
					char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
					TCHAR colName16[MAX_COLUMN_LENGTH + 1];
					Header_GetItemText(hHeader, colNo, colName16, MAX_COLUMN_LENGTH + 1);
					char* colName8 = utf16to8(colName16);
					char query8[1024 + strlen(tablename8)];
					sprintf(query8, "select %s from \"%s\" where rowid = %i", colName8, tablename8, _ttoi(cache[rowNo - *pCacheOffset][colCount]));
					free(colName8);
					
					TCHAR path16[MAX_PATH] = {0};
					TCHAR filter16[] = TEXT("Images (*.jpg, *.gif, *.png, *.bmp)\0*.jpg;*.jpeg;*.gif;*.png;*.bmp\0Binary(*.bin,*.dat)\0*.bin,*.dat\0All\0*.*\0");
					if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0) && 
						SQLITE_ROW == sqlite3_step(stmt) &&
						saveFile(path16, filter16, TEXT(""), hWnd)) {
						FILE *fp = _tfopen (path16, TEXT("wb"));
						if (fp) {
							fwrite(sqlite3_column_blob(stmt, 0), sqlite3_column_bytes(stmt, 0), 1, fp);
							fclose(fp);						
						} else {
							MessageBox(hWnd, TEXT("Opening the file for writing is failed"), NULL, MB_OK);
						}
					}
					sqlite3_finalize(stmt);
				}	
			}

			if (cmd == IDM_HIDE_COLUMN) {
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				SendMessage(hWnd, WMU_HIDE_COLUMN, colNo, 0);
			}						

			if (cmd == IDM_DROP && IDYES == MessageBox(hWnd, TEXT("Are you sure?"), TEXT("Confirmation"), MB_YESNO)) {		
				sqlite3* db = (sqlite3*)GetProp(hWnd, TEXT("DB"));
				char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));

				HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
				int pos = ListBox_GetCurSel(hListWnd);
				BOOL isTable = SendMessage(hListWnd, LB_GETITEMDATA, pos, 0);

				char query8[1024];
				snprintf(query8, 1023, "drop %s \"%s\"", isTable ? "table" : "view", tablename8);
				if (SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0)) {
					ListBox_DeleteString(hListWnd, ListBox_GetCurSel(hListWnd));
					if (LB_ERR != ListBox_SetCurSel(hListWnd, 0)) {
						SendMessage(hWnd, WM_COMMAND, MAKELPARAM(IDC_TABLELIST, LBN_SELCHANGE), (LPARAM)hListWnd);
					} else {
						SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
					}
					SendMessage(hWnd, WMU_UPDATE_COUNTS, 0, 0);
				} else {
					MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
				}
			}
			
			if (cmd == IDM_FILTER_ROW || cmd == IDM_DARK_THEME) {
				HMENU hMenu = (HMENU)GetProp(hWnd, TEXT("GRIDMENU"));
				int* pOpt = (int*)GetProp(hWnd, cmd == IDM_FILTER_ROW ? TEXT("FILTERROW") : TEXT("DARKTHEME"));
				*pOpt = (*pOpt + 1) % 2;
				Menu_SetItemState(hMenu, cmd, *pOpt ? MFS_CHECKED : 0);
				
				UINT msg = cmd == IDM_FILTER_ROW ? WMU_SET_HEADER_FILTERS : WMU_SET_THEME;
				SendMessage(hWnd, msg, 0, 0);				
			}
		}
		break;

		case WM_NOTIFY : {
			NMHDR* pHdr = (LPNMHDR)lParam;
			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_GETDISPINFO) {
				HWND hGridWnd = pHdr->hwndFrom;
				sqlite3* db = (sqlite3*)GetProp(hWnd, TEXT("DB"));
				int orderBy = *(int*)GetProp(hWnd, TEXT("ORDERBY"));

				LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
				LV_ITEM* pItem= &(pDispInfo)->item;
				int cacheSize = *(int*)GetProp(hWnd, TEXT("CACHESIZE"));
				int* pCacheOffset = (int*)GetProp(hWnd, TEXT("CACHEOFFSET"));
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));

			   	if ((pItem->iItem < *pCacheOffset) || (pItem->iItem >= *pCacheOffset + cacheSize) || (*pCacheOffset == -1)) {
			   		SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			   		*pCacheOffset = pItem->iItem;

			   		HWND hHeader = ListView_GetHeader(hGridWnd);
			   		int colCount = Header_GetItemCount(hHeader);
					sqlite3_stmt *stmt;

					char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));
					char* where8 = (char*)GetProp(hWnd, TEXT("WHERE8"));
					int hasRowid = *(int*)GetProp(hWnd, TEXT("HASROWID"));

					char query8[1024 + strlen(tablename8) + strlen(where8)];
					char orderBy8[32] = {0};
					if (orderBy > 0)
						sprintf(orderBy8, "order by %i", orderBy);
					if (orderBy < 0)
						sprintf(orderBy8, "order by %i desc", -orderBy);

					sprintf(query8, "select *%s from \"%s\" %s %s limit %i offset %i", hasRowid ? ", rowid" : "", tablename8, where8, orderBy8, cacheSize, *pCacheOffset);
					if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
						bindFilterValues(hHeader, stmt);

						int rowNo = 0;
						while (SQLITE_ROW == sqlite3_step(stmt)) {
							cache[rowNo] = (TCHAR**)calloc (colCount + hasRowid, sizeof (TCHAR*));

							for (int colNo = 0; colNo < colCount; colNo++) {
								cache[rowNo][colNo] = sqlite3_column_type(stmt, colNo) != SQLITE_BLOB ?
									utf8to16((const char*)sqlite3_column_text(stmt, colNo)) :
									utf8to16(BLOB_VALUE);
							}
							
							if (hasRowid)
								cache[rowNo][colCount] = utf8to16((const char*)sqlite3_column_text(stmt, colCount));

							rowNo++;
						}
					}
					sqlite3_finalize(stmt);
			   	}

				if(pItem->mask & LVIF_TEXT)
					pItem->pszText = cache[pItem->iItem - *pCacheOffset][pItem->iSubItem];
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_COLUMNCLICK) {
				NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
				if (HIWORD(GetKeyState(VK_CONTROL))) 
					return SendMessage(hWnd, WMU_HIDE_COLUMN, lv->iSubItem, 0);
					
				int colNo = lv->iSubItem + 1;
				int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
				int orderBy = *pOrderBy;
				*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
				SendMessage(hWnd, WMU_UPDATE_DATA, 0, 0);				
			}			

			if (pHdr->idFrom == IDC_GRID && (pHdr->code == (DWORD)NM_CLICK || pHdr->code == (DWORD)NM_RCLICK)) {
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				SendMessage(hWnd, WMU_SET_CURRENT_CELL, ia->iItem, ia->iSubItem);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_DBLCLK) {
				SendMessage(hWnd, WMU_EDIT_CELL, 1, 0);
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_CLICK && HIWORD(GetKeyState(VK_MENU))) {	
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int cacheOffset = *(int*)GetProp(hWnd, TEXT("CACHEOFFSET"));

				TCHAR* url = extractUrl(cache[ia->iItem - cacheOffset][ia->iSubItem]);
				ShellExecute(0, TEXT("open"), url, 0, 0 , SW_SHOW);
				free(url);
			}			

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK) {				
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				
				int colCount = Header_GetItemCount(hHeader);
				int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));

				BOOL isBlob = FALSE;
				if (rowNo != -1 && colNo != -1 && rowNo < rowCount && colCount != 0 && colNo < colCount) {
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
					int cacheSize = *(int*)GetProp(hWnd, TEXT("CACHESIZE"));
					int cacheOffset = *(int*)GetProp(hWnd, TEXT("CACHEOFFSET"));
					isBlob = rowNo >= cacheOffset && rowNo < cacheOffset + cacheSize && (_tcscmp(cache[rowNo - cacheOffset][colNo], TEXT(BLOB_VALUE)) == 0);
				}
				
				POINT p;
				GetCursorPos(&p);					
				TrackPopupMenu(GetProp(hWnd, isBlob ? TEXT("BLOBMENU") : TEXT("GRIDMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
				NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
				if (lv->uOldState != lv->uNewState && (lv->uNewState & LVIS_SELECTED))				
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, lv->iItem, *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));	
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				HWND hGridWnd = pHdr->hwndFrom;
				BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
				
				if (kd->wVKey == 0x43) { // C
					BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
					BOOL isShift = HIWORD(GetKeyState(VK_SHIFT)); 
					BOOL isCopyColumn = getStoredValue(TEXT("copy-column"), 0) && ListView_GetSelectedCount(hGridWnd) > 1;
					if (!isCtrl && !isShift)
						return FALSE;
						
					int action = !isShift && !isCopyColumn ? IDM_COPY_CELL : isCtrl || isCopyColumn ? IDM_COPY_COLUMN : IDM_COPY_ROWS;
					SendMessage(hWnd, WM_COMMAND, action, 0);
					return TRUE;
				}

				if (kd->wVKey == 0x41 && isCtrl) { // Ctrl + A
					SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));					
					ListView_SetItemState(hGridWnd, -1, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED);
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
					SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
					InvalidateRect(hGridWnd, NULL, TRUE);
				}
								
				if (kd->wVKey == 0x20 && isCtrl) { // Ctrl + Space				
					SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);					
					return TRUE;
				}				

				if ((kd->wVKey == VK_F2 || kd->wVKey == VK_SPACE && !isCtrl)) { // F2 or Space
					SendMessage(hWnd, WMU_EDIT_CELL, kd->wVKey == VK_F2, 0);
					return TRUE;
				}
				
				if (kd->wVKey == VK_LEFT || kd->wVKey == VK_RIGHT) {
					int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) + (kd->wVKey == VK_RIGHT ? 1 : -1);
					colNo = colNo < 0 ? colCount - 1 : colNo > colCount - 1 ? 0 : colNo;
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, *(int*)GetProp(hWnd, TEXT("CURRENTROWNO")), colNo);
					return TRUE;
				}
				
				if (kd->wVKey == VK_DELETE)
					SendMessage(hWnd, WM_COMMAND, IDM_DELETE_ROW, 0);
			}

			if (pHdr->code == HDN_ITEMCHANGED && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(hWnd, IDC_GRID)))
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
				
			if (pHdr->code == (UINT)NM_SETFOCUS)
				SetProp(hWnd, TEXT("LASTFOCUS"), pHdr->hwndFrom);
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (UINT)NM_CUSTOMDRAW) {
				int result = CDRF_DODEFAULT;
				
				NMLVCUSTOMDRAW* pCustomDraw = (LPNMLVCUSTOMDRAW)lParam;				
				if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT) 
					result = CDRF_NOTIFYITEMDRAW;
	
				if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
					if (ListView_GetItemState(pHdr->hwndFrom, pCustomDraw->nmcd.dwItemSpec, LVIS_SELECTED)) {
						pCustomDraw->nmcd.uItemState &= ~CDIS_SELECTED;
						result = CDRF_NOTIFYSUBITEMDRAW;
					} else {
						pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, pCustomDraw->nmcd.dwItemSpec % 2 == 0 ? TEXT("BACKCOLOR") : TEXT("BACKCOLOR2"));
					}				
				}
				
				if (pCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM)) {
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
					BOOL isCurrCell = (pCustomDraw->nmcd.dwItemSpec == (DWORD)rowNo) && (pCustomDraw->iSubItem == colNo);
					pCustomDraw->clrText = *(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
					pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, isCurrCell ? TEXT("CURRENTCELLCOLOR") : TEXT("SELECTIONBACKCOLOR"));
				}
	
				return result;
			}				
		}
		break;
		
		// wParam = colNo
		case WMU_HIDE_COLUMN: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);		
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colNo = (int)wParam;

			HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
			SetWindowLongPtr(hEdit, GWLP_USERDATA, (LONG_PTR)ListView_GetColumnWidth(hGridWnd, colNo));				
			ListView_SetColumnWidth(hGridWnd, colNo, 0); 
			InvalidateRect(hHeader, NULL, TRUE);
		}
		break;
		
		case WMU_SHOW_COLUMNS: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			for (int colNo = 0; colNo < colCount; colNo++) {
				if (ListView_GetColumnWidth(hGridWnd, colNo) == 0) {
					HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
					ListView_SetColumnWidth(hGridWnd, colNo, (int)GetWindowLongPtr(hEdit, GWLP_USERDATA));
				}
			}

			InvalidateRect(hGridWnd, NULL, TRUE);		
		}
		break;
		
		case WMU_EDIT_CELL: {
			BOOL withText = (BOOL)wParam;
			if (getStoredValue(TEXT("editable"), 1) == 0)
				return FALSE;
				
			if (*(int*)GetProp(hWnd, TEXT("HASROWID")) == 0)
				return MessageBox(hWnd, TEXT("Can't edit the value"), NULL, MB_OK);
			
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);	
			int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
			int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));					
			int cacheSize = *(int*)GetProp(hWnd, TEXT("CACHESIZE"));
			int cacheOffset = *(int*)GetProp(hWnd, TEXT("CACHEOFFSET"));				
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));

			if (rowNo == - 1 || rowNo < cacheOffset || rowNo >= cacheOffset + cacheSize)
				return FALSE;
				
			if (_tcscmp(cache[rowNo - cacheOffset][colNo], TEXT(BLOB_VALUE)) == 0) 
				return MessageBeep(0);	
							
			RECT rect;
			ListView_GetSubItemRect(hGridWnd, rowNo, colNo, LVIR_BOUNDS, &rect);	
			int h = rect.bottom - rect.top;
			int w = ListView_GetColumnWidth(hGridWnd, colNo);
			
			HWND hEdit = CreateWindowEx(0, WC_EDIT, withText ? cache[rowNo - cacheOffset][colNo] : 0, 
				WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, rect.left, rect.top, w, h, 
				hGridWnd, (HMENU)IDC_CELL_EDIT, GetModuleHandle(NULL), NULL);
			SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
			SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)&cbNewCellEdit));
			if (withText) {
				int end = GetWindowTextLength(hEdit);
				SendMessage(hEdit, EM_SETSEL, end, end);				
			}
			SetFocus(hEdit);		
		}
		break;		

		case WMU_UPDATE_GRID: {
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			sqlite3* db = (sqlite3*)GetProp(hWnd, TEXT("DB"));
			int* pHasRowid = (int*)GetProp(hWnd, TEXT("HASROWID"));
			int filterAlign = *(int*)GetProp(hWnd, TEXT("FILTERALIGN"));
			
			SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hGridWnd);

			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			ListView_SetItemCount(hGridWnd, 0);
			SendMessage(hWnd, WMU_SET_CURRENT_CELL, 0, 0);

			int colCount = Header_GetItemCount(hHeader);
			for (int colNo = 0; colNo < colCount; colNo++) 
				DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

			for (int colNo = 0; colNo < colCount; colNo++)
				ListView_DeleteColumn(hGridWnd, colCount - colNo - 1);

			char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));
			int pos = ListBox_GetCurSel(hListWnd);
			if (pos == -1) {
				SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
				return TRUE;
			}

			TCHAR name16[256];
			ListBox_GetText(hListWnd, pos, name16);
			char* name8 = utf16to8(name16);
			sprintf(tablename8, name8);
			free(name8);

			TCHAR buf[255];
			int type = SendMessage(hListWnd, LB_GETITEMDATA, pos, 0);
			_sntprintf(buf, 255, type ? TEXT(" TABLE"): TEXT("  VIEW"));
			SendMessage(hStatusWnd, SB_SETTEXT, SB_TYPE, (LPARAM)buf);

			ListView_SetItemCount(hGridWnd, 0);

			sqlite3_stmt *stmt;
			char query8[1024 + strlen(tablename8)];
			sprintf(query8, "select rowid from \"%s\" where 1 = 2", tablename8);
			*pHasRowid = sqlite3_exec(db, query8, 0, 0, 0) == SQLITE_OK;
			
			sprintf(query8, "select *%s from \"%s\" limit 1", *pHasRowid ? ", rowid" : "", tablename8);
			if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
				colCount = sqlite3_column_count(stmt) - *pHasRowid;
				for (int colNo = 0; colNo < colCount; colNo++) {
					TCHAR* name16 = utf8to16(sqlite3_column_name(stmt, colNo));
					const char* decltype8 = sqlite3_column_decltype(stmt, colNo);
					int fmt = LVCFMT_LEFT;
					if (decltype8) {
						char type8[strlen(decltype8) + 1];
						sprintf(type8, decltype8);
						strlwr(type8);

						if (strstr(type8, "int") || strstr(type8, "float") || strstr(type8, "double") ||
							strstr(type8, "real") || strstr(type8, "decimal") || strstr(type8, "numeric"))
							fmt = LVCFMT_RIGHT;
					}
					
					ListView_AddColumn(hGridWnd, name16, fmt);
					free(name16);
				}

				int align = filterAlign == -1 ? ES_LEFT : filterAlign == 1 ? ES_RIGHT : ES_CENTER;
				for (int colNo = 0; colNo < colCount; colNo++) {
					RECT rc;
					Header_GetItemRect(hHeader, colNo, &rc);
					HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, align | ES_AUTOHSCROLL | WS_CHILD | WS_BORDER | WS_TABSTOP, 0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
					SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
					SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
				}
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
			}
			sqlite3_finalize(stmt);
			
			if (getStoredValue(TEXT("editable"), 1)) {
				HMENU hGridMenu = GetProp(hWnd, TEXT("GRIDMENU"));
				DeleteMenu(hGridMenu, IDM_ADD_ROW, MF_BYCOMMAND);
				DeleteMenu(hGridMenu, IDM_DELETE_ROW, MF_BYCOMMAND);
				DeleteMenu(hGridMenu, IDM_SEPARATOR1, MF_BYCOMMAND);				
				if (*pHasRowid) {
					InsertMenu(hGridMenu, IDM_SEPARATOR2, MF_BYCOMMAND | MF_STRING, IDM_SEPARATOR1, NULL);						
					InsertMenu(hGridMenu, IDM_SEPARATOR2, MF_BYCOMMAND | MF_STRING, IDM_ADD_ROW, TEXT("Add row"));
					InsertMenu(hGridMenu, IDM_SEPARATOR2, MF_BYCOMMAND | MF_STRING, IDM_DELETE_ROW, TEXT("Delete row(s)"));
				}
			}

			*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;
			SendMessage(hWnd, WMU_UPDATE_ROW_COUNT, 0, 0);
			SendMessage(hWnd, WMU_UPDATE_DATA, 0, 0);
			SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);

			SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_DATA: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			*(int*)GetProp(hWnd, TEXT("CACHEOFFSET")) = -1;
			ListView_SetItemCount(hGridWnd, ListView_GetItemCount(hGridWnd));
			UpdateWindow(hGridWnd);
		}
		break;

		case WMU_UPDATE_ROW_COUNT: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			sqlite3* db = (sqlite3*)GetProp(hWnd, TEXT("DB"));
			char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));
			char* where8 = (char*)GetProp(hWnd, TEXT("WHERE8"));
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));			

			TCHAR where16[MAX_TEXT_LENGTH] = {0};
			_tcscat(where16, TEXT("where (1 = 1) "));

			for (int colNo = 0; colNo < colCount; colNo++) {
				HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
				if (GetWindowTextLength(hEdit) > 0) {
					TCHAR colname16[256] = {0};
					Header_GetItemText(hHeader, colNo, colname16, 255);
					_tcscat(where16, TEXT(" and \""));
					_tcscat(where16, colname16);

					TCHAR buf16[2] = {0};
					GetWindowText(hEdit, buf16, 2);
					_tcscat(where16,
						buf16[0] == TEXT('=') ? TEXT("\" = ? ") :
						buf16[0] == TEXT('!') ? TEXT("\" not like '%%' || ? || '%%' ") :
						buf16[0] == TEXT('>') ? TEXT("\" > ? ") :
						buf16[0] == TEXT('<') ? TEXT("\" < ? ") :
						TEXT("\" like '%%' || ? || '%%' "));
				}
			}
			char* _where8 = utf16to8(where16);
			sprintf(where8, _where8);
			free(_where8);

			sqlite3_stmt *stmt;
			char query8[1024 + strlen(tablename8) + strlen(where8)];
			sprintf(query8, "select count(*) from \"%s\" %s", tablename8, where8);
			int rowCount = -1;
			if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
				bindFilterValues(hHeader, stmt);

				if (SQLITE_ROW == sqlite3_step(stmt))
					rowCount = sqlite3_column_int(stmt, 0);
			}
			sqlite3_finalize(stmt);

			if (strcmp(where8, "where (1 = 1) ") == 0)
				*pTotalRowCount = rowCount;
			*pRowCount = rowCount;				

			if (rowCount == -1)
				MessageBox(hWnd, TEXT("Error occured while fetching data"), NULL, MB_OK);
			else
				ListView_SetItemCount(hGridWnd, rowCount);

			TCHAR buf[255];
			_sntprintf(buf, 255, TEXT(" Rows: %i/%i"), rowCount, *pTotalRowCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)buf);

			return rowCount;
		}
		break;

		case WMU_UPDATE_FILTER_SIZE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			SendMessage(hHeader, WM_SIZE, 0, 0);
			for (int colNo = 0; colNo < colCount; colNo++) {
				RECT rc;
				Header_GetItemRect(hHeader, colNo, &rc);
				int h2 = round((rc.bottom - rc.top) / 2);
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left, h2, rc.right - rc.left, h2 + 1, SWP_NOZORDER);
			}
		}
		break;
		
		case WMU_SET_HEADER_FILTERS: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int isFilterRow = *(int*)GetProp(hWnd, TEXT("FILTERROW"));
			int colCount = Header_GetItemCount(hHeader);
			
			SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
			LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
			styles = isFilterRow ? styles | HDS_FILTERBAR : styles & (~HDS_FILTERBAR);
			SetWindowLongPtr(hHeader, GWL_STYLE, styles);
					
			for (int colNo = 0; colNo < colCount; colNo++) 		
				ShowWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), isFilterRow ? SW_SHOW : SW_HIDE);

			// Bug fix: force Windows to redraw header
			SetWindowPos(hGridWnd, 0, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOMOVE);
			SendMessage(getMainWindow(hWnd), WM_SIZE, 0, 0);			

			if (isFilterRow)				
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);											
						
			SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;

		case WMU_AUTO_COLUMN_SIZE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);

			for (int colNo = 0; colNo < colCount - 1; colNo++)
				ListView_SetColumnWidth(hGridWnd, colNo, colNo < colCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

			if (colCount == 1 && ListView_GetColumnWidth(hGridWnd, 0) < 100)
				ListView_SetColumnWidth(hGridWnd, 0, 100);
				
			int maxWidth = getStoredValue(TEXT("max-column-width"), 300);
			if (colCount > 1) {
				for (int colNo = 0; colNo < colCount; colNo++) {
					if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
						ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
				}
			}

			// Fix last column				
			if (colCount > 1) {
				int colNo = colCount - 1;
				ListView_SetColumnWidth(hGridWnd, colNo, LVSCW_AUTOSIZE);
				TCHAR name16[MAX_COLUMN_LENGTH + 1];
				Header_GetItemText(hHeader, colNo, name16, MAX_COLUMN_LENGTH);

				SIZE s = {0};
				HDC hDC = GetDC(hHeader);
				HFONT hOldFont = (HFONT)SelectObject(hDC, (HFONT)GetProp(hWnd, TEXT("FONT")));
				GetTextExtentPoint32(hDC, name16, _tcslen(name16), &s);
				SelectObject(hDC, hOldFont);
				ReleaseDC(hHeader, hDC);

				int w = s.cx + 12;
				if (ListView_GetColumnWidth(hGridWnd, colNo) < w)
					ListView_SetColumnWidth(hGridWnd, colNo, w);
					
				if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
					ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);					
			}

			SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hGridWnd, NULL, TRUE);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;
		
		// wParam = rowNo, lParam = colNo
		case WMU_SET_CURRENT_CELL: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)0);
						
			int *pRowNo = (int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
			int *pColNo = (int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
			if (*pRowNo == wParam && *pColNo == lParam)
				return 0;
			
			RECT rc, rc2;
			ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
			if (*pColNo == 0)
				rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);			
			InvalidateRect(hGridWnd, &rc, TRUE);
			
			*pRowNo = wParam;
			*pColNo = lParam;
			ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
			if (*pColNo == 0)
				rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);
			InvalidateRect(hGridWnd, &rc, FALSE);
			
			GetClientRect(hGridWnd, &rc2);
			int w = rc.right - rc.left;
			int dx = rc2.right < rc.right ? rc.left - rc2.right + w : rc.left < 0 ? rc.left : 0;

			ListView_Scroll(hGridWnd, dx, 0);
			
			TCHAR buf[32] = {0};
			if (*pColNo != - 1 && *pRowNo != -1)
				_sntprintf(buf, 32, TEXT(" %i:%i"), *pRowNo + 1, *pColNo + 1);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, (LPARAM)buf);
		}
		break;		

		case WMU_RESET_CACHE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			int cacheSize = *(int*)GetProp(hWnd, TEXT("CACHESIZE"));
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			*(int*)GetProp(hWnd, TEXT("CACHEOFFSET")) = -1;
			int hasRowid = *(int*)GetProp(hWnd, TEXT("HASROWID"));
			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd)) + hasRowid;
			if (colCount == 0)
				return 0;
			for (int rowNo = 0;  rowNo < cacheSize; rowNo++) {
				if (cache[rowNo]) {
					for (int colNo = 0; colNo < colCount; colNo++)
						if (cache[rowNo][colNo])
							free(cache[rowNo][colNo]);

					free(cache[rowNo]);
				}
				cache[rowNo] = 0;
			}
		}
		break;
		
		// wParam - size delta
		case WMU_SET_FONT: {
			int* pFontSize = (int*)GetProp(hWnd, TEXT("FONTSIZE"));
			if (*pFontSize + wParam < 10 || *pFontSize + wParam > 48)
				return 0;
			*pFontSize += wParam;
			DeleteFont(GetProp(hWnd, TEXT("FONT")));

			HFONT hFont = CreateFont (*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, (TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
			HWND hBtnWnd = GetDlgItem(hWnd, IDC_BUTTON);
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SendMessage(hBtnWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hListWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hGridWnd, WM_SETFONT, (LPARAM)hFont, TRUE);

			HWND hHeader = ListView_GetHeader(hGridWnd);
			for (int colNo = 0; colNo < Header_GetItemCount(hHeader); colNo++)
				SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), WM_SETFONT, (LPARAM)hFont, TRUE);
				
			int w = 0;
			HDC hDC = GetDC(hListWnd);
			HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);
			TCHAR buf[MAX_TABLE_LENGTH]; 
			for (int i = 0; i < ListBox_GetCount(hListWnd); i++) {
				SIZE s = {0};
				ListBox_GetText(hListWnd, i, buf);
				GetTextExtentPoint32(hDC, buf, _tcslen(buf), &s);			
				if (w < s.cx)
					w = s.cx;
			}
			SelectObject(hDC, hOldFont);
			ReleaseDC(hHeader, hDC);
			SendMessage(hListWnd, LB_SETHORIZONTALEXTENT, w, 0);
				
			SetProp(hWnd, TEXT("FONT"), hFont);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
			PostMessage(hWnd, WM_SIZE, 0, 0);
		}
		break;
		
		case WMU_SET_THEME: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			BOOL isDark = *(int*)GetProp(hWnd, TEXT("DARKTHEME"));
			
			int textColor = !isDark ? getStoredValue(TEXT("text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("text-color-dark"), RGB(220, 220, 220));
			int backColor = !isDark ? getStoredValue(TEXT("back-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("back-color-dark"), RGB(32, 32, 32));
			int backColor2 = !isDark ? getStoredValue(TEXT("back-color2"), RGB(240, 240, 240)) : getStoredValue(TEXT("back-color2-dark"), RGB(52, 52, 52));
			int filterTextColor = !isDark ? getStoredValue(TEXT("filter-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("filter-text-color-dark"), RGB(255, 255, 255));
			int filterBackColor = !isDark ? getStoredValue(TEXT("filter-back-color"), RGB(240, 240, 240)) : getStoredValue(TEXT("filter-back-color-dark"), RGB(60, 60, 60));
			int currCellColor = !isDark ? getStoredValue(TEXT("current-cell-back-color"), RGB(70, 96, 166)) : getStoredValue(TEXT("current-cell-back-color-dark"), RGB(32, 62, 62));
			int selectionTextColor = !isDark ? getStoredValue(TEXT("selection-text-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("selection-text-color-dark"), RGB(220, 220, 220));
			int selectionBackColor = !isDark ? getStoredValue(TEXT("selection-back-color"), RGB(10, 36, 106)) : getStoredValue(TEXT("selection-back-color-dark"), RGB(72, 102, 102));			
			int splitterColor = !isDark ? getStoredValue(TEXT("splitter-color"), GetSysColor(COLOR_BTNFACE)) : getStoredValue(TEXT("splitter-color-dark"), GetSysColor(COLOR_BTNFACE));			
			
			*(int*)GetProp(hWnd, TEXT("TEXTCOLOR")) = textColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR")) = backColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR2")) = backColor2;			
			*(int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")) = filterTextColor;
			*(int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")) = filterBackColor;
			*(int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")) = currCellColor;
			*(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")) = selectionTextColor;
			*(int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")) = selectionBackColor;
			*(int*)GetProp(hWnd, TEXT("SPLITTERCOLOR")) = splitterColor;

			DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));			
			DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));			
			DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));			
			SetProp(hWnd, TEXT("BACKBRUSH"), CreateSolidBrush(backColor));
			SetProp(hWnd, TEXT("FILTERBACKBRUSH"), CreateSolidBrush(filterBackColor));
			SetProp(hWnd, TEXT("SPLITTERBRUSH"), CreateSolidBrush(splitterColor));			

			ListView_SetTextColor(hGridWnd, textColor);			
			ListView_SetBkColor(hGridWnd, backColor);
			ListView_SetTextBkColor(hGridWnd, backColor);
			InvalidateRect(hWnd, NULL, TRUE);	
		}
		break;
		
		case WMU_HOT_KEYS: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
			if (wParam == VK_TAB) {
				HWND hFocus = GetFocus();
				HWND wnds[1000] = {0};
				EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumTabStopChildren, (LPARAM)wnds);

				int no = 0;
				while(wnds[no] && wnds[no] != hFocus)
					no++;

				int cnt = no;
				while(wnds[cnt])
					cnt++;

				no += isCtrl ? -1 : 1;
				SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isCtrl ? wnds[cnt - 1] : wnds[0]));
			}
			
			if (wParam == VK_F1) {
				SendMessage(hWnd, WM_COMMAND, IDM_ABOUT, 0);
				return TRUE;
			}
			
			if (wParam == VK_INSERT)
				SendMessage(hWnd, WM_COMMAND, IDM_ADD_ROW, 0);
			
			if (wParam == 0x20 && isCtrl) { // Ctrl + Space
				SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
				return TRUE;
			}
			
			if (wParam == VK_ESCAPE && getStoredValue(TEXT("exit-by-escape"), 1)) {
				SendMessage(GetParent(hWnd), WM_DESTROY, 0, 0);
				return TRUE;
			}			
			
			return FALSE;					
		}
		break;
		
		case WMU_HOT_CHARS: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
			return !_istprint(wParam) && (
				wParam == VK_ESCAPE || wParam == VK_F11 || wParam == VK_F1 ||
				wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7) ||
				wParam == VK_TAB || wParam == VK_RETURN ||
				isCtrl && (wParam == 0x46 || wParam == 0x20);
		}
		break;	

		case WMU_UPDATE_COUNTS: {
			int tCount = 0, vCount = 0;
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			for (int pos = 0; pos < ListBox_GetCount(hListWnd); pos++) {
				BOOL isTable = SendMessage(hListWnd, LB_GETITEMDATA, pos, 0);
				tCount += isTable;
				vCount += !isTable; 
			}	

			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			TCHAR buf[255];	
			_sntprintf(buf, 255, TEXT(" Tables: %i"), tCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_TABLE_COUNT, (LPARAM)buf);
			_sntprintf(buf, 255, TEXT(" Views: %i"), vCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_VIEW_COUNT, (LPARAM)buf);
		}				
		break;
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}


LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_KEYDOWN && SendMessage(getMainWindow(hWnd), WMU_HOT_KEYS, wParam, lParam))
		return 0;

	// Prevent beep
	if (msg == WM_CHAR && SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
		return 0;	

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_CTLCOLOREDIT) {
		HWND hMainWnd = getMainWindow(hWnd);
		SetBkColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
		SetTextColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR")));
		return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));	
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetClientRect(hWnd, &rc);
			HWND hMainWnd = getMainWindow(hWnd);
			BOOL isDark = *(int*)GetProp(hMainWnd, TEXT("DARKTHEME")); 

			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			LineTo(hDC, rc.right - 1, rc.bottom - 1);
			
			if (isDark) {
				DeleteObject(hPen);
				hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
				SelectObject(hDC, hPen);
				
				MoveToEx(hDC, 0, 0, 0);
				LineTo(hDC, 0, rc.bottom);
				MoveToEx(hDC, 0, rc.bottom - 1, 0);
				LineTo(hDC, rc.right, rc.bottom - 1);
				MoveToEx(hDC, 0, rc.bottom - 2, 0);
				LineTo(hDC, rc.right, rc.bottom - 2);
			}
			
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);
			
			return 0;
		}
		break;
		
		case WM_SETFOCUS: {
			SetProp(getMainWindow(hWnd), TEXT("LASTFOCUS"), hWnd);
		}
		break;

		case WM_KEYDOWN: {
			HWND hMainWnd = getMainWindow(hWnd);
			if (wParam == VK_RETURN) {
				SendMessage(hMainWnd, WMU_UPDATE_ROW_COUNT, 0, 0);
				SendMessage(hMainWnd, WMU_UPDATE_DATA, 0, 0);
				return 0;			
			}
			
			if (SendMessage(hMainWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;
		}
		break;
	
		// Prevent beep
		case WM_CHAR: {
			if (SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
				return 0;	
		}
		break;

		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewCellEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_DESTROY: {
			HWND hGridWnd = GetParent(hWnd);
			BOOL isChanged = !!GetProp(hWnd, TEXT("ISCHANGED"));
			if (isChanged) {			
				HWND hMainWnd = getMainWindow(hWnd);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = *(int*)GetProp(hMainWnd, TEXT("CURRENTROWNO"));
				int colNo = *(int*)GetProp(hMainWnd, TEXT("CURRENTCOLNO"));					
				int cacheSize = *(int*)GetProp(hMainWnd, TEXT("CACHESIZE"));
				int cacheOffset = *(int*)GetProp(hMainWnd, TEXT("CACHEOFFSET"));				
				TCHAR*** cache = (TCHAR***)GetProp(hMainWnd, TEXT("CACHE"));
				int colCount = Header_GetItemCount(hHeader);
				sqlite3* db = (sqlite3*)GetProp(hMainWnd, TEXT("DB"));
				char* tablename8 = (char*)GetProp(hMainWnd, TEXT("TABLENAME8"));
			
				int len = GetWindowTextLength(hWnd);
				TCHAR* value16 = (TCHAR*)calloc (len + 1, sizeof (TCHAR));
				GetWindowText(hWnd, value16, len + 1);
				
				TCHAR colName16[MAX_COLUMN_LENGTH + 1];
				Header_GetItemText(hHeader, colNo, colName16, MAX_COLUMN_LENGTH + 1);
				char* colName8 = utf16to8(colName16);
				char query8[1024 + strlen(tablename8) + strlen(colName8)];
				sprintf(query8, "update \"%s\" set \"%s\" = ?1 where rowid = %i", tablename8, colName8, _ttoi(cache[rowNo - cacheOffset][colCount]));
				free(colName8);

				BOOL rc = FALSE;
				sqlite3_stmt *stmt;
				if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
					char* value8 = utf16to8(value16);
					bindValue(stmt, 1, value8);
					free(value8);
					rc = SQLITE_DONE == sqlite3_step(stmt);					 
				}

				if (rc) {
					free(cache[rowNo - cacheOffset][colNo]);
					cache[rowNo - cacheOffset][colNo] = value16;
				} else {
					MessageBoxA(hMainWnd, sqlite3_errmsg(db), "Error", MB_OK);
					free(value16);
				}
				
				sqlite3_finalize(stmt);
			}
			
			SetFocus(hGridWnd);
			RemoveProp(hWnd, TEXT("ISCHANGED"));
		}
		break;

		case WM_KILLFOCUS: {
			DestroyWindow(hWnd);
		}
		break;

		case WM_KEYDOWN: {
			SetProp(hWnd, TEXT("ISCHANGED"), (HANDLE)1);
			
			if (wParam == VK_RETURN) {
				DestroyWindow(hWnd);
				return TRUE;
			}

			if (wParam == VK_ESCAPE) {
				SetProp(hWnd, TEXT("ISCHANGED"), 0);
				DestroyWindow(hWnd);
				return TRUE;
			}

			if (wParam == 0x41 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + A
				SendMessage(hWnd, EM_SETSEL, 0, -1);
				return TRUE;
			}
		}
		break;

		case WM_PASTE:
		case WM_CUT: {
			SetProp(hWnd, TEXT("ISCHANGED"), (HANDLE)1);
		}
		break;
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

// lParam = USERDATA = hMainWnd
INT_PTR CALLBACK cbDlgAddRow(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_INITDIALOG: {
			SetWindowLongPtr(hWnd, GWLP_USERDATA, lParam);
			SetWindowLongPtr(hWnd, GWL_STYLE, WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU);
			SetWindowLongPtr(hWnd, GWL_EXSTYLE, WS_EX_DLGMODALFRAME);
			SetWindowText(hWnd, TEXT("Add row"));
			
			HWND hGridWnd = GetDlgItem((HWND)lParam, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			int w = 120;
			int h = 20;
			for (int colNo = 0; colNo < colCount; colNo++) {
				TCHAR colName16[256];
				Header_GetItemText(hHeader, colNo, colName16, 255);
				HWND hLabel = CreateWindow(WC_STATIC, colName16, WS_VISIBLE | WS_CHILD | SS_WORDELLIPSIS | SS_RIGHT, 
					5, 10 + colNo * h, w - 5, h, hWnd, (HMENU)IntToPtr(IDC_DLG_LABEL + colNo), GetModuleHandle(0), 0);
				HWND hEdit = CreateWindow(WC_EDIT, NULL, WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_BORDER | WS_CLIPSIBLINGS | ES_AUTOHSCROLL, 
					w + 5, 10 + colNo * h - 2, 200, h - 2, hWnd, (HMENU)IntToPtr(IDC_DLG_EDIT + colNo), GetModuleHandle(0), 0);
			}	
			
			RECT rc;
			GetWindowRect(hGridWnd, &rc);
			SetWindowPos(hWnd, 0, rc.left + 50, rc.top + 50, 340, colCount * h + 70, SWP_NOZORDER);	
		
			HWND hBtn = CreateWindow(WC_BUTTON, TEXT("OK"), WS_VISIBLE | WS_CHILD | WS_TABSTOP, 245, colCount * h + 15, 80, 22, hWnd, (HMENU)IDC_DLG_OK, GetModuleHandle(0), 0);
			SendMessage(hWnd, DM_SETDEFID, IDC_DLG_OK, 0);

			EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumSetFont, (LPARAM)(HFONT)GetStockObject(DEFAULT_GUI_FONT));
			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;
		
		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDC_DLG_OK) {
				HWND hMainWnd = (HWND)GetWindowLongPtr(hWnd, GWLP_USERDATA);
				sqlite3* db = (sqlite3*)GetProp(hMainWnd, TEXT("DB"));
				char* tablename8 = (char*)GetProp(hMainWnd, TEXT("TABLENAME8"));
				
				HWND hGridWnd = GetDlgItem(hMainWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int colCount = Header_GetItemCount(hHeader);
				
				int len = 128;
				for (int colNo = 0; colNo < colCount; colNo++) 				
					len += GetWindowTextLength(GetDlgItem(hWnd, IDC_DLG_LABEL + colNo)) + GetWindowTextLength(GetDlgItem(hWnd, IDC_DLG_EDIT + colNo)) + 10;
					
				char* query8 = (char*)calloc(len, sizeof(char));
				sprintf(query8, "insert into \"%s\" (\"", tablename8);
				for (int colNo = 0; colNo < colCount; colNo++) {
					TCHAR colName16[255];
					Header_GetItemText(hHeader, colNo, colName16, 255);
					char* colName8 = utf16to8(colName16);
					strcat(query8, colName8);
					strcat(query8, colNo != colCount - 1 ? "\", \"" : "\") values (");
					free(colName8);
				}
				
				for (int colNo = 0; colNo < colCount; colNo++) {
					int len = GetWindowTextLength(GetDlgItem(hWnd, IDC_DLG_EDIT + colNo)) + 1;
					TCHAR value16[len + 1];
					GetWindowText(GetDlgItem(hWnd, IDC_DLG_EDIT + colNo), value16, len);
					
					if (_tcslen(value16)) {
						char* value8 = utf16to8(value16);
						char* qvalue8 = (char*)calloc(2 * strlen(value8) + 1, sizeof(char));
						int j = 0;
						for (int i = 0; i < strlen(value8); i++) {
							if (value8[i] == '\'') {
								qvalue8[i + j] = value8[i];
								j++;
							}
							qvalue8[i + j] = value8[i];
						}	
						strcat(query8, "\'");
						strcat(query8, qvalue8);
						strcat(query8, "\'");
						free(value8);
						free(qvalue8);
					} else {
						strcat(query8, "null");
					}
					
					strcat(query8, colNo != colCount - 1 ? ", " : ")");
				}

				if (SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0)) {
					free(query8);
					EndDialog(hWnd, IDOK);
				} else {
					free(query8);	
					MessageBoxA(hWnd, sqlite3_errmsg(db), "Error", MB_OK);
				}
			}
			
			if (cmd == IDCANCEL) {
				EndDialog (hWnd, IDCANCEL);
			}
		}
		break;
		
		case WM_CLOSE: {
			EndDialog(hWnd, IDCANCEL);
		}
		break;
	}
	
	return FALSE;
}

// lParam = USERDATA = hMainWnd
INT_PTR CALLBACK cbDlgAddTable(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_INITDIALOG: {
			SetWindowLongPtr(hWnd, GWLP_USERDATA, lParam);
			SetWindowLongPtr(hWnd, GWL_STYLE, WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU);
			SetWindowLongPtr(hWnd, GWL_EXSTYLE, WS_EX_DLGMODALFRAME);
			SetWindowText(hWnd, TEXT("New table..."));
			
			
			CreateWindow(WC_STATIC, TEXT("Table name"), WS_VISIBLE | WS_CHILD, 
				10, 5, 300, 16, hWnd, (HMENU)IntToPtr(IDC_DLG_LABEL), GetModuleHandle(0), 0);

			CreateWindow(WC_EDIT, NULL, WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_BORDER | WS_CLIPSIBLINGS | ES_AUTOHSCROLL, 
				10, 22, 300, 20, hWnd, (HMENU)IntToPtr(IDC_DLG_EDIT), GetModuleHandle(0), 0);

			CreateWindow(WC_STATIC, TEXT("Column names separated by commas"), WS_VISIBLE | WS_CHILD, 
				10, 55, 300, 16, hWnd, (HMENU)IntToPtr(IDC_DLG_LABEL), GetModuleHandle(0), 0);

			CreateWindow(WC_EDIT, NULL, WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_BORDER | WS_CLIPSIBLINGS | ES_AUTOHSCROLL, 
				10, 72, 300, 50, hWnd, (HMENU)IntToPtr(IDC_DLG_EDIT + 1), GetModuleHandle(0), 0);

			CreateWindow(WC_BUTTON, TEXT("OK"), WS_VISIBLE | WS_CHILD | WS_TABSTOP, 240, 130, 70, 22, hWnd, (HMENU)IDC_DLG_OK, GetModuleHandle(0), 0);

			RECT rc;
			GetWindowRect((HWND)lParam, &rc);
			SetWindowPos(hWnd, 0, rc.left + 50, rc.top + 50, 325, 185, SWP_NOZORDER);	
		
			EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumSetFont, (LPARAM)(HFONT)GetStockObject(DEFAULT_GUI_FONT));	
			InvalidateRect(hWnd, NULL, TRUE);
			
		}
		break;
		
		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDC_DLG_OK) {
				HWND hMainWnd = (HWND)GetWindowLongPtr(hWnd, GWLP_USERDATA);

				TCHAR tablename16[MAX_TABLE_LENGTH];
				GetDlgItemText(hWnd, IDC_DLG_EDIT, tablename16, MAX_TABLE_LENGTH);
				if (_tcslen(tablename16) == 0)
					return MessageBox(hWnd, TEXT("Table name is empty"), NULL, MB_OK);

				TCHAR columns16[MAX_TEXT_LENGTH];
				GetDlgItemText(hWnd, IDC_DLG_EDIT + 1, columns16, MAX_TEXT_LENGTH);
				if (_tcslen(columns16) == 0)
					return MessageBox(hWnd, TEXT("You should provide at least one column name"), NULL, MB_OK);

				int len = _tcslen(tablename16) + _tcslen(columns16) + 256;
				TCHAR query16[len + 1];
				_sntprintf(query16, len, TEXT("create table \"%ls\" (%ls)"), tablename16, columns16);

				sqlite3* db = (sqlite3*)GetProp(hMainWnd, TEXT("DB"));
				char* query8 = utf16to8(query16);
				BOOL rc = SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0);
				free(query8);

				if (!rc)
					return MessageBoxA(hWnd, sqlite3_errmsg(db), "Error", MB_OK);

				HWND hListWnd = GetDlgItem(hMainWnd, IDC_TABLELIST);
				int pos = ListBox_AddString(hListWnd, tablename16);
				SendMessage(hListWnd, LB_SETITEMDATA, pos, 1);

				EndDialog(hWnd, IDOK);
			}
			
			if (cmd == IDCANCEL) {
				EndDialog (hWnd, IDCANCEL);
			}
		}
		break;
		
		case WM_CLOSE: {
			EndDialog(hWnd, IDCANCEL);
		}
		break;
	}
	
	return FALSE;
}

BOOL CALLBACK cbEnumSetFont (HWND hWnd, LPARAM hFont) {
	SendMessage(hWnd, WM_SETFONT, (WPARAM)hFont, 0);
	return TRUE;
}

void bindValue(sqlite3_stmt* stmt, int i, const char* value8) {
	int len = strlen(value8);
	BOOL isNum = TRUE;
	int dotCount = 0;
	for (int i = +(value8[0] == '-' /* is negative */); i < len; i++) {
		isNum = isNum && (isdigit(value8[i]) || value8[i] == '.' || value8[i] == ',');
		dotCount += value8[i] == '.' || value8[i] == ',';
	}

	if (isNum && dotCount == 0 && len < 10) // 2147483647
		sqlite3_bind_int(stmt, i, atoi(value8));
	else if (isNum && dotCount == 0 && len >= 10)
			sqlite3_bind_int64(stmt, i, atoll(value8));
	else if (isNum && dotCount == 1) {
		double d;
		char *endptr;
		errno = 0;
		d = strtold(value8, &endptr);
		sqlite3_bind_double(stmt, i, d);
	} else {
		sqlite3_bind_text(stmt, i, value8, len, SQLITE_TRANSIENT);
	}
}

void bindFilterValues(HWND hHeader, sqlite3_stmt* stmt) {
	int colCount = Header_GetItemCount(hHeader);
	int bindNo = 1;
	for (int colNo = 0; colNo < colCount; colNo++) {
		HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
		int size = GetWindowTextLength(hEdit);
		if (size > 0) {
			TCHAR value16[size + 1];
			GetWindowText(hEdit, value16, size + 1);
			char* value8 = utf16to8(value16[0] == TEXT('=') || value16[0] == TEXT('/') || value16[0] == TEXT('!') || value16[0] == TEXT('<') || value16[0] == TEXT('>') ? value16 + 1 : value16);
			bindValue(stmt, bindNo, value8);
			free(value8);
			bindNo++;
		}
	}
}

HWND getMainWindow(HWND hWnd) {
	HWND hMainWnd = hWnd;
	while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
		hMainWnd = GetParent(hMainWnd);
	return hMainWnd;	
}

void setStoredValue(TCHAR* name, int value) {
	TCHAR buf[128];
	_sntprintf(buf, 128, TEXT("%i"), value);
	WritePrivateProfileString(APP_NAME, name, buf, iniPath);
}

int getStoredValue(TCHAR* name, int defValue) {
	TCHAR buf[128];
	return GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) ? _ttoi(buf) : defValue;
}

TCHAR* getStoredString(TCHAR* name, TCHAR* defValue) { 
	TCHAR* buf = calloc(256, sizeof(TCHAR));
	if (0 == GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) && defValue)
		_tcsncpy(buf, defValue, 255);
	return buf;	
}

int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam) {
	if (GetWindowLong(hWnd, GWL_STYLE) & WS_TABSTOP && IsWindowVisible(hWnd)) {
		int no = 0;
		HWND* wnds = (HWND*)lParam;
		while (wnds[no])
			no++;
		wnds[no] = hWnd;
	}

	return TRUE;
}

TCHAR* utf8to16(const char* in) {
	TCHAR *out;
	if (!in || strlen(in) == 0) {
		out = (TCHAR*)calloc (1, sizeof (TCHAR));
	} else  {
		DWORD size = MultiByteToWideChar(CP_UTF8, 0, in, -1, NULL, 0);
		out = (TCHAR*)calloc (size, sizeof (TCHAR));
		MultiByteToWideChar(CP_UTF8, 0, in, -1, out, size);
	}
	return out;
}

char* utf16to8(const TCHAR* in) {
	char* out;
	if (!in || _tcslen(in) == 0) {
		out = (char*)calloc (1, sizeof(char));
	} else  {
		int len = WideCharToMultiByte(CP_UTF8, 0, in, -1, NULL, 0, 0, 0);
		out = (char*)calloc (len, sizeof(char));
		WideCharToMultiByte(CP_UTF8, 0, in, -1, out, len, 0, 0);
	}
	return out;
}

int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords) {
	if (!text || !word)
		return -1;
		
	int res = -1;
	int tlen = _tcslen(text);
	int wlen = _tcslen(word);	
	if (!tlen || !wlen)
		return res;
	
	if (!isMatchCase) {
		TCHAR* ltext = calloc(tlen + 1, sizeof(TCHAR));
		_tcsncpy(ltext, text, tlen);
		text = _tcslwr(ltext);

		TCHAR* lword = calloc(wlen + 1, sizeof(TCHAR));
		_tcsncpy(lword, word, wlen);
		word = _tcslwr(lword);
	}

	if (isWholeWords) {
		for (int pos = 0; (res  == -1) && (pos <= tlen - wlen); pos++) 
			res = (pos == 0 || pos > 0 && !_istalnum(text[pos - 1])) && 
				!_istalnum(text[pos + wlen]) && 
				_tcsncmp(text + pos, word, wlen) == 0 ? pos : -1;
	} else {
		TCHAR* s = _tcsstr(text, word);
		res = s != NULL ? s - text : -1;
	}
	
	if (!isMatchCase) {
		free(text);
		free(word);
	}

	return res; 
}

TCHAR* extractUrl(TCHAR* data) {
	int len = data ? _tcslen(data) : 0;
	int start = len;
	int end = len;
	
	TCHAR* url = calloc(len + 10, sizeof(TCHAR));
	
	TCHAR* slashes = _tcsstr(data, TEXT("://"));
	if (slashes) {
		start = len - _tcslen(slashes);
		end = start + 3;
		for (; start > 0 && _istalpha(data[start - 1]); start--);
		for (; end < len && data[end] != TEXT(' ') && data[end] != TEXT('"') && data[end] != TEXT('\''); end++);
		_tcsncpy(url, data + start, end - start);
		
	} else if (_tcschr(data, TEXT('.'))) {
		_sntprintf(url, len + 10, TEXT("https://%ls"), data);
	}
	
	return url;
}	

void setClipboardText(const TCHAR* text) {
	int len = (_tcslen(text) + 1) * sizeof(TCHAR);
	HGLOBAL hMem =  GlobalAlloc(GMEM_MOVEABLE, len);
	memcpy(GlobalLock(hMem), text, len);
	GlobalUnlock(hMem);
	OpenClipboard(0);
	EmptyClipboard();
	SetClipboardData(CF_UNICODETEXT, hMem);
	CloseClipboard();
}

int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt) {
	int colNo = Header_GetItemCount(ListView_GetHeader(hListWnd));
	LVCOLUMN lvc = {0};
	lvc.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM | LVCF_FMT;
	lvc.iSubItem = colNo;
	lvc.pszText = colName;
	lvc.cchTextMax = _tcslen(colName) + 1;
	lvc.cx = 100;
	lvc.fmt = fmt;
	return ListView_InsertColumn(hListWnd, colNo, &lvc);
}

int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax) {
	if (i < 0)
		return FALSE;

	TCHAR buf[cchTextMax];

	HDITEM hdi = {0};
	hdi.mask = HDI_TEXT;
	hdi.pszText = buf;
	hdi.cchTextMax = cchTextMax;
	int rc = Header_GetItem(hWnd, i, &hdi);

	_tcsncpy(pszText, buf, cchTextMax);
	return rc;
}

void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState) {
	MENUITEMINFO mii = {0};
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STATE;
	mii.fState = fState;
	SetMenuItemInfo(hMenu, wID, FALSE, &mii);
}

BOOL saveFile(TCHAR* path, const TCHAR* filter, const TCHAR* defExt, HWND hWnd) {
	OPENFILENAME ofn = {0};
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = hWnd;
	ofn.lpstrFile = path;
	ofn.lpstrFile[_tcslen(path) + 1] = '\0';
	ofn.nMaxFile = MAX_PATH;
	ofn.lpstrFilter = filter;
	ofn.nFilterIndex = 1;
	ofn.lpstrFileTitle = NULL;
	ofn.nMaxFileTitle = 0;
	ofn.lpstrInitialDir = NULL;
	ofn.Flags = OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR;

	if (!GetSaveFileName(&ofn))
		return FALSE;

	if (_tcslen(path) == 0)
		return saveFile(path, filter, defExt, hWnd);

	TCHAR ext[32] = {0};
	TCHAR path2[MAX_PATH] = {0};

	_tsplitpath(path, NULL, NULL, NULL, ext);
	if (_tcslen(ext) > 0)
		_tcscpy(path2, path);
	else
		_sntprintf(path2, MAX_PATH, TEXT("%ls%ls%ls"), path, _tcslen(defExt) > 0 ? TEXT(".") : TEXT(""), defExt);

	BOOL isFileExists = _taccess(path2, 0) == 0;
	if (isFileExists && IDYES != MessageBox(hWnd, TEXT("Overwrite file?"), TEXT("Confirmation"), MB_YESNO))
		return saveFile(path, filter, defExt, hWnd);

	if (isFileExists)
		DeleteFile(path2);

	_tcscpy(path, path2);

	return TRUE;
}