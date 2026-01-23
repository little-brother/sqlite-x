#define UNICODE
#define _UNICODE
#define _WIN32_WINNT 0x0600

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <richedit.h>
#include <uxtheme.h>
#include <locale.h>
#include <tchar.h>
#include <shlwapi.h>
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#include <math.h>
#include <stdarg.h>
#include "resource.h"
#include "include/sqlite3.h"

#define LVS_EX_AUTOSIZECOLUMNS 0x10000000
#define EM_SETTARGETDEVICE	   (WM_USER + 72)

#define SPLITTER_WIDTH         5
#define MAX_TEXT_LENGTH        32000
#define MAX_COLUMN_LENGTH      2000
#define MAX_TABLENAME_LENGTH   2000
#define BLOB_VALUE             "(BLOB)"
#define MAX_RECENT_FILES       10
#define CELL_BUFFER_LENGTH     512
#define KEEP_PREV_VALUE        -2

#define MIN(a,b) (((a)<(b))?(a):(b))
#define MAX(a,b) (((a)>(b))?(a):(b))

LRESULT CALLBACK cbMainWnd(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewDataGrid(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewEditor(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
BOOL cbNewAutoCompleteEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewCellEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewAutoComplete(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewColumns(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewAddRowEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK cbDlgAddTable(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK cbDlgAddRow(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

void bindValue(sqlite3_stmt* stmt, int i, const char* value8);
char* buildFilterWhere(HWND hHeader);
void bindFilterValues(HWND hHeader, sqlite3_stmt* stmt);
void setStoredValue(char* name, int value);
int getStoredValue(char* name);
TCHAR* getStoredString(TCHAR* name, TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords);
TCHAR* extractUrl(TCHAR* data);
void setClipboardText(const TCHAR* text);
TCHAR* getClipboardText();
int getFontHeight(HWND hWnd);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt);
void ListView_EnsureVisibleCell(HWND hWnd, int rowNo, int colNo);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);
void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState);
BOOL saveFile(TCHAR* path, const TCHAR* filter, const TCHAR* defExt, HWND hWnd);
TCHAR* loadString(UINT id, ...);
TCHAR* loadFilters(UINT id);

static TCHAR iniPath[MAX_PATH] = {0};
sqlite3 *db = 0;
HWND hAutoComplete = 0;

const char* storedKeys[] = {
	"position-x", "position-y", "width", "height", "show-maximized", "tablelist-width", "editor-width", "show-row-maximized",

	"font-size", "filter-row", "dark-theme", "disable-grid-lines", "max-column-width", "show-editor", "editor-word-wrap",
	"delete-journal", "filter-align", "read-only", "open-last-db", "enable-foreign-keys",
	"max-autocomplete-rowid", "max-autocomplete-results", "hide-service-tables", "dpi-aware", "cache-size",
	"inline-char-limit", "editor-char-limit", "hide-editor-on-save", "exit-by-escape", "auto-filters",
	"enable-load-extensions",

	"text-color", "back-color", "back-color2", "filter-text-color", "filter-back-color",
	"selection-text-color", "selection-back-color", "current-cell-back-color", "splitter-color",

	"text-color-dark", "back-color-dark", "back-color2-dark", "filter-text-color-dark", "filter-back-color-dark",
	"selection-text-color-dark", "selection-back-color-dark", "current-cell-back-color-dark", "splitter-color-dark",
	0
};

int storedValues[] = {
	200, 100, 1200, 800, 0, 200, 350, 0,

	11, 0, 0, 0, 300, 0, 1,
	0, 0, 0, 1, 0,
	10000, 1000, 0, 0, 2000,
	64, 64000, 0, 1, 1,
	1,

	0, 16777215, 15790320, 0, 15790320,
	16777215, 6956042, 10903622, 0,

	14474460, 2105376, 3421236, 16777215, 3947580,
	14474460, 6710856, 4079136, 0,
	0
};

int WINAPI WinMain (HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	WNDCLASS wc = {0};
	wc.lpfnWndProc = cbMainWnd;
	wc.hInstance = hInstance;
	wc.hCursor = LoadCursor (NULL, IDC_ARROW);
	wc.lpszClassName = WND_CLASS;
	wc.hbrBackground = (HBRUSH) GetSysColorBrush(COLOR_APPWORKSPACE);
	wc.style = CS_DBLCLKS;
	wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON));
	RegisterClass(&wc);

	TCHAR* title16 = loadString(IDS_NO_DATABASE);
	HWND hWnd = CreateWindowEx(0, WND_CLASS, title16, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, (HMENU)0, hInstance, NULL);
	free(title16);
	if (hWnd == NULL)
		return 0;

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&icex);

	TCHAR buf16[MAX_PATH];
	GetModuleFileName(0, buf16, MAX_PATH);
	PathRemoveFileSpec(buf16);
	_sntprintf(iniPath, MAX_PATH, TEXT("%ls/%ls.ini"), buf16, APP_NAME);

	setStoredValue("splitter-color", GetSysColor(COLOR_BTNFACE));
	setStoredValue("splitter-color-dark", GetSysColor(COLOR_BTNFACE));
	if (_taccess(iniPath, 0) == 0) {
		TCHAR buf16[128];
		for (int i = 0; storedKeys[i] != 0; i++) {
			TCHAR* key16 = utf8to16(storedKeys[i]);
			if (GetPrivateProfileString(APP_NAME, key16, NULL, buf16, 127, iniPath))
				storedValues[i] = _ttoi(buf16);
			free(key16);
		}
	}

	setlocale(LC_CTYPE, "");
	if (getStoredValue("dpi-aware"))
		SetProcessDPIAware();

	HMENU hMainMenu = LoadMenu(hInstance, MAKEINTRESOURCE(IDC_MENU_MAIN));
	ModifyMenu(hMainMenu, 2, MF_BYPOSITION | MFT_RIGHTJUSTIFY, 0, TEXT("?"));
	SetMenu(hWnd, hMainMenu);

	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, NULL, hWnd, IDC_STATUSBAR);
	HDC hDC = GetDC(hWnd);
	float z = GetDeviceCaps(hDC, LOGPIXELSX) / 96.0; // 96 = 100%, 120 = 125%, 144 = 150%
	ReleaseDC(hWnd, hDC);
	int sizes[7] = {35 * z, 110 * z, 180 * z, 225 * z, 420 * z, 500 * z, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 7, (LPARAM)&sizes);

	TCHAR buf[255];
	_sntprintf(buf, 32, TEXT(" %ls"), APP_VERSION);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)buf);

	HWND hBtnWnd = CreateWindow(WC_BUTTON, TEXT("Add table"), WS_CHILD,
		0, 0, 100, 100, hWnd, (HMENU)IDC_ADD_TABLE, GetModuleHandle(0), NULL);
	SendMessage(hBtnWnd, WM_SETFONT, (LPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);

	CreateWindow(WC_LISTBOX, NULL, WS_CHILD | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_HSCROLL | LBS_WANTKEYBOARDINPUT | WS_TABSTOP,
		0, 20, 100, 100, hWnd, (HMENU)IDC_TABLELIST, hInstance, NULL);

	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hWnd, (HMENU)IDC_GRID, hInstance, NULL);

	int noLines = getStoredValue("disable-grid-lines");
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | (noLines ? 0 : LVS_EX_GRIDLINES) | LVS_EX_LABELTIP | LVS_EX_HEADERDRAGDROP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbNewDataGrid));
	SetProp(hGridWnd, TEXT("TABLENAME8"), (HANDLE)calloc(MAX_TABLENAME_LENGTH + 1, sizeof(char)));
	SetProp(hGridWnd, TEXT("CELLBUFFER16"), (HANDLE)calloc(CELL_BUFFER_LENGTH + 1, sizeof(TCHAR)));

	LoadLibrary(TEXT("msftedit.dll"));
	HWND hEditorWnd = CreateWindow(TEXT("RICHEDIT50W"), NULL, WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN | WS_VSCROLL | ES_NOHIDESEL | WS_HSCROLL | WS_TABSTOP,
		0, 0, 100, 100, hWnd, (HMENU)IDC_EDITOR, hInstance, NULL);
	SetProp(hEditorWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEditorWnd, GWLP_WNDPROC, (LONG_PTR)cbNewEditor));
	SendMessage(hEditorWnd, EM_SETEVENTMASK, 0, ENM_CHANGE);
	SendMessage(hEditorWnd, EM_SETTARGETDEVICE,(WPARAM)NULL, (LPARAM)!getStoredValue("editor-word-wrap"));

	HWND hCellEdit = CreateWindowEx(0, WC_EDIT, NULL, WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, hGridWnd, (HMENU)IDC_CELL_EDIT, GetModuleHandle(NULL), NULL);
	SetProp(hCellEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hCellEdit, GWLP_WNDPROC, (LONG_PTR)&cbNewCellEdit));

	hAutoComplete = CreateWindowEx(WS_EX_TOPMOST, WC_LISTBOX, NULL, WS_POPUP | WS_BORDER | WS_VSCROLL, 0, 0, 170, 200, hWnd, (HMENU)0, hInstance, NULL);
	SetProp(hAutoComplete, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hAutoComplete, GWLP_WNDPROC, (LONG_PTR)&cbNewAutoComplete));
	SetProp(hAutoComplete, TEXT("CURRENTINPUT"), 0);

	HWND hHeader = ListView_GetHeader(hGridWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);
	SetWindowTheme(hHeader, TEXT(" "), TEXT(" "));
	SetProp(hHeader, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hHeader, GWLP_WNDPROC, (LONG_PTR)cbNewHeader));

	HMENU hListMenu = GetSubMenu(LoadMenu(hInstance, MAKEINTRESOURCE(IDC_MENU_LIST)), 0);
	SetProp(hWnd, TEXT("LISTMENU"), hListMenu);

	HMENU hGridMenu = GetSubMenu(LoadMenu(hInstance, MAKEINTRESOURCE(IDC_MENU_GRID)), 0);
	if (getStoredValue("read-only")) {
		DeleteMenu(hGridMenu, IDM_PASTE_CELL, MF_BYCOMMAND);
		DeleteMenu(hGridMenu, IDM_SEPARATOR1, MF_BYCOMMAND);
		DeleteMenu(hGridMenu, IDM_ADD_ROW, MF_BYCOMMAND);
		DeleteMenu(hGridMenu, IDM_DUPLICATE_ROWS, MF_BYCOMMAND);
		DeleteMenu(hGridMenu, IDM_DELETE_ROWS, MF_BYCOMMAND);
	}
	SetProp(hWnd, TEXT("GRIDMENU"), hGridMenu);

	HMENU hBlobMenu = GetSubMenu(LoadMenu(hInstance, MAKEINTRESOURCE(IDC_MENU_BLOB)), 0);
	SetProp(hWnd, TEXT("BLOBMENU"), hBlobMenu);

	HMENU hEditorWndMenu = GetSubMenu(LoadMenu(hInstance, MAKEINTRESOURCE(IDC_MENU_EDITOR)), 0);
	SetProp(hWnd, TEXT("EDITORMENU"), hEditorWndMenu);

	SendMessage(hWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hWnd, WMU_SET_THEME, 0, 0);

	int nArgs = 0;
	TCHAR** args = CommandLineToArgvW(GetCommandLine(), &nArgs);

	if (nArgs > 1) {
		TCHAR path16[MAX_PATH + 1] = {0};
		_tcsncpy(path16, args[1], MAX_PATH);
		SendMessage(hWnd, WMU_OPEN_DB, (WPARAM)path16, 0);
	}

	SetWindowPos(hWnd, 0,
		getStoredValue("position-x"),
		getStoredValue("position-y"),
		getStoredValue("width"),
		getStoredValue("height"),
		SWP_NOZORDER);

	ShowWindow(hWnd, getStoredValue("show-maximized") ? SW_MAXIMIZE : nCmdShow);
	SendMessage(hWnd, WMU_UPDATE_RECENT, 0, 0);
	SendMessage(hWnd, WMU_UPDATE_SETTINGS, 0, 0);

	if (nArgs < 2 && getStoredValue("open-last-db"))
		SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(IDM_RECENT, 0), 0);

	HACCEL hAccel = LoadAccelerators(0, MAKEINTRESOURCE(IDA_ACCEL));

	MSG msg = {};
	while (GetMessage(&msg, NULL, 0, 0) > 0) {
		if (TranslateAccelerator(hWnd, hAccel, &msg))
			continue;

		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return 0;
}

LRESULT CALLBACK cbMainWnd(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
	switch (uMsg) {
		case WM_ACTIVATE: {
			if (LOWORD(wParam) == WA_INACTIVE)
				ShowWindow(hAutoComplete, SW_HIDE);
		}
		break;

		case WM_CLOSE: {
			if (IsWindowVisible(hWnd)) {
				WINDOWPLACEMENT wp = {};
				wp.length = sizeof(WINDOWPLACEMENT);
				GetWindowPlacement(hWnd, &wp);
				if (wp.showCmd == SW_SHOWMAXIMIZED) {
					setStoredValue("show-maximized", 1);
				} else {
					RECT rc;
					GetWindowRect(hWnd, &rc);
					setStoredValue("position-x", rc.left);
					setStoredValue("position-y", rc.top);
					setStoredValue("width", rc.right - rc.left);
					setStoredValue("height", rc.bottom - rc.top);
					setStoredValue("show-maximized", 0);
				}
			}

			SendMessage(hWnd, WMU_CLOSE_DB, 0, 0);

			DeleteFont((HFONT)GetProp(hWnd, TEXT("FONT")));
			DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));
			DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
			DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));

			ShowWindow(hWnd, SW_HIDE);

			TCHAR buf16[128];
			for (int i = 0; storedKeys[i] != 0; i++) {
				TCHAR* key16 = utf8to16(storedKeys[i]);
				_sntprintf(buf16, 128, TEXT("%i"), storedValues[i]);
				WritePrivateProfileString(APP_NAME, key16, buf16, iniPath);
				free(key16);
			}

			PostQuitMessage(0);
			return 0;
		}
		break;

		case WM_SIZE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int statusH = rc.bottom;

			int tablelistWidth = getStoredValue("tablelist-width");
			int editorWidth = getStoredValue("editor-width");

			GetClientRect(hWnd, &rc);
			HWND hBtnWnd = GetDlgItem(hWnd, IDC_ADD_TABLE);
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hEditorWnd = GetDlgItem(hWnd, IDC_EDITOR);

			int h = 0;
			if (getStoredValue("read-only") == 0) {
				HWND hHeader = ListView_GetHeader(hGridWnd);
				BOOL isFilterRow = getStoredValue("filter-row");
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
				if (isFilterRow || !IsWindowVisible(hGridWnd))
					h = round(h/2);

				if (!hasColumns) {
					SetWindowLongPtr(hHeader, GWL_STYLE, styles);
					ListView_DeleteColumn(hGridWnd, 0);
				}
			}

			BOOL isShowEditor = getStoredValue("show-editor");
			int w = isShowEditor ? editorWidth + SPLITTER_WIDTH : 0;

			SetWindowPos(hBtnWnd, 0, 0, 0, tablelistWidth, h, SWP_NOMOVE | SWP_NOZORDER);
			SetWindowPos(hListWnd, 0, 0, h, tablelistWidth, rc.bottom - statusH - h, SWP_NOZORDER);
			SetWindowPos(hGridWnd, 0, tablelistWidth + SPLITTER_WIDTH, 0, rc.right - tablelistWidth - SPLITTER_WIDTH - w, rc.bottom - statusH, SWP_NOZORDER);
			SetWindowPos(hEditorWnd, 0, rc.right - w + SPLITTER_WIDTH, 0, w - SPLITTER_WIDTH, rc.bottom - statusH, SWP_NOZORDER | (db && isShowEditor ? SWP_SHOWWINDOW : SWP_HIDEWINDOW));

			rc = (RECT){0, 0, w - SPLITTER_WIDTH, rc.bottom - statusH};
			InflateRect(&rc, -10, -10);
			SendMessage(hEditorWnd, EM_SETRECT, 0, (LPARAM)&rc);
		}
		break;

		case WM_PAINT: {
			if (!db)
				break;

			PAINTSTRUCT ps = {0};
			HDC hDC = BeginPaint(hWnd, &ps);

			RECT rc;
			GetClientRect(hWnd, &rc);
			int w = rc.right;
			HBRUSH hBrush = (HBRUSH)GetProp(hWnd, TEXT("SPLITTERBRUSH"));

			int tablelistWidth = getStoredValue("tablelist-width");
			rc.left = tablelistWidth;
			rc.right = rc.left + SPLITTER_WIDTH;
			FillRect(hDC, &rc, hBrush);

			if (getStoredValue("show-editor")) {
				int editorWidth = getStoredValue("editor-width");
				rc.left = w - editorWidth - SPLITTER_WIDTH;
				rc.right = w - editorWidth;
				FillRect(hDC, &rc, hBrush);
			}

			EndPaint(hWnd, &ps);
			return 0;
		}
		break;

		case WM_SETCURSOR: {
			if (db && GetProp(hWnd, TEXT("ISMOUSEHOVER"))) {
				SetCursor(LoadCursor(0, IDC_SIZEWE));
				return TRUE;
			}
		}
		break;

		case WM_SETFOCUS: {
			SetFocus(GetProp(hWnd, TEXT("LASTFOCUS")));
		}
		break;

		case WM_LBUTTONDOWN: {
			int x = GET_X_LPARAM(lParam);

			int w = getStoredValue("tablelist-width");
			if (x >= w && x <= w + SPLITTER_WIDTH) {
				SetProp(hWnd, TEXT("ISMOUSEDOWN"), IntToPtr(1));
				SetProp(hWnd, TEXT("ISMOVEEDITOR"), 0);
				SetCapture(hWnd);
			}

			RECT rc;
			GetClientRect(hWnd, &rc);
			w = getStoredValue("editor-width");
			if (x >= rc.right - w - SPLITTER_WIDTH && x <= rc.right - w) {
				SetProp(hWnd, TEXT("ISMOUSEDOWN"), IntToPtr(1));
				SetProp(hWnd, TEXT("ISMOVEEDITOR"), IntToPtr(1));
				SetCapture(hWnd);
			}

			return 0;
		}
		break;

		case WM_LBUTTONUP: {
			ReleaseCapture();
			RemoveProp(hWnd, TEXT("ISMOUSEDOWN"));
			RemoveProp(hWnd, TEXT("ISMOUSEHOVER"));
			RemoveProp(hWnd, TEXT("ISMOVEEDITOR"));
		}
		break;

		case WM_MOUSEMOVE: {
			DWORD x = GET_X_LPARAM(lParam);

			RECT rc = {0};
			GetClientRect(hWnd, &rc);

			if (!GetProp(hWnd, TEXT("ISMOUSEHOVER"))) {
				int tw = getStoredValue("tablelist-width");
				int ew = getStoredValue("editor-width");

				if ((tw <= x && x <= tw + SPLITTER_WIDTH) || (rc.right - ew - SPLITTER_WIDTH <= x && x <= rc.right - ew)) {
					TRACKMOUSEEVENT tme = {sizeof(TRACKMOUSEEVENT), TME_LEAVE, hWnd, 0};
					TrackMouseEvent(&tme);
					SetProp(hWnd, TEXT("ISMOUSEHOVER"), IntToPtr(1));
				}
			}

			if (GetProp(hWnd, TEXT("ISMOUSEDOWN")) && (x > 0) && (x < 32000)) {
				BOOL isEditor = !!GetProp(hWnd, TEXT("ISMOVEEDITOR"));
				if (isEditor) {
					setStoredValue("editor-width", rc.right - x);
				} else {
					setStoredValue("tablelist-width", x);
				}

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

				if (pos != -1 && currPos != -1 && getStoredValue("read-only") == 0) {
					ClientToScreen(hListWnd, &p);
					HMENU hMenu = GetProp(hWnd, TEXT("LISTMENU"));
					TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
				}
			}

			if ((HWND)wParam == GetDlgItem(hWnd, IDC_EDITOR)) {
				HWND hEditorWnd = (HWND)wParam;

				POINT p;
				GetCursorPos(&p);

				HMENU hMenu = (HMENU)GetProp(hWnd, TEXT("EDITORMENU"));
				Menu_SetItemState(hMenu, IDM_EDITOR_WORDWRAP, getStoredValue("editor-word-wrap") ? MF_CHECKED : 0);
				Menu_SetItemState(hMenu, IDM_EDITOR_SAVE, GetProp(hEditorWnd, TEXT("ISCHANGED")) ? MF_ENABLED : MF_DISABLED | MF_GRAYED );
				TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hEditorWnd, NULL);
			}
		}
		break;

		case WM_CTLCOLORLISTBOX: {
			SetBkColor((HDC)wParam, (COLORREF)PtrToInt(GetProp(hWnd, TEXT("BACKCOLOR"))));
			SetTextColor((HDC)wParam, (COLORREF)PtrToInt(GetProp(hWnd, TEXT("TEXTCOLOR"))));
			return (LRESULT)(HBRUSH)GetProp(hWnd, TEXT("BACKBRUSH"));
		}
		break;

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);

			if (cmd == IDM_OPEN_DB) {
				TCHAR path16[MAX_PATH];
				TCHAR* filters16 = loadFilters(IDR_SQLITE_FILTERS);

				OPENFILENAME ofn = {0};
				ofn.lStructSize = sizeof(ofn);
				ofn.hwndOwner = hWnd;
				ofn.lpstrFile = path16;
				ofn.lpstrFile[0] = '\0';
				ofn.nMaxFile = MAX_PATH;
				ofn.lpstrFilter = filters16;
				ofn.nFilterIndex = 1;
				ofn.lpstrFileTitle = NULL;
				ofn.nMaxFileTitle = 0;
				ofn.lpstrInitialDir = NULL;
				ofn.Flags = OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;

				int rc = GetOpenFileName(&ofn);
				free(filters16);
				if (!rc)
					return 0;

				TCHAR ext16[32] = {0};
				_tsplitpath(path16, NULL, NULL, NULL, ext16);
				if (_tcslen(ext16) == 0)
					_tcscat(path16, TEXT(".sqlite"));

				SendMessage(hWnd, WMU_OPEN_DB, (WPARAM)path16, 0);
			}

			if (cmd >= IDM_RECENT && cmd < IDM_RECENT + MAX_RECENT_FILES) {
				TCHAR path16[MAX_PATH + 1];
				if (GetMenuString(GetSubMenu(GetMenu(hWnd), 0), cmd, path16, MAX_PATH, MF_BYCOMMAND)) {
					SendMessage(hWnd, WMU_OPEN_DB, (WPARAM)path16, 0);
				}
			}

			if (cmd == IDM_EXIT) {
				WORD isAccel = HIWORD(wParam);

				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hCellEdit = GetDlgItem(hGridWnd, IDC_CELL_EDIT);
				if (IsWindowVisible(hCellEdit))
					return SendMessage(hCellEdit, WM_KEYDOWN, VK_ESCAPE, 0);

				if (IsWindowVisible(hAutoComplete)) {
					HWND hEdit = GetProp(hAutoComplete, TEXT("CURRENTINPUT"));
					SendMessage(hEdit, WM_KEYDOWN, VK_ESCAPE, 0);
					return TRUE;
				}

				if (!isAccel || (isAccel && getStoredValue("exit-by-escape")))
					SendMessage(hWnd, WM_CLOSE, 0, 0);
			}

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


				TCHAR* about16 = loadString(IDS_ABOUT,
					APP_VERSION,
					APP_PLATFORM == 32 ? TEXT("x86") : TEXT("x86-64"),
					sqlite3_libversion(),
					TEXT(__DATE__),
					vi.dwMajorVersion, vi.dwMinorVersion, vi.dwBuildNumber, vi.szCSDVersion,
					isWin64 ? TEXT("x64") : TEXT("x32"));
				MessageBox(hWnd, about16, APP_NAME, MB_OK);
				free(about16);
			}

			if (cmd == IDM_HOMEPAGE) {
				ShellExecute(0, 0, TEXT("https://github.com/little-brother/sqlite-x"), 0, 0 , SW_SHOW);
			}

			if (cmd == IDM_WIKI) {
				ShellExecute(0, 0, TEXT("https://github.com/little-brother/sqlite-x/wiki"), 0, 0 , SW_SHOW);
			}

			if (cmd == IDM_ADD_TABLE || cmd == IDC_ADD_TABLE) {
				HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
				if (IDOK == DialogBoxParam ((HINSTANCE)GetModuleHandle(0), MAKEINTRESOURCE(IDD_TABLE), hWnd, &cbDlgAddTable, (LPARAM)hWnd)) {
					ListBox_SetCurSel(hListWnd, ListBox_GetCount(hListWnd) - 1);
					SendMessage(hWnd, WMU_SB_SET_COUNTS, 0, 0);
					SendMessage(hWnd, WM_COMMAND, MAKELPARAM(IDC_TABLELIST, LBN_SELCHANGE), (LPARAM)hListWnd);
				}
				SetFocus(hListWnd);
			}

			if (cmd == IDM_DROP_TABLE) {
				HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				int pos = ListBox_GetCurSel(hListWnd);
				BOOL isTable = SendMessage(hListWnd, LB_GETITEMDATA, pos, 0);
				TCHAR tablename16[MAX_TABLENAME_LENGTH + 1];
				ListBox_GetText(hListWnd, pos, tablename16);
				char* tablename8 = utf16to8(tablename16);

				TCHAR* msg16 = loadString(IDS_DROP_TABLE, tablename16);
				TCHAR* caption16 = loadString(IDS_CONFIRMATION);
				if(IDYES == MessageBox(hWnd, msg16, caption16, MB_YESNO)) {
					char query8[MAX_TABLENAME_LENGTH + 64];
					snprintf(query8, MAX_TABLENAME_LENGTH + 64, "drop %s \"%s\"", isTable ? "table" : "view", tablename8);
					if (SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0)) {
						ListBox_DeleteString(hListWnd, ListBox_GetCurSel(hListWnd));
						if (LB_ERR != ListBox_SetCurSel(hListWnd, 0)) {
							SendMessage(hWnd, WM_COMMAND, MAKELPARAM(IDC_TABLELIST, LBN_SELCHANGE), (LPARAM)hListWnd);
						} else {
							ShowWindow(hGridWnd, SW_HIDE);
						}
					} else {
						MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
					}
					free(tablename8);
				}

				free(msg16);
				free(caption16);
			}

			if (cmd == IDM_DARK_THEME) {
				BOOL isDarkTheme = !getStoredValue("dark-theme");
				setStoredValue("dark-theme", isDarkTheme);

				SendMessage(hWnd, WMU_SET_THEME, 0, 0);
				SendMessage(hWnd, WMU_UPDATE_SETTINGS, 0, 0);
			}

			if (cmd == IDM_FILTER_ROW) {
				BOOL isFilterRow = !getStoredValue("filter-row");
				setStoredValue("filter-row", isFilterRow);

				SendMessage(GetDlgItem(hWnd, IDC_GRID), WMU_SET_HEADER_FILTERS, 0, 0);
				SendMessage(hWnd, WMU_UPDATE_SETTINGS, 0, 0);
			}

			if (cmd == IDM_SEARCH) {
				BOOL isFilterRow = getStoredValue("filter-row");
				if (!isFilterRow)
					SendMessage(hWnd, WM_COMMAND, IDM_FILTER_ROW, 0);

				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				SetFocus(GetDlgItem(hHeader, IDC_HEADER_EDIT));
			}

			if (cmd == IDM_GRID_LINES) {
				BOOL noLines = !getStoredValue("disable-grid-lines");
				setStoredValue("disable-grid-lines", noLines);

				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | (noLines ? 0 : LVS_EX_GRIDLINES) | LVS_EX_LABELTIP | LVS_EX_HEADERDRAGDROP);
				InvalidateRect(hGridWnd, NULL, TRUE);
				SendMessage(hWnd, WMU_UPDATE_SETTINGS, 0, 0);
			}

			if (cmd == IDM_EDITOR) {
				BOOL isShowEditor = !getStoredValue("show-editor");
				setStoredValue("show-editor", isShowEditor);

				SendMessage(hWnd, WM_SIZE, 0, 0);
				SendMessage(hWnd, WMU_UPDATE_SETTINGS, 0, 0);

				if (!db)
					return TRUE;

				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				if (isShowEditor) {
					PostMessage(hGridWnd, WMU_EDIT_CELL, 3, 0);
				} else {
					SetFocus(hGridWnd);
				}
			}

			if (cmd == IDM_AUTOFILTERS) {
				BOOL isAutoFilters = !getStoredValue("auto-filters");
				setStoredValue("auto-filters", isAutoFilters);
				SendMessage(hWnd, WMU_UPDATE_SETTINGS, 0, 0);
			}

			if (cmd == IDM_EDITOR_AUTOHIDE) {
				BOOL isAutoHide = !getStoredValue("hide-editor-on-save");
				setStoredValue("hide-editor-on-save", isAutoHide);
				SendMessage(hWnd, WMU_UPDATE_SETTINGS, 0, 0);
			}

			if (cmd == IDM_ESCAPE_EXIT) {
				BOOL isExitByEscape = !getStoredValue("exit-by-escape");
				setStoredValue("exit-by-escape", isExitByEscape);

				SendMessage(hWnd, WMU_UPDATE_SETTINGS, 0, 0);
			}

			if (cmd == IDM_PREV_CONTROL || cmd == IDM_NEXT_CONTROL) {
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hCellEdit = GetDlgItem(hGridWnd, IDC_CELL_EDIT);
				if (IsWindowVisible(hCellEdit))
					return SendMessage(hCellEdit, WM_KEYDOWN, VK_TAB, 0);

				if (IsWindowVisible(hAutoComplete)) {
					HWND hEdit = GetProp(hAutoComplete, TEXT("CURRENTINPUT"));
					SendMessage(hEdit, WM_KEYDOWN, VK_TAB, 0);
					return TRUE;
				}

				BOOL isBack = cmd == IDM_PREV_CONTROL;
				HWND hFocus = GetFocus();
				HWND hNextWnd = GetNextDlgTabItem(hWnd, hFocus, isBack);

				int idc = GetDlgCtrlID(hFocus);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int colCount = Header_GetItemCount(hHeader);
				if(idc >= IDC_HEADER_EDIT && idc < IDC_HEADER_EDIT + colCount) {
					idc = IDC_HEADER_EDIT + (idc - IDC_HEADER_EDIT + (isBack ? -1 : 1) + colCount) % colCount;
					hNextWnd = GetDlgItem(hHeader, idc);
				}

				if (hNextWnd == hWnd)
					hNextWnd = GetDlgItem(hHeader, IDC_TABLELIST);
				SetFocus(hNextWnd);
			}

			if (cmd == IDC_EDITOR && HIWORD(wParam) == EN_CHANGE) {
				SetProp((HWND)lParam, TEXT("ISCHANGED"), IntToPtr(1));
			}

			if (cmd == IDC_TABLELIST && HIWORD(wParam) == LBN_SELCHANGE) {
				HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				int pos = ListBox_GetCurSel(hListWnd);
				if (pos == -1)
					return TRUE;

				TCHAR tablename16[MAX_TABLENAME_LENGTH + 1];
				ListBox_GetText(hListWnd, pos, tablename16);
				BOOL isTable = SendMessage(hListWnd, LB_GETITEMDATA, pos, 0);

				SendMessage(hGridWnd, WMU_SET_SOURCE, (WPARAM)tablename16, (LPARAM)isTable);

				HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
				TCHAR* type16 = loadString(isTable ? IDS_SB_TABLE : IDS_SB_VIEW);
				SendMessage(hStatusWnd, SB_SETTEXT, SB_TYPE, (LPARAM)type16);
				free(type16);

				SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, 0);

				HWND hEditorWnd = GetDlgItem(hWnd, IDC_EDITOR);
				SetWindowText(hEditorWnd, 0);
				SendMessage(hEditorWnd, EM_SETREADONLY, !isTable, 0);
				InvalidateRect(hEditorWnd, NULL, TRUE);

				RemoveProp(hEditorWnd, TEXT("SID"));
			}
		}
		break;

		case WM_NOTIFY: {
			NMHDR* pHdr = (LPNMHDR)lParam;
			if (pHdr->idFrom == IDC_GRID) {
				HWND hGridWnd = pHdr->hwndFrom;

				if (pHdr->code == LVN_GETDISPINFO) {
					LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
					LV_ITEM* pItem = &(pDispInfo)->item;
					if (pItem->mask & LVIF_TEXT) {
						TCHAR* value16 = (TCHAR*)SendMessage(hGridWnd, WMU_GET_CELL, pItem->iItem, pItem->iSubItem);

						pItem->pszText = _tcslen(value16) > CELL_BUFFER_LENGTH ? _tcsncpy(GetProp(hGridWnd, TEXT("CELLBUFFER16")), value16, CELL_BUFFER_LENGTH) : value16;
					}
				}

				if (pHdr->code == LVN_COLUMNCLICK) {
					NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
					return SendMessage(hGridWnd, HIWORD(GetKeyState(VK_CONTROL)) ? WMU_HIDE_COLUMN : WMU_SORT_COLUMN, lv->iSubItem, 0);
				}

				if (pHdr->code == (DWORD)NM_CLICK || pHdr->code == (DWORD)NM_RCLICK) {
					NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
					SendMessage(hGridWnd, WMU_SET_CURRENT_CELL, ia->iItem, ia->iSubItem);
				}

				if (pHdr->code == (DWORD)NM_DBLCLK) {
					SendMessage(hGridWnd, WMU_EDIT_CELL, 1, 0);
				}

				if (pHdr->code == (DWORD)NM_CLICK && HIWORD(GetKeyState(VK_MENU))) {
					NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
					TCHAR* value16 = (TCHAR*)SendMessage(hGridWnd, WMU_GET_CELL, ia->iItem, ia->iSubItem);
					TCHAR* url16 = extractUrl(value16);
					ShellExecute(0, TEXT("open"), url16, 0, 0 , SW_SHOW);
					free(url16);
				}

				if (pHdr->code == (DWORD)NM_RCLICK) {
					TCHAR* value16 = (TCHAR*)SendMessage(hGridWnd, WMU_GET_CELL, -1, -1);
					BOOL isBlob = value16 && _tcscmp(value16, TEXT(BLOB_VALUE)) == 0;

					POINT p;
					GetCursorPos(&p);
					TrackPopupMenu(GetProp(hWnd, isBlob ? TEXT("BLOBMENU") : TEXT("GRIDMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hGridWnd, NULL);
				}

				if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
					NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
					if (lv->uOldState != lv->uNewState && (lv->uNewState & LVIS_SELECTED)) {
						SendMessage(hGridWnd, WMU_SET_CURRENT_CELL, lv->iItem, (LPARAM)KEEP_PREV_VALUE);
					}
				}

				if (pHdr->code == (UINT)NM_CUSTOMDRAW) {
					int result = CDRF_DODEFAULT;

					NMLVCUSTOMDRAW* pCustomDraw = (LPNMLVCUSTOMDRAW)lParam;
					if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT)
						result = CDRF_NOTIFYITEMDRAW;

					if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
						if (ListView_GetItemState(pHdr->hwndFrom, pCustomDraw->nmcd.dwItemSpec, LVIS_SELECTED)) {
							pCustomDraw->nmcd.uItemState &= ~CDIS_SELECTED;
							result = CDRF_NOTIFYSUBITEMDRAW;
						} else {
							pCustomDraw->clrTextBk = (COLORREF)PtrToInt(GetProp(hWnd, pCustomDraw->nmcd.dwItemSpec % 2 == 0 ? TEXT("BACKCOLOR") : TEXT("BACKCOLOR2")));
						}
					}

					if (pCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM)) {
						int rowNo = PtrToInt(GetProp(pHdr->hwndFrom, TEXT("CURRENTROWNO")));
						int colNo = PtrToInt(GetProp(pHdr->hwndFrom, TEXT("CURRENTCOLNO")));
						BOOL isCurrCell = (pCustomDraw->nmcd.dwItemSpec == (DWORD)rowNo) && (pCustomDraw->iSubItem == colNo);
						pCustomDraw->clrText = (COLORREF)PtrToInt(GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")));
						pCustomDraw->clrTextBk = (COLORREF)PtrToInt(GetProp(hWnd, isCurrCell ? TEXT("CURRENTCELLCOLOR") : TEXT("SELECTIONBACKCOLOR")));
					}

					return result;
				}
			}

			if ((pHdr->code == HDN_ITEMCHANGED || pHdr->code == HDN_ENDDRAG) && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(hWnd, IDC_GRID)))
				PostMessage(GetDlgItem(hWnd, IDC_GRID), WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		case WM_VKEYTOITEM: {
			WORD key = LOWORD(wParam);
			if ((HWND)lParam == GetDlgItem(hWnd, IDC_TABLELIST)) {
				if (key == VK_DELETE) {
					SendMessage(hWnd, WM_COMMAND, IDM_DROP_TABLE, 0);
				}

				if (key == VK_INSERT) {
					SendMessage(hWnd, WM_COMMAND, IDM_ADD_TABLE, 0);
				}
			}
		}
		break;

		case WMU_UPDATE_RECENT: {
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
		break;

		case WMU_UPDATE_SETTINGS: {
			HMENU hMenu = GetMenu(hWnd);
			Menu_SetItemState(hMenu, IDM_DARK_THEME, getStoredValue("dark-theme") ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_FILTER_ROW, getStoredValue("filter-row") ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_GRID_LINES, !getStoredValue("disable-grid-lines") ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_EDITOR, getStoredValue("show-editor") ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_AUTOFILTERS, getStoredValue("auto-filters") ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_EDITOR_AUTOHIDE, getStoredValue("hide-editor-on-save") ? MFS_CHECKED : 0);
			Menu_SetItemState(hMenu, IDM_ESCAPE_EXIT, getStoredValue("exit-by-escape") ? MFS_CHECKED : 0);
		}
		break;

		case WMU_OPEN_DB: {
			SendMessage(hWnd, WMU_CLOSE_DB, 0, 0);

			TCHAR* dbPath16 = (TCHAR*)wParam;
			char* dbPath8 = utf16to8(dbPath16);

			BOOL isReadOnly = getStoredValue("read-only");
			if (SQLITE_OK != sqlite3_open_v2(dbPath8, &db, (isReadOnly ? SQLITE_OPEN_READONLY : (SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE)) | SQLITE_OPEN_URI, NULL) ||
				SQLITE_OK != sqlite3_exec(db, "select 1 from sqlite_master", 0, 0, 0)) {
				TCHAR* err16 = loadString(IDS_ERR_DB_OPEN, dbPath16);
				MessageBox(hWnd, err16, 0, 0);
				free(err16);

				sqlite3_close(db);
				db = 0;
			}
			free(dbPath8);

			if (!db)
				return FALSE;

			SetWindowText(hWnd, dbPath16);

			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			ListBox_ResetContent(hListWnd);
			ShowWindow(GetDlgItem(hWnd, IDC_ADD_TABLE), SW_SHOW);
			ShowWindow(hListWnd, SW_SHOW);

			sqlite3_stmt *stmt;
			sqlite3_prepare_v2(db, getStoredValue("hide-service-tables") ?
				"select name, type = 'table' from sqlite_master where type in ('table', 'view') and instr(name, '__') != 1 and instr(name, '$') != 1 order by type, name" :
				"select name, type = 'table' from sqlite_master where type in ('table', 'view') order by type, name",
				-1, &stmt, 0);
			while(SQLITE_ROW == sqlite3_step(stmt)) {
				TCHAR* name16 = utf8to16((char*)sqlite3_column_text(stmt, 0));
				int pos = ListBox_AddString(hListWnd, name16);
				free(name16);

				BOOL isTable = sqlite3_column_int(stmt, 1);
				SendMessage(hListWnd, LB_SETITEMDATA, pos, isTable);
			}
			sqlite3_finalize(stmt);

			if (getStoredValue("enable-foreign-keys"))
				sqlite3_exec(db, "pragma foreign_keys = on", 0, 0, 0);

			TCHAR* paths16 = calloc(MAX_RECENT_FILES * MAX_PATH, sizeof(TCHAR));
			_tcscat(paths16, dbPath16);
			_tcscat(paths16, TEXT("?"));

			TCHAR* buf16 = calloc(MAX_RECENT_FILES * MAX_PATH, sizeof(TCHAR));
			if (GetPrivateProfileString(APP_NAME, TEXT("recent"), NULL, buf16, MAX_RECENT_FILES * MAX_PATH, iniPath)) {
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
			SendMessage(hWnd, WMU_UPDATE_RECENT, 0, 0);
			SendMessage(hWnd, WMU_SB_SET_COUNTS, 0, 0);

			if (getStoredValue("enable-load-extensions"))
				sqlite3_enable_load_extension(db, TRUE);

			TCHAR* query16 = getStoredString(TEXT("execute-on-connect"), TEXT(""));
			char* query8 = utf16to8(query16);
			if (SQLITE_OK != sqlite3_exec(db, query8, 0, 0, 0))
				MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
			free(query16);
			free(query8);

			SetFocus(hWnd);
			ListBox_SetCurSel(hListWnd, 0);
			SendMessage(hWnd, WM_COMMAND, MAKELPARAM(IDC_TABLELIST, LBN_SELCHANGE), (LPARAM)hListWnd);
			InvalidateRect(hWnd, NULL, TRUE);

			return TRUE;
		}
		break;

		case WMU_CLOSE_DB: {
			if (!db)
				return TRUE;

			TCHAR* dbPath16 = utf8to16(sqlite3_db_filename(db, 0));

			sqlite3_close(db);
			db = 0;

			int len = _tcslen(dbPath16) + 10;
			TCHAR* journalPath16 = calloc(len + 1, sizeof(TCHAR));
			_sntprintf(journalPath16, len, TEXT("%ls-wal"), dbPath16);
			BOOL hasJournal = _taccess(journalPath16, 0) == 0;

			int deleteJournalMode = getStoredValue("delete-journal");
			if (!hasJournal && deleteJournalMode) {
				TCHAR* msg16 = loadString(IDS_DELETE_JOURNAL);
				TCHAR* caption16 = loadString(IDS_CONFIRMATION);

				if (_taccess(journalPath16, 0) == 0 && (deleteJournalMode == 2 || (deleteJournalMode == 1 && IDYES == MessageBox(hWnd, msg16, caption16, MB_YESNO | MB_ICONQUESTION)))) {
					DeleteFile(journalPath16);
					_sntprintf(journalPath16, len, TEXT("%ls-shm"), dbPath16);
					DeleteFile(journalPath16);
				}
				free(journalPath16);
				free(msg16);
				free(caption16);
			}
			free(dbPath16);
			free(journalPath16);

			ShowWindow(GetDlgItem(hWnd, IDC_ADD_TABLE), SW_HIDE);
			ShowWindow(GetDlgItem(hWnd, IDC_TABLELIST), SW_HIDE);
			ShowWindow(GetDlgItem(hWnd, IDC_GRID), SW_HIDE);
			ShowWindow(GetDlgItem(hWnd, IDC_EDITOR), SW_HIDE);

			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_TYPE, 0);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_TABLE_COUNT, 0);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_VIEW_COUNT, 0);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, 0);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, 0);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, 0);

			TCHAR* title16 = loadString(IDS_NO_DATABASE);
			SetWindowText(hWnd, title16);
			free(title16);

			InvalidateRect(hWnd, NULL, TRUE);
			UpdateWindow(hWnd);

			return TRUE;
		}
		break;

		case WMU_SB_SET_COUNTS: {
			int tCount = 0, vCount = 0;
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			for (int pos = 0; pos < ListBox_GetCount(hListWnd); pos++) {
				BOOL isTable = SendMessage(hListWnd, LB_GETITEMDATA, pos, 0);
				tCount += isTable;
				vCount += !isTable;
			}

			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);

			TCHAR* tableCount16 = loadString(IDS_SB_TABLE_COUNT, tCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_TABLE_COUNT, (LPARAM)tableCount16);
			free(tableCount16);

			TCHAR* viewCount16 = loadString(IDS_SB_VIEW_COUNT, vCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_VIEW_COUNT, (LPARAM)viewCount16);
			free(viewCount16);
		}
		break;

		// wParam = total count, lParam = count
		case WMU_SB_SET_ROW_COUNTS: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			int rowCount = (int)wParam;
			int totalRowCount = (int)lParam;

			TCHAR* rows16 = loadString(IDS_SB_ROWS, rowCount, totalRowCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)rows16);
			free(rows16);
		}
		break;

		// wParam = rowNo, lParam = colNo
		case WMU_SB_SET_CURRENT_CELL: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			int rowNo = (int)wParam;
			int colNo = (int)lParam;

			TCHAR buf[32] = {0};
			if (colNo != - 1 && rowNo != -1)
				_sntprintf(buf, 32, TEXT(" %i:%i"), rowNo + 1, colNo + 1);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, (LPARAM)buf);
		}
		break;

		// wParam - size delta
		case WMU_SET_FONT: {
			int fontSize = getStoredValue("font-size") + wParam;
			fontSize = fontSize < 8 ? 8 : fontSize > 36 ? 36 : fontSize;
			setStoredValue("font-size", fontSize);

			HFONT hFont = (HFONT)GetProp(hWnd, TEXT("FONT"));
			if (hFont)
				DeleteFont(hFont);

			TCHAR* fontFamily16 = getStoredString(TEXT("font"), TEXT("Arial"));
			HDC hDC = GetDC(hWnd);
			int fontHeight = -MulDiv(fontSize, GetDeviceCaps(hDC, LOGPIXELSY), 72);
			ReleaseDC(hWnd, hDC);
			hFont = CreateFont (fontHeight, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, fontFamily16);
			free(fontFamily16);

			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hEditorWnd = GetDlgItem(hWnd, IDC_EDITOR);
			HWND hCellEdit = GetDlgItem(hGridWnd, IDC_CELL_EDIT);

			SendMessage(hWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hListWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hGridWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hEditorWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hCellEdit, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hAutoComplete, WM_SETFONT, (LPARAM)hFont, TRUE);

			HWND hHeader = ListView_GetHeader(hGridWnd);
			for (int colNo = 0; colNo < Header_GetItemCount(hHeader); colNo++)
				SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), WM_SETFONT, (LPARAM)hFont, TRUE);

			if (TRUE) {
				int w = 0;
				HDC hDC = GetDC(hListWnd);
				HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);
				TCHAR buf[MAX_TABLENAME_LENGTH];
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
			}

			SetProp(hWnd, TEXT("FONT"), hFont);

			PostMessage(hGridWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
			PostMessage(hWnd, WM_SIZE, 0, 0);
		}
		break;

		case WMU_SET_THEME: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hEditorWnd = GetDlgItem(hWnd, IDC_EDITOR);

			BOOL isDark = getStoredValue("dark-theme");

			int textColor = !isDark ? getStoredValue("text-color") : getStoredValue("text-color-dark");
			int backColor = !isDark ? getStoredValue("back-color") : getStoredValue("back-color-dark");
			int backColor2 = !isDark ? getStoredValue("back-color2") : getStoredValue("back-color2-dark");
			int filterTextColor = !isDark ? getStoredValue("filter-text-color") : getStoredValue("filter-text-color-dark");
			int filterBackColor = !isDark ? getStoredValue("filter-back-color") : getStoredValue("filter-back-color-dark");
			int currCellColor = !isDark ? getStoredValue("current-cell-back-color") : getStoredValue("current-cell-back-color-dark");
			int selectionTextColor = !isDark ? getStoredValue("selection-text-color") : getStoredValue("selection-text-color-dark");
			int selectionBackColor = !isDark ? getStoredValue("selection-back-color") : getStoredValue("selection-back-color-dark");
			int splitterColor = !isDark ? getStoredValue("splitter-color") : getStoredValue("splitter-color-dark");

			SetProp(hWnd, TEXT("TEXTCOLOR"), IntToPtr(textColor));
			SetProp(hWnd, TEXT("BACKCOLOR"), IntToPtr(backColor));
			SetProp(hWnd, TEXT("BACKCOLOR2"), IntToPtr(backColor2));
			SetProp(hWnd, TEXT("FILTERTEXTCOLOR"), IntToPtr(filterTextColor));
			SetProp(hWnd, TEXT("FILTERBACKCOLOR"), IntToPtr(filterBackColor));
			SetProp(hWnd, TEXT("CURRENTCELLCOLOR"), IntToPtr(currCellColor));
			SetProp(hWnd, TEXT("SELECTIONTEXTCOLOR"), IntToPtr(selectionTextColor));
			SetProp(hWnd, TEXT("SELECTIONBACKCOLOR"), IntToPtr(selectionBackColor));
			SetProp(hWnd, TEXT("SPLITTERCOLOR"), IntToPtr(splitterColor));

			DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));
			DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
			DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));
			SetProp(hWnd, TEXT("BACKBRUSH"), CreateSolidBrush(backColor));
			SetProp(hWnd, TEXT("FILTERBACKBRUSH"), CreateSolidBrush(filterBackColor));
			SetProp(hWnd, TEXT("SPLITTERBRUSH"), CreateSolidBrush(splitterColor));

			ListView_SetTextColor(hGridWnd, textColor);
			ListView_SetBkColor(hGridWnd, backColor);
			ListView_SetTextBkColor(hGridWnd, backColor);

			SendMessage(hEditorWnd, EM_SETBKGNDCOLOR, 0, (LPARAM)backColor);
			CHARFORMAT2 cf = {0};
			cf.cbSize = sizeof(cf);
			cf.dwMask = CFM_COLOR;
			cf.crTextColor = textColor;
			cf.dwEffects = 0;
			SendMessage(hEditorWnd, EM_SETCHARFORMAT, (WPARAM)SCF_ALL, (LPARAM)&cf);


			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;
	}
	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK cbNewDataGrid(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch(msg){
		case WM_DESTROY: {
			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);

			free((char*)GetProp(hWnd, TEXT("TABLENAME8")));

			RemoveProp(hWnd, TEXT("TABLENAME8"));
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (HIWORD(wParam) == EN_CHANGE && cmd == IDC_CELL_EDIT) {
				SendMessage((HWND)lParam, WMU_AUTOCOMPLETE, 0, 0);
			}

			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROWS || cmd == IDM_COPY_COLUMN || cmd == IDM_SAVE_AS) {
				HWND hHeader = ListView_GetHeader(hWnd);
				int rowNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTROWNO")));
				int colNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTCOLNO")));

				int colCount = Header_GetItemCount(hHeader);
				int rowCount = PtrToInt(GetProp(hWnd, TEXT("ROWCOUNT")));
				int selCount = ListView_GetSelectedCount(hWnd);

				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int cacheSize = getStoredValue("cache-size");
				int cacheOffset = PtrToInt(GetProp(hWnd, TEXT("CACHEOFFSET")));

				if (rowNo == -1 ||
					rowNo >= rowCount ||
					colCount == 0 ||
					(cmd == IDM_COPY_CELL && colNo == -1) ||
					(cmd == IDM_COPY_CELL && colNo >= colCount) ||
					(cmd == IDM_COPY_COLUMN && colNo == -1) ||
					(cmd == IDM_COPY_COLUMN && colNo >= colCount) ||
					(cmd == IDM_COPY_ROWS && selCount == 0)) {
					setClipboardText(TEXT(""));
					return 0;
				}

				int minRowNo = ListView_GetNextItem(hWnd, -1, LVNI_SELECTED);
				int maxRowNo = minRowNo;
				do {
					int rowNo = ListView_GetNextItem(hWnd, maxRowNo, LVNI_SELECTED);
					if (rowNo == -1)
						break;
					maxRowNo = rowNo;
				} while (maxRowNo <= cacheOffset + cacheSize);

				if (rowNo < cacheOffset ||
					rowNo >= cacheOffset + cacheSize ||
					(cmd == IDM_COPY_COLUMN && selCount == 1 && (0 < cacheOffset || rowCount > cacheSize)) ||
					(cmd == IDM_COPY_COLUMN && selCount != 1 && (minRowNo < cacheOffset || maxRowNo >= cacheOffset + cacheSize)) ||
					((cmd == IDM_COPY_ROWS || cmd == IDM_SAVE_AS) && selCount > 1 && (minRowNo < cacheOffset || maxRowNo >= cacheOffset + cacheSize))
				) {
					TCHAR* err16 = loadString(IDS_ERR_CACHE_LIMIT, rowCount, cacheSize);
					MessageBox(hWnd, err16, 0, 0);
					free(err16);
					return 0;
				}

				TCHAR* delimiter = getStoredString(TEXT("column-delimiter"), TEXT("\t"));
				UINT searchNext = cmd == IDM_SAVE_AS && ListView_GetSelectedCount(hWnd) < 2 ? LVNI_ALL : LVNI_SELECTED;

				int len = 0;
				if (cmd == IDM_COPY_CELL)
					len = _tcslen(cache[rowNo - cacheOffset][colNo]);

				if (cmd == IDM_COPY_ROWS || cmd == IDM_SAVE_AS) {
					int rowNo = ListView_GetNextItem(hWnd, -1, searchNext);
					while (rowNo != -1) {
						for (int colNo = 0; colNo < colCount; colNo++) {
							if (ListView_GetColumnWidth(hWnd, colNo))
								len += _tcslen(cache[rowNo - cacheOffset][colNo]) + 1; /* column delimiter */
						}

						len++; /* \n */
						rowNo = ListView_GetNextItem(hWnd, rowNo, searchNext);
					}
				}

				if (cmd == IDM_COPY_COLUMN) {
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hWnd, -1, searchNext);
					while (rowNo != -1 && rowNo < rowCount) {
						len += _tcslen(cache[rowNo - cacheOffset][colNo]) + 1 /* \n */;
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hWnd, rowNo, searchNext);
					}
				}

				TCHAR* buf = calloc(len + 1, sizeof(TCHAR));
				if (cmd == IDM_COPY_CELL)
					_tcscat(buf, cache[rowNo - cacheOffset][colNo]);

				if (cmd == IDM_COPY_ROWS || cmd == IDM_SAVE_AS) {
					int pos = 0;
					int rowNo = ListView_GetNextItem(hWnd, -1, searchNext);

					int* colOrder = calloc(colCount, sizeof(int));
					Header_GetOrderArray(hHeader, colCount, colOrder);

					while (rowNo != -1) {
						for (int idx = 0; idx < colCount; idx++) {
							int colNo = colOrder[idx];
							if (ListView_GetColumnWidth(hWnd, colNo)) {
								int len = _tcslen(cache[rowNo - cacheOffset][colNo]);
								_tcsncpy(buf + pos, cache[rowNo - cacheOffset][colNo], len);
								buf[pos + len] = delimiter[0];
								pos += len + 1;
							}
						}

						buf[pos - (pos > 0)] = TEXT('\n');
						rowNo = ListView_GetNextItem(hWnd, rowNo, searchNext);
					}
					buf[pos - 1] = 0; // remove last \n

					free(colOrder);
				}

				if (cmd == IDM_COPY_COLUMN) {
					int pos = 0;
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hWnd, -1, searchNext);
					while (rowNo != -1 && rowNo < rowCount) {
						int len = _tcslen(cache[rowNo - cacheOffset][colNo]);
						_tcsncpy(buf + pos, cache[rowNo - cacheOffset][colNo], len);
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hWnd, rowNo, searchNext);
						if (rowNo != -1 && rowNo < rowCount)
							buf[pos + len] = TEXT('\n');
						pos += len + 1;
					}
				}

				if (cmd != IDM_SAVE_AS) {
					setClipboardText(buf);
				} else {
					TCHAR path16[MAX_PATH] = {0};
					TCHAR filter16[] = TEXT("CSV files (*.csv, *.tsv)\0*.csv;*.tsv\0All\0*.*\0");
					if (saveFile(path16, filter16, TEXT(""), hWnd)) {
						FILE* fp = _tfopen (path16, TEXT("wb"));
						if (fp) {
							char* delimiter8 = utf16to8(delimiter);
							TCHAR colName16[MAX_COLUMN_LENGTH];

							for(int colNo = 0; colNo < colCount; colNo++) {
								Header_GetItemText(hHeader, colNo, colName16, MAX_COLUMN_LENGTH);
								char* colName8 = utf16to8(colName16);
								fwrite(colName8, strlen(colName8), 1, fp);
								free(colName8);

								fwrite(colNo < colCount - 1 ? delimiter8 : "\n", 1, 1, fp);
							}
							free(delimiter8);

							char* buf8 = utf16to8(buf);
							fwrite(buf8, strlen(buf8), 1, fp);
							fclose(fp);
							free(buf8);
						} else {
							TCHAR* err16 = loadString(IDS_ERR_FILE_OPEN, path16);
							MessageBox(hWnd, err16, 0, 0);
							free(err16);
						}
					}
				}
				free(buf);
				free(delimiter);
			}

			if (cmd == IDM_BLOB) {
				HWND hHeader = ListView_GetHeader(hWnd);

				if (!GetProp(hWnd, TEXT("HASROWID"))) {
					TCHAR* err16 = loadString(IDS_ERR_NO_ROWID);
					MessageBox(hWnd, err16, 0, 0);
					free(err16);
					return 0;
				}

				int rowNo = ListView_GetNextItem(hWnd, -1, LVNI_SELECTED);
				if (rowNo != -1) {
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
					int cacheSize = getStoredValue("cache-size");
					int cacheOffset = PtrToInt(GetProp(hWnd, TEXT("CACHEOFFSET")));

					if (rowNo < cacheOffset || rowNo >= cacheOffset + cacheSize)
						return 0;

					int colCount = Header_GetItemCount(hHeader);
					sqlite3_stmt *stmt;

					char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));
					int colNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTCOLNO")));
					TCHAR colName16[MAX_COLUMN_LENGTH + 1];
					Header_GetItemText(hHeader, colNo, colName16, MAX_COLUMN_LENGTH + 1);
					char* colName8 = utf16to8(colName16);
					char query8[1024 + strlen(tablename8)];
					sprintf(query8, "select %s from \"%s\" where rowid = %i", colName8, tablename8, _ttoi(cache[rowNo - cacheOffset][colCount]));
					free(colName8);

					TCHAR path16[MAX_PATH] = {0};
					TCHAR* filters16 = loadFilters(IDR_BLOB_FILTERS);
					if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0) &&
						SQLITE_ROW == sqlite3_step(stmt) &&
						saveFile(path16, filters16, TEXT(""), hWnd)) {
						FILE *fp = _tfopen (path16, TEXT("wb"));
						if (fp) {
							fwrite(sqlite3_column_blob(stmt, 0), sqlite3_column_bytes(stmt, 0), 1, fp);
							fclose(fp);
						} else {
							TCHAR* err16 = loadString(IDS_ERR_FILE_WRITE, path16);
							MessageBox(hWnd, err16, 0, 0);
							free(err16);
						}
					}
					sqlite3_finalize(stmt);
					free(filters16);
				}
			}

			if (cmd == IDM_PASTE_CELL && getStoredValue("read-only") == 0) {
				SendMessage(hWnd, WMU_UPDATE_CELL, 0, 0);
				SetFocus(hWnd);
			}

			if (cmd == IDM_ADD_ROW && getStoredValue("read-only") == 0) {
				if (IDOK == DialogBoxParam ((HINSTANCE)GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_ROW), hWnd, &cbDlgAddRow, (LPARAM)hWnd)) {
					SendMessage(hWnd, WMU_UPDATE_DATA, 1, 0);
				}
				SetFocus(hWnd);
			}

			if ((cmd == IDM_DUPLICATE_ROWS || cmd == IDM_DELETE_ROWS) && getStoredValue("read-only") == 0) {
				SendMessage(hWnd, cmd == IDM_DUPLICATE_ROWS ? WMU_DUPLICATE_ROWS : WMU_DELETE_ROWS, 0, 0);
				SendMessage(hWnd, WMU_UPDATE_DATA, 0, 0);
				SetFocus(hWnd);
			}

			if (cmd == IDM_HIDE_COLUMN) {
				int colNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTCOLNO")));
				SendMessage(hWnd, WMU_HIDE_COLUMN, colNo, 0);
			}
		}
		break;

		case WM_KEYDOWN: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
			BOOL isShift = HIWORD(GetKeyState(VK_SHIFT));
			int key = wParam;

			if (key == 0x43) { // C
				UINT cmd = isCtrl && !isShift ? IDM_COPY_CELL :
					isCtrl && isShift ? IDM_COPY_COLUMN :
					!isCtrl && isShift ? IDM_COPY_ROWS :
					0;

				SendMessage(hWnd, WM_COMMAND, cmd, 0);
				return 0;
			}

			if (key == 0x56 && isCtrl) { // Ctrl + V
				SendMessage(hWnd, WM_COMMAND, IDM_PASTE_CELL, 0);
				return 0;
			}

			if (key == 0x44 && isCtrl) { // Ctrl + D
				SendMessage(hWnd, WM_COMMAND, IDM_DUPLICATE_ROWS, 0);
				return 0;
			}

			if (key == 0x41 && isCtrl) { // Ctrl + A
				ListView_SetItemState(hWnd, -1, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED);
				InvalidateRect(hWnd, NULL, TRUE);
				return 0;
			}

			if (key == 0x20 && isCtrl) { // Ctrl + Space
				SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
				return 0;
			}

			if ((key == VK_F2 || (key == VK_SPACE && !isCtrl))) { // F2 or Space
				SendMessage(hWnd, WMU_EDIT_CELL, key == VK_F2, 0);
				return 0;
			}

			if (!isCtrl &&
				((key == VK_OEM_PLUS) || (key == VK_OEM_MINUS) || (key == VK_OEM_COMMA) || (key == VK_OEM_PERIOD) || // +-,.
				(key >= 0x30 && key <= 0x5A) || // 0, 1, ..., y, z
				(key >= 0x60 && key <= 0x6F) // Numpad
			)) {
				SendMessage(hWnd, WMU_EDIT_CELL, 0, key);
				return TRUE;
			}

			if (key == VK_LEFT || key == VK_RIGHT) {
				HWND hHeader = ListView_GetHeader(hWnd);

				int colCount = Header_GetItemCount(ListView_GetHeader(hWnd));
				int rowNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTROWNO")));
				int colNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTCOLNO")));

				int* colOrder = calloc(colCount, sizeof(int));
				Header_GetOrderArray(hHeader, colCount, colOrder);

				int dir = key == VK_RIGHT ? 1 : -1;
				int idx;
				for (idx = 0; colOrder[idx] != colNo; idx++);
				do {
					idx = (colCount + idx + dir) % colCount;
				} while (ListView_GetColumnWidth(hWnd, colOrder[idx]) == 0);

				colNo = colOrder[idx];
				free(colOrder);

				SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
				return 0;
			}

			if (key == VK_DELETE)
				SendMessage(hWnd, WM_COMMAND, IDM_DELETE_ROWS, 0);

			if (key == VK_INSERT)
				SendMessage(hWnd, WM_COMMAND, IDM_ADD_ROW, 0);
		}
		break;

		case WM_CHAR: {
			return 0;
		}
		break;

		// wParam = tablename16, lParam = is table
		case WMU_SET_SOURCE: {
			TCHAR* tablename16 = (TCHAR*)wParam;
			char* tablename8 = utf16to8(tablename16);
			strncpy(GetProp(hWnd, TEXT("TABLENAME8")), tablename8, MAX_TABLENAME_LENGTH);
			BOOL isTable = !!lParam;

			ShowWindow(hWnd, SW_SHOW);

			HWND hHeader = ListView_GetHeader(hWnd);
			int colCount = Header_GetItemCount(hHeader);
			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			for (int colNo = 0; colNo < colCount; colNo++)
				ListView_DeleteColumn(hWnd, colCount - colNo - 1);

			for (int colNo = 0; colNo < colCount; colNo++)
				DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

			ListView_SetItemCount(hWnd, 0);

			char query8[MAX_TABLENAME_LENGTH + 128];
			sprintf(query8, "select rowid from \"%s\" where 1 = 2", tablename8);
			BOOL hasRowid = sqlite3_exec(db, query8, 0, 0, 0) == SQLITE_OK;

			sprintf(query8, "select *%s from \"%s\" limit 1", hasRowid ? ", rowid" : "", tablename8);

			sqlite3_stmt *stmt;
			if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
				int colCount = sqlite3_column_count(stmt) - hasRowid;
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

					ListView_AddColumn(hWnd, name16, fmt);
					free(name16);
				}

				HFONT hFont = (HFONT)SendMessage(hWnd, WM_GETFONT, 0, 0);
				int filterAlign = getStoredValue("filter-align");
				int align = filterAlign == -1 ? ES_LEFT : filterAlign == 1 ? ES_RIGHT : ES_CENTER;
				for (int colNo = 0; colNo < colCount; colNo++) {
					RECT rc;
					Header_GetItemRect(hHeader, colNo, &rc);
					HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, align | ES_AUTOHSCROLL | WS_CHILD | WS_BORDER | WS_TABSTOP | WS_GROUP, 0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
					SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
					SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
					SetProp(hEdit, TEXT("COLNO"), IntToPtr(colNo));
				}
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
			}
			sqlite3_finalize(stmt);

			int rowCount = 0;
			sprintf(query8, "select count(*) from \"%s\"", tablename8);
			if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
				sqlite3_step(stmt);
				rowCount = sqlite3_column_int(stmt, 0);
			}
			sqlite3_finalize(stmt);

			SetProp(hWnd, TEXT("TOTALROWCOUNT"), IntToPtr(rowCount));
			SetProp(hWnd, TEXT("HASROWID"), IntToPtr(hasRowid));
			SetProp(hWnd, TEXT("ISTABLE"), IntToPtr(isTable));
			SetProp(hWnd, TEXT("ORDERBY"), 0);
			SetProp(hWnd, TEXT("CACHEOFFSET"), 0);

			SendMessage(hWnd, WMU_UPDATE_DATA, 0, 0);

			SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);
			SendMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;

		// wParam = Is store current cell?
		case WMU_UPDATE_DATA: {
			char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));
			HWND hHeader = ListView_GetHeader(hWnd);
			char* where8 = buildFilterWhere(hHeader);
			char* query8 = calloc(MAX_TABLENAME_LENGTH + strlen(where8) + 256, sizeof(char));
			sprintf(query8, "select count(1) from \"%s\" %s", tablename8, where8);

			int rowCount = 0;
			sqlite3_stmt *stmt;
			if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
				bindFilterValues(hHeader, stmt);
				sqlite3_step(stmt);
				rowCount = sqlite3_column_int(stmt, 0);
			}

			sqlite3_finalize(stmt);
			free(where8);
			free(query8);

			ListView_SetItemCount(hWnd, rowCount);
			SetProp(hWnd, TEXT("ROWCOUNT"), IntToPtr(rowCount));
			if (wParam) {
				SetProp(hWnd, TEXT("CURRENTROWNO"), IntToPtr(-1));
				SetProp(hWnd, TEXT("CURRENTCOLNO"), IntToPtr(-1));
				ListView_SetItemState(hWnd, -1, 0, LVIS_FOCUSED | LVIS_SELECTED);
			}

			SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
			InvalidateRect(hWnd, NULL, TRUE);

			SendMessage(GetParent(hWnd), WMU_SB_SET_ROW_COUNTS, (WPARAM)GetProp(hWnd, TEXT("TOTALROWCOUNT")), (LPARAM)rowCount);
		}
		break;

		// only intenal use
		case WMU_UPDATE_CACHE: {
			char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));
			BOOL hasRowid = !!GetProp(hWnd, TEXT("HASROWID"));
			int orderBy = PtrToInt(GetProp(hWnd, TEXT("ORDERBY")));
			int cacheOffset = PtrToInt(GetProp(hWnd, TEXT("CACHEOFFSET")));
			int cacheSize = getStoredValue("cache-size");

			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);

			HWND hHeader = ListView_GetHeader(hWnd);
			int colCount = Header_GetItemCount(hHeader);
			char* where8 = buildFilterWhere(hHeader);
			char* query8 = calloc(MAX_TABLENAME_LENGTH + strlen(where8) + 256, sizeof(char));

			char orderBy8[32] = {0};
			if (orderBy > 0)
				sprintf(orderBy8, "order by %i", orderBy);
			if (orderBy < 0)
				sprintf(orderBy8, "order by %i desc", -orderBy);

			sprintf(query8, "select *%s from \"%s\" %s %s limit %i offset %i", hasRowid ? ", rowid" : "", tablename8, where8, orderBy8, cacheSize, cacheOffset);

			TCHAR*** cache = (TCHAR***)calloc(cacheSize, sizeof(TCHAR*));

			sqlite3_stmt *stmt;
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
			free(where8);
			free(query8);

			SetProp(hWnd, TEXT("CACHE"), cache);
		}
		break;

		case WMU_RESET_CACHE: {
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int cacheSize = getStoredValue("cache-size");
			BOOL hasRowid = !!GetProp(hWnd, TEXT("HASROWID"));

			SetProp(hWnd, TEXT("CACHE"), 0);
			SetProp(hWnd, TEXT("EDITSID"), 0);
			SetProp(hWnd, TEXT("CURRENTROWNO"), IntToPtr(-1));
			SetProp(hWnd, TEXT("CURRENTCOLNO"), IntToPtr(-1));

			if (!cache)
				return 0;

			int colCount = Header_GetItemCount(ListView_GetHeader(hWnd)) + hasRowid;
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

		// wParam = rowNo, lParam = colNo
		case WMU_GET_CELL: {
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int cacheOffset = PtrToInt(GetProp(hWnd, TEXT("CACHEOFFSET")));
			int cacheSize = getStoredValue("cache-size");

			if (!cache)
				return 0;

			int rowNo = wParam;
			int colNo = lParam;

			if (rowNo == -1)
				rowNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTROWNO")));

			if (colNo == -1)
				colNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTCOLNO")));

			if (rowNo == -1 || colNo == -1)
				return 0;

			if ((rowNo < cacheOffset) || (rowNo >= cacheOffset + cacheSize)) {
				cacheOffset = rowNo;
				SetProp(hWnd, TEXT("CACHEOFFSET"), IntToPtr(cacheOffset));
				SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
				cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			}

			return (LRESULT)(cache[rowNo - cacheOffset][colNo]);
		}
		break;

		// wParam = rowNo, lParam = colNo
		case WMU_SET_CURRENT_CELL: {
			int rowNo = (int)wParam;
			int colNo = (int)lParam;
			HWND hHeader = ListView_GetHeader(hWnd);
			int colCount = Header_GetItemCount(hHeader);

			int prevRowNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTROWNO")));
			int prevColNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTCOLNO")));

			if (rowNo == KEEP_PREV_VALUE)
				rowNo = prevRowNo;

			if (colNo == KEEP_PREV_VALUE)
				colNo = prevColNo;

			if (rowNo < -1 || rowNo >= PtrToInt(GetProp(hWnd, TEXT("ROWCOUNT"))))
				rowNo = -1;

			if (colNo < -1 || colNo >= colCount)
				colNo = -1;

			PostMessage(GetParent(hWnd), WMU_SB_SET_CURRENT_CELL, rowNo, colNo);

			if (prevRowNo == rowNo && prevColNo == colNo)
				return 0;

			SetProp(hWnd, TEXT("CURRENTROWNO"), IntToPtr(rowNo));
			SetProp(hWnd, TEXT("CURRENTCOLNO"), IntToPtr(colNo));

			if ((prevRowNo != rowNo || prevColNo != colNo) &&  getStoredValue("show-editor")) {
				if (rowNo == -1 || colNo == -1) {
					SetWindowText(GetDlgItem(GetParent(hWnd), IDC_EDITOR), 0);
				} else {
					SendMessage(hWnd, WMU_EDIT_CELL, 2, 0);
				}
			}

			ListView_EnsureVisibleCell(hWnd, rowNo, colNo);

			RECT rc;
			if (prevRowNo != -1 && prevColNo != - 1) {
				ListView_GetSubItemRect(hWnd, prevRowNo, prevColNo, LVIR_BOUNDS, &rc);
				InvalidateRect(hWnd, &rc, TRUE);
			}

			if (rowNo != -1 && colNo != - 1) {
				ListView_GetSubItemRect(hWnd, rowNo, colNo, LVIR_BOUNDS, &rc);
				InvalidateRect(hWnd, &rc, TRUE);
			}
		}
		break;

		// wParam:
		// 0 - try to edit inside cell, don't keep the text, instead use editor
		// 1 - edit inside cell, keep the text
		// 2 - edit in the editor, keep the text
		// lParam = The kb key that caused the message
		case WMU_EDIT_CELL: {
			BOOL withText = wParam > 0;
			BOOL isTable = PtrToInt(GetProp(hWnd, TEXT("ISTABLE")));
			if (getStoredValue("read-only"))
				return FALSE;

			if (!GetProp(hWnd, TEXT("HASROWID")) && isTable)
				return MessageBox(hWnd, TEXT("Can't edit the value"), NULL, MB_OK);

			int rowNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTROWNO")));
			int colNo = PtrToInt(GetProp(hWnd, TEXT("CURRENTCOLNO")));

			if (rowNo == -1 || colNo == -1)
				return 0;

			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int cacheOffset = PtrToInt(GetProp(hWnd, TEXT("CACHEOFFSET")));
			int cacheSize = getStoredValue("cache-size");
			HWND hCellEdit = GetDlgItem(hWnd, IDC_CELL_EDIT);
			HWND hEditorWnd = GetDlgItem(GetParent(hWnd), IDC_EDITOR);
			BOOL isEditor = wParam > 1 || !isTable;

			if (rowNo == - 1 || rowNo < cacheOffset || rowNo >= cacheOffset + cacheSize)
				return FALSE;

			TCHAR* text16 = cache[rowNo - cacheOffset][colNo];
			if (isEditor)
				SetWindowText(hEditorWnd, text16);

			if (_tcscmp(text16, TEXT(BLOB_VALUE)) == 0) {
				SetProp(hEditorWnd, TEXT("SID"), UIntToPtr(-1));
				MessageBeep(0);
				return 0;
			}

			if (withText && (_tcslen(text16) >= getStoredValue("inline-char-limit") || _tcschr(text16, TEXT('\n')))) {
				int limit = getStoredValue("editor-char-limit");
				if (limit == 0)
					limit = SendMessage(hEditorWnd, EM_GETLIMITTEXT, 0, 0);
				if (_tcslen(text16) >= limit) {
					TCHAR* err16 = loadString(IDS_ERR_EDIT_LIMIT, _tcslen(text16), limit);
					MessageBox(hWnd, err16, 0, 0);
					free(err16);
					return 0;
				}

				if (!getStoredValue("show-editor")) {
					SendMessage(GetParent(hWnd), WM_COMMAND, IDM_EDITOR, 0);
				}

				isEditor = TRUE;
			}

			HWND hEdit = isEditor ? hEditorWnd : hCellEdit;
			if (withText) {
				SetWindowText(hEdit, text16);
				int pos = isEditor && wParam == 2 ? 0 : _tcslen(text16);
				SendMessage(hEdit, EM_SETSEL, pos, pos);
			} else {
				SetWindowText(hEdit, NULL);
			}

			DWORD sid = GetTickCount();
			SetProp(hCellEdit, TEXT("SID"), UIntToPtr(sid));
			SetProp(hCellEdit, TEXT("ISCHANGED"), IntToPtr(0));
			SetProp(hCellEdit, TEXT("COLNO"), IntToPtr(colNo));
			SetProp(hEditorWnd, TEXT("SID"), UIntToPtr(sid));
			SetProp(hEditorWnd, TEXT("ISCHANGED"), IntToPtr(0));
			SetProp(hEditorWnd, TEXT("COLNO"), IntToPtr(colNo));

			SetProp(hWnd, TEXT("EDITSID"), UIntToPtr(sid));
			SetProp(hWnd, TEXT("EDITROWNO"), IntToPtr(rowNo));
			SetProp(hWnd, TEXT("EDITCOLNO"), IntToPtr(colNo));

			ListView_EnsureVisibleCell(hWnd, rowNo, colNo);

			if (isEditor && wParam < 2)
				SetFocus(hEdit);

			if (isEditor)
				return 0;

			RECT rc;
			ListView_GetSubItemRect(hWnd, rowNo, colNo, LVIR_BOUNDS, &rc);
			int h = rc.bottom - rc.top;
			int w = ListView_GetColumnWidth(hWnd, colNo);
			SetWindowPos(hEdit, 0, rc.left, rc.top, w, h, 0);
			SetProp(hEdit, TEXT("ISMOVENEXT"), IntToPtr(0));

			ShowWindow(hEdit, SW_SHOW);
			SetFocus(hEdit);

			if (lParam && (lParam != VK_SPACE))
				keybd_event(lParam, 0, 0, 0);
		}
		break;

		// wParam = is move to the next, lParam = HWND-Source, 0 = Clipboard
		case WMU_UPDATE_CELL: {
			BOOL isMoveNext = !!wParam;
			HWND hEdit = (HWND)lParam;

			if (hEdit && GetProp(hWnd, TEXT("EDITSID")) != GetProp(hEdit, TEXT("SID")))
				return 0;

			int rowNo = PtrToInt(GetProp(hWnd, hEdit ? TEXT("EDITROWNO") : TEXT("CURRENTROWNO")));
			int colNo = PtrToInt(GetProp(hWnd, hEdit ? TEXT("EDITCOLNO") : TEXT("CURRENTCOLNO")));
			char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));

			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int cacheOffset = PtrToInt(GetProp(hWnd, TEXT("CACHEOFFSET")));
			HWND hHeader = ListView_GetHeader(hWnd);
			int colCount = Header_GetItemCount(hHeader);

			TCHAR* value16 = 0;
			if (hEdit) {
				int len = GetWindowTextLength(hEdit);
				value16 = (TCHAR*)calloc (len + 1, sizeof (TCHAR));
				GetWindowText(hEdit, value16, len + 1);
			} else {
				value16 = getClipboardText();
			}

			BOOL rc = TRUE;
			if (_tcscmp(value16, cache[rowNo - cacheOffset][colNo])) {
				TCHAR colName16[MAX_COLUMN_LENGTH + 1];
				Header_GetItemText(hHeader, colNo, colName16, MAX_COLUMN_LENGTH + 1);
				char* colName8 = utf16to8(colName16);
				char query8[1024 + strlen(tablename8) + strlen(colName8)];
				sprintf(query8, "update \"%s\" set \"%s\" = ?1 where rowid = %i", tablename8, colName8, _ttoi(cache[rowNo - cacheOffset][colCount]));
				free(colName8);

				rc = FALSE;
				sqlite3_stmt *stmt;
				if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
					char* value8 = utf16to8(value16);
					bindValue(stmt, 1, value8);
					free(value8);
					rc = SQLITE_DONE == sqlite3_step(stmt);
				}
				sqlite3_finalize(stmt);

				if (rc) {
					if (sqlite3_changes(db) == 1) {
						free(cache[rowNo - cacheOffset][colNo]);
						cache[rowNo - cacheOffset][colNo] = value16;
					}

					RECT rc;
					ListView_GetSubItemRect(hWnd, rowNo, colNo, LVIR_BOUNDS, &rc);
					InvalidateRect(hWnd, &rc, TRUE);
				} else {
					MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
					free(value16);
				}
			}

			if (rc && isMoveNext) {
				int nextColNo = (colNo + 1) % colCount;
				int nextRowNo = rowNo + (nextColNo == 0);

				ListView_SetItemState(hWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
				ListView_SetItemState(hWnd, nextRowNo, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
				PostMessage(hWnd, WMU_SET_CURRENT_CELL, nextRowNo, nextColNo);
				PostMessage(hWnd, WMU_EDIT_CELL, 1, 0);
			}

			if (rc && GetDlgCtrlID(hEdit) == IDC_EDITOR && getStoredValue("hide-editor-on-save"))
				SendMessage(GetParent(hWnd), WM_COMMAND, IDM_EDITOR, 0);

			return rc;
		}
		break;

		case WMU_DUPLICATE_ROWS: {
			if (!GetProp(hWnd, TEXT("HASROWID"))) {
				TCHAR* err16 = loadString(IDS_ERR_NO_ROWID);
				MessageBox(hWnd, err16, 0, 0);
				free(err16);
				return 0;
			}

			HWND hHeader = ListView_GetHeader(hWnd);
			int selCount = ListView_GetSelectedCount(hWnd);
			int rowCount = PtrToInt(GetProp(hWnd, TEXT("ROWCOUNT")));
			char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));

			char* query8 = 0;

			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int cacheSize = getStoredValue("cache-size");
			int cacheOffset = PtrToInt(GetProp(hWnd, TEXT("CACHEOFFSET")));

			int minRowNo = ListView_GetNextItem(hWnd, -1, LVNI_SELECTED);
			int maxRowNo = minRowNo;
			while (TRUE) {
				int rowNo = ListView_GetNextItem(hWnd, maxRowNo, LVNI_SELECTED);
				if (rowNo == -1)
					break;
				maxRowNo = rowNo;
			}

			if (minRowNo < cacheOffset || maxRowNo >= cacheOffset + cacheSize) {
				TCHAR* err16 = loadString(IDS_ERR_CACHE_LIMIT, rowCount, cacheSize);
				MessageBox(hWnd, err16, 0, 0);
				free(err16);
				return 0;
			}

			char* columns8 = 0;
			sqlite3_stmt* stmt;
			if (SQLITE_OK == sqlite3_prepare_v2(db, "select '\"' || group_concat(name, '\", \"') || '\"' from pragma_table_info(?1) where pk <> 1 order by cid", -1, &stmt, 0)) {
				sqlite3_bind_text(stmt, 1, tablename8, -1, SQLITE_TRANSIENT);
				if (SQLITE_ROW == sqlite3_step(stmt)) {
					int len = sqlite3_column_bytes(stmt, 0);
					columns8 = (char*)calloc(len + 1, sizeof(char));
					strncpy(columns8, (const char*)sqlite3_column_text(stmt, 0), len);
				}
			}
			sqlite3_finalize(stmt);

			if (columns8 == 0) {
				MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
				return 0;
			}

			if (strlen(columns8) == 0) {
				TCHAR* err16 = loadString(IDS_ERR_ALL_COLUMNS_PK);
				MessageBox(hWnd, err16, 0, 0);
				free(columns8);

				return 0;
			}

			query8 = (char*)calloc (64 + 2 * strlen(columns8) + 2 * strlen(tablename8) + selCount * 25, sizeof (char));
			sprintf(query8, "insert into \"%s\" (%s) select %s from \"%s\" where rowid in (0", tablename8, columns8, columns8, tablename8);
			free(columns8);


			int colCount = Header_GetItemCount(hHeader);
			int rowNo = ListView_GetNextItem(hWnd, -1, LVNI_SELECTED);
			while (rowNo != -1) {
				char buf8[64];
				sprintf(buf8, ", %i", _ttoi(cache[rowNo - cacheOffset][colCount]));
				strcat(query8, buf8);
				rowNo = ListView_GetNextItem(hWnd, rowNo, LVNI_SELECTED);
			}
			strcat(query8, ")");

			if (SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0)) {
				ListView_SetItemState(hWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
			} else {
				MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
			}

			free(query8);
			return 0;
		}
		break;

		case WMU_DELETE_ROWS: {
			if (!GetProp(hWnd, TEXT("HASROWID"))) {
				TCHAR* err16 = loadString(IDS_ERR_NO_ROWID);
				MessageBox(hWnd, err16, 0, 0);
				free(err16);
				return 0;
			}

			HWND hHeader = ListView_GetHeader(hWnd);
			int selCount = ListView_GetSelectedCount(hWnd);
			int rowCount = PtrToInt(GetProp(hWnd, TEXT("ROWCOUNT")));
			int totalRowCount = PtrToInt(GetProp(hWnd, TEXT("TOTALROWCOUNT")));
			char* tablename8 = (char*)GetProp(hWnd, TEXT("TABLENAME8"));

			char* query8 = 0;

			if (selCount == totalRowCount) {
				query8 = (char*)calloc (64 + strlen(tablename8), sizeof (char));
				sprintf(query8, "delete from \"%s\"", tablename8);
			} else {
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int cacheSize = getStoredValue("cache-size");
				int cacheOffset = PtrToInt(GetProp(hWnd, TEXT("CACHEOFFSET")));

				int minRowNo = ListView_GetNextItem(hWnd, -1, LVNI_SELECTED);
				int maxRowNo = minRowNo;
				while (TRUE) {
					int rowNo = ListView_GetNextItem(hWnd, maxRowNo, LVNI_SELECTED);
					if (rowNo == -1)
						break;
					maxRowNo = rowNo;
				}

				if (minRowNo < cacheOffset || maxRowNo >= cacheOffset + cacheSize) {
					TCHAR* err16 = loadString(IDS_ERR_CACHE_LIMIT, rowCount, cacheSize);
					MessageBox(hWnd, err16, 0, 0);
					free(err16);
					return 0;
				}

				query8 = (char*)calloc (64 + strlen(tablename8) + selCount * 25, sizeof (char));
				sprintf(query8, "delete from \"%s\" where rowid in (0", tablename8);

				int colCount = Header_GetItemCount(hHeader);
				int rowNo = ListView_GetNextItem(hWnd, -1, LVNI_SELECTED);
				while (rowNo != -1) {
					char buf8[64];
					sprintf(buf8, ", %i", _ttoi(cache[rowNo - cacheOffset][colCount]));
					strcat(query8, buf8);
					rowNo = ListView_GetNextItem(hWnd, rowNo, LVNI_SELECTED);
				}
				strcat(query8, ")");
			}

			if (SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0)) {
				ListView_SetItemState(hWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
			} else {
				MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
			}

			free(query8);
			return 0;
		}
		break;

		// wParam = colNo
		case WMU_SORT_COLUMN: {
			int colNo = wParam + 1;
			if (colNo <= 0)
				return FALSE;

			int orderBy = PtrToInt(GetProp(hWnd, TEXT("ORDERBY")));
			orderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
			SetProp(hWnd, TEXT("ORDERBY"), IntToPtr(orderBy));

			SendMessage(hWnd, WMU_UPDATE_DATA, 0, 0);
			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;

		case WMU_UPDATE_FILTER_SIZE: {
			HWND hHeader = ListView_GetHeader(hWnd);
			int colCount = Header_GetItemCount(hHeader);
			SendMessage(hHeader, WM_SIZE, 0, 0);

			int* colOrder = calloc(colCount, sizeof(int));
			Header_GetOrderArray(hHeader, colCount, colOrder);

			for (int idx = 0; idx < colCount; idx++) {
				int colNo = colOrder[idx];
				RECT rc;
				Header_GetItemRect(hHeader, colNo, &rc);
				int h2 = round((rc.bottom - rc.top) / 2);
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left, h2, rc.right - rc.left, h2 + 1, SWP_NOZORDER);
			}

			free(colOrder);
		}
		break;

		case WMU_SET_HEADER_FILTERS: {
			HWND hHeader = ListView_GetHeader(hWnd);
			int isFilterRow = getStoredValue("filter-row");
			int colCount = Header_GetItemCount(hHeader);

			LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
			styles = isFilterRow ? styles | HDS_FILTERBAR : styles & (~HDS_FILTERBAR);
			SetWindowLongPtr(hHeader, GWL_STYLE, styles);

			for (int colNo = 0; colNo < colCount; colNo++)
				ShowWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), isFilterRow ? SW_SHOW : SW_HIDE);


			SetWindowPos(hWnd, 0, 0, 0, 100, 100, SWP_NOZORDER | SWP_NOMOVE); // issue #7: width and height must be greater than 0
			SendMessage(GetParent(hWnd), WM_SIZE, 0, 0);

			if (isFilterRow)
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);

			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;

		case WMU_AUTO_COLUMN_SIZE: {
			SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hWnd);
			int colCount = Header_GetItemCount(hHeader);

			for (int colNo = 0; colNo < colCount - 1; colNo++)
				ListView_SetColumnWidth(hWnd, colNo, colNo < colCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

			if (colCount == 1 && ListView_GetColumnWidth(hWnd, 0) < 100)
				ListView_SetColumnWidth(hWnd, 0, 100);

			int maxWidth = getStoredValue("max-column-width");
			if (colCount > 1) {
				for (int colNo = 0; colNo < colCount; colNo++) {
					if (ListView_GetColumnWidth(hWnd, colNo) > maxWidth)
						ListView_SetColumnWidth(hWnd, colNo, maxWidth);
				}
			}

			// Fix last column
			if (colCount > 1) {
				int colNo = colCount - 1;
				ListView_SetColumnWidth(hWnd, colNo, LVSCW_AUTOSIZE);
				TCHAR name16[MAX_COLUMN_LENGTH + 1];
				Header_GetItemText(hHeader, colNo, name16, MAX_COLUMN_LENGTH);

				SIZE s = {0};
				HDC hDC = GetDC(hHeader);
				HFONT hOldFont = (HFONT)SelectObject(hDC, (HFONT)SendMessage(hWnd, WM_GETFONT, 0, 0));
				GetTextExtentPoint32(hDC, name16, _tcslen(name16), &s);
				SelectObject(hDC, hOldFont);
				ReleaseDC(hHeader, hDC);

				int w = s.cx + 12;
				if (ListView_GetColumnWidth(hWnd, colNo) < w)
					ListView_SetColumnWidth(hWnd, colNo, w);

				if (ListView_GetColumnWidth(hWnd, colNo) > maxWidth)
					ListView_SetColumnWidth(hWnd, colNo, maxWidth);
			}

			SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hWnd, NULL, TRUE);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		// wParam = colNo
		case WMU_HIDE_COLUMN: {
			int colNo = (int)wParam;
			if (colNo < 0)
				return 0;

			HWND hHeader = ListView_GetHeader(hWnd);
			HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
			SetWindowLongPtr(hEdit, GWLP_USERDATA, (LONG_PTR)ListView_GetColumnWidth(hWnd, colNo));
			ListView_SetColumnWidth(hWnd, colNo, 0);
			InvalidateRect(hHeader, NULL, TRUE);
		}
		break;
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg){
		case WM_CTLCOLOREDIT:
		case WM_CTLCOLORSTATIC: {
			HWND hMainWnd = GetAncestor(hWnd, GA_ROOT);
			SetBkColor((HDC)wParam, (COLORREF)PtrToInt(GetProp(hMainWnd, TEXT("FILTERBACKCOLOR"))));
			SetTextColor((HDC)wParam, (COLORREF)PtrToInt(GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"))));
			return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));
		}
		break;

		case WM_COMMAND: {
			if (HIWORD(wParam) == EN_CHANGE && LOWORD(wParam) >= IDC_HEADER_EDIT && LOWORD(wParam) < IDC_HEADER_EDIT + Header_GetItemCount(hWnd)) {
				if (getStoredValue("auto-filters"))
					SendMessage(GetParent(hWnd), WMU_UPDATE_DATA, 0, 0);
				else
					SendMessage((HWND)lParam, WMU_AUTOCOMPLETE, 0, 0);
			}
		}
		break;
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (cbNewAutoCompleteEdit(hWnd, msg, wParam, lParam))
		return TRUE;

	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetClientRect(hWnd, &rc);
			HWND hMainWnd = GetAncestor(hWnd, GA_ROOT);
			BOOL isDark = !!getStoredValue("dark-theme");

			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, (COLORREF)PtrToInt(GetProp(hMainWnd, TEXT("FILTERBACKCOLOR"))));
			HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
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
			SetProp(GetParent(hWnd), TEXT("LASTFOCUS"), hWnd);
		}
		break;

		case WM_KEYDOWN: {
			HWND hMainWnd = GetAncestor(hWnd, GA_ROOT);
			HWND hGridWnd = GetDlgItem(hMainWnd, IDC_GRID);
			if (wParam == VK_RETURN) {
				SendMessage(hGridWnd, WMU_UPDATE_DATA, 0, 0);
				InvalidateRect(hGridWnd, NULL, TRUE);
				return 0;
			}

			if (wParam == VK_UP || wParam == VK_DOWN) {
				SetFocus(hGridWnd);
				SendMessage(hMainWnd, msg, wParam, lParam);
				return 0;
			}
		}
		break;

		case WM_CHAR: {
			if (wParam == VK_RETURN)
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

LRESULT CALLBACK cbNewEditor(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch(msg){
		case WM_KEYDOWN: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
			int key = wParam;

			if (key == 0x53 && isCtrl) { // Ctrl + S
				SendMessage(hWnd, WM_COMMAND, IDM_EDITOR_SAVE, 0);
				return 0;
			}

			if (key == 0x57 && isCtrl) { // Ctrl + W
				SendMessage(hWnd, WM_COMMAND, IDM_EDITOR_WORDWRAP, 0);
				return 0;
			}
		}
		break;

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);

			if (cmd == IDM_EDITOR_COPY) {
				SendMessage(hWnd, WM_COPY, 0, 0);
			}

			if (cmd == IDM_EDITOR_PASTE) {
				SendMessage(hWnd, WM_PASTE, 0, 0);
			}

			if (cmd == IDM_EDITOR_SAVE) {
				HWND hMainWnd = GetParent(hWnd);
				HWND hGridWnd = GetDlgItem(hMainWnd, IDC_GRID);

				if (SendMessage(hGridWnd, WMU_UPDATE_CELL, 0, (LPARAM)hWnd) && getStoredValue("hide-editor-on-save")) {
					SendMessage(hMainWnd, IDM_EDITOR, 0, 0);
				}
			}

			if (cmd == IDM_EDITOR_WORDWRAP) {
				BOOL isWordWrap = !getStoredValue("editor-word-wrap");
				setStoredValue("editor-word-wrap", isWordWrap);

				SendMessage(hWnd, EM_SETTARGETDEVICE,(WPARAM)NULL, (LPARAM)!isWordWrap);
				InvalidateRect(hWnd, NULL, TRUE);
			}
		}
		break;
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

BOOL cbNewAutoCompleteEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		// wParam = Is force?
		case WMU_AUTOCOMPLETE: {
			if (hWnd != GetFocus())
				return FALSE;

			HWND hMainWnd = GetAncestor(hWnd, GA_ROOT);
			if (GetProp(hWnd, TEXT("CHANGEBYAUTOCOMPLETE"))) {
				SetProp(hWnd, TEXT("CHANGEBYAUTOCOMPLETE"), 0);
				return FALSE;
			}

			if (hWnd != GetProp(hAutoComplete, TEXT("CURRENTINPUT"))) {
				SetProp(hAutoComplete, TEXT("CURRENTINPUT"), (HANDLE)hWnd);

				if (GetWindowTextLength(hWnd) < 3 && !wParam) {
					ShowWindow(hAutoComplete, SW_HIDE);
					return FALSE;
				}
			}

			HWND hGridWnd = GetDlgItem(hMainWnd, IDC_GRID);
			if (!hGridWnd)
				hGridWnd = (HWND)GetWindowLongPtr(hMainWnd, GWLP_USERDATA);
			HWND hHeader = ListView_GetHeader(hGridWnd);

			char* tablename8 = (char*)GetProp(hGridWnd, TEXT("TABLENAME8"));
			int colNo = PtrToInt(GetProp(hWnd, TEXT("COLNO")));

			sqlite3_stmt* stmt;
			char query8[1024 + 2 * MAX_TABLENAME_LENGTH  + 4 * MAX_COLUMN_LENGTH];
			BOOL isEnable = TRUE;
			sprintf(query8, "select max(rowid) from \"%s\"", tablename8);
			if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0) && SQLITE_ROW == sqlite3_step(stmt))
				isEnable = getStoredValue("max-autocomplete-rowid") > sqlite3_column_int(stmt, 0);
			sqlite3_finalize(stmt);

			if (!isEnable) {
				ShowWindow(hAutoComplete, SW_HIDE);
				return FALSE;
			}

			TCHAR colName16[MAX_COLUMN_LENGTH + 1] = {0};
			Header_GetItemText(hHeader, colNo, colName16, MAX_COLUMN_LENGTH + 1);
			char* colName8 = utf16to8(colName16);

			char fkTableName8[MAX_TABLENAME_LENGTH + 1] = {0};
			char fkColName8[MAX_COLUMN_LENGTH + 1] = {0};
			if (SQLITE_OK == sqlite3_prepare_v2(db, "select \"table\", \"to\" from pragma_foreign_key_list(?1) where \"from\" = ?2", -1, &stmt, 0)) {
				sqlite3_bind_text(stmt, 1, tablename8, -1, SQLITE_TRANSIENT);
				sqlite3_bind_text(stmt, 2, colName8, -1, SQLITE_TRANSIENT);

				if (SQLITE_ROW == sqlite3_step(stmt)) {
					strncpy(fkTableName8, (const char*)sqlite3_column_text(stmt, 0), MAX_TABLENAME_LENGTH);
					strncpy(fkColName8, (const char*)sqlite3_column_text(stmt, 1), MAX_TABLENAME_LENGTH);
				}
			}
			sqlite3_finalize(stmt);

			int maxRows = getStoredValue("max-autocomplete-results");
			if (strlen(fkTableName8)) {
				sprintf(query8, "select distinct coalesce(trim(val), '') from (" \
					"select \"%s\" as val from \"%s\" where instr(lower(\"%s\"), coalesce(lower(?1), '')) > 0 and length(trim(\"%s\")) > 0" \
					" union " \
					"select \"%s\" as val from \"%s\" where instr(lower(\"%s\"), coalesce(lower(?1), '')) > 0 and length(trim(\"%s\")) > 0" \
					") order by 1 limit %i",
					colName8, tablename8, colName8, colName8,
					fkColName8, fkTableName8, fkColName8, fkColName8,
					maxRows);
			} else {
				sprintf(query8,
					"select distinct coalesce(trim(\"%s\"), '') from \"%s\" where instr(lower(\"%s\"), coalesce(lower(?1), '')) > 0 and length(trim(\"%s\")) > 0 limit %i",
					colName8, tablename8, colName8, colName8,
					maxRows);
			}
			free(colName8);

			SendMessage(hAutoComplete, LB_RESETCONTENT, 0, 0);
			if (SQLITE_OK == sqlite3_prepare_v2(db, query8, -1, &stmt, 0)) {
				TCHAR* value16 = calloc(MAX_TEXT_LENGTH, sizeof(TCHAR));
				GetWindowText(hWnd, value16, MAX_TEXT_LENGTH);

				char* value8 = utf16to8(value16);
				bindValue(stmt, 1, value8);
				free(value8);
				free(value16);

				while (SQLITE_ROW == sqlite3_step(stmt)) {
					TCHAR* value16 = utf8to16((char*)sqlite3_column_text(stmt, 0));
					ListBox_AddString(hAutoComplete, value16);
					free(value16);
				}
			} else {
				MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
			}
			sqlite3_finalize(stmt);

			int cnt = ListBox_GetCount(hAutoComplete);
			if (cnt == 0) {
				ShowWindow(hAutoComplete, SW_HIDE);
				return TRUE;
			}

			RECT rc;
			GetWindowRect(hWnd, &rc);
			int h = ListBox_GetItemHeight(hAutoComplete, SW_HIDE);
			SetWindowPos(hAutoComplete, 0, rc.left, rc.bottom - 1, rc.right - rc.left, (cnt < 8 ? h * cnt : h * 7) + h/2, SWP_NOZORDER | SWP_NOACTIVATE);

			if (!IsWindowVisible(hAutoComplete))
				ShowWindow(hAutoComplete, SW_SHOWNOACTIVATE);

			ListBox_SetCurSel(hAutoComplete, 0);
		}
		break;

		case WM_KILLFOCUS: {
			if ((HWND)wParam != hAutoComplete)
				ShowWindow(hAutoComplete, SW_HIDE);
		}
		break;

		case WM_KEYDOWN: {
			if (wParam == 0x20 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + Space
				SendMessage(hWnd, WMU_AUTOCOMPLETE, 1, 0);
				return TRUE;
			}

			if (wParam == VK_UP || wParam == VK_DOWN) {
				if (!IsWindowVisible(hAutoComplete))
					return FALSE;

				int cnt = ListBox_GetCount(hAutoComplete);
				int pos = ListBox_GetCurSel(hAutoComplete);
				ListBox_SetCurSel(hAutoComplete, (cnt + pos + (wParam == VK_UP ? -1 : 1)) % cnt);

				return TRUE;
			}

			if (wParam == VK_RETURN || wParam == VK_TAB) {
				if (IsWindowVisible(hAutoComplete)) {
					SendMessage(hAutoComplete, WM_LBUTTONUP, 0, 0);

					return TRUE;
				}
			}

			if (wParam == VK_ESCAPE) {
				if (IsWindowVisible(hAutoComplete)) {
					ShowWindow(hAutoComplete, SW_HIDE);
					return TRUE;
				}
			}
		}
		break;

		case WM_CHAR: {
			return (wParam == 0x20 && HIWORD(GetKeyState(VK_CONTROL))) || wParam == VK_ESCAPE;
		}
		break;
	}

	return FALSE;
}

LRESULT CALLBACK cbNewAutoComplete(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_LBUTTONUP: {
			TCHAR buf[1024] = {0};
			ListBox_GetText(hWnd, ListBox_GetCurSel(hWnd), buf);

			RECT rc;
			GetWindowRect(hWnd, &rc);
			ShowWindow(hWnd, SW_HIDE);

			HWND hEdit = (HWND)GetProp(hWnd, TEXT("CURRENTINPUT"));
			SetProp(hEdit, TEXT("CHANGEBYAUTOCOMPLETE"), IntToPtr(1));
			SetWindowText(hEdit, buf);
			int len = _tcslen(buf);
			SendMessage(hEdit, EM_SETSEL, len, len);

			SendMessage(hEdit, WMU_UPDATE_CELL, 0, 0);
		}
		break;
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewCellEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (cbNewAutoCompleteEdit(hWnd, msg, wParam, lParam))
		return TRUE;

	switch (msg) {
		case WM_KILLFOCUS: {
			if ((HWND)wParam != hAutoComplete && IsWindowVisible(hWnd)) {
				SendMessage(hWnd, WMU_UPDATE_CELL, 0, 0);
			}
		}
		break;

		case WM_KEYDOWN: {
			if (wParam == VK_RETURN || wParam == VK_TAB) {
				SetProp(hWnd, TEXT("ISMOVENEXT"), IntToPtr(1));
				SetFocus(GetParent(hWnd));
				return TRUE;
			}

			if (wParam == VK_ESCAPE) {
				ShowWindow(hWnd, SW_HIDE);
				return TRUE;
			}

			if (wParam == 0x41 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + A
				SendMessage(hWnd, EM_SETSEL, 0, -1);
				return TRUE;
			}
		}
		break;

		case WM_CHAR: {
			if (wParam == VK_RETURN || wParam == VK_TAB)
				return 0;
		}
		break;

		case WMU_UPDATE_CELL: {
			HWND hGridWnd = GetParent(hWnd);
			if (SendMessage(hGridWnd, WMU_UPDATE_CELL, (WPARAM)GetProp(hWnd, TEXT("ISMOVENEXT")), (LPARAM)hWnd)) {
				ShowWindow(hWnd, SW_HIDE);
				SetFocus(hGridWnd);
			}

			return 0;
		}
		break;
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewColumns(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_COMMAND: {
			SendMessage(GetParent(hWnd), WM_COMMAND, wParam, lParam);
		}
		break;

		case WM_SETFOCUS: {
			BOOL isShift = HIWORD(GetKeyState(VK_SHIFT));
			int colCount = PtrToInt(GetProp(GetParent(hWnd), TEXT("COLCOUNT")));
			SetFocus(GetDlgItem(hWnd, IDC_DLG_EDIT + (isShift ? colCount - 1 : 0)));
		}
		break;

		case WM_MOUSEWHEEL: {
			SendMessage(hWnd, WM_VSCROLL, MAKEWPARAM(GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? SB_LINEUP: SB_LINEDOWN, 0), 0);
			return 0;
		}
		break;

		case WM_VSCROLL: {
			WORD action = LOWORD(wParam);
			int scrollY = PtrToInt(GetProp(hWnd, TEXT("SCROLLY")));

			int pos = action == SB_THUMBPOSITION ? HIWORD(wParam) :
				action == SB_THUMBTRACK ? HIWORD(wParam) :
				action == SB_LINEDOWN ? scrollY + 30 :
				action == SB_LINEUP ? scrollY - 30 :
				action == SB_TOP ? 0 :
				action == SB_BOTTOM ? 100 :
				-1;

			if (pos == -1)
				return 0;

			if (!(GetWindowLong(hWnd, GWL_STYLE) & WS_VSCROLL))
				return 0;

			int currPos = GetScrollPos(hWnd, SB_VERT);
			SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);

			SCROLLINFO si = {0};
			si.cbSize = sizeof(SCROLLINFO);
			si.fMask = SIF_POS;
			si.nPos = pos;
			si.nTrackPos = 0;
			SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
			GetScrollInfo(hWnd, SB_VERT, &si);
			pos = si.nPos;
			POINT p = {0, pos - scrollY};
			HDC hdc = GetDC(hWnd);
			LPtoDP(hdc, &p, 1);
			ReleaseDC(hWnd, hdc);
			ScrollWindow(hWnd, 0, -p.y, NULL, NULL);
			SetProp(hWnd, TEXT("SCROLLY"), IntToPtr(pos));

			SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
			if (currPos != pos)
				InvalidateRect(hWnd, NULL, TRUE);

			return 0;
		}
		break;

		case WM_ERASEBKGND: {
			HDC hDC = (HDC)wParam;
			RECT rc = {0};
			GetClientRect(hWnd, &rc);
			HBRUSH hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE));
			FillRect(hDC, &rc, hBrush);
			DeleteObject(hBrush);

			return 1;
		}
		break;

		case WM_CTLCOLOREDIT: {
			BOOL isDark = !!getStoredValue("dark-theme");
			if (!isDark)
				return 0;

			HWND hGridWnd = (HWND)GetWindowLongPtr(GetParent(hWnd), GWLP_USERDATA);
			HWND hMainWnd = GetParent(hGridWnd);

			SetBkColor((HDC)wParam, (COLORREF)PtrToInt(GetProp(hMainWnd, TEXT("FILTERBACKCOLOR"))));
			SetTextColor((HDC)wParam, (COLORREF)PtrToInt(GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"))));
			return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));
		}
		break;

		case WMU_SET_SCROLL_HEIGHT: {
			RECT rc = {0};
			GetWindowRect(hWnd, &rc);

			SCROLLINFO si = {0};
			si.cbSize = sizeof(SCROLLINFO);
			si.fMask = SIF_ALL;
			si.nMin = 0;
			si.nMax = wParam;
			si.nPage = rc.bottom - rc.top;
			si.nPos = 0;
			si.nTrackPos = 0;
			SetScrollInfo(hWnd, SB_VERT, &si, TRUE);

			SetProp(hWnd, TEXT("SCROLLHEIGHT"), IntToPtr((int)wParam));
		}
		break;

		case WMU_UPDATE_SCROLL_POS: {
			int colNo = GetDlgCtrlID(GetFocus()) - IDC_DLG_EDIT;
			int lineHeight = PtrToInt(GetProp(hWnd, TEXT("LINEHEIGHT")));
			int H = 2 + (colNo + 1) * lineHeight;

			RECT rc;
			GetClientRect(hWnd, &rc);
			int h = rc.bottom;

			SendMessage(hWnd, WM_VSCROLL, MAKEWPARAM(SB_THUMBPOSITION, MAX(0, H - h)), 0);
			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;

		case WM_SIZE: {
			int H = HIWORD(lParam);
			int maxH = PtrToInt(GetProp(hWnd, TEXT("SCROLLHEIGHT")));

			if (H > maxH) {
				SendMessage(hWnd, WM_VSCROLL, SB_TOP, 0);
				SetProp(hWnd, TEXT("SCROLLY"), 0);
			}
			SendMessage(hWnd, WMU_SET_SCROLL_HEIGHT, maxH, 0);
			SendMessage(hWnd, WMU_UPDATE_SCROLL_POS, 0, 0);
			ShowScrollBar(hWnd, SB_VERT, H < maxH);
		}
		break;

		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("LINEHEIGHT"));
			RemoveProp(hWnd, TEXT("SCROLLHEIGHT"));
			RemoveProp(hWnd, TEXT("SCROLLY"));
		}
		break;
	}

	return CallWindowProc(DefWindowProc, hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewAddRowEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (cbNewAutoCompleteEdit(hWnd, msg, wParam, lParam))
		return TRUE;

	if (msg == WM_GETDLGCODE)
		return (DLGC_WANTTAB | CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam));

	if (msg == WM_KEYDOWN && wParam == VK_TAB) {
		BOOL isShift = HIWORD(GetKeyState(VK_SHIFT));

		int idc = GetDlgCtrlID(hWnd);
		HWND hColumnsWnd = GetParent(hWnd);
		HWND hNextWnd = GetDlgItem(hColumnsWnd, idc + (isShift ? -1: 1));
		SetFocus(hNextWnd ? hNextWnd : GetDlgItem(GetParent(hColumnsWnd), IDC_DLG_OK));
		return TRUE;
	}

	if (msg == WM_CHAR && wParam == VK_TAB)
		return FALSE;

	if (msg == WM_SETFOCUS) {
		RECT rc;
		GetWindowRect(hWnd, &rc);
		BOOL isTopVisible = WindowFromPoint((POINT){rc.left + 1, rc.top + 1}) == hWnd;
		BOOL isBottomVisible = WindowFromPoint((POINT){rc.left + 1, rc.bottom - 1}) == hWnd;

		if (!IsWindowVisible(hWnd) || !isTopVisible || !isBottomVisible)
			SendMessage(GetParent(hWnd), WMU_UPDATE_SCROLL_POS, 0, 0);
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

// lParam = USERDATA = hMainWnd
INT_PTR CALLBACK cbDlgAddTable(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_INITDIALOG: {
			SetWindowLongPtr(hWnd, GWLP_USERDATA, lParam);
			SetFocus(GetDlgItem(hWnd, IDC_DLG_TABLENAME));
		}
		break;

		case WM_CTLCOLOREDIT: {
			HWND hMainWnd = (HWND)GetWindowLongPtr(hWnd, GWLP_USERDATA);
			BOOL isDark = !!GetProp(hMainWnd, TEXT("DARKTHEME"));
			if (!isDark)
				return 0;

			SetBkColor((HDC)wParam, (COLORREF)PtrToInt(GetProp(hMainWnd, TEXT("FILTERBACKCOLOR"))));
			SetTextColor((HDC)wParam, (COLORREF)PtrToInt(GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"))));
			return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));
		}
		break;

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDC_DLG_OK) {
				HWND hMainWnd = (HWND)GetWindowLongPtr(hWnd, GWLP_USERDATA);

				TCHAR tablename16[MAX_TABLENAME_LENGTH];
				GetDlgItemText(hWnd, IDC_DLG_TABLENAME, tablename16, MAX_TABLENAME_LENGTH);
				if (_tcslen(tablename16) == 0) {
					TCHAR* err16 = loadString(IDS_ERR_NO_TABLENAME);
					MessageBox(hWnd, err16, 0, 0);
					free(err16);
					return 0;
				}

				TCHAR columns16[MAX_TEXT_LENGTH];
				GetDlgItemText(hWnd, IDC_DLG_COLUMNS, columns16, MAX_TEXT_LENGTH);
				if (_tcslen(columns16) == 0) {
					TCHAR* err16 = loadString(IDS_ERR_NO_COLUMNS);
					MessageBox(hWnd, err16, 0, 0);
					free(err16);
					return 0;
				}

				int len = _tcslen(tablename16) + _tcslen(columns16) + 256;
				TCHAR query16[len + 1];
				_sntprintf(query16, len, TEXT("create table \"%ls\" (%ls)"), tablename16, columns16);

				char* query8 = utf16to8(query16);
				BOOL rc = SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0);
				free(query8);

				if (!rc)
					return MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);

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

// lParam = USERDATA = hGridWnd
INT_PTR CALLBACK cbDlgAddRow(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_INITDIALOG: {
			SetWindowLongPtr(hWnd, GWLP_USERDATA, lParam);

			HWND hGridWnd = (HWND)lParam;
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			SetProp(hWnd, TEXT("COLCOUNT"), IntToPtr(colCount));

			float z = (getFontHeight(hGridWnd)) / 15.;
			int w = 120 * z;
			int h = 20 * z;

			HWND hColumnsWnd = CreateWindow(WC_STATIC, NULL, WS_VISIBLE | WS_CHILD | WS_TABSTOP, 0, 10, 0, 0, hWnd, (HMENU)IDC_DLG_COLUMNS, GetModuleHandle(0), 0);
			SetWindowLongPtr(hColumnsWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewColumns);
			SendMessage(hColumnsWnd, WMU_SET_SCROLL_HEIGHT, colCount * h, 0);
			SetProp(hColumnsWnd, TEXT("LINEHEIGHT"), IntToPtr(h));

			HFONT hFont = (HFONT)SendMessage((HWND)lParam, WM_GETFONT, 0, 0);

			for (int colNo = 0; colNo < colCount; colNo++) {
				TCHAR colName16[256];
				Header_GetItemText(hHeader, colNo, colName16, 255);
				HWND hLabel = CreateWindow(WC_STATIC, colName16, WS_VISIBLE | WS_CHILD | SS_WORDELLIPSIS | SS_RIGHT,
					5, colNo * h + 2 * z, w - 5, h, hColumnsWnd, (HMENU)IntToPtr(IDC_DLG_LABEL + colNo), GetModuleHandle(0), 0);
				SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, 0);
				HWND hEdit = CreateWindow(WC_EDIT, NULL, WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_BORDER | WS_CLIPSIBLINGS | ES_AUTOHSCROLL,
					w + 5 * z, 2 + colNo * h - 2, 0, 0, hColumnsWnd, (HMENU)IntToPtr(IDC_DLG_EDIT + colNo), GetModuleHandle(0), 0);
				SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, 0);
				SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)&cbNewAddRowEdit));
				SetProp(hEdit, TEXT("COLNO"), (HANDLE)(LONG_PTR)colNo);
			}

			char* tablename8 = (char*)GetProp(hGridWnd, TEXT("TABLENAME8"));

			sqlite3_stmt* stmt;
			if (SQLITE_OK == sqlite3_prepare_v2(db, "select replace(dflt_value, '\"\"', '\"') from pragma_table_info(?1) order by cid", -1, &stmt, 0)) {
				sqlite3_bind_text(stmt, 1, tablename8, -1, SQLITE_TRANSIENT);

				int colNo = 0;
				while (SQLITE_ROW == sqlite3_step(stmt)) {
					TCHAR* default16 = utf8to16((const char*)sqlite3_column_text(stmt, 0));
					int len = _tcslen(default16);
					if (len > 0) {
						BOOL isQ = default16[0] == TEXT('"') && default16[len - 1] == TEXT('"');
						default16[len - 1] = isQ ? 0 : default16[len - 1];
						Edit_SetCueBannerText(GetDlgItem(hColumnsWnd, IDC_DLG_EDIT + colNo), isQ ? default16 + 1 : default16);
					}
					free(default16);

					colNo++;
				}
			}
			sqlite3_finalize(stmt);

			RECT rc;
			GetWindowRect(hGridWnd, &rc);
			SetWindowPos(hWnd, 0,
					rc.left + 50, rc.top + 50, 340 * z,
					MIN(GetSystemMetrics(SM_CYCAPTION) + 2 * GetSystemMetrics(SM_CYBORDER) + colCount * h + 40 * z + 50 * z, rc.bottom - rc.top), SWP_NOZORDER);
			SetWindowPos(GetDlgItem(hWnd, IDC_DLG_OK), 0, 0, 0, 80 * z, 22 * z, SWP_NOMOVE | SWP_NOZORDER);

			GetClientRect(hWnd, &rc);
			BOOL isMaximized = getStoredValue("show-row-maximized");
			SendMessage(hWnd, WM_SIZE, isMaximized ? SIZE_MAXIMIZED : SIZE_RESTORED, MAKELPARAM(rc.right, rc.bottom));
			ShowWindow(hWnd, isMaximized ? SW_MAXIMIZE : SW_NORMAL);

			InvalidateRect(hWnd, NULL, TRUE);

			SetFocus(GetDlgItem(hColumnsWnd, IDC_DLG_EDIT));
		}
		break;

		case WM_GETMINMAXINFO: {
			LPMINMAXINFO lpMMI = (LPMINMAXINFO)lParam;
			lpMMI->ptMinTrackSize.x = 200;
			lpMMI->ptMinTrackSize.y = 100;
		}
		break;

		case WM_SIZE: {
			HWND hGridWnd = (HWND)GetWindowLongPtr(hWnd, GWLP_USERDATA);
			HWND hColumnsWnd = GetDlgItem(hWnd, IDC_DLG_COLUMNS);
			int colCount = PtrToInt(GetProp(hWnd, TEXT("COLCOUNT")));
			int W = (int)LOWORD(lParam);
			int H = (int)HIWORD(lParam);
			float z = (getFontHeight(hGridWnd)) / 15.;

			SetWindowPos(GetDlgItem(hWnd, IDC_DLG_COLUMNS), 0, 0, 0, W, H - 32 * z - 10, SWP_NOMOVE | SWP_NOZORDER);
			SetWindowPos(GetDlgItem(hWnd, IDC_DLG_OK), 0, W - 85 * z, H - 27 * z, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

			RECT rc;
			GetClientRect(hColumnsWnd, &rc);
			for (int colNo = 0; colNo < colCount; colNo++)
				SetWindowPos(GetDlgItem(hColumnsWnd, IDC_DLG_EDIT + colNo), 0, 0, 0, rc.right - 125 * z - 5 * z, 20 * z - 1, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;

		case WM_MOUSEWHEEL: {
			return SendMessage(GetDlgItem(hWnd, IDC_DLG_COLUMNS), WM_MOUSEWHEEL, wParam, lParam);
		}
		break;

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);

			if (cmd == IDC_DLG_OK) {
				HWND hGridWnd = (HWND)GetWindowLongPtr(hWnd, GWLP_USERDATA);
				HWND hColumnsWnd = GetDlgItem(hWnd, IDC_DLG_COLUMNS);

				HWND hHeader = ListView_GetHeader(hGridWnd);
				int colCount = Header_GetItemCount(hHeader);
				char* tablename8 = (char*)GetProp(hGridWnd, TEXT("TABLENAME8"));

				int len = 128;
				for (int colNo = 0; colNo < colCount; colNo++)
					len += GetWindowTextLength(GetDlgItem(hColumnsWnd, IDC_DLG_LABEL + colNo)) + GetWindowTextLength(GetDlgItem(hColumnsWnd, IDC_DLG_EDIT + colNo)) + 10;

				char* query8 = (char*)calloc(len, sizeof(char));
				sprintf(query8, "insert into \"%s\" (", tablename8);

				int valCount = 0;
				for (int colNo = 0; colNo < colCount; colNo++) {
					if (!GetWindowTextLength(GetDlgItem(hColumnsWnd, IDC_DLG_EDIT + colNo)))
						continue;

					TCHAR colName16[255];
					Header_GetItemText(hHeader, colNo, colName16, 255);
					char* colName8 = utf16to8(colName16);
					strcat(query8, valCount != 0 ? "\", \"" : "\"");
					strcat(query8, colName8);
					free(colName8);

					valCount++;
				}
				strcat(query8, "\") values (");

				valCount = 0;
				for (int colNo = 0; colNo < colCount; colNo++) {
					int len = GetWindowTextLength(GetDlgItem(hColumnsWnd, IDC_DLG_EDIT + colNo)) + 1;
					TCHAR value16[len + 1];
					GetWindowText(GetDlgItem(hColumnsWnd, IDC_DLG_EDIT + colNo), value16, len);

					if (!_tcslen(value16))
						continue;

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
					strcat(query8, valCount != 0 ? ", " : "");
					strcat(query8, "\'");
					strcat(query8, qvalue8);
					strcat(query8, "\'");
					free(value8);
					free(qvalue8);

					valCount++;
				}
				strcat(query8, ")");

				if (valCount == 0)
					sprintf(query8, "insert into \"%s\" default values", tablename8);

				if (SQLITE_OK == sqlite3_exec(db, query8, 0, 0, 0)) {
					free(query8);
					EndDialog(hWnd, IDOK);
				} else {
					free(query8);
					MessageBoxA(hWnd, sqlite3_errmsg(db), NULL, MB_OK);
				}
			}

			if (cmd == IDCANCEL || cmd == IDOK) {
				RemoveProp(hWnd, TEXT("COLCOUNT"));

				WINDOWPLACEMENT wp = {};
				wp.length = sizeof(WINDOWPLACEMENT);
				GetWindowPlacement(hWnd, &wp);
				setStoredValue("show-row-maximized", wp.showCmd == SW_SHOWMAXIMIZED);
			}

			if (cmd == IDCANCEL) {
				return IsWindowVisible(hAutoComplete) ? ShowWindow(hAutoComplete, SW_HIDE) : EndDialog (hWnd, IDCANCEL);
			}

			if (cmd == IDOK) {
				return IsWindowVisible(hAutoComplete) ? SendMessage(hAutoComplete, WM_LBUTTONUP, 0, 0) : SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(IDC_DLG_OK, 0), 0);
			}

			if (HIWORD(wParam) == EN_CHANGE && cmd >= IDC_DLG_EDIT && cmd < IDC_DLG_EDIT + 100) {
				SendMessage((HWND)lParam, WMU_AUTOCOMPLETE, 0, 0);
			}
		}
		break;
	}

	return FALSE;
}


void bindValue(sqlite3_stmt* stmt, int i, const char* value8) {
	int len = strlen(value8);
	if (len == 0) {
		sqlite3_bind_null(stmt, i);
		return;
	}

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
		sqlite3_bind_text(stmt, i, value8, -1, SQLITE_TRANSIENT);
	}
}

char* buildFilterWhere(HWND hHeader) {
	int colCount = Header_GetItemCount(hHeader);
	TCHAR where16[MAX_TEXT_LENGTH] = {0};
	_tcscat(where16, TEXT("where (1 = 1) "));

	TCHAR colname16[MAX_COLUMN_LENGTH + 1];
	for (int colNo = 0; colNo < colCount; colNo++) {
		HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
		if (GetWindowTextLength(hEdit) > 0) {
			Header_GetItemText(hHeader, colNo, colname16, MAX_COLUMN_LENGTH);
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
	return utf16to8(where16);
}

void bindFilterValues(HWND hHeader, sqlite3_stmt* stmt) {
	int colCount = Header_GetItemCount(hHeader);
	int bindNo = 1;
	for (int colNo = 0; colNo < colCount; colNo++) {
		HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
		int size = GetWindowTextLength(hEdit);
		if (size == 0)
			continue;

		TCHAR value16[size + 1];
		GetWindowText(hEdit, value16, size + 1);
		char* value8 = utf16to8(value16[0] == TEXT('=') || value16[0] == TEXT('/') || value16[0] == TEXT('!') || value16[0] == TEXT('<') || value16[0] == TEXT('>') ? value16 + 1 : value16);
		bindValue(stmt, bindNo, value8);
		free(value8);
		bindNo++;
	}
}

void setStoredValue(char* name, int value) {
	for (int i = 0; storedKeys[i] != 0; i++) {
		if (strcmp(name, storedKeys[i]) == 0)
			storedValues[i] = value;
	}
}

int getStoredValue(char* name) {
	for (int i = 0; storedKeys[i] != 0; i++) {
		if (strcmp(name, storedKeys[i]) == 0)
			return storedValues[i];
	}

	printf("Unknown key: %s\n", name);
	return 0;
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
			res = (pos == 0 || (pos > 0 && !_istalnum(text[pos - 1]))) &&
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

TCHAR* getClipboardText() {
	TCHAR* res = 0;
	if (OpenClipboard(NULL)) {
		HANDLE clip = GetClipboardData(CF_UNICODETEXT);
		TCHAR* str = (LPWSTR)GlobalLock(clip);
		if (str != 0) {
			int len = _tcslen(str);
			res = (TCHAR*)calloc(len + 1, sizeof(TCHAR));
			_tcsncpy(res, str, len);
		}
		GlobalUnlock(clip);
		CloseClipboard();
	}

	return res ? res : (TCHAR*)calloc(1, sizeof(TCHAR));
}

int getFontHeight(HWND hWnd) {
	TEXTMETRIC tm;
	HDC hDC = (HDC)GetDC(hWnd);
	HFONT hFont = (HFONT)SendMessage(hWnd, WM_GETFONT, 0, 0);

	HFONT hOldFont = SelectObject(hDC, hFont);
	GetTextMetrics(hDC, &tm);
	SelectObject(hDC, hOldFont);
	ReleaseDC(hWnd, hDC);

	return tm.tmHeight;
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

void ListView_EnsureVisibleCell(HWND hWnd, int rowNo, int colNo) {
	if (rowNo == -1 || colNo == -1)
		return;

	ListView_EnsureVisible(hWnd, rowNo, TRUE);

	RECT rc, rcCell;
	GetClientRect(hWnd, &rc);
	ListView_GetSubItemRect(hWnd, rowNo, colNo, LVIR_BOUNDS, &rcCell);
	if (rcCell.left <= rc.left || rcCell.right >= rc.right) {
		int scrollX = rcCell.left <= rc.left ? rcCell.left - rc.left : rcCell.right >= rc.right ? rcCell.right - rc.right : 0;
		ListView_Scroll(hWnd, scrollX, 0);
	}
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

	if (isFileExists) {
		TCHAR* msg16 = loadString(IDS_OVERWRITE_FILE, path2);
		TCHAR* caption16 = loadString(IDS_CONFIRMATION);
		int rc = MessageBox(hWnd, msg16, caption16, MB_YESNO);
		free(msg16);
		free(caption16);
		if (IDYES != rc)
			return saveFile(path, filter, defExt, hWnd);
	}

	if (isFileExists)
		DeleteFile(path2);

	_tcscpy(path, path2);

	return TRUE;
}

TCHAR* loadString(UINT id, ...) {
	TCHAR templateBuffer[1024];
	if (!LoadString(GetModuleHandle(NULL), id, templateBuffer, sizeof(templateBuffer) / sizeof(TCHAR)))
		return (TCHAR*)calloc(1, sizeof(TCHAR));

	va_list args;
	va_start(args, id);
	TCHAR localBuffer[1024];
	int formattedLength = _vsntprintf(localBuffer, sizeof(localBuffer) / sizeof(TCHAR) - 1, templateBuffer, args);
	va_end(args);

	if (formattedLength < 0) {
		va_start(args, id);
		formattedLength = _vsctprintf(templateBuffer, args);
		va_end(args);

		if (formattedLength < 0)
			formattedLength = 0;
	}

	TCHAR* result = (TCHAR*)calloc((formattedLength + 1), sizeof(TCHAR));
	if (formattedLength < (int)(sizeof(localBuffer) / sizeof(TCHAR))) {
		_tcscpy_s(result, formattedLength + 1, localBuffer);
	} else {
		va_start(args, id);
		_vsntprintf(result, formattedLength + 1, templateBuffer, args);
		va_end(args);
	}

	return result;
}

TCHAR* loadFilters(UINT id) {
	HINSTANCE hInstance = GetModuleHandle(NULL);
	HRSRC hRes = FindResource(hInstance, MAKEINTRESOURCE(id), RT_RCDATA);
	HGLOBAL hData = LoadResource(hInstance, hRes);
	DWORD size8 = SizeofResource(hInstance, hRes);

	if (!hData || size8 == 0)
		return (TCHAR*)calloc (1, sizeof (TCHAR));

	char* res8 = (char*)LockResource(hData);
	DWORD size16 = MultiByteToWideChar(CP_UTF8, 0, res8, size8, NULL, 0);
	TCHAR* res16 = (TCHAR*)calloc (size16 + 1, sizeof (TCHAR));
	MultiByteToWideChar(CP_UTF8, 0, res8, size8, res16, size16);

	return res16;
}