#define SB_VERSION             0
#define SB_TABLE_COUNT         1
#define SB_VIEW_COUNT          2
#define SB_TYPE                3
#define SB_ROW_COUNT           4
#define SB_CURRENT_CELL        5
#define SB_AUXILIARY           6

#define IDD_TABLE              10
#define IDD_ROW                11

#define IDC_MAIN               100
#define IDC_TABLELIST          101
#define IDC_GRID               102
#define IDC_EDITOR             103
#define IDC_STATUSBAR          104
#define IDC_CELL_EDIT          105
#define IDC_ADD_TABLE          107
#define IDC_HEADER_EDIT        1000
#define IDC_DLG_OK             2001
#define IDC_DLG_COLUMNS        2002
#define IDC_DLG_TABLENAME      2003
#define IDC_DLG_LABEL          2004
#define IDC_DLG_EDIT           2100 // iterable + 50
#define IDC_MENU_MAIN          2501
#define IDC_MENU_LIST          2502
#define IDC_MENU_GRID          2503
#define IDC_MENU_BLOB          2504
#define IDC_MENU_EDITOR        2505

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROWS          5001
#define IDM_COPY_COLUMN        5002
#define IDM_PASTE_CELL         5003
#define IDM_DARK_THEME         5010
#define IDM_FILTER_ROW         5011
#define IDM_SEARCH             5012
#define IDM_GRID_LINES         5013
#define IDM_EDITOR             5014
#define IDM_ESCAPE_EXIT        5015
#define IDM_AUTOFILTERS        5016
#define IDM_PREV_CONTROL       5017
#define IDM_NEXT_CONTROL       5018

#define IDM_EDITOR_COPY        5020
#define IDM_EDITOR_PASTE       5021
#define IDM_EDITOR_SAVE        5022
#define IDM_EDITOR_WORDWRAP    5023
#define IDM_EDITOR_AUTOHIDE    5024

#define IDM_HIDE_COLUMN        5030
#define IDM_ADD_TABLE          5031
#define IDM_DROP_TABLE         5032
#define IDM_BLOB               5033
#define IDM_ADD_ROW            5034
#define IDM_DUPLICATE_ROWS     5035
#define IDM_DELETE_ROWS        5036
#define IDM_SEPARATOR1         5037
#define IDM_SEPARATOR2         5038
#define IDM_SAVE_AS            5039
#define IDM_OPEN_DB            5040
#define IDM_EXIT               5050
#define IDM_ABOUT              5053
#define IDM_HOMEPAGE           5054
#define IDM_WIKI               5055
#define IDM_RECENT             5200 // iterable + 50

#define IDI_ICON               6000
#define IDA_ACCEL              7000

#define IDS_ABOUT              8001
#define IDS_NO_DATABASE        8002
#define IDS_CONFIRMATION       8003
#define IDS_DROP_TABLE         8004
#define IDS_DELETE_JOURNAL     8005
#define IDS_OVERWRITE_FILE     8006
#define IDS_SB_TABLE_COUNT     8041
#define IDS_SB_VIEW_COUNT      8042
#define IDS_SB_TABLE           8043
#define IDS_SB_VIEW            8044
#define IDS_SB_ROWS            8045
#define IDS_ERR_DB_OPEN        8050
#define IDS_ERR_FILE_OPEN      8051
#define IDS_ERR_FILE_WRITE     8052
#define IDS_ERR_NO_TABLENAME   8060
#define IDS_ERR_NO_COLUMNS     8061
#define IDS_ERR_NO_ROWID       8062
#define IDS_ERR_CACHE_LIMIT    8063
#define IDS_ERR_EDIT_LIMIT     8064
#define IDS_ERR_ALL_COLUMNS_PK 8065

#define IDR_SQLITE_FILTERS     9001
#define IDR_BLOB_FILTERS       9002

#define WMU_SET_SOURCE          WM_USER + 700
#define WMU_UPDATE_DATA         WM_USER + 702
#define WMU_UPDATE_CACHE        WM_USER + 703
#define WMU_RESET_CACHE         WM_USER + 704
#define WMU_GET_CELL            WM_USER + 705

#define WMU_UPDATE_ROW_COUNT    WM_USER + 710
#define WMU_UPDATE_FILTER_SIZE  WM_USER + 711
#define WMU_SET_HEADER_FILTERS  WM_USER + 712
#define WMU_AUTO_COLUMN_SIZE    WM_USER + 713
#define WMU_SET_CURRENT_CELL    WM_USER + 714
#define WMU_SET_FONT            WM_USER + 715
#define WMU_SET_THEME           WM_USER + 716
#define WMU_HIDE_COLUMN         WM_USER + 717
#define WMU_SHOW_COLUMNS        WM_USER + 718
#define WMU_SORT_COLUMN         WM_USER + 719
#define WMU_EDIT_CELL           WM_USER + 720
#define WMU_UPDATE_CELL         WM_USER + 721
#define WMU_DUPLICATE_ROWS      WM_USER + 725
#define WMU_DELETE_ROWS         WM_USER + 726

#define WMU_OPEN_DB             WM_USER + 750
#define WMU_CLOSE_DB            WM_USER + 751
#define WMU_UPDATE_RECENT       WM_USER + 752
#define WMU_UPDATE_SETTINGS     WM_USER + 753
#define WMU_SB_SET_COUNTS       WM_USER + 754
#define WMU_SB_SET_CURRENT_CELL WM_USER + 755
#define WMU_SB_SET_ROW_COUNTS   WM_USER + 760
#define WMU_AUTOCOMPLETE        WM_USER + 761
#define WMU_SET_SCROLL_HEIGHT   WM_USER + 762
#define WMU_UPDATE_SCROLL_POS   WM_USER + 763
#define WMU_TOGGLE_WORDWRAP     WM_USER + 764
#define WMU_PROCESS_KEY         WM_USER + 765


#define APP_NAME               TEXT("sqlite-x")
#define APP_VERSION            TEXT("1.2.3")
#define WND_CLASS              TEXT("sqlite-x-class")
#ifdef __MINGW64__
#define APP_PLATFORM           64
#else
#define APP_PLATFORM           32
#endif

