/* For availability of getting combobox info with _WIN32_WINNT defined as 0x0400 */
/* Extracted from WinUser.h */

#define CB_GETCOMBOBOXINFO          0x0164

/*
 * Combobox information
 */
typedef struct tagCOMBOBOXINFO
{
    DWORD cbSize;
    RECT rcItem;
    RECT rcButton;
    DWORD stateButton;
    HWND hwndCombo;
    HWND hwndItem;
    HWND hwndList;
} COMBOBOXINFO, *PCOMBOBOXINFO, *LPCOMBOBOXINFO;