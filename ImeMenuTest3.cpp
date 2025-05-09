// License: MIT
#define _CRTDBG_MAP_ALLOC
#include <crtdbg.h>

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <imm.h>
#include "MenuFromImeMenu.h"

BOOL OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam)
{
    return TRUE;
}

void OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
    switch (id)
    {
    case IDOK:
    case IDCANCEL:
        EndDialog(hwnd, id);
        break;
    }
}

void ShowImeMenu(HWND hWnd, HIMC hIMC, BOOL bRight)
{
    PIMEMENUNODE pMenu = CreateImeMenu(hWnd, hIMC, NULL, bRight);
    HMENU hMenu = MenuFromImeMenu(pMenu);
    if (hMenu)
    {
        DWORD dwPos = (DWORD)GetMessagePos();
        INT nCmd = TrackPopupMenuEx(hMenu,
                                    TPM_RETURNCMD | TPM_NONOTIFY |
                                    TPM_LEFTBUTTON | TPM_LEFTALIGN | TPM_TOPALIGN,
                                    LOWORD(dwPos), HIWORD(dwPos),
                                    hWnd, NULL);
        if (nCmd)
        {
            MENUITEMINFO mii = { sizeof(mii), MIIM_DATA };
            GetMenuItemInfo(hMenu, nCmd, FALSE, &mii);

            ImmNotifyIME(hIMC, NI_IMEMENUSELECTED, nCmd, mii.dwItemData);
        }
    }
    else
    {
        MessageBoxA(NULL, "FAILED", NULL, MB_ICONERROR);
    }

    CleanupImeMenus();
    DestroyMenu(hMenu);
}

void OnLButtonDown(HWND hwnd, BOOL fDoubleClick, int x, int y, UINT keyFlags)
{
    POINT pt = { x, y };
    ClientToScreen(hwnd, &pt);

    if (HIMC hIMC = ImmGetContext(hwnd))
    {
        ShowImeMenu(hwnd, hIMC, FALSE);
        ImmReleaseContext(hwnd, hIMC);
    }
}

void OnRButtonDown(HWND hwnd, BOOL fDoubleClick, int x, int y, UINT keyFlags)
{
    POINT pt = { x, y };
    ClientToScreen(hwnd, &pt);

    if (HIMC hIMC = ImmGetContext(hwnd))
    {
        ShowImeMenu(hwnd, hIMC, TRUE);
        ImmReleaseContext(hwnd, hIMC);
    }
}

INT_PTR CALLBACK
DialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        HANDLE_MSG(hwnd, WM_INITDIALOG, OnInitDialog);
        HANDLE_MSG(hwnd, WM_COMMAND, OnCommand);
        HANDLE_MSG(hwnd, WM_LBUTTONDOWN, OnLButtonDown);
        HANDLE_MSG(hwnd, WM_RBUTTONDOWN, OnRButtonDown);
    }
    return 0;
}

INT WINAPI
WinMain(HINSTANCE   hInstance,
        HINSTANCE   hPrevInstance,
        LPSTR       lpCmdLine,
        INT         nCmdShow)
{
    InitCommonControls();
    DialogBox(hInstance, MAKEINTRESOURCE(1), NULL, DialogProc);
    return 0;
}
