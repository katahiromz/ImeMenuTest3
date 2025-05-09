// License: MIT
#define _CRTDBG_MAP_ALLOC
#include <crtdbg.h>

#include <windows.h>
#include <imm.h>
#include "MenuFromImeMenu.h"

#if 0
/* fType of IMEMENUITEMINFO structure */
#define IMFT_RADIOCHECK         0x00001
#define IMFT_SEPARATOR          0x00002
#define IMFT_SUBMENU            0x00004

#define IMEMENUITEM_STRING_SIZE 80

typedef struct tagIMEMENUITEMINFOA {
    UINT        cbSize;
    UINT        fType;
    UINT        fState;
    UINT        wID;
    HBITMAP     hbmpChecked;
    HBITMAP     hbmpUnchecked;
    DWORD       dwItemData;
    CHAR        szString[IMEMENUITEM_STRING_SIZE];
    HBITMAP     hbmpItem;
} IMEMENUITEMINFOA, *PIMEMENUITEMINFOA, NEAR *NPIMEMENUITEMINFOA, FAR *LPIMEMENUITEMINFOA;

typedef struct tagIMEMENUITEMINFOW {
    UINT        cbSize;
    UINT        fType;
    UINT        fState;
    UINT        wID;
    HBITMAP     hbmpChecked;
    HBITMAP     hbmpUnchecked;
    DWORD       dwItemData;
    WCHAR       szString[IMEMENUITEM_STRING_SIZE];
    HBITMAP     hbmpItem;
} IMEMENUITEMINFOW, *PIMEMENUITEMINFOW, NEAR *NPIMEMENUITEMINFOW, FAR *LPIMEMENUITEMINFOW;
#endif

BOOL MakeImeMenu(IN HMENU hMenu, IN const IMEMENUNODE *pMenu);

static void *Alloc(size_t size)
{
    return calloc(1, size);
}

static void Free(void *ptr)
{
    free(ptr);
}

PIMEMENUNODE g_pMenuList = NULL;
INT g_nNextMenuID = 0;

VOID FillImeMenuItem(OUT LPMENUITEMINFO pItemInfo, IN const IMEMENUITEM *pItem)
{
    ZeroMemory(pItemInfo, sizeof(MENUITEMINFO));
    pItemInfo->cbSize = sizeof(MENUITEMINFO);
    pItemInfo->fMask = MIIM_ID | MIIM_STATE | MIIM_DATA;
    pItemInfo->wID = pItem->m_Info.wID;
    pItemInfo->fState = pItem->m_Info.fState;
    pItemInfo->dwItemData = pItem->m_Info.dwItemData;

    if (pItem->m_Info.fType)
    {
        pItemInfo->fMask |= MIIM_FTYPE;
        pItemInfo->fType = 0;
        if (pItem->m_Info.fType & IMFT_RADIOCHECK)
            pItemInfo->fType |= MFT_RADIOCHECK;
        if (pItem->m_Info.fType & IMFT_SEPARATOR)
            pItemInfo->fType |= MFT_SEPARATOR;
    }

    if (pItem->m_Info.fType & IMFT_SUBMENU)
    {
        pItemInfo->fMask |= MIIM_SUBMENU;
        pItemInfo->hSubMenu = CreatePopupMenu();
        MakeImeMenu(pItemInfo->hSubMenu, pItem->m_pSubMenu);
    }

    if (pItem->m_Info.hbmpChecked && pItem->m_Info.hbmpUnchecked)
    {
        pItemInfo->fMask |= MIIM_CHECKMARKS;
        pItemInfo->hbmpChecked = pItem->m_Info.hbmpChecked;
        pItemInfo->hbmpUnchecked = pItem->m_Info.hbmpUnchecked;
    }

    if (pItem->m_Info.hbmpItem)
    {
        pItemInfo->fMask |= MIIM_BITMAP;
        pItemInfo->hbmpItem = pItem->m_Info.hbmpItem;
    }

    PCTSTR szString = pItem->m_Info.szString;
    if (szString && szString[0])
    {
        pItemInfo->fMask |= MIIM_STRING;
        pItemInfo->dwTypeData = (PTSTR)szString;
        pItemInfo->cch = lstrlen(szString);
    }
}

BOOL MakeImeMenu(IN HMENU hMenu, IN const IMEMENUNODE *pMenu)
{
    if (!pMenu->m_nItems)
        return FALSE;

    for (INT iItem = 0; iItem < pMenu->m_nItems; ++iItem)
    {
        MENUITEMINFO mi = { sizeof(mi) };
        FillImeMenuItem(&mi, &pMenu->m_Items[iItem]);
        InsertMenuItem(hMenu, iItem, TRUE, &mi);
    }

    return TRUE;
}

void AddImeMenuNode(PIMEMENUNODE pMenu)
{
    if (!g_pMenuList)
    {
        g_pMenuList = pMenu;
        return;
    }

    pMenu->m_pNext = g_pMenuList;
    g_pMenuList = pMenu;
}

PIMEMENUNODE AllocateImeMenu(DWORD itemCount)
{
    SIZE_T cbMenu = sizeof(IMEMENUNODE) + (itemCount - 1) * sizeof(IMEMENUITEM);
    PIMEMENUNODE pMenu = (PIMEMENUNODE)Alloc(cbMenu);
    if (!pMenu)
        return NULL;
    AddImeMenuNode(pMenu);
    pMenu->m_nItems = itemCount;
    return pMenu;
}

void GetImeMenuItem(IN HWND hWnd, IN HIMC hIMC, OUT PIMEMENUITEMINFO lpImeParentMenu, IN BOOL bRightMenu, OUT PIMEMENUITEM pItem)
{
    ZeroMemory(pItem, sizeof(IMEMENUITEM));
    pItem->m_Info = *lpImeParentMenu;

    if (lpImeParentMenu->fType & IMFT_SUBMENU)
        pItem->m_pSubMenu = CreateImeMenu(hWnd, hIMC, lpImeParentMenu, bRightMenu);

    pItem->m_nRealID = pItem->m_Info.wID;
    pItem->m_Info.wID = ID_STARTIMEMENU + g_nNextMenuID++;
}

#ifdef USE_CUSTOM
#include "CustomImeMenu.cpp"
#endif

DWORD
GetImeMenuItems(
    IN HIMC hIMC,
    IN DWORD dwFlags,
    IN DWORD dwType,
    OUT PIMEMENUITEMINFO lpImeParentMenu OPTIONAL,
    OUT PIMEMENUITEMINFO pImeMenuItems OPTIONAL,
    IN DWORD cbItems)
{
#ifdef USE_CUSTOM
    return ImeGetImeMenuItems(hIMC, dwFlags, dwType, lpImeParentMenu, pImeMenuItems, cbItems);
#else
    return ImmGetImeMenuItems(hIMC, dwFlags, dwType, lpImeParentMenu, pImeMenuItems, cbItems);
#endif
}

PIMEMENUNODE
CreateImeMenu(IN HWND hWnd, IN HIMC hIMC, OUT PIMEMENUITEMINFO lpImeParentMenu OPTIONAL, IN BOOL bRightMenu)
{
    DWORD itemCount = GetImeMenuItems(hIMC, bRightMenu, 0x3F, lpImeParentMenu, NULL, 0);
    if (!itemCount)
        return NULL;

    PIMEMENUNODE pMenu = AllocateImeMenu(itemCount);
    if (!pMenu)
        return NULL;

    DWORD cbItems = sizeof(IMEMENUITEMINFO) * itemCount;
    PIMEMENUITEMINFO pImeMenuItems = (PIMEMENUITEMINFO)Alloc(cbItems);
    if (!pImeMenuItems)
        return NULL;

    DWORD newItemCount = GetImeMenuItems(hIMC, bRightMenu, 0x3F, lpImeParentMenu, pImeMenuItems, cbItems);
    if (!newItemCount)
    {
        Free(pImeMenuItems);
        return NULL;
    }

    PIMEMENUITEM pItems = pMenu->m_Items;
    for (DWORD iItem = 0; iItem < newItemCount; ++iItem)
    {
        GetImeMenuItem(hWnd, hIMC, &pImeMenuItems[iItem], bRightMenu, &pItems[iItem]);
    }

    Free(pImeMenuItems);
    return pMenu;
}

HMENU MenuFromImeMenu(const IMEMENUNODE *pMenu)
{
    if (!pMenu)
        return NULL;
    HMENU hMenu = CreatePopupMenu();
    if (!MakeImeMenu(hMenu, pMenu))
    {
        DestroyMenu(hMenu);
        return NULL;
    }
    return hMenu;
}

BOOL FreeMenuNode(PIMEMENUNODE pMenuNode)
{
    if (!pMenuNode)
        return FALSE;

    for (INT iItem = 0; iItem < pMenuNode->m_nItems; ++iItem)
    {
        PIMEMENUITEM pItem = &pMenuNode->m_Items[iItem];
        if (pItem->m_Info.hbmpChecked)
            DeleteObject(pItem->m_Info.hbmpChecked);
        if (pItem->m_Info.hbmpUnchecked)
            DeleteObject(pItem->m_Info.hbmpUnchecked);
        if (pItem->m_Info.hbmpItem)
            DeleteObject(pItem->m_Info.hbmpItem);
    }

    Free(pMenuNode);
    return TRUE;
}

VOID CleanupImeMenus(VOID)
{
    if (!g_pMenuList)
        return;

    PIMEMENUNODE pNext;
    for (PIMEMENUNODE pNode = g_pMenuList; pNode; pNode = pNext)
    {
        pNext = pNode->m_pNext;
        FreeMenuNode(pNode);
    }

    g_pMenuList = NULL;
    g_nNextMenuID = 0;
}

INT RemapImeMenuID(IN const IMEMENUNODE *pMenu, INT nID)
{
    if (!pMenu || !pMenu->m_nItems || nID < ID_STARTIMEMENU)
        return 0;

    for (INT iItem = 0; iItem < pMenu->m_nItems; ++iItem)
    {
        const IMEMENUITEM *pItem = &pMenu->m_Items[iItem];
        if (pItem->m_Info.wID == nID)
            return pItem->m_nRealID;

        if (pItem->m_pSubMenu)
        {
            INT nRealID = RemapImeMenuID(pItem->m_pSubMenu, nID);
            if (nRealID)
                return nRealID;
        }
    }

    return 0;
}
