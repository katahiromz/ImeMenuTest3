// License: MIT
#define _CRTDBG_MAP_ALLOC
#include <crtdbg.h>

#include <windows.h>
#include <imm.h>
#include <stdio.h>
#include <assert.h>
#include <strsafe.h>
#include "ImeMenu.h"

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

LPVOID ImmLocalAlloc(_In_ DWORD dwFlags, _In_ DWORD dwBytes)
{
    return calloc(1, dwBytes);
}
#define ImmLocalFree(ptr) free(ptr)

#define IMEMENUINFO_BUFFER_SIZE 0x20000

typedef struct tagIMEMENUINFO
{
    DWORD dwVersion;           // Must be 1
    DWORD dwBufferSize;        // Always 0x20000 (128KB)
    DWORD dwFlags;             // dwFlags of ImmGetImeMenuItems
    DWORD dwType;              // dwType of ImmGetImeMenuItems
    DWORD_PTR dwParentOffset;  // Relative offset to parent menu data (or pointer)
    DWORD_PTR dwItemsOffset;   // Relative offset to menu items (or pointer)
    DWORD dwCount;             // # of items
    DWORD_PTR dwSubMenuOffset; // Relative offset to the sub-menu (or pointer)
    DWORD_PTR dwEndOffset;     // Offset to the bottom of this data (or pointer)
} IMEMENUINFO, *PIMEMENUINFO;

typedef struct tagBITMAPNODE
{
    DWORD_PTR dwNext;       // Relative offset to next node (or pointer)
    HBITMAP hbmpCached;     // Cached HBITMAP
    DWORD_PTR dibBitsPtr;   // Relative offset to DIB pixel data (or pointer)
    BITMAPINFOHEADER bmih;  // BITMAPINFOHEADER
} BITMAPNODE, *PBITMAPNODE;

#define IMEMENUINFO_MAPPING_NAME  L"ImmMenuInfo"
#define IMEMENUINFO_BUFFER_SIZE   0x20000 // 0x20000 (128KB)

BOOL MakeImeMenu(IN HMENU hMenu, IN const IMEMENUNODE *pMenu);

void *Alloc(size_t size)
{
    return calloc(1, size);
}

void Free(void *ptr)
{
    free(ptr);
}

PIMEMENUNODE g_pMenuList = NULL;
INT g_nNextMenuID = 0;

DWORD
GetImeMenuItemsBase(
    IN HIMC hIMC,
    IN DWORD dwFlags,
    IN DWORD dwType,
    IN OUT PIMEMENUITEMINFO lpImeParentMenu OPTIONAL,
    OUT PIMEMENUITEMINFO pImeMenuItems OPTIONAL,
    IN DWORD cbItems)
{
#ifdef USE_CUSTOM
    return ImeGetImeMenuItems(hIMC, dwFlags, dwType, lpImeParentMenu, pImeMenuItems, cbItems);
#else
    return ImmGetImeMenuItems(hIMC, dwFlags, dwType, lpImeParentMenu, pImeMenuItems, cbItems);
#endif
}

static PVOID
Imm32WriteHBitmapToNode(
    _In_  HDC          hDC,
    _In_  HBITMAP      hbmp,
    _In_  PBITMAPNODE  pNode,
    _In_  PIMEMENUINFO pView)
{
    BITMAPINFOHEADER *pBmih;
    DWORD nodeExtraSize;
    PVOID dibBitsPtr;
    HBITMAP hbmpTemp;
    HGDIOBJ hbmpOld;
    DWORD totalNeeded;
    int scanlines;

    if (!hbmp)
        return NULL;

    pBmih = &pNode->bmih;
    pBmih->biSize     = sizeof(BITMAPINFOHEADER);
    pBmih->biBitCount = 0;
    if (!GetDIBits(hDC, hbmp, 0, 0, NULL, (BITMAPINFO*)pBmih, DIB_RGB_COLORS))
        return NULL;

    switch (pBmih->biBitCount)
    {
        case 16: case 32:
            nodeExtraSize = sizeof(BITMAPINFOHEADER) + 3 * sizeof(DWORD);
            break;
        case 24:
            nodeExtraSize = sizeof(BITMAPINFOHEADER);
            break;
        default:
            nodeExtraSize = sizeof(BITMAPINFOHEADER);
            nodeExtraSize += (1 << pBmih->biBitCount) * sizeof(RGBQUAD);
            break;
    }

    totalNeeded = pBmih->biSizeImage + nodeExtraSize;
    if ((PBYTE)pBmih + totalNeeded + 4 > (PBYTE)pView + pView->dwBufferSize)
        return NULL;

    dibBitsPtr = (PBYTE)pBmih + nodeExtraSize;
    pNode->dibBitsPtr = (DWORD_PTR)dibBitsPtr; // Absolute address; convert to relative later

    hbmpTemp = CreateCompatibleBitmap(hDC, pBmih->biWidth, pBmih->biHeight);
    if (!hbmpTemp)
        return NULL;

    hbmpOld = SelectObject(hDC, hbmpTemp);
    scanlines = GetDIBits(hDC, hbmp, 0, pBmih->biHeight, dibBitsPtr, (PBITMAPINFO)pBmih,
                          DIB_RGB_COLORS);
    SelectObject(hDC, hbmpOld);
    DeleteObject(hbmpTemp);

    if (!scanlines)
        return NULL;

    return (PBYTE)pBmih + nodeExtraSize + pBmih->biSizeImage;
}

void
Imm32InitImeMenuView(
    PIMEMENUINFO pView,
    DWORD dwFlags,
    DWORD dwType,
    PIMEMENUITEMINFO lpImeParentMenu,
    PIMEMENUITEMINFO lpImeMenuItems,
    DWORD dwSize)
{
    ZeroMemory(pView, sizeof(*pView));
    pView->dwVersion    = 1;
    pView->dwBufferSize = IMEMENUINFO_BUFFER_SIZE;
    pView->dwFlags      = dwFlags;
    pView->dwType       = dwType;
    pView->dwCount      = dwSize / sizeof(IMEMENUITEMINFOW);

    if ((dwSize == 0 || !lpImeMenuItems) && !lpImeParentMenu)
    {
        // No parents nor items
    }
    else if (lpImeParentMenu)
    {
        // There is parent
        PIMEMENUITEMINFOW pParentSlot = (PIMEMENUITEMINFOW)(pView + 1);
        RtlCopyMemory(pParentSlot, lpImeParentMenu, sizeof(IMEMENUITEMINFOW));
        // Clear parent bitmaps (unused)
        pParentSlot->hbmpChecked   = NULL;
        pParentSlot->hbmpUnchecked = NULL;
        pParentSlot->hbmpItem      = NULL;
        // Offset (relative)
        pView->dwParentOffset = sizeof(*pView);
        pView->dwItemsOffset  = sizeof(*pView) + sizeof(IMEMENUITEMINFOW);
    }
    else
    {
        // No parent but some items
        pView->dwParentOffset = 0;
        pView->dwItemsOffset  = sizeof(*pView);
    }
}

static PBITMAPNODE
Imm32SerializeBitmap(
    _In_ PIMEMENUINFO pView,
    _In_ HBITMAP      hbmp)
{
    PBITMAPNODE pListHead, pNode;
    PVOID pNewEnd;
    HDC hDC;

    // Check bitmap caches
    pNode = (PBITMAPNODE)pView->dwSubMenuOffset;
    while (pNode)
    {
        if (pNode->hbmpCached == hbmp)
            return pNode; // Cache hit
        pNode = (PBITMAPNODE)pNode->dwNext;
    }

    // Boundary check
    pNode = (PBITMAPNODE)pView->dwEndOffset;
    if (!pNode || (PBYTE)pNode < (PBYTE)pView ||
        (PBYTE)pNode >= (PBYTE)pView + pView->dwBufferSize)
    {
        return NULL; // Insufficient or incorrect space
    }

    hDC = GetDC(GetDesktopWindow());
    if (!hDC)
        return NULL;

    pListHead = (PBITMAPNODE)pView->dwSubMenuOffset;
    pNode->dwNext = (ULONG_PTR)pListHead;

    pNewEnd = Imm32WriteHBitmapToNode(hDC, hbmp, pNode, pView);

    ReleaseDC(GetDesktopWindow(), hDC);

    if (!pNewEnd) // Failure
    {
        pView->dwEndOffset = (ULONG_PTR)pNode;
        pView->dwSubMenuOffset = pNode->dwNext;
        return NULL;
    }

    pNode->hbmpCached = hbmp;
    pNode->dwNext = (ULONG_PTR)pListHead;
    pView->dwSubMenuOffset = (DWORD)pNode;
    pView->dwEndOffset     = (ULONG_PTR)pNewEnd;

    return pNode;
}

static HBITMAP
Imm32DeserializeBitmap(
    _Inout_ PBITMAPNODE pNode)
{
    if (!pNode || !pNode->dibBitsPtr)
        return NULL;

    if (pNode->hbmpCached)
        return pNode->hbmpCached;

    HDC hDC = GetDC(GetDesktopWindow());
    if (!hDC)
        return NULL;

    HDC hCompatDC = CreateCompatibleDC(hDC);
    if (!hCompatDC)
    {
        ReleaseDC(GetDesktopWindow(), hDC);
        return NULL;
    }

    HBITMAP hbmp = CreateDIBitmap(hCompatDC, &pNode->bmih, CBM_INIT, (PVOID)pNode->dibBitsPtr,
                                  (PBITMAPINFO)&pNode->bmih, DIB_RGB_COLORS);
    pNode->hbmpCached = hbmp;

    DeleteDC(hCompatDC);
    ReleaseDC(GetDesktopWindow(), hDC);

    return hbmp;
}

BOOL Imm32SerializeImeMenu(HIMC hIMC, PIMEMENUINFO pView)
{
    PIMEMENUITEMINFOW pParent, pItems, pTempBuf;
    DWORD i, dwCount, dwTempBufSize;
    BOOL ret = FALSE;
    PBITMAPNODE pNode;

    if (pView->dwVersion != 1)
        return FALSE;

    if (pView->dwParentOffset)
    {
        // Convert relative to absolute
        pParent = (PIMEMENUITEMINFOW)((PBYTE)pView + pView->dwParentOffset);
        if ((PBYTE)pParent < (PBYTE)pView || (PBYTE)pParent >= (PBYTE)pView + pView->dwBufferSize)
            return FALSE;
        pView->dwParentOffset = (ULONG_PTR)pParent;
    }
    else
    {
        pParent = NULL;
    }

    if (pView->dwItemsOffset)
    {
        // Convert relative to absolute
        pItems = (PIMEMENUITEMINFOW)((PBYTE)pView + pView->dwItemsOffset);
        if ((PBYTE)pItems < (PBYTE)pView || (PBYTE)pItems >= (PBYTE)pView + pView->dwBufferSize)
            return FALSE;
        pView->dwItemsOffset = (ULONG_PTR)pItems;
    }
    else
    {
        pItems = NULL;
    }

    // Allocate items buffer
    if (pView->dwCount > 0)
    {
        dwTempBufSize = pView->dwCount * sizeof(IMEMENUITEMINFOW);
        pTempBuf = (PIMEMENUITEMINFOW)ImmLocalAlloc(0, dwTempBufSize);
        if (!pTempBuf)
            return FALSE;
    }
    else
    {
        pTempBuf = NULL;
        dwTempBufSize = 0;
    }

#if 1
    dwCount = GetImeMenuItemsBase(hIMC, pView->dwFlags, pView->dwType, pParent,
                                  pTempBuf, dwTempBufSize);
#else
    dwCount = ImmGetImeMenuItemsW(hIMC, pView->dwFlags, pView->dwType, pParent,
                                  pTempBuf, dwTempBufSize);
#endif
    pView->dwCount = dwCount;

    if (dwCount == 0) // No items?
    {
        if (pTempBuf)
            ImmLocalFree(pTempBuf);
        ret = TRUE;
        goto ConvertBack;
    }

    if (!pTempBuf)
    {
        ret = TRUE;
        goto ConvertBack;
    }

    pView->dwSubMenuOffset = 0;
    pView->dwEndOffset = (ULONG_PTR)((PBYTE)pView + (dwCount + 1) * sizeof(IMEMENUITEMINFOW));

    // Serialize items
    for (i = 0; i < dwCount; i++)
    {
        PIMEMENUITEMINFOW pSrc  = &pTempBuf[i];
        PIMEMENUITEMINFOW pDest = pItems + i;
        *pDest = *pSrc;

        // Serialize hbmpChecked
        if (pSrc->hbmpChecked)
        {
            pNode = Imm32SerializeBitmap(pView, pSrc->hbmpChecked);
            if (!pNode)
                goto FreeAndCleanup;
            pDest->hbmpChecked = (HBITMAP)(ULONG_PTR)pNode; // Absolute pointer
        }
        else
        {
            pDest->hbmpChecked = NULL;
        }

        // Serialize hbmpUnchecked
        if (pSrc->hbmpUnchecked)
        {
            pNode = Imm32SerializeBitmap(pView, pSrc->hbmpUnchecked);
            if (!pNode)
                goto FreeAndCleanup;
            pDest->hbmpUnchecked = (HBITMAP)(ULONG_PTR)pNode;
        }
        else
        {
            pDest->hbmpUnchecked = NULL;
        }

        // Serialize hbmpItem
        if (pSrc->hbmpItem)
        {
            pNode = Imm32SerializeBitmap(pView, pSrc->hbmpItem);
            if (!pNode)
                goto FreeAndCleanup;
            pDest->hbmpItem = (HBITMAP)(ULONG_PTR)pNode;
        }
        else
        {
            pDest->hbmpItem = NULL;
        }
    }

    ret = TRUE;
FreeAndCleanup:
    ImmLocalFree(pTempBuf);

ConvertBack:
    // Convert dwItemsOffset to relative
    if (pView->dwItemsOffset)
        pView->dwItemsOffset = pView->dwItemsOffset - (ULONG_PTR)pView;

    if (pView->dwCount > 0 && pItems)
    {
        // Convert bitmaps to relative
        for (i = 0; i < pView->dwCount; i++)
        {
            PIMEMENUITEMINFOW pItem = pItems + i;
            if (pItem->hbmpChecked)
                pItem->hbmpChecked = (HBITMAP)((PBYTE)pItem->hbmpChecked - (PBYTE)pView);
            if (pItem->hbmpUnchecked)
                pItem->hbmpUnchecked = (HBITMAP)((PBYTE)pItem->hbmpUnchecked - (PBYTE)pView);
            if (pItem->hbmpItem)
                pItem->hbmpItem = (HBITMAP)((PBYTE)pItem->hbmpItem - (PBYTE)pView);
        }
    }

    // Convert dwSubMenuOffset to relative
    if (pView->dwSubMenuOffset)
    {
        PBITMAPNODE pCur = (PBITMAPNODE)pView->dwSubMenuOffset;
        pView->dwSubMenuOffset = (ULONG_PTR)pCur - (ULONG_PTR)pView;

        while (pCur)
        {
            PBITMAPNODE pNext = (PBITMAPNODE)pCur->dwNext;

            if (pCur->dibBitsPtr)
                pCur->dibBitsPtr = pCur->dibBitsPtr - (ULONG_PTR)pView;
            if (pCur->dwNext)
                pCur->dwNext = (ULONG_PTR)pNext - (ULONG_PTR)pView;

            pCur = pNext;
        }
    }

    // Convert dwParentOffset to relative
    if (pView->dwParentOffset && (PBYTE)pView->dwParentOffset >= (PBYTE)pView)
        pView->dwParentOffset = pView->dwParentOffset - (ULONG_PTR)pView;

    return ret;
}

DWORD WINAPI
Imm32DeserializeImeMenu(
    _In_ PIMEMENUINFO pView,
    _Out_opt_ PIMEMENUITEMINFOW lpImeMenuItems,
    _In_ DWORD dwSize)
{
    DWORD i, dwCount;
    PBYTE pViewBase = (PBYTE)pView;
    PBITMAPNODE pNode;
    PIMEMENUITEMINFOW pItemsBase;

    dwCount = pView->dwCount;

    if (!lpImeMenuItems)
        return dwCount;

    if (dwCount == 0)
        return 0;

    if (pView->dwSubMenuOffset)
    {
        PBITMAPNODE pCur = (PBITMAPNODE)(pViewBase + pView->dwSubMenuOffset);
        pView->dwSubMenuOffset = (ULONG_PTR)pCur;

        while (pCur)
        {
            PBITMAPNODE pNextCur;

            pCur->hbmpCached = NULL;

            // dibBitsPtr: relative to absolute
            if (pCur->dibBitsPtr)
                pCur->dibBitsPtr = (ULONG_PTR)(pViewBase + pCur->dibBitsPtr);

            // dwNext: relative to absolute
            if (pCur->dwNext)
            {
                pNextCur = (PBITMAPNODE)(pViewBase + pCur->dwNext);
                if ((PBYTE)pNextCur < pViewBase ||
                    (PBYTE)pNextCur >= pViewBase + pView->dwBufferSize)
                    return 0;
                pCur->dwNext = (ULONG_PTR)pNextCur;
                pCur = pNextCur;
            }
            else
            {
                pCur->dwNext = 0;
                pCur = NULL;
            }
        }
    }

    // dwItemsOffset: relative to absolute
    if (pView->dwItemsOffset)
    {
        pItemsBase = (PIMEMENUITEMINFOW)(pViewBase + pView->dwItemsOffset);
        // Boundary check
        if ((PBYTE)pItemsBase < pViewBase || (PBYTE)pItemsBase >= pViewBase + pView->dwBufferSize)
            return 0;
        pView->dwItemsOffset = (ULONG_PTR)pItemsBase;
    }
    else
    {
        return 0;
    }

    // Bitmap offsets: relative to absolute
    for (i = 0; i < dwCount; i++)
    {
        PIMEMENUITEMINFOW pItem =
            (PIMEMENUITEMINFOW)((PBYTE)pView->dwItemsOffset + i * sizeof(IMEMENUITEMINFOW));

        // hbmpChecked
        if (pItem->hbmpChecked)
        {
            PBITMAPNODE pN = (PBITMAPNODE)(pViewBase + (DWORD)(ULONG_PTR)pItem->hbmpChecked);
            if ((PBYTE)pN < pViewBase || (PBYTE)pN >= pViewBase + pView->dwBufferSize)
                return 0;
            pItem->hbmpChecked = (HBITMAP)(ULONG_PTR)pN;
        }

        // hbmpUnchecked
        if (pItem->hbmpUnchecked)
        {
            PBITMAPNODE pN = (PBITMAPNODE)(pViewBase + (DWORD)(ULONG_PTR)pItem->hbmpUnchecked);
            if ((PBYTE)pN < pViewBase || (PBYTE)pN >= pViewBase + pView->dwBufferSize)
                return 0;
            pItem->hbmpUnchecked = (HBITMAP)(ULONG_PTR)pN;
        }

        // hbmpItem
        if (pItem->hbmpItem)
        {
            PBITMAPNODE pN = (PBITMAPNODE)(pViewBase + (DWORD)(ULONG_PTR)pItem->hbmpItem);
            if ((PBYTE)pN < pViewBase || (PBYTE)pN >= pViewBase + pView->dwBufferSize)
                return 0;
            pItem->hbmpItem = (HBITMAP)(ULONG_PTR)pN;
        }
    }

    // De-serialize items
    for (i = 0; i < dwCount; i++)
    {
        PIMEMENUITEMINFOW pSrc =
            (PIMEMENUITEMINFOW)((PBYTE)pView->dwItemsOffset + i * sizeof(IMEMENUITEMINFOW));
        PIMEMENUITEMINFOW pDst = &lpImeMenuItems[i];

        // Copy scalar fields (excluding HBITMAP fields)
        pDst->cbSize    = pSrc->cbSize;
        pDst->fType     = pSrc->fType;
        pDst->fState    = pSrc->fState;
        pDst->wID       = pSrc->wID;
        pDst->dwItemData = pSrc->dwItemData;

        // Copy szString
        StringCbCopyW(pDst->szString, sizeof(pDst->szString), pSrc->szString);

        // De-serialize hbmpChecked
        pNode = (PBITMAPNODE)(ULONG_PTR)pSrc->hbmpChecked;
        pDst->hbmpChecked = Imm32DeserializeBitmap(pNode);

        // De-serialize hbmpUnchecked
        pNode = (PBITMAPNODE)(ULONG_PTR)pSrc->hbmpUnchecked;
        pDst->hbmpUnchecked = Imm32DeserializeBitmap(pNode);

        // De-serialize hbmpItem
        pNode = (PBITMAPNODE)(ULONG_PTR)pSrc->hbmpItem;
        pDst->hbmpItem = Imm32DeserializeBitmap(pNode);
    }

    return dwCount;
}

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
#include "CustomImeMenu.c"
#endif

DWORD
GetImeMenuItems(
    IN HIMC hIMC,
    IN DWORD dwFlags,
    IN DWORD dwType,
    IN OUT PIMEMENUITEMINFO lpImeParentMenu OPTIONAL,
    OUT PIMEMENUITEMINFO lpImeMenuItems OPTIONAL,
    IN DWORD dwSize)
{
#ifdef DO_TRANSPORT
    PBYTE pb = Alloc(IMEMENUINFO_BUFFER_SIZE);
    if (!pb)
    {
        assert(0);
        return 0;
    }
    Imm32InitImeMenuView((PIMEMENUINFO)pb, dwFlags, dwType, lpImeParentMenu, lpImeMenuItems, dwSize);
    DWORD dwItemCount = Imm32SerializeImeMenu(hIMC, (PIMEMENUINFO)pb);
    if (!dwItemCount)
    {
        assert(0);
        return 0;
    }
    ZeroMemory(lpImeMenuItems, dwSize);
    DWORD ret = Imm32DeserializeImeMenu((PIMEMENUINFO)pb, lpImeMenuItems, dwSize);
    Free(pb);
    return ret;
#else
    return GetImeMenuItemsBase(hIMC, dwFlags, dwType, lpImeParentMenu, pImeMenuItems, dwSize);
#endif
}

PIMEMENUNODE
CreateImeMenu(IN HWND hWnd, IN HIMC hIMC, IN OUT PIMEMENUITEMINFO lpImeParentMenu OPTIONAL, IN BOOL bRightMenu)
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

INT GetRealImeMenuID(IN const IMEMENUNODE *pMenu, INT nID)
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
            INT nRealID = GetRealImeMenuID(pItem->m_pSubMenu, nID);
            if (nRealID)
                return nRealID;
        }
    }

    return 0;
}
