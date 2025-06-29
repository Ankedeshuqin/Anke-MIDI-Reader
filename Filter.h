/* Filter flags */
#define FILT_CHECKED 1
#define FILT_AVAILABLE 0x10000

typedef struct filtstate {
    DWORD dwFiltChn[17];
    DWORD dwFiltEvtType[10];
    DWORD dwFiltTrk[1];
} FILTERSTATES;

int EvtGetTrkIndex(EVENT *pevt) {
    return pevt->wTrk;
}

int EvtGetChnIndex(EVENT *pevt) {
    if(pevt->bStatus < 0xF0)
        return (pevt->bStatus & 0xF) + 1;
    return 0;
}

int EvtGetEvtTypeIndex(EVENT *pevt) {
    switch(pevt->bStatus >> 4) {
    case 0x8:
        return 0;
    case 0x9:
        if(pevt->bData2 == 0)
            return 1;
        return 2;
    case 0xA:
    case 0xB:
    case 0xC:
    case 0xD:
    case 0xE:
        return (pevt->bStatus >> 4) - 7;
    }
    if(pevt->bStatus == 0xF0 || pevt->bStatus == 0xF7)
        return 8;
    return 9;
}

BOOL IsEvtUnfiltered(EVENT *pevt, FILTERSTATES *pfiltstate) {
    UINT uTrkIndex, uChnIndex, uEvtTypeIndex;

    uTrkIndex = EvtGetTrkIndex(pevt);
    uChnIndex = EvtGetChnIndex(pevt);
    uEvtTypeIndex = EvtGetEvtTypeIndex(pevt);
    if(
        !(pfiltstate->dwFiltTrk[uTrkIndex] & FILT_CHECKED) ||
        !(pfiltstate->dwFiltChn[uChnIndex] & FILT_CHECKED) ||
        !(pfiltstate->dwFiltEvtType[uEvtTypeIndex] & FILT_CHECKED)
    ) {
        return FALSE;
    }
    return TRUE;
}

void MakeCBFiltLists(MIDIFILE *pmf, FILTERSTATES *pfiltstate, HWND hwndCBFiltTrk, HWND hwndCBFiltChn, HWND hwndCBFiltEvtType) {
    EVENT *pevtCur;
    UINT uTrkIndex, uChnIndex, uEvtTypeIndex;
    UINT uFiltIndex;

    int i, cCBFiltItem;
    WCHAR szBuf[128];
    LPWSTR lpszBuf;

    for(uFiltIndex = 0; uFiltIndex < pmf->cTrk; uFiltIndex++)
        pfiltstate->dwFiltTrk[uFiltIndex] &= ~FILT_AVAILABLE;
    for(uFiltIndex = 0; uFiltIndex < 17; uFiltIndex++)
        pfiltstate->dwFiltChn[uFiltIndex] &= ~FILT_AVAILABLE;
    for(uFiltIndex = 0; uFiltIndex < 10; uFiltIndex++)
        pfiltstate->dwFiltEvtType[uFiltIndex] &= ~FILT_AVAILABLE;

    pevtCur = pmf->pevtHead;
    while(pevtCur) {
        uTrkIndex = EvtGetTrkIndex(pevtCur);
        uChnIndex = EvtGetChnIndex(pevtCur);
        uEvtTypeIndex = EvtGetEvtTypeIndex(pevtCur);
        if(!IsEvtUnfiltered(pevtCur, pfiltstate)) {
            if(
                !(pfiltstate->dwFiltTrk[uTrkIndex] & FILT_CHECKED) &&
                pfiltstate->dwFiltChn[uChnIndex] & FILT_CHECKED &&
                pfiltstate->dwFiltEvtType[uEvtTypeIndex] & FILT_CHECKED
            ) {
                pfiltstate->dwFiltTrk[uTrkIndex] |= FILT_AVAILABLE;
            }
            if(
                pfiltstate->dwFiltTrk[uTrkIndex] & FILT_CHECKED &&
                !(pfiltstate->dwFiltChn[uChnIndex] & FILT_CHECKED) &&
                pfiltstate->dwFiltEvtType[uEvtTypeIndex] & FILT_CHECKED
            ) {
                pfiltstate->dwFiltChn[uChnIndex] |= FILT_AVAILABLE;
            }
            if(
                pfiltstate->dwFiltTrk[uTrkIndex] & FILT_CHECKED &&
                pfiltstate->dwFiltChn[uChnIndex] & FILT_CHECKED &&
                !(pfiltstate->dwFiltEvtType[uEvtTypeIndex] & FILT_CHECKED)
            ) {
                pfiltstate->dwFiltEvtType[uEvtTypeIndex] |= FILT_AVAILABLE;
            }
            pevtCur = pevtCur->pevtNext;
            continue;
        }
        pfiltstate->dwFiltTrk[uTrkIndex] |= FILT_AVAILABLE;
        pfiltstate->dwFiltChn[uChnIndex] |= FILT_AVAILABLE;
        pfiltstate->dwFiltEvtType[uEvtTypeIndex] |= FILT_AVAILABLE;
        pevtCur = pevtCur->pevtNext;
    }

    cCBFiltItem = SendMessage(hwndCBFiltTrk, CB_GETCOUNT, 0, 0);
    for(i = 2; i < cCBFiltItem; i++)
        SendMessage(hwndCBFiltTrk, CB_DELETESTRING, 2, 0);
    cCBFiltItem = SendMessage(hwndCBFiltChn, CB_GETCOUNT, 0, 0);
    for(i = 2; i < cCBFiltItem; i++)
        SendMessage(hwndCBFiltChn, CB_DELETESTRING, 2, 0);
    cCBFiltItem = SendMessage(hwndCBFiltEvtType, CB_GETCOUNT, 0, 0);
    for(i = 2; i < cCBFiltItem; i++)
        SendMessage(hwndCBFiltEvtType, CB_DELETESTRING, 2, 0);
    
    for(uFiltIndex = 0, i = 2; uFiltIndex < pmf->cTrk; uFiltIndex++) {
        if(pfiltstate->dwFiltTrk[uFiltIndex] & FILT_AVAILABLE) {
            if(pmf->ppevtTrkName[uFiltIndex]) {
                lpszBuf = (LPWSTR)malloc((pmf->ppevtTrkName[uFiltIndex]->cbData + 128) * sizeof(WCHAR));
                ZeroMemory(lpszBuf, (pmf->ppevtTrkName[uFiltIndex]->cbData + 128) * sizeof(WCHAR));
                wsprintf(lpszBuf, L"Track #%u (", uFiltIndex);
                MultiByteToWideChar(CP_ACP, MB_COMPOSITE, (LPCSTR)pmf->ppevtTrkName[uFiltIndex]->abData, pmf->ppevtTrkName[uFiltIndex]->cbData, lpszBuf + lstrlen(lpszBuf), pmf->ppevtTrkName[uFiltIndex]->cbData * sizeof(WCHAR));
                lstrcat(lpszBuf, L")");
                SendMessage(hwndCBFiltTrk, CB_ADDSTRING, 0, (LPARAM)lpszBuf);
                free(lpszBuf);
            } else {
                wsprintf(szBuf, L"Track #%u", uFiltIndex);
                SendMessage(hwndCBFiltTrk, CB_ADDSTRING, 0, (LPARAM)szBuf);
            }
            i++;
        }
    }
    for(uFiltIndex = 0, i = 2; uFiltIndex < 17; uFiltIndex++) {
        if(pfiltstate->dwFiltChn[uFiltIndex] & FILT_AVAILABLE) {
            if(uFiltIndex == 0) {
                SendMessage(hwndCBFiltChn, CB_ADDSTRING, 0, (LPARAM)L"Non-channel");
            } else {
                wsprintf(szBuf, L"Channel #%u", uFiltIndex);
                SendMessage(hwndCBFiltChn, CB_ADDSTRING, 0, (LPARAM)szBuf);
            }
            i++;
        }
    }
    for(uFiltIndex = 0, i = 2; uFiltIndex < 10; uFiltIndex++) {
        if(pfiltstate->dwFiltEvtType[uFiltIndex] & FILT_AVAILABLE) {
            switch(uFiltIndex){
            case 0: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Note off"); break;
            case 1: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Note on (with no velocity)"); break;
            case 2: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Note on (with velocity)"); break;
            case 3: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Note aftertouch"); break;
            case 4: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Controller"); break;
            case 5: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Program change"); break;
            case 6: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Channel aftertouch"); break;
            case 7: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Pitch bend"); break;
            case 8: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"System exclusive"); break;
            case 9: SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Meta event"); break;
            }
            i++;
        }
    }
}