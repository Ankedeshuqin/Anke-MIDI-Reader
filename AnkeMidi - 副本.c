#define _WIN32_WINNT 0x0400
#include <Windows.h>
#include "cbinfo.h"
#include <CommCtrl.h>
#include <stdio.h> // For float number outputting
#include "StrList.h"
#include "MidiRead.h"
#include "resource.h"

#pragma comment(lib, "comctl32.lib")

#define DLUX(x) ((x) * cxChar / 4) // x in DLUs
#define DLUY(y) ((y) * cyChar / 8) // y in DLUs


#define CPAGE 4

enum ePage {
    PAGEEVTLIST, PAGETEMPOLIST, PAGEPLAYCTL, PAGETONALITYANLYZER
};

enum eIDC {
    /* Controls of common pages */
    IDC_STATUS, IDC_TC, IDC_BTNOPEN, IDC_BTNCLOSE, IDC_EDITFILEPATH, IDC_BTNPLAY, IDC_BTNSTOP, IDC_TBTIME, IDC_STATICTIME, IDC_CHECKHEX,
    /* Controls of event list page */
    IDC_STATICTYPE, IDC_STATICCTRK, IDC_STATICTB, IDC_STATICCEVT, IDC_CBFILTTRK, IDC_CBFILTCHN, IDC_CBFILTEVTTYPE, IDC_LVEVTLIST,
    /* Controls of tempo list page */
    IDC_STATICINITEMPO, IDC_STATICAVGTEMPO, IDC_STATICDUR, IDC_LVTEMPOLIST,
    /* Controls of play control page */
    IDC_EDITTRANSP, IDC_UDTRANSP, IDC_BTNTRANSPRESET, IDC_EDITTEMPOR, IDC_UDTEMPOR, IDC_BTNTEMPORRESET, IDC_TVMUT,
    /* Controls of tonality analyzer page */
    IDC_EDITCNOTEFIRST, IDC_EDITCNOTETOTAL = IDC_EDITCNOTEFIRST + 12, IDC_BTNANLYZTONALITY, IDC_STATICMOSTPROBTONALITY, IDC_EDITTONALITYFIRST, IDC_STATICTONALITYBARFIRST = IDC_EDITTONALITYFIRST + 24, IDC_STATICTONALITYBARLABELFIRST = IDC_STATICTONALITYBARFIRST + 24
};


enum eAppMsg {
    WM_APP_OPENFILE = WM_APP, // lParam: lpszFilePath
    WM_APP_CLOSEFILE,
    WM_APP_PLAYFROMEVT, // lParam: pevt
    WM_APP_STOP,
    WM_APP_FILTCBITEMCLK, // wParam: iIndex; lParam: hwndCheckedCB
    WM_APP_ANLYZTONALITY
};


COLORREF acrEvtType[] = {
    RGB(128, 128, 128), RGB(128, 128, 128), RGB(0, 0, 0), RGB(0, 0, 128), RGB(0, 128, 0),
    RGB(192, 128, 0), RGB(0, 0, 255), RGB(255, 0, 255), RGB(255, 0, 0), RGB(192, 128, 128)
};

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

DWORD GetEvtList(HWND hwndLVEvtList, MIDIFILE *pmf, FILTERSTATES *pfiltstate, BOOL fHex) {
    LVITEM lvitem;

    EVENT *pevtCur;
    DWORD dwRow = 0;
    WCHAR szBuf[128];
    LPWSTR lpszData, lpszDataCmt; // For meta or sys-ex events

    int cSharp = 0; // For key signature events

    UINT u;

    SendMessage(hwndLVEvtList, WM_SETREDRAW, FALSE, 0);
    ListView_DeleteAllItems(hwndLVEvtList);

    lvitem.mask = LVIF_TEXT | LVIF_PARAM;
    pevtCur = pmf->pevtHead;
    while(pevtCur) {
        if(!IsEvtUnfiltered(pevtCur, pfiltstate)) {
            pevtCur = pevtCur->pevtNext;
            continue;
        }

        lvitem.iItem = dwRow;
        lvitem.iSubItem = 0;
        wsprintf(szBuf, L"%u", dwRow + 1);
        lvitem.pszText = szBuf;
        lvitem.lParam = (LPARAM)pevtCur;
        ListView_InsertItem(hwndLVEvtList, &lvitem);

        wsprintf(szBuf, L"%u", pevtCur->wTrk);
        ListView_SetItemText(hwndLVEvtList, dwRow, 1, szBuf);

        wsprintf(szBuf, L"%u", pevtCur->dwTk);
        ListView_SetItemText(hwndLVEvtList, dwRow, 3, szBuf);

        if(pevtCur->bStatus < 0xF0) {
            wsprintf(szBuf, L"%u", (pevtCur->bStatus & 0xF) + 1);
            ListView_SetItemText(hwndLVEvtList, dwRow, 2, szBuf);

            switch(pevtCur->bStatus >> 4) {
            case 0x8:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Note off");
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszNote[pevtCur->bData1]);
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);
                break;
            case 0x9:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Note on");
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszNote[pevtCur->bData1]);
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);
                break;
            case 0xA:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Note aftertouch");
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszNote[pevtCur->bData1]);
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);
                break;
            case 0xB:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Controller");
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszCtl[pevtCur->bData1]);
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);
                break;
            case 0xC:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Program change");
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszPrg[pevtCur->bData1]);
                break;
            case 0xD:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Channel aftertouch");
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                break;
            case 0xE:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Pitch bend");
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);

                /* Generate comment string for pitch bend events */
                wsprintf(szBuf, L"%d", (int)pevtCur->bData1 | (pevtCur->bData2 << 7) - 8192);
                ListView_SetItemText(hwndLVEvtList, dwRow, 8, szBuf);
                break;
            }
        } else {
            lpszData = (LPWSTR)malloc((pevtCur->cbData * 4 + 1) * sizeof(WCHAR));
            lpszData[0] = '\0';
            for(u = 0; u < pevtCur->cbData; u++) {
                wsprintf(lpszData, fHex ? L"%s%02X" : L"%s%u", lpszData, pevtCur->abData[u]);
                if(u < pevtCur->cbData - 1)
                    wsprintf(lpszData, L"%s ", lpszData);
            }

            switch(pevtCur->bStatus) {
            case 0xF0:
            case 0xF7:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"System exclusive");
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, lpszData);
                free(lpszData);
                break;
            case 0xFF:
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Meta event");
                wsprintf(szBuf, fHex ? L"%02X" : L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszMeta[pevtCur->bData1]);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, lpszData);
                free(lpszData);

                /* Generate comment string for text-typed meta events */
                if(pevtCur->bData1 >= 0x1 && pevtCur->bData1 <= 0x9) {
                    lpszDataCmt = (LPWSTR)malloc((pevtCur->cbData + 3) * sizeof(WCHAR));
                    ZeroMemory(lpszDataCmt, (pevtCur->cbData + 3) * sizeof(WCHAR));
                    lpszDataCmt[0] = '\"';
                    MultiByteToWideChar(CP_ACP, MB_COMPOSITE, (LPCSTR)pevtCur->abData, pevtCur->cbData, lpszDataCmt + 1, (pevtCur->cbData + 2) * sizeof(WCHAR));
                    lstrcat(lpszDataCmt, L"\"");
                    ListView_SetItemText(hwndLVEvtList, dwRow, 8, lpszDataCmt);
                    free(lpszDataCmt);
                }

                /* Generate comment string for tempo events */
                if(pevtCur->bData1 == 0x51) {
                    swprintf(szBuf, 128, L"%f bpm", 60000000. / (pevtCur->abData[0] << 16 | pevtCur->abData[1] << 8 | pevtCur->abData[2]));
                    ListView_SetItemText(hwndLVEvtList, dwRow, 8, szBuf);
                }

                /* Generate comment string for time signature events */
                if(pevtCur->bData1 == 0x58) {
                    if(pevtCur->cbData >= 2) {
                        wsprintf(szBuf, L"%d/%d", pevtCur->abData[0], 1 << pevtCur->abData[1]);
                        ListView_SetItemText(hwndLVEvtList, dwRow, 8, szBuf);
                    }
                }

                /* Generate comment string for key signature events */
                if(pevtCur->bData1 == 0x59) {
                    if(pevtCur->cbData >= 2 && (pevtCur->abData[0] <= 7 || pevtCur->abData[0] >= 121) && pevtCur->abData[1] <= 1){
                        cSharp = (pevtCur->abData[0] + 64) % 128 - 64;

                        if(!pevtCur->abData[1]) { // Major
                            wsprintf(szBuf, L"%c", 'A' + (4 * cSharp + 30) % 7);
                            if (cSharp >= 6)
                                lstrcat(szBuf, L"-sharp");
                            if (cSharp <= -2)
                                lstrcat(szBuf, L"-flat");
                            lstrcat(szBuf, L" major");
                        } else { // Minor
                            wsprintf(szBuf, L"%c", 'A' + (4 * cSharp + 28) % 7);
                            if (cSharp >= 3)
                                lstrcat(szBuf, L"-sharp");
                            if (cSharp <= -5)
                                lstrcat(szBuf, L"-flat");
                            lstrcat(szBuf, L" minor");
                        }
                        ListView_SetItemText(hwndLVEvtList, dwRow, 8, szBuf);
                    }
                }
                break;
            }
        }
        
        pevtCur = pevtCur->pevtNext;
        dwRow++;
    }
    
    SendMessage(hwndLVEvtList, WM_SETREDRAW, TRUE, 0);
    return dwRow;
}

DWORD GetTempoList(HWND hwndLVTempoList, MIDIFILE *pmf, BOOL fHex) {
    LVITEM lvitem;

    TEMPOEVENT *ptempoevtCur;
    DWORD dwRow = 0;
    WCHAR szBuf[128];

    SendMessage(hwndLVTempoList, WM_SETREDRAW, FALSE, 0);
    ListView_DeleteAllItems(hwndLVTempoList);

    lvitem.mask = LVIF_TEXT | LVIF_PARAM;
    ptempoevtCur = pmf->ptempoevtHead;
    while(ptempoevtCur) {
        lvitem.iItem = dwRow;
        lvitem.iSubItem = 0;
        wsprintf(szBuf, L"%u", dwRow + 1);
        lvitem.pszText = szBuf;
        lvitem.lParam = (LPARAM)ptempoevtCur;
        ListView_InsertItem(hwndLVTempoList, &lvitem);

        wsprintf(szBuf, L"%u", ptempoevtCur->dwTk);
        ListView_SetItemText(hwndLVTempoList, dwRow, 1, szBuf);

        wsprintf(szBuf, L"%u", ptempoevtCur->cTk);
        ListView_SetItemText(hwndLVTempoList, dwRow, 2, szBuf);

        wsprintf(szBuf, fHex ? L"%06X" : L"%u", ptempoevtCur->dwData);
        ListView_SetItemText(hwndLVTempoList, dwRow, 3, szBuf);

        swprintf(szBuf, 128, L"%f bpm", 60000000. / ptempoevtCur->dwData);
        ListView_SetItemText(hwndLVTempoList, dwRow, 4, szBuf);

        ptempoevtCur = ptempoevtCur->ptempoevtNext;
        dwRow++;
    }

    SendMessage(hwndLVTempoList, WM_SETREDRAW, TRUE, 0);
    return dwRow;
}

void SetTimeText(HWND hwndStaticTime, double dCurTimeSec, double dDurSec, int iTempoR) {
    WCHAR szBuf[64];
    static int iCurTimeOld = 0, iDurOld = 0;
    int iCurTimeNew, iDurNew;

    iCurTimeNew = (int)(dCurTimeSec * 100 / iTempoR + 0.5);
    iDurNew = (int)(dDurSec * 100 / iTempoR + 0.5);
    if(iCurTimeNew == iCurTimeOld && iDurNew == iDurOld)
        return;

    wsprintf(szBuf, L"%02d:%02d / %02d:%02d", iCurTimeNew / 60, iCurTimeNew % 60, iDurNew / 60, iDurNew % 60);
    SetWindowText(hwndStaticTime, szBuf);

    iCurTimeOld = iCurTimeNew;
    iDurOld = iDurNew;
}


#pragma comment(lib, "winmm.lib")

enum eStrmStatus{
    STRM_STOP, STRM_PLAY
};

#define STRMBUFLEN 1536

EVENT *FillStrmBuf(MIDIHDR *pmhdr, EVENT *pevtCurBuf, DWORD *pdwPrevBufEvtTk, int iTransp, int iTempoR, WORD *pwTrkChnUnmuted) {
    UINT uStrmDataOffset = 0;
    MIDIEVENT mevt;
    UINT cbLongEvtStrmData;

    while(pevtCurBuf) {
        if(pevtCurBuf->bStatus == 0xF0 || pevtCurBuf->bStatus == 0xF7)
            cbLongEvtStrmData = (pevtCurBuf->cbData + 3) & ~3;
        else
            cbLongEvtStrmData = 0;

        if(uStrmDataOffset + 12 + cbLongEvtStrmData > STRMBUFLEN)
            break;

        mevt.dwDeltaTime = pevtCurBuf->dwTk - *pdwPrevBufEvtTk;
        mevt.dwStreamID = 0;
        if(cbLongEvtStrmData) { // Sys-ex events
            mevt.dwEvent = pevtCurBuf->cbData | MEVT_LONGMSG << 24;
        } else if(pevtCurBuf->bStatus == 0xFF) {
            if(pevtCurBuf->bData1 == 0x51) { // Tempo events
                mevt.dwEvent = max(min((pevtCurBuf->abData[0] << 16 | pevtCurBuf->abData[1] << 8 | pevtCurBuf->abData[2]) * 100 / iTempoR, 0xFFFFFF), 1) | MEVT_TEMPO << 24;
            } else { // Other meta events
                mevt.dwEvent = MEVT_NOP << 24;
            }
        } else if(pwTrkChnUnmuted[pevtCurBuf->wTrk] & (1 << (pevtCurBuf->bStatus & 0xF))) { // Unmuted channel events
            if((pevtCurBuf->bStatus & 0xF) != 0x9 && pevtCurBuf->bStatus >= 0x80 && pevtCurBuf->bStatus < 0xB0) { // Note events
                if((int)pevtCurBuf->bData1 + iTransp >= 0 && (int)pevtCurBuf->bData1 + iTransp < 128) {
                    mevt.dwEvent = MAKEFOURCC(pevtCurBuf->bStatus, pevtCurBuf->bData1 + iTransp, pevtCurBuf->bData2, MEVT_SHORTMSG);
                } else {
                    mevt.dwEvent = MEVT_NOP << 24;
                }
            } else { // Other channel events
                mevt.dwEvent = MAKEFOURCC(pevtCurBuf->bStatus, pevtCurBuf->bData1, pevtCurBuf->bData2, MEVT_SHORTMSG);
            }
        } else { // Muted channel events
            mevt.dwEvent = MEVT_NOP << 24;
        }
        mevt.dwEvent |= MEVT_F_CALLBACK;

        memcpy(pmhdr->lpData + uStrmDataOffset, &mevt, 12);
        if(cbLongEvtStrmData)
            memcpy(pmhdr->lpData + uStrmDataOffset + 12, pevtCurBuf->abData, pevtCurBuf->cbData);
        uStrmDataOffset += 12 + cbLongEvtStrmData;

        *pdwPrevBufEvtTk = pevtCurBuf->dwTk;
        pevtCurBuf = pevtCurBuf->pevtNext;
    }
    pmhdr->dwBytesRecorded = uStrmDataOffset;
    return pevtCurBuf;
}


WCHAR szCmdFilePath[MAX_PATH] = L"";
HINSTANCE hInstanceMain;
HWND hwndMain;

WNDPROC DefStaticProc, DefLBProc;

LRESULT CALLBACK PassToMainWndStaticProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch(uMsg) {
    case WM_COMMAND:
    case WM_NOTIFY:
    case WM_DRAWITEM:
    case WM_APP_FILTCBITEMCLK:
        return SendMessage(hwndMain, uMsg, wParam, lParam);
    }
    return CallWindowProc(DefStaticProc, hwnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK LBCBFiltProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    HWND hwndCheckedCB = (HWND)GetWindowLong(hwnd, GWL_USERDATA); // The user data of list boxes from filter combo boxes are set to the combo boxes' window handles
    int iIndex;
    BOOL fNonClient;
    HWND hwndParent;

    switch(uMsg) {
    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK:
        fNonClient = HIWORD(SendMessage(hwnd, LB_ITEMFROMPOINT, 0, lParam));
        if(fNonClient)
            break;

        iIndex = LOWORD(SendMessage(hwnd, LB_ITEMFROMPOINT, 0, lParam));
        hwndParent = GetParent(hwndCheckedCB);
        SendMessage(hwndParent, WM_APP_FILTCBITEMCLK, iIndex, (LPARAM)hwndCheckedCB);

        InvalidateRect(hwnd, NULL, FALSE);
        return 0;
    }
    return CallWindowProc(DefLBProc, hwnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK WndProc(HWND hwnd,UINT uMsg,WPARAM wParam,LPARAM lParam){
    LPCWSTR alpszNoteLabel[] = {L"C", L"C#", L"D", L"D#", L"E", L"F", L"F#", L"G", L"G#", L"A", L"A#", L"B"};
    LPCWSTR alpszTonalityLabel[] = {L"C", L"Cm", L"Db", L"C#m", L"D", L"Dm", L"Eb", L"Ebm", L"E", L"Em", L"F", L"Fm", L"F#", L"F#m", L"G", L"Gm", L"Ab", L"G#m", L"A", L"Am", L"Bb", L"Bbm", L"B", L"Bm"};


    static OSVERSIONINFO osversioninfo;

    static HFONT hfontGUI;
    static TEXTMETRIC tm;
    static int cxChar, cyChar;
    NONCLIENTMETRICS ncm;
    HDC hdc;

    static HWND hwndStatus;
    static long cyStatus;
    RECT rc;
    
    static HWND hwndTool;
    static TOOLINFO ti;

    static HWND hwndTC;
    static TCITEM tci;
    static int iCurPage = 0;
    static HWND hwndStaticPage[CPAGE];

    static HWND hwndBtnOpen, hwndBtnClose, hwndEditFilePath, hwndBtnPlay, hwndBtnStop, hwndTBTime, hwndStaticTime, hwndCheckHex,
        hwndStaticType, hwndStaticCTrk, hwndStaticTb, hwndStaticCEvt, hwndCBFiltTrk, hwndCBFiltChn, hwndCBFiltEvtType, hwndLBCBFiltTrk, hwndLBCBFiltChn, hwndLBCBFiltEvtType, hwndLVEvtList,
        hwndStaticIniTempo, hwndStaticAvgTempo, hwndStaticDur, hwndLVTempoList,
        hwndEditTransp, hwndUdTransp, hwndBtnTranspReset, hwndEditTempoR, hwndUdTempoR, hwndBtnTempoRReset, hwndTVMut,
        hwndEditCNote[12], hwndEditCNoteTotal, hwndBtnAnlyzTonality, hwndStaticMostProbTonality, hwndEditTonality[24], hwndStaticTonalityBar[24], hwndStaticTonalityBarLabel[24];
    HWND hwndTemp;
    static BOOL fHex = FALSE;

    COMBOBOXINFO cbi;

    LVCOLUMN lvcol;


    static MIDIFILE mf;
    
    static FILTERSTATES *pfiltstate;
    LPDRAWITEMSTRUCT lpdis;
    RECT rcCheckBox, rcText;
    UINT uState;
    int i, cCBFiltItem;
    UINT uFiltIndex;
    static DWORD cEvtListRow = 0, cTempoListRow = 0;
    int iTopIndex;
    LPNMLVCUSTOMDRAW lplvcd;
    

    static BOOL fAnlyzTonality = FALSE;

    UINT cNoteTotal = 0;
    static UINT acNote[12];
    static double dTonalityProprtn[24];
    static BOOL fTonalityMax[24];
    double dTonalityProprtnMax;
    HBRUSH hbr;


    static int iTransp;
    static int iTempoR;

    static WORD *pwTrkChnUnmuted;
    TVINSERTSTRUCT tvis;
    static HTREEITEM htiRoot;
    HTREEITEM htiTrk, htiChn, htiTemp;
    TVHITTESTINFO tvhti;
    int iTVMutIndex, iMutTrk, iMutChn;
    BOOL f;


    static BOOL fTBIsTracking = FALSE;

    static HMIDISTRM hms;
    static MIDIHDR mhdr[2];
    static UINT uDevID = MIDI_MAPPER;
    static UINT uStrmDataOffset;
    static BYTE *pbStrmData;
    static UINT uCurOutputEvt, uCurEvtListRow, uCurTempoListRow;
    MIDIPROPTIMEDIV mptd;
    MIDIPROPTEMPO mptempo;
    static EVENT *pevtCurBuf, *pevtCurOutput;
    static DWORD dwPrevBufEvtTk;
    static BOOL fDone;
    static DWORD dwCurTk;
    MMTIME mmt;
    static int iCurStrmStatus = STRM_STOP;
    static DWORD dwStartTime;
    static DWORD dwCurTempoData;
    EVENT *pevtTemp;
    static int iPlayCtlBufRemaining; // See the process of the MM_MOM_DONE message


    HDROP hdrop;
    WCHAR szDragFilePath[MAX_PATH];
    static OPENFILENAME ofn;
    static WCHAR szFilePath[MAX_PATH], szFileName[MAX_PATH];

    UINT u;
    WCHAR szBuf[128];
    LPWSTR lpszBuf;

    
    switch(uMsg){
    case WM_CREATE:
        InitCommonControls();

        osversioninfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
        GetVersionEx(&osversioninfo);

        ncm.cbSize = sizeof(NONCLIENTMETRICS);
        SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
        hfontGUI = CreateFontIndirect(&ncm.lfMessageFont);
        hdc = GetDC(hwnd);
        SelectObject(hdc, hfontGUI);
        GetTextMetrics(hdc, &tm);
        cxChar = tm.tmAveCharWidth;
        cyChar = tm.tmHeight + tm.tmExternalLeading;
        ReleaseDC(hwnd, hdc);


        /* Controls for main window */
        hwndStatus = CreateStatusWindow(WS_CHILD | WS_VISIBLE | CCS_BOTTOM | SBARS_SIZEGRIP,
            L"", hwnd, IDC_STATUS);
        GetWindowRect(hwndStatus, &rc);
        cyStatus = rc.bottom - rc.top;

        hwndTool = CreateWindow(TOOLTIPS_CLASS, NULL,
            TTS_ALWAYSTIP,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            hwnd, NULL, hInstanceMain, NULL);
        ti.cbSize = sizeof(TOOLINFO);
        ti.uFlags = TTF_SUBCLASS;
        ti.hwnd = hwnd;

        hwndBtnOpen = CreateWindow(L"BUTTON", L"Open",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(7), DLUY(7), DLUX(28), DLUY(14),
            hwnd, (HMENU)IDC_BTNOPEN, hInstanceMain, NULL);
        SendMessage(hwndBtnOpen, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnClose = CreateWindow(L"BUTTON", L"Close",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(39), DLUY(7), DLUX(28), DLUY(14),
            hwnd, (HMENU)IDC_BTNCLOSE, hInstanceMain, NULL);
        SendMessage(hwndBtnClose, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndEditFilePath = CreateWindow(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_READONLY | ES_AUTOHSCROLL,
            0, 0, 0, 0,
            hwnd, (HMENU)IDC_EDITFILEPATH, hInstanceMain, NULL);
        SendMessage(hwndEditFilePath, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnPlay = CreateWindow(L"BUTTON", L"Play",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(7), DLUY(25), DLUX(28), DLUY(14),
            hwnd, (HMENU)IDC_BTNPLAY, hInstanceMain, NULL);
        SendMessage(hwndBtnPlay, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnStop = CreateWindow(L"BUTTON", L"Stop",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(39), DLUY(25), DLUX(28), DLUY(14),
            hwnd, (HMENU)IDC_BTNSTOP, hInstanceMain, NULL);
        SendMessage(hwndBtnStop, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndTBTime = CreateWindow(TRACKBAR_CLASS, NULL,
            WS_CHILD | WS_VISIBLE,
            0, 0, 0, 0,
            hwnd, (HMENU)IDC_TBTIME, hInstanceMain, NULL);
        SendMessage(hwndTBTime, TBM_SETRANGEMIN, TRUE, 0);
        SendMessage(hwndTBTime, TBM_SETRANGEMAX, TRUE, 0);
        SendMessage(hwndTBTime, TBM_SETPOS, TRUE, 0);
        SendMessage(hwndTBTime, TBM_SETPAGESIZE, 0, 3000);
        hwndStaticTime = CreateWindow(L"STATIC", L"00:00 / 00:00",
            WS_CHILD | WS_VISIBLE | SS_RIGHT,
            0, 0, 0, 0,
            hwnd, (HMENU)IDC_STATICTIME, hInstanceMain, NULL);
        SendMessage(hwndStaticTime, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndCheckHex = CreateWindow(L"BUTTON", L"Show event data in hex",
            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX | BS_RIGHT | BS_RIGHTBUTTON,
            0, 0, 0, 0,
            hwnd, (HMENU)IDC_CHECKHEX, hInstanceMain, NULL);
        SendMessage(hwndCheckHex, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);

        
        hwndTC = CreateWindow(WC_TABCONTROL, NULL,
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
            0, 0, 0, 0,
            hwnd, (HMENU)IDC_TC, hInstanceMain, NULL);
        SendMessage(hwndTC, WM_SETFONT, (WPARAM)hfontGUI, TRUE);

        tci.mask = TCIF_TEXT;
        tci.pszText = L"Event list";
        SendMessage(hwndTC, TCM_INSERTITEM, PAGEEVTLIST, (LPARAM)&tci);
        tci.pszText = L"Tempo list";
        SendMessage(hwndTC, TCM_INSERTITEM, PAGETEMPOLIST, (LPARAM)&tci);
        tci.pszText = L"Play control";
        SendMessage(hwndTC, TCM_INSERTITEM, PAGEPLAYCTL, (LPARAM)&tci);
        tci.pszText = L"Tonality analyzer";
        SendMessage(hwndTC, TCM_INSERTITEM, PAGETONALITYANLYZER, (LPARAM)&tci);
        for(i = 0; i < CPAGE; i++) {
            hwndStaticPage[i] = CreateWindow(L"STATIC", NULL,
                WS_CHILD,
                0, 0, 0, 0,
                hwndTC, (HMENU)-1, hInstanceMain, NULL);
            if(i == 0)
                DefStaticProc = (WNDPROC)GetWindowLong(hwndStaticPage[i], GWL_WNDPROC);
            SetWindowLong(hwndStaticPage[i], GWL_WNDPROC, (LONG)PassToMainWndStaticProc);
        }
        iCurPage = PAGEEVTLIST;
        ShowWindow(hwndStaticPage[iCurPage], SW_SHOW);

        /* Controls for event list page */
        hwndStaticType = CreateWindow(L"STATIC", L"Type: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(7), DLUY(7), DLUX(56), DLUY(8),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_STATICTYPE, hInstanceMain, NULL);
        SendMessage(hwndStaticType, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticCTrk = CreateWindow(L"STATIC", L"Number of tracks: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(67), DLUY(7), DLUX(104), DLUY(8),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_STATICCTRK, hInstanceMain, NULL);
        SendMessage(hwndStaticCTrk, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticTb = CreateWindow(L"STATIC", L"Timebase: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(175), DLUY(7), DLUX(72), DLUY(8),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_STATICTB, hInstanceMain, NULL);
        SendMessage(hwndStaticTb, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticCEvt = CreateWindow(L"STATIC", L"Number of events: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(251), DLUY(7), DLUX(104), DLUY(8),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_STATICCEVT, hInstanceMain, NULL);
        SendMessage(hwndStaticCEvt, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);

        hwndCBFiltTrk = CreateWindow(L"COMBOBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            DLUX(7), DLUY(22), DLUX(122), DLUY(162),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_CBFILTTRK, hInstanceMain, NULL);
        SendMessage(hwndCBFiltTrk, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndCBFiltChn = CreateWindow(L"COMBOBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            DLUX(133), DLUY(22), DLUX(122), DLUY(162),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_CBFILTCHN, hInstanceMain, NULL);
        SendMessage(hwndCBFiltChn, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndCBFiltEvtType = CreateWindow(L"COMBOBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            DLUX(259), DLUY(22), DLUX(122), DLUY(162),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_CBFILTEVTTYPE, hInstanceMain, NULL);
        SendMessage(hwndCBFiltEvtType, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);

        SendMessage(hwndCBFiltTrk, CB_ADDSTRING, 0, (LPARAM)L"Select all");
        SendMessage(hwndCBFiltTrk, CB_ADDSTRING, 0, (LPARAM)L"Select none");
        SendMessage(hwndCBFiltChn, CB_ADDSTRING, 0, (LPARAM)L"Select all");
        SendMessage(hwndCBFiltChn, CB_ADDSTRING, 0, (LPARAM)L"Select none");
        SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Select all");
        SendMessage(hwndCBFiltEvtType, CB_ADDSTRING, 0, (LPARAM)L"Select none");

        cbi.cbSize = sizeof(COMBOBOXINFO);
        SendMessage(hwndCBFiltTrk, CB_GETCOMBOBOXINFO, 0, (LPARAM)&cbi);
        hwndLBCBFiltTrk = cbi.hwndList;
        DefLBProc = (WNDPROC)GetWindowLong(hwndLBCBFiltTrk, GWL_WNDPROC);
        SendMessage(hwndCBFiltChn, CB_GETCOMBOBOXINFO, 0, (LPARAM)&cbi);
        hwndLBCBFiltChn = cbi.hwndList;
        SendMessage(hwndCBFiltEvtType, CB_GETCOMBOBOXINFO, 0, (LPARAM)&cbi);
        hwndLBCBFiltEvtType = cbi.hwndList;
        SetWindowLong(hwndLBCBFiltTrk, GWL_WNDPROC, (LONG)LBCBFiltProc);
        SetWindowLong(hwndLBCBFiltTrk, GWL_USERDATA, (LONG)hwndCBFiltTrk);
        SetWindowLong(hwndLBCBFiltChn, GWL_WNDPROC, (LONG)LBCBFiltProc);
        SetWindowLong(hwndLBCBFiltChn, GWL_USERDATA, (LONG)hwndCBFiltChn);
        SetWindowLong(hwndLBCBFiltEvtType, GWL_WNDPROC, (LONG)LBCBFiltProc);
        SetWindowLong(hwndLBCBFiltEvtType, GWL_USERDATA, (LONG)hwndCBFiltEvtType);

        hwndLVEvtList = CreateWindow(WC_LISTVIEW, NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            0, 0, 0, 0,
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_LVEVTLIST, hInstanceMain, NULL);
        ListView_SetExtendedListViewStyle(hwndLVEvtList, LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);

        lvcol.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT;
        lvcol.fmt = LVCFMT_LEFT;
        lvcol.cx = cxChar * 10;
        lvcol.pszText = L"#";
        ListView_InsertColumn(hwndLVEvtList, 0, &lvcol);
        lvcol.pszText = L"Track";
        ListView_InsertColumn(hwndLVEvtList, 1, &lvcol);
        lvcol.pszText = L"Channel";
        ListView_InsertColumn(hwndLVEvtList, 2, &lvcol);
        lvcol.cx = cxChar * 15;
        lvcol.pszText = L"Start tick";
        ListView_InsertColumn(hwndLVEvtList, 3, &lvcol);
        lvcol.cx = cxChar * 25;
        lvcol.pszText = L"Event type";
        ListView_InsertColumn(hwndLVEvtList, 4, &lvcol);
        lvcol.cx = cxChar * 10;
        lvcol.pszText = L"Data 1";
        ListView_InsertColumn(hwndLVEvtList, 5, &lvcol);
        lvcol.cx = cxChar * 25;
        lvcol.pszText = L"Data 1 type";
        ListView_InsertColumn(hwndLVEvtList, 6, &lvcol);
        lvcol.cx = cxChar * 10;
        lvcol.pszText = L"Data 2";
        ListView_InsertColumn(hwndLVEvtList, 7, &lvcol);
        lvcol.cx = cxChar * 50;
        lvcol.pszText = L"Comment";
        ListView_InsertColumn(hwndLVEvtList, 8, &lvcol);

        /* Controls for tempo list page */
        hwndStaticIniTempo = CreateWindow(L"STATIC", L"Initial tempo: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(7), DLUY(7), DLUX(136), DLUY(8),
            hwndStaticPage[PAGETEMPOLIST], (HMENU)IDC_STATICINITEMPO, hInstanceMain, NULL);
        SendMessage(hwndStaticIniTempo, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticAvgTempo = CreateWindow(L"STATIC", L"Average tempo: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(147), DLUY(7), DLUX(136), DLUY(8),
            hwndStaticPage[PAGETEMPOLIST], (HMENU)IDC_STATICAVGTEMPO, hInstanceMain, NULL);
        SendMessage(hwndStaticAvgTempo, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticDur = CreateWindow(L"STATIC", L"Duration: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(287), DLUY(7), DLUX(108), DLUY(8),
            hwndStaticPage[PAGETEMPOLIST], (HMENU)IDC_STATICDUR, hInstanceMain, NULL);
        SendMessage(hwndStaticDur, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);

        hwndLVTempoList = CreateWindow(WC_LISTVIEW, NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            0, 0, 0, 0,
            hwndStaticPage[PAGETEMPOLIST], (HMENU)IDC_LVTEMPOLIST, hInstanceMain, NULL);
        ListView_SetExtendedListViewStyle(hwndLVTempoList, LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);

        lvcol.cx = cxChar * 10;
        lvcol.pszText = L"#";
        ListView_InsertColumn(hwndLVTempoList, 0, &lvcol);
        lvcol.cx = cxChar * 15;
        lvcol.pszText = L"Start tick";
        ListView_InsertColumn(hwndLVTempoList, 1, &lvcol);
        lvcol.pszText = L"Lasting ticks";
        ListView_InsertColumn(hwndLVTempoList, 2, &lvcol);
        lvcol.cx = cxChar * 20;
        lvcol.pszText = L"Tempo data";
        ListView_InsertColumn(hwndLVTempoList, 3, &lvcol);
        lvcol.cx = cxChar * 30;
        lvcol.pszText = L"Tempo";
        ListView_InsertColumn(hwndLVTempoList, 4, &lvcol);

        /* Controls for play control page */
        hwndTemp = CreateWindow(L"STATIC", L"Transposition:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(7), DLUY(10), DLUX(64), DLUY(8),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndEditTransp = CreateWindow(L"EDIT", NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            DLUX(74), DLUY(7), DLUX(34), DLUY(14),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_EDITTRANSP, hInstanceMain, NULL);
        SendMessage(hwndEditTransp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndUdTransp = CreateWindow(UPDOWN_CLASS, NULL,
            WS_CHILD | WS_VISIBLE | UDS_ALIGNRIGHT | UDS_ARROWKEYS | UDS_SETBUDDYINT,
            0, 0, 0, 0,
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_UDTRANSP, hInstanceMain, NULL);
        SendMessage(hwndUdTransp, UDM_SETBUDDY, (WPARAM)hwndEditTransp, 0);
        SendMessage(hwndUdTransp, UDM_SETRANGE, 0, MAKELPARAM(36, -36));
        SendMessage(hwndUdTransp, UDM_SETPOS, 0, 0);
        hwndBtnTranspReset = CreateWindow(L"BUTTON", L"Reset",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(112), DLUY(7), DLUX(28), DLUY(14),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_BTNTRANSPRESET, hInstanceMain, NULL);
        SendMessage(hwndBtnTranspReset, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndTemp = CreateWindow(L"STATIC", L"Tempo ratio:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(159), DLUY(10), DLUX(56), DLUY(8),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndEditTempoR = CreateWindow(L"EDIT", NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            DLUX(218), DLUY(7), DLUX(34), DLUY(14),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_EDITTEMPOR, hInstanceMain, NULL);
        SendMessage(hwndEditTempoR, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndUdTempoR = CreateWindow(UPDOWN_CLASS, NULL,
            WS_CHILD | WS_VISIBLE | UDS_ALIGNRIGHT | UDS_ARROWKEYS | UDS_SETBUDDYINT,
            0, 0, 0, 0,
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_UDTEMPOR, hInstanceMain, NULL);
        SendMessage(hwndUdTempoR, UDM_SETBUDDY, (WPARAM)hwndEditTempoR, 0);
        SendMessage(hwndUdTempoR, UDM_SETRANGE, 0, MAKELPARAM(500, 20));
        SendMessage(hwndUdTempoR, UDM_SETPOS, 0, 100);
        hwndTemp = CreateWindow(L"STATIC", L"%",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(255), DLUY(10), DLUX(12), DLUY(8),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnTempoRReset = CreateWindow(L"BUTTON", L"Reset",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(267), DLUY(7), DLUX(28), DLUY(14),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_BTNTEMPORRESET, hInstanceMain, NULL);
        SendMessage(hwndBtnTempoRReset, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndTemp = CreateWindow(L"STATIC", L"Muting:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(7), DLUY(28), DLUX(36), DLUY(8),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndTVMut = CreateWindow(WC_TREEVIEW, NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | TVS_HASBUTTONS | TVS_LINESATROOT | TVS_HASLINES,
            0, 0, 0, 0,
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_TVMUT, hInstanceMain, NULL);
        SetWindowLong(hwndTVMut, GWL_STYLE, GetWindowLong(hwndTVMut, GWL_STYLE) | TVS_CHECKBOXES); // Due to Windows API bugs, we have to do so
        
        /* Controls for tonality analyzer page */
        hwndTemp = CreateWindow(L"STATIC", L"Note count:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(7), DLUY(7), DLUX(68), DLUY(8),
            hwndStaticPage[PAGETONALITYANLYZER], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        for(i = 0; i < 12; i++) {
            hwndTemp = CreateWindow(L"STATIC", alpszNoteLabel[i],
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                DLUX(7 + 36 * i), DLUY(18), DLUX(32), DLUY(8),
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)-1, hInstanceMain, NULL);
            SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
            hwndEditCNote[i] = CreateWindow(L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_AUTOHSCROLL,
                DLUX(7 + 36 * i), DLUY(29), DLUX(32), DLUY(14),
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)(IDC_EDITCNOTEFIRST + i), hInstanceMain, NULL);
            SendMessage(hwndEditCNote[i], WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        }
        hwndTemp = CreateWindow(L"STATIC", L"Total",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            DLUX(439), DLUY(18), DLUX(32), DLUY(8),
            hwndStaticPage[PAGETONALITYANLYZER], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndEditCNoteTotal = CreateWindow(L"EDIT", L"0",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            DLUX(439), DLUY(29), DLUX(32), DLUY(14),
            hwndStaticPage[PAGETONALITYANLYZER], (HMENU)IDC_EDITCNOTETOTAL, hInstanceMain, NULL);
        SendMessage(hwndEditCNoteTotal, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnAnlyzTonality = CreateWindow(L"BUTTON", L"Analyze tonality",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(475), DLUY(29), DLUX(72), DLUY(14),
            hwndStaticPage[PAGETONALITYANLYZER], (HMENU)IDC_BTNANLYZTONALITY, hInstanceMain, NULL);
        SendMessage(hwndBtnAnlyzTonality, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticMostProbTonality = CreateWindow(L"STATIC", L"Most probable tonality: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            0, 0, 0, 0,
            hwndStaticPage[PAGETONALITYANLYZER], (HMENU)IDC_STATICMOSTPROBTONALITY, hInstanceMain, NULL);
        SendMessage(hwndStaticMostProbTonality, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        for(i = 0; i < 24; i++) {
            hwndTemp = CreateWindow(L"STATIC", alpszTonalityLabel[i],
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                DLUX(7 + 36 * i), DLUY(65), DLUX(32), DLUY(8),
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)-1, hInstanceMain, NULL);
            SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
            hwndEditTonality[i] = CreateWindow(L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_READONLY | ES_AUTOHSCROLL,
                DLUX(7 + 36 * i), DLUY(76), DLUX(32), DLUY(14),
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)(IDC_EDITTONALITYFIRST + i), hInstanceMain, NULL);
            SendMessage(hwndEditTonality[i], WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
            hwndStaticTonalityBar[i] = CreateWindow(L"STATIC", L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | SS_OWNERDRAW,
                0, 0, 0, 0,
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)(IDC_STATICTONALITYBARFIRST + i), hInstanceMain, NULL);
            SetWindowLong(hwndStaticTonalityBar[i], GWL_WNDPROC, (LONG)PassToMainWndStaticProc);
            hwndStaticTonalityBarLabel[i] = CreateWindow(L"STATIC", alpszTonalityLabel[i],
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                0, 0, 0, 0,
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)(IDC_STATICTONALITYBARLABELFIRST + i), hInstanceMain, NULL);
            SendMessage(hwndStaticTonalityBarLabel[i], WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        }
        

        ZeroMemory(&mf,sizeof(MIDIFILE));

        iTransp = 0;
        iTempoR = 100;


        ofn.lStructSize = sizeof(OPENFILENAME);
        ofn.hwndOwner = hwnd;
        ofn.lpstrFilter = L"MIDI Files (*.mid;*.midi;*.kar;*.rmi)\0*.mid;*.midi;*.kar;*.rmi\0"
            L"All Files (*.*)\0*.*\0\0";
        ofn.lpstrCustomFilter = NULL;
        ofn.nFilterIndex = 0;
        ofn.lpstrFile = szFilePath;
        ofn.nMaxFile = MAX_PATH;
        ofn.lpstrFileTitle = szFileName;
        ofn.nMaxFileTitle = MAX_PATH;
        ofn.Flags = OFN_FILEMUSTEXIST;

        if(szCmdFilePath[0])
            PostMessage(hwnd, WM_APP_OPENFILE, 0, (LPARAM)szCmdFilePath);
        return 0;

    case WM_SIZE:
        /* Controls for main window */
        SendMessage(hwndStatus, uMsg, wParam, lParam);
        MoveWindow(hwndTC, 0, DLUY(46), LOWORD(lParam), HIWORD(lParam) - cyStatus - DLUY(46), TRUE);
        MoveWindow(hwndEditFilePath, DLUX(71), DLUY(7), LOWORD(lParam) - DLUX(78), DLUY(14), TRUE);
        MoveWindow(hwndTBTime, DLUX(71), DLUY(25), LOWORD(lParam) - DLUX(134), DLUY(14), TRUE);
        MoveWindow(hwndStaticTime, LOWORD(lParam) - DLUX(59), DLUY(28), DLUX(52), DLUY(8), TRUE);
        MoveWindow(hwndCheckHex, LOWORD(lParam) - DLUX(115), DLUY(46), DLUX(108), DLUY(10), TRUE);
        
        GetClientRect(hwnd, &rc);
        rc.bottom -= cyStatus + DLUY(46);
        TabCtrl_AdjustRect(hwndTC, FALSE, &rc);
        for(i = 0; i < CPAGE; i++) {
            MoveWindow(hwndStaticPage[i], rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, TRUE);
        }
        
        /* Controls for event list page */
        MoveWindow(hwndLVEvtList, DLUX(7), DLUY(osversioninfo.dwMajorVersion >= 6 ? 36 : 40), rc.right - rc.left - DLUX(14), rc.bottom - rc.top - DLUY(osversioninfo.dwMajorVersion >= 6 ? 43 : 47), TRUE);

        /* Controls for tempo list page */
        MoveWindow(hwndLVTempoList, DLUX(7), DLUY(22), rc.right - rc.left - DLUX(14), rc.bottom - rc.top - DLUY(29), TRUE);

        /* Controls for play control page */
        MoveWindow(hwndTVMut, DLUX(7), DLUY(39), rc.right - rc.left - DLUX(14), rc.bottom - rc.top - DLUY(46), TRUE);

        /* Controls for tonality analyzer page */
        MoveWindow(hwndStaticMostProbTonality, DLUX(7), DLUY(50), rc.right - rc.left - DLUX(7), DLUY(8), TRUE);
        for(i = 0; i < 24; i++) {
            MoveWindow(hwndStaticTonalityBar[i], DLUX(7 + 24 * i), DLUY(94), DLUX(24), rc.bottom - rc.top - DLUY(112), TRUE);
            MoveWindow(hwndStaticTonalityBarLabel[i], DLUX(7 + 24 * i), rc.bottom - rc.top - DLUY(15), DLUX(24), DLUY(8), TRUE);

            ti.uId = i;
            GetWindowRect(hwndStaticTonalityBar[i], &ti.rect);
            MapWindowPoints(HWND_DESKTOP, hwnd, (LPPOINT)&ti.rect, 2);
            SendMessage(hwndTool, TTM_NEWTOOLRECT, 0, (LPARAM)&ti);
        }
        return 0;

    case WM_DROPFILES:
        hdrop = (HDROP)wParam;
        DragQueryFile(hdrop, 0, szDragFilePath, MAX_PATH + 1);
        DragFinish(hdrop);
        SendMessage(hwnd, WM_APP_OPENFILE, 0, (LPARAM)szDragFilePath);
        return 0;

    case WM_NOTIFY:
        switch(((NMHDR *)lParam)->idFrom) {
        case IDC_TC:
            switch(((NMHDR *)lParam)->code) {
            case TCN_SELCHANGE:
                ShowWindow(hwndStaticPage[iCurPage], SW_HIDE);
                iCurPage = TabCtrl_GetCurSel(hwndTC);
                ShowWindow(hwndStaticPage[iCurPage], SW_SHOW);

                if(mf.fOpened) {
                    switch(iCurPage) {
                    case PAGEEVTLIST:
                        wsprintf(szBuf, L"%u event(s) in total.", cEvtListRow);
                        SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)szBuf);
                        break;
                    case PAGETEMPOLIST:
                        wsprintf(szBuf, L"%u tempo event(s) in total.", cTempoListRow);
                        SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)szBuf);
                        break;
                    default:
                        SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)L"");
                        break;
                    }
                }

                if(fAnlyzTonality) {
                    for(i = 0; i < 24; i++) {
                        ti.uId = i;
                        if(iCurPage == PAGETONALITYANLYZER) {
                            swprintf(szBuf, 128, L"%.2f%%", dTonalityProprtn[i] * 100);
                            GetWindowRect(hwndStaticTonalityBar[i], &ti.rect);
                            MapWindowPoints(HWND_DESKTOP, hwnd, (LPPOINT)&ti.rect, 2);
                            ti.lpszText = szBuf;
                            SendMessage(hwndTool, TTM_ADDTOOL, 0, (LPARAM)&ti);
                        } else {
                            SendMessage(hwndTool, TTM_DELTOOL, 0, (LPARAM)&ti);
                        }
                    }
                }
                break;
            }
            break;

        case IDC_LVEVTLIST:
            switch(((NMHDR *)lParam)->code) {
            /* Event list custom draw */
            case NM_CUSTOMDRAW:
                lplvcd = (LPNMLVCUSTOMDRAW)lParam;
                switch(lplvcd->nmcd.dwDrawStage) {
                case CDDS_PREPAINT:
                    return CDRF_NOTIFYITEMDRAW;
                case CDDS_ITEMPREPAINT:
                    lplvcd->clrText = acrEvtType[EvtGetEvtTypeIndex((EVENT *)(lplvcd->nmcd.lItemlParam))]; // The user data of the list view item is set to the pointer to the corresponding event
                    return CDRF_NEWFONT;
                }
                break;
            }
            break;

        case IDC_TVMUT:
            switch(((NMHDR *)lParam)->code) {
            case NM_CLICK:
            case TVN_KEYDOWN:
                if(((NMHDR *)lParam)->code == NM_CLICK) {
                    GetCursorPos(&tvhti.pt);
                    MapWindowPoints(NULL, hwndTVMut, &tvhti.pt, 1);
                    TreeView_HitTest(hwndTVMut, &tvhti);
                    if(tvhti.flags != TVHT_ONITEMSTATEICON)
                        break;
                    htiTemp = tvhti.hItem;
                } else {
                    if(((NMTVKEYDOWN *)lParam)->wVKey != VK_SPACE)
                        break;
                    if(!(htiTemp = TreeView_GetSelection(hwndTVMut)))
                        break;
                }

                if(!TreeView_GetParent(hwndTVMut, htiTemp)) { // Mute or unmute all
                    f = !TreeView_GetCheckState(hwndTVMut, htiRoot);
                    for(htiTrk = TreeView_GetChild(hwndTVMut, htiRoot), iMutTrk = 0; htiTrk; htiTrk = TreeView_GetNextSibling(hwndTVMut, htiTrk), iMutTrk++) {
                        while(!mf.pwTrkChnUsage[iMutTrk])
                            iMutTrk++;
                        TreeView_SetCheckState(hwndTVMut, htiTrk, f);

                        for(htiChn = TreeView_GetChild(hwndTVMut, htiTrk), iMutChn = 0; htiChn; htiChn = TreeView_GetNextSibling(hwndTVMut, htiChn), iMutChn++) {
                            while(!(mf.pwTrkChnUsage[iMutTrk] & (1 << iMutChn)))
                                iMutChn++;
                            if(f)
                                pwTrkChnUnmuted[iMutTrk] |= (1 << iMutChn);
                            else
                                pwTrkChnUnmuted[iMutTrk] &= ~(1 << iMutChn);
                            TreeView_SetCheckState(hwndTVMut, htiChn, f);
                        }
                    }
                } else if(TreeView_GetParent(hwndTVMut, htiTemp) == htiRoot) { // Mute or unmute a track
                    htiTrk = htiTemp;

                    iTVMutIndex = 0;
                    while(htiTemp = TreeView_GetPrevSibling(hwndTVMut, htiTemp))
                        iTVMutIndex++;
                    for(iMutTrk = 0, i = 0; ; iMutTrk++, i++) {
                        while(!mf.pwTrkChnUsage[iMutTrk])
                            iMutTrk++;
                        if(i == iTVMutIndex)
                            break;
                    }

                    f = !TreeView_GetCheckState(hwndTVMut, htiTrk);
                    for(htiChn = TreeView_GetChild(hwndTVMut, htiTrk), iMutChn = 0; htiChn; htiChn = TreeView_GetNextSibling(hwndTVMut, htiChn), iMutChn++) {
                        while(!(mf.pwTrkChnUsage[iMutTrk] & (1 << iMutChn)))
                            iMutChn++;
                        if(f)
                            pwTrkChnUnmuted[iMutTrk] |= (1 << iMutChn);
                        else
                            pwTrkChnUnmuted[iMutTrk] &= ~(1 << iMutChn);
                        TreeView_SetCheckState(hwndTVMut, htiChn, f);
                    }

                    f = TRUE;
                    for(u = 0; u < mf.cTrk; u++) {
                        if (pwTrkChnUnmuted[u] != 0xFFFF) {
                            f = FALSE;
                            break;
                        }
                    }
                    TreeView_SetCheckState(hwndTVMut, htiRoot, f);
                } else { // Mute or unmute a channel within a track
                    htiTrk = TreeView_GetParent(hwndTVMut, htiTemp);
                    htiChn = htiTemp;

                    iTVMutIndex = 0;
                    htiTemp = htiTrk;
                    while(htiTemp = TreeView_GetPrevSibling(hwndTVMut, htiTemp))
                        iTVMutIndex++;
                    for(iMutTrk = 0, i = 0;; iMutTrk++, i++) {
                        while(!mf.pwTrkChnUsage[iMutTrk])
                            iMutTrk++;
                        if(i == iTVMutIndex)
                            break;
                    }

                    iTVMutIndex = 0;
                    htiTemp = htiChn;
                    while(htiTemp = TreeView_GetPrevSibling(hwndTVMut, htiTemp))
                        iTVMutIndex++;
                    for(iMutChn = 0, i = 0;; iMutChn++, i++) {
                        while(!(mf.pwTrkChnUsage[iMutTrk] & (1 << iMutChn)))
                            iMutChn++;
                        if(i == iTVMutIndex)
                            break;
                    }
                    pwTrkChnUnmuted[iMutTrk] ^= (1 << iMutChn);

                    TreeView_SetCheckState(hwndTVMut, htiTrk, pwTrkChnUnmuted[iMutTrk] == 0xFFFF);

                    f = TRUE;
                    for(u = 0; u < mf.cTrk; u++) {
                        if(pwTrkChnUnmuted[u] != 0xFFFF) {
                            f = FALSE;
                            break;
                        }
                    }
                    TreeView_SetCheckState(hwndTVMut, htiRoot, f);
                }

                if(iCurStrmStatus != STRM_PLAY)
                    break;

                mmt.wType = TIME_MS;
                midiStreamPosition(hms, &mmt, sizeof(MMTIME));
                dwStartTime += mmt.u.ms * iTempoR / 100;
                SendMessage(hwndTBTime, TBM_SETPOS, TRUE, dwStartTime);

                iPlayCtlBufRemaining = 2; // See the process of the MM_MOM_DONE message
                midiStreamStop(hms);
                break;
            }
            break;
        }
        return 0;

    case WM_DRAWITEM:
        lpdis = (LPDRAWITEMSTRUCT)lParam;

        /* Tonality bars custom draw */
        if(wParam >= IDC_STATICTONALITYBARFIRST && wParam < IDC_STATICTONALITYBARFIRST + 24) {
            FillRect(lpdis->hDC, &lpdis->rcItem, (HBRUSH)GetStockObject(WHITE_BRUSH));
            hbr = CreateSolidBrush((wParam - IDC_STATICTONALITYBARFIRST) % 2 == 0 ? RGB(255, 0, 255) : RGB(0, 0, 255));
            lpdis->rcItem.top += (int)((double)(lpdis->rcItem.bottom - lpdis->rcItem.top) * (1. - dTonalityProprtn[wParam - IDC_STATICTONALITYBARFIRST]));
            FillRect(lpdis->hDC, &lpdis->rcItem, hbr);
            DeleteObject(hbr);
        }

        /* Filter combo boxes custom draw */
        if(wParam >= IDC_CBFILTTRK && wParam <= IDC_CBFILTEVTTYPE) {
            hdc = lpdis->hDC;
            if(lpdis->itemState & ODS_COMBOBOXEDIT){
                rcText = lpdis->rcItem;
                if(wParam == IDC_CBFILTTRK)
                    DrawText(hdc, L"Filter by track", -1, &rcText, DT_VCENTER);
                if(wParam == IDC_CBFILTCHN)
                    DrawText(hdc, L"Filter by channel", -1, &rcText, DT_VCENTER);
                if(wParam == IDC_CBFILTEVTTYPE)
                    DrawText(hdc, L"Filter by event type", -1, &rcText, DT_VCENTER);
                return 0;
            }

            if(lpdis->itemID == 0xFFFFFFFF)
                return 0;

            rcText = lpdis->rcItem;
            if(wParam == IDC_CBFILTTRK) {
                lpszBuf = (LPWSTR)malloc((SendMessage(hwndCBFiltTrk, CB_GETLBTEXTLEN, lpdis->itemID, 0) + 1) * sizeof(WCHAR));
                SendMessage(hwndCBFiltTrk, CB_GETLBTEXT, lpdis->itemID, (LPARAM)lpszBuf);
            }
            if(wParam == IDC_CBFILTCHN) {
                lpszBuf = (LPWSTR)malloc((SendMessage(hwndCBFiltChn, CB_GETLBTEXTLEN, lpdis->itemID, 0) + 1) * sizeof(WCHAR));
                SendMessage(hwndCBFiltChn, CB_GETLBTEXT, lpdis->itemID, (LPARAM)lpszBuf);
            }
            if(wParam == IDC_CBFILTEVTTYPE) {
                lpszBuf = (LPWSTR)malloc((SendMessage(hwndCBFiltEvtType, CB_GETLBTEXTLEN, lpdis->itemID, 0) + 1) * sizeof(WCHAR));
                SendMessage(hwndCBFiltEvtType, CB_GETLBTEXT, lpdis->itemID, (LPARAM)lpszBuf);
            }

            if(lpdis->itemID >= 2) {
                rcCheckBox = lpdis->rcItem;
                rcCheckBox.right = rcCheckBox.left + rcCheckBox.bottom - rcCheckBox.top;
                rcText.left = rcCheckBox.right;

                uState = DFCS_BUTTONCHECK;
                if(wParam == IDC_CBFILTTRK) {
                    for(i = 2, uFiltIndex = 0; ; i++, uFiltIndex++) {
                        while(!(pfiltstate->dwFiltTrk[uFiltIndex] & FILT_AVAILABLE))
                            uFiltIndex++;
                        if(i == lpdis->itemID)
                            break;
                    }
                    if(pfiltstate->dwFiltTrk[uFiltIndex] & FILT_CHECKED)
                        uState |= DFCS_CHECKED;
                }
                if(wParam == IDC_CBFILTCHN) {
                    for(i = 2, uFiltIndex = 0; ; i++, uFiltIndex++) {
                        while(!(pfiltstate->dwFiltChn[uFiltIndex] & FILT_AVAILABLE))
                            uFiltIndex++;
                        if(i == lpdis->itemID)
                            break;
                    }
                    if(pfiltstate->dwFiltChn[uFiltIndex] & FILT_CHECKED)
                        uState |= DFCS_CHECKED;
                }
                if(wParam == IDC_CBFILTEVTTYPE) {
                    for(i = 2, uFiltIndex = 0; ; i++, uFiltIndex++) {
                        while(!(pfiltstate->dwFiltEvtType[uFiltIndex] & FILT_AVAILABLE))
                            uFiltIndex++;
                        if(i == lpdis->itemID)
                            break;
                    }
                    if(pfiltstate->dwFiltEvtType[uFiltIndex] & FILT_CHECKED)
                        uState |= DFCS_CHECKED;

                    SetTextColor(hdc, acrEvtType[uFiltIndex]);
                }
                DrawFrameControl(hdc, &rcCheckBox, DFC_BUTTON, uState);
            }
            DrawText(hdc, lpszBuf, -1, &rcText, DT_VCENTER | DT_NOPREFIX);
            free(lpszBuf);
            SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));
        }
        return TRUE;

    case WM_COMMAND:
        if(LOWORD(wParam) >= IDC_EDITCNOTEFIRST && LOWORD(wParam) < IDC_EDITCNOTEFIRST + 12 && HIWORD(wParam) == EN_UPDATE) {
            cNoteTotal = 0;
            for(u = 0; u < 12; u++) {
                cNoteTotal += GetDlgItemInt(hwndStaticPage[PAGETONALITYANLYZER], IDC_EDITCNOTEFIRST + u, NULL, FALSE);
            }
            SetDlgItemInt(hwndStaticPage[PAGETONALITYANLYZER], IDC_EDITCNOTETOTAL, cNoteTotal, FALSE);
        }

        switch(LOWORD(wParam)) {
        case IDC_BTNOPEN:
            if(!GetOpenFileName(&ofn))
                break;

            SendMessage(hwnd, WM_APP_OPENFILE, 0, (LPARAM)szFilePath);
            break;

        case IDC_BTNCLOSE:
            SendMessage(hwnd, WM_APP_CLOSEFILE, 0, 0);
            break;

        case IDC_BTNPLAY:
            if(!mf.fOpened)
                break;

            switch(iCurStrmStatus){
            case STRM_STOP: // Play
                if(!(pevtCurOutput = GetEvtByMs(&mf, dwStartTime, &dwPrevBufEvtTk, &dwCurTempoData)))
                    break;
                SendMessage(hwnd, WM_APP_PLAYFROMEVT, 0, (LPARAM)pevtCurOutput);
                iCurStrmStatus = STRM_PLAY;
                SetWindowText(hwndBtnPlay, L"Pause");
                break;
            case STRM_PLAY: // Pause
                mmt.wType = TIME_MS;
                midiStreamPosition(hms, &mmt, sizeof(MMTIME));
                dwStartTime += mmt.u.ms * iTempoR / 100;
                SendMessage(hwndTBTime, TBM_SETPOS, TRUE, dwStartTime);

                pevtCurOutput = NULL; // See the process of the MM_MOM_POSITIONCB message
                fDone = TRUE;
                midiStreamStop(hms);
                iCurStrmStatus = STRM_STOP;
                SetWindowText(hwndBtnPlay, L"Play");
                break;
            }
            break;

        case IDC_BTNSTOP:
            SendMessage(hwnd, WM_APP_STOP, 0, 0);
            break;

        case IDC_CHECKHEX:
            fHex = !fHex;
            if(!mf.fOpened)
                break;
            i = ListView_GetNextItem(hwndLVEvtList, -1, LVNI_SELECTED);
            iTopIndex = ListView_GetTopIndex(hwndLVEvtList);
            GetEvtList(hwndLVEvtList, &mf, pfiltstate, fHex);
            ListView_SetItemState(hwndLVEvtList, i, LVIS_SELECTED, LVIS_SELECTED);
            ListView_EnsureVisible(hwndLVEvtList, iTopIndex + ListView_GetCountPerPage(hwndLVEvtList) - 1, FALSE);
            i = ListView_GetNextItem(hwndLVTempoList, -1, LVNI_SELECTED);
            iTopIndex = ListView_GetTopIndex(hwndLVTempoList);
            GetTempoList(hwndLVTempoList, &mf, fHex);
            ListView_SetItemState(hwndLVTempoList, i, LVIS_SELECTED, LVIS_SELECTED);
            ListView_EnsureVisible(hwndLVTempoList, iTopIndex + ListView_GetCountPerPage(hwndLVTempoList) - 1, FALSE);
            break;

        case IDC_EDITTRANSP:
        case IDC_EDITTEMPOR:
            if(HIWORD(wParam) != EN_UPDATE)
                break;

            if(iCurStrmStatus == STRM_PLAY) {
                mmt.wType = TIME_MS;
                midiStreamPosition(hms, &mmt, sizeof(MMTIME));
                dwStartTime += mmt.u.ms * iTempoR / 100;
                SendMessage(hwndTBTime, TBM_SETPOS, TRUE, dwStartTime);
            }

            if(LOWORD(wParam) == IDC_EDITTRANSP) {
                iTransp = GetDlgItemInt(hwndStaticPage[PAGEPLAYCTL], IDC_EDITTRANSP, NULL, TRUE);
            } else {
                iTempoR = GetDlgItemInt(hwndStaticPage[PAGEPLAYCTL], IDC_EDITTEMPOR, NULL, FALSE);
                SetTimeText(hwndStaticTime, (double)dwStartTime / 1000, mf.dDur, iTempoR);
            }

            if(iCurStrmStatus != STRM_PLAY)
                break;

            iPlayCtlBufRemaining = 2; // See the process of the MM_MOM_DONE message
            midiStreamStop(hms);
            break;

        case IDC_BTNTRANSPRESET:
            SendMessage(hwndUdTransp, UDM_SETPOS, 0, 0);
            break;

        case IDC_BTNTEMPORRESET:
            SendMessage(hwndUdTempoR, UDM_SETPOS, 0, 100);
            break;

        case IDC_BTNANLYZTONALITY:
            SendMessage(hwnd, WM_APP_ANLYZTONALITY, 0, 0);
            break;

        case ID_PLAY:
            SendMessage(hwndBtnPlay, BM_CLICK, 0, 0);
            break;
        }
        return 0;

    case WM_HSCROLL:
        if((HWND)lParam == hwndTBTime) {
            switch(LOWORD(wParam)) {
            case SB_THUMBTRACK:
                fTBIsTracking = TRUE;
                SetTimeText(hwndStaticTime, (double)SendMessage(hwndTBTime, TBM_GETPOS, 0, 0) / 1000, mf.dDur, iTempoR);
                break;
            case SB_THUMBPOSITION:
                fTBIsTracking = FALSE;
            case SB_PAGELEFT:
            case SB_PAGERIGHT:
                dwStartTime = SendMessage(hwndTBTime, TBM_GETPOS, 0, 0);
                SetTimeText(hwndStaticTime, (double)dwStartTime / 1000, mf.dDur, iTempoR);
                if(iCurStrmStatus != STRM_PLAY)
                    break;

                if(!fTBIsTracking) {
                    iPlayCtlBufRemaining = 2; // See the process of the MM_MOM_DONE message
                    midiStreamStop(hms);
                }
                break;
            }
        }
        return 0;

    case WM_KEYDOWN:
        switch(wParam) {
        case VK_LEFT: // Use left and right keys to control the time bar
            fTBIsTracking = TRUE;
            SendMessage(hwndTBTime, WM_KEYDOWN, VK_PRIOR, lParam);
            break;
        case VK_RIGHT:
            fTBIsTracking = TRUE;
            SendMessage(hwndTBTime, WM_KEYDOWN, VK_NEXT, lParam);
            break;
        }
        return 0;

    case WM_KEYUP:
        switch(wParam) {
        case VK_LEFT:
        case VK_RIGHT:
            SendMessage(hwnd, WM_HSCROLL, SB_THUMBPOSITION, (LPARAM)hwndTBTime);
            break;
        }
        return 0;

    case WM_APP_OPENFILE:
        SendMessage(hwnd, WM_APP_CLOSEFILE, 0, 0);

        if(!ReadMidi((LPWSTR)lParam, &mf)) {
            wsprintf(szBuf, L"Cannot read file \"%s\"!", (LPWSTR)lParam);
            MessageBox(hwnd, szBuf, L"Error", MB_ICONEXCLAMATION);
            break;
        }

        if(szFileName[0] == '\0'){
            i = lstrlen((LPWSTR)lParam);
            while(((LPWSTR)lParam)[i-1] != '\\')
                i--;
            lstrcpy(szFileName, ((LPWSTR)lParam) + i);
        }

        SetWindowText(hwndEditFilePath, (LPWSTR)lParam);

        wsprintf(szBuf, L"Type: %u", mf.wType);
        SetWindowText(hwndStaticType, szBuf);
        wsprintf(szBuf, L"Number of tracks: %u", mf.cTrk);
        SetWindowText(hwndStaticCTrk, szBuf);
        wsprintf(szBuf, L"Timebase: %u", mf.wTb);
        SetWindowText(hwndStaticTb, szBuf);
        wsprintf(szBuf, L"Number of events: %u", mf.cEvt);
        SetWindowText(hwndStaticCEvt, szBuf);

        swprintf(szBuf, 128, L"Initial tempo: %f bpm", mf.dIniTempo);
        SetWindowText(hwndStaticIniTempo, szBuf);
        swprintf(szBuf, 128, L"Average tempo: %f bpm", mf.dAvgTempo);
        SetWindowText(hwndStaticAvgTempo, szBuf);
        swprintf(szBuf, 128, L"Duration: %f s", mf.dDur);
        SetWindowText(hwndStaticDur, szBuf);

        pfiltstate = (FILTERSTATES *)malloc(sizeof(FILTERSTATES) + (mf.cTrk - 1) * sizeof(DWORD));
        for(u = 0; u < mf.cTrk; u++)
            pfiltstate->dwFiltTrk[u] = FILT_CHECKED;
        for(u = 0; u < 17; u++)
            pfiltstate->dwFiltChn[u] = FILT_CHECKED;
        for(u = 0; u < 10; u++)
            pfiltstate->dwFiltEvtType[u] = FILT_CHECKED;

        cEvtListRow = GetEvtList(hwndLVEvtList, &mf, pfiltstate, fHex);
        cTempoListRow = GetTempoList(hwndLVTempoList, &mf, fHex);

        MakeCBFiltLists(&mf, pfiltstate, hwndCBFiltTrk, hwndCBFiltChn, hwndCBFiltEvtType);

        switch(iCurPage) {
        case PAGEEVTLIST:
            wsprintf(szBuf, L"%u event(s) in total.", cEvtListRow);
            SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)szBuf);
            break;
        case PAGETEMPOLIST:
            wsprintf(szBuf, L"%u tempo event(s) in total.", cTempoListRow);
            SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)szBuf);
            break;
        default:
            SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)L"");
            break;
        }

        dwStartTime = 0;
        SendMessage(hwndTBTime, TBM_SETRANGEMAX, TRUE, (int)(mf.dDur * 1000));
        SendMessage(hwndTBTime, TBM_SETPOS, TRUE, 0);
        SetTimeText(hwndStaticTime, 0, mf.dDur, iTempoR);

        pwTrkChnUnmuted = (WORD *)malloc(mf.cTrk * sizeof(WORD));
        tvis.hInsertAfter = TVI_LAST;
        tvis.item.mask = TVIF_STATE | TVIF_TEXT;
        tvis.item.state = TVIS_EXPANDED;
        tvis.item.stateMask = TVIS_EXPANDED;
        tvis.hParent = TVI_ROOT;
        tvis.item.pszText = szFileName;
        htiRoot = TreeView_InsertItem(hwndTVMut, &tvis);
        TreeView_SetCheckState(hwndTVMut, htiRoot, TRUE);
        for(u = 0; u < mf.cTrk; u++) {
            pwTrkChnUnmuted[u] = 0xFFFF;

            if(!mf.pwTrkChnUsage[u])
                continue;
            tvis.hParent = htiRoot;
            if(mf.ppevtTrkName[u]) {
                lpszBuf = (LPWSTR)malloc((mf.ppevtTrkName[u]->cbData + 128) * sizeof(WCHAR));
                ZeroMemory(lpszBuf, (mf.ppevtTrkName[u]->cbData + 128) * sizeof(WCHAR));
                wsprintf(lpszBuf, L"Track #%u (", u);
                MultiByteToWideChar(CP_ACP, MB_COMPOSITE, (LPCSTR)mf.ppevtTrkName[u]->abData, mf.ppevtTrkName[u]->cbData, lpszBuf + lstrlen(lpszBuf), mf.ppevtTrkName[u]->cbData * sizeof(WCHAR));
                lstrcat(lpszBuf, L")");
                tvis.item.pszText = lpszBuf;
                htiTrk = TreeView_InsertItem(hwndTVMut, &tvis);
                free(lpszBuf);
            } else {
                wsprintf(szBuf, L"Track #%u", u);
                tvis.item.pszText = szBuf;
                htiTrk = TreeView_InsertItem(hwndTVMut, &tvis);
            }
            TreeView_SetCheckState(hwndTVMut, htiTrk, TRUE);

            tvis.hParent = htiTrk;
            for(i = 0; i < 16; i++) {
                if(mf.pwTrkChnUsage[u] & (1 << i)) {
                    wsprintf(szBuf, L"Channel #%d", i + 1);
                    tvis.item.pszText = szBuf;
                    htiChn = TreeView_InsertItem(hwndTVMut, &tvis);
                    TreeView_SetCheckState(hwndTVMut, htiChn, TRUE);
                }
            }
        }

        for(i = 0; i < 12; i++)
            SetDlgItemInt(hwndStaticPage[PAGETONALITYANLYZER], IDC_EDITCNOTEFIRST + i, mf.acNote[i], FALSE);
        SendMessage(hwnd, WM_APP_ANLYZTONALITY, 0, 0);

        midiStreamOpen(&hms, &uDevID, 1, (DWORD)hwnd, 0, CALLBACK_WINDOW);
        for(i = 0; i < 2; i++) {
            mhdr[i].lpData = (LPSTR)malloc(STRMBUFLEN);
            mhdr[i].dwBufferLength = STRMBUFLEN;
            mhdr[i].dwFlags = 0;
            midiOutPrepareHeader((HMIDIOUT)hms, &mhdr[i], sizeof(MIDIHDR));
        }
        
        mptd.cbStruct = sizeof(MIDIPROPTIMEDIV);
        mptd.dwTimeDiv = mf.wTb;
        midiStreamProperty(hms, (LPBYTE)&mptd, MIDIPROP_SET | MIDIPROP_TIMEDIV);

        dwPrevBufEvtTk = 0;

        iPlayCtlBufRemaining = 0;
        return 0;

    case WM_APP_CLOSEFILE:
        if(!mf.fOpened)
            break;

        szFileName[0] = '\0';

        cCBFiltItem = SendMessage(hwndCBFiltTrk, CB_GETCOUNT, 0, 0);
        for(i = 2; i < cCBFiltItem; i++)
            SendMessage(hwndCBFiltTrk, CB_DELETESTRING, 2, 0);
        cCBFiltItem = SendMessage(hwndCBFiltChn, CB_GETCOUNT, 0, 0);
        for(i = 2; i < cCBFiltItem; i++)
            SendMessage(hwndCBFiltChn, CB_DELETESTRING, 2, 0);
        cCBFiltItem = SendMessage(hwndCBFiltEvtType, CB_GETCOUNT, 0, 0);
        for(i = 2; i < cCBFiltItem; i++)
            SendMessage(hwndCBFiltEvtType, CB_DELETESTRING, 2, 0);

        SendMessage(hwnd, WM_APP_STOP, 0, 0);
        for(i=0;i<2;i++){
            midiOutUnprepareHeader((HMIDIOUT)hms,&mhdr[i],sizeof(MIDIHDR));
            free(mhdr[i].lpData);
        }
        midiStreamClose(hms);

        free(pfiltstate);

        SetWindowText(hwndEditFilePath, L"");
        FreeMidi(&mf);
        SendMessage(hwndTBTime, TBM_SETRANGEMAX, TRUE, 0);
        SetTimeText(hwndStaticTime, 0, 0, iTempoR);

        SetWindowText(hwndStaticType, L"Type: ");
        SetWindowText(hwndStaticCTrk, L"Number of tracks: ");
        SetWindowText(hwndStaticTb, L"Timebase: ");
        SetWindowText(hwndStaticCEvt, L"Number of events: ");

        SetWindowText(hwndStaticIniTempo, L"Initial tempo: ");
        SetWindowText(hwndStaticAvgTempo, L"Average tempo: ");
        SetWindowText(hwndStaticDur, L"Duration: ");

        ListView_DeleteAllItems(hwndLVEvtList);
        ListView_DeleteAllItems(hwndLVTempoList);

        SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)L"");

        fAnlyzTonality = FALSE;

        for(i = 0; i < 12; i++)
            SetWindowText(hwndEditCNote[i], L"");

        for(i = 0; i < 24; i++) {
            SetWindowText(hwndEditTonality[i], L"");
            dTonalityProprtn[i] = 0;
            InvalidateRect(hwndStaticTonalityBar[i], NULL, TRUE);

            ti.uId = i;
            SendMessage(hwndTool, TTM_DELTOOL, 0, (LPARAM)&ti);
        }

        SetWindowText(hwndStaticMostProbTonality, L"Most probable tonality: ");

        free(pwTrkChnUnmuted);
        TreeView_DeleteAllItems(hwndTVMut);
        return 0;

    case WM_APP_FILTCBITEMCLK:
        if(!mf.fOpened)
            return 0;
        
        cCBFiltItem = SendMessage((HWND)lParam, CB_GETCOUNT, 0, 0);

        if(wParam >= 2) { // Check change
            // Update filter states
            if((HWND)lParam == hwndCBFiltTrk) {
                for(i = 2, uFiltIndex = 0; ; i++, uFiltIndex++) {
                    while(!(pfiltstate->dwFiltTrk[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if(i == wParam)
                        break;
                }
                pfiltstate->dwFiltTrk[uFiltIndex] ^= FILT_CHECKED;
            }
            if((HWND)lParam == hwndCBFiltChn) {
                for(i = 2, uFiltIndex = 0; ; i++, uFiltIndex++) {
                    while(!(pfiltstate->dwFiltChn[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if(i == wParam)
                        break;
                }
                pfiltstate->dwFiltChn[uFiltIndex] ^= FILT_CHECKED;
            }
            if((HWND)lParam == hwndCBFiltEvtType) {
                for(i = 2, uFiltIndex = 0; ; i++, uFiltIndex++) {
                    while(!(pfiltstate->dwFiltEvtType[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if (i == wParam)
                        break;
                }
                pfiltstate->dwFiltEvtType[uFiltIndex] ^= FILT_CHECKED;
            }
        } else { // Check all or check none
            // Update filter states
            if((HWND)lParam == hwndCBFiltTrk) {
                for(uFiltIndex = 0;; uFiltIndex++) {
                    while(uFiltIndex < mf.cTrk && !(pfiltstate->dwFiltTrk[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if(uFiltIndex == mf.cTrk)
                        break;
                    if(!wParam) // Check all
                        pfiltstate->dwFiltTrk[uFiltIndex] |= FILT_CHECKED;
                    else
                        pfiltstate->dwFiltTrk[uFiltIndex] &= ~FILT_CHECKED;
                }
            }
            if((HWND)lParam == hwndCBFiltChn) {
                for(uFiltIndex = 0;; uFiltIndex++) {
                    while(uFiltIndex < 17 && !(pfiltstate->dwFiltChn[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if(uFiltIndex == 17)
                        break;
                    if(!wParam) // Check all
                        pfiltstate->dwFiltChn[uFiltIndex] |= FILT_CHECKED;
                    else
                        pfiltstate->dwFiltChn[uFiltIndex] &= ~FILT_CHECKED;
                }
            }
            if((HWND)lParam == hwndCBFiltEvtType) {
                for(uFiltIndex = 0;; uFiltIndex++) {
                    while(uFiltIndex < 10 && !(pfiltstate->dwFiltEvtType[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if(uFiltIndex == 10)
                        break;
                    if(!wParam) // Check all
                        pfiltstate->dwFiltEvtType[uFiltIndex] |= FILT_CHECKED;
                    else
                        pfiltstate->dwFiltEvtType[uFiltIndex] &= ~FILT_CHECKED;
                }
            }
        }

        iTopIndex = SendMessage((HWND)lParam, CB_GETTOPINDEX, 0, 0);
        SendMessage((HWND)lParam, WM_SETREDRAW, FALSE, 0);

        MakeCBFiltLists(&mf, pfiltstate, hwndCBFiltTrk, hwndCBFiltChn, hwndCBFiltEvtType);

        SendMessage((HWND)lParam, CB_SETTOPINDEX, iTopIndex, 0);
        SendMessage((HWND)lParam, WM_SETREDRAW, TRUE, 0);

        // Get event list again
        cEvtListRow = GetEvtList(hwndLVEvtList, &mf, pfiltstate, fHex);
        wsprintf(szBuf, L"%u event(s) in total.", cEvtListRow);
        SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)szBuf);

        if(iCurStrmStatus == STRM_PLAY) {
            pevtTemp = mf.pevtHead;
            uCurEvtListRow = -1;
            while(pevtTemp != pevtCurOutput) {
                if(IsEvtUnfiltered(pevtTemp, pfiltstate))
                    uCurEvtListRow++;
                pevtTemp = pevtTemp->pevtNext;
            }
            ListView_SetItemState(hwndLVEvtList, uCurEvtListRow, LVIS_SELECTED, LVIS_SELECTED);
            ListView_EnsureVisible(hwndLVEvtList, uCurEvtListRow, FALSE);
        }
        return 0;

    case WM_APP_PLAYFROMEVT:
        mptempo.cbStruct = sizeof(MIDIPROPTEMPO);
        mptempo.dwTempo = max(min(dwCurTempoData * 100 / iTempoR, 0xFFFFFF), 1);
        midiStreamProperty(hms, (LPBYTE)&mptempo, MIDIPROP_SET | MIDIPROP_TEMPO);

        pevtCurBuf = mf.pevtHead;
        uCurOutputEvt = 0;
        uCurEvtListRow = -1;
        uCurTempoListRow = -1;
        while(pevtCurBuf != (EVENT *)lParam) {
            uCurOutputEvt++;
            if(IsEvtUnfiltered(pevtCurBuf, pfiltstate))
                uCurEvtListRow++;
            if(pevtCurBuf->bStatus == 0xFF && pevtCurBuf->bData1 == 0x51)
                uCurTempoListRow++;
            pevtCurBuf = pevtCurBuf->pevtNext;
        }
        ListView_SetItemState(hwndLVEvtList, uCurEvtListRow, LVIS_SELECTED, LVIS_SELECTED);
        ListView_EnsureVisible(hwndLVEvtList, uCurEvtListRow, FALSE);
        ListView_SetItemState(hwndLVTempoList, uCurTempoListRow, LVIS_SELECTED, LVIS_SELECTED);
        ListView_EnsureVisible(hwndLVTempoList, uCurTempoListRow, FALSE);
        fDone = FALSE;

        /* Firstly fill and output the two buffers at once */
        pevtCurBuf = FillStrmBuf(&mhdr[0], pevtCurBuf, &dwPrevBufEvtTk, iTransp, iTempoR, pwTrkChnUnmuted);
        midiStreamOut(hms, &mhdr[0], sizeof(MIDIHDR));
        pevtCurBuf = FillStrmBuf(&mhdr[1], pevtCurBuf, &dwPrevBufEvtTk, iTransp, iTempoR, pwTrkChnUnmuted);
        midiStreamOut(hms, &mhdr[1], sizeof(MIDIHDR));

        midiStreamRestart(hms);
        return 0;

    case MM_MOM_DONE: // When one buffer has been finished outputting
        /* After play controlling while playing, only when the MM_MOM_DONE message from all buffers has been received can restart playing. Use iPlayCtlBufRemaining to mark this */
        if(iPlayCtlBufRemaining) {
            iPlayCtlBufRemaining--;
            if(iPlayCtlBufRemaining == 0) {
                if(!(pevtCurOutput = GetEvtByMs(&mf, dwStartTime, &dwPrevBufEvtTk, &dwCurTempoData)))
                    break;
                SendMessage(hwnd, WM_APP_PLAYFROMEVT, 0, (LPARAM)pevtCurOutput);
            }
            return 0;
        }

        if(fDone)
            return 0;

        /* Refill and output the buffer that has just been finished outputting */
        pevtCurBuf = FillStrmBuf((MIDIHDR *)lParam, pevtCurBuf, &dwPrevBufEvtTk, iTransp, iTempoR, pwTrkChnUnmuted);
        midiStreamOut(hms, (MIDIHDR *)lParam, sizeof(MIDIHDR));
        return 0;

    case MM_MOM_POSITIONCB:
        if(!pevtCurOutput) {
            /* To prevent potentially continue outputting after pausing or stopping, which may cause bugs */
            break;
        }

        uCurOutputEvt++;
        if(IsEvtUnfiltered(pevtCurOutput, pfiltstate)) {
            uCurEvtListRow++;
            ListView_SetItemState(hwndLVEvtList, uCurEvtListRow, LVIS_SELECTED, LVIS_SELECTED);
            ListView_EnsureVisible(hwndLVEvtList, uCurEvtListRow, FALSE);
        }
        if(pevtCurOutput->bStatus == 0xFF && pevtCurOutput->bData1 == 0x51) {
            uCurTempoListRow++;
            ListView_SetItemState(hwndLVTempoList, uCurTempoListRow, LVIS_SELECTED, LVIS_SELECTED);
            ListView_EnsureVisible(hwndLVTempoList, uCurTempoListRow, FALSE);
        }
        if(!fTBIsTracking) {
            mmt.wType = TIME_MS;
            midiStreamPosition(hms, &mmt, sizeof(MMTIME));
            SendMessage(hwndTBTime, TBM_SETPOS, TRUE, mmt.u.ms * iTempoR / 100 + dwStartTime);
            SetTimeText(hwndStaticTime, (double)(mmt.u.ms * iTempoR / 100 + dwStartTime) / 1000, mf.dDur, iTempoR);
        }

        pevtCurOutput = pevtCurOutput->pevtNext;

        if(uCurOutputEvt == mf.cEvt) { // All events have been finished outputting
            SendMessage(hwnd, WM_APP_STOP, 0, 0);
        }
        return 0;

    case WM_APP_STOP:
        if(!mf.fOpened)
            return 0;

        pevtCurOutput = NULL; // See the process of the MM_MOM_POSITIONCB message
        fDone = TRUE;
        midiStreamStop(hms);
        if(iCurStrmStatus != STRM_STOP)
            SetWindowText(hwndBtnPlay, L"Play");
        iCurStrmStatus = STRM_STOP;
        dwStartTime = 0;
        SendMessage(hwndTBTime, TBM_SETPOS, TRUE, 0);
        SetTimeText(hwndStaticTime, 0, mf.dDur, iTempoR);
        return 0;

    case WM_APP_ANLYZTONALITY:
        fAnlyzTonality = TRUE;
        dTonalityProprtnMax = 0;

        for(i = 0; i < 12; i++)
            acNote[i] = GetDlgItemInt(hwndStaticPage[PAGETONALITYANLYZER], IDC_EDITCNOTEFIRST + i, NULL, FALSE);

        AnalyzeTonality(acNote, dTonalityProprtn, fTonalityMax);
        for(i = 0; i < 24; i++) {
            swprintf(szBuf, 128, L"%.2f%%", dTonalityProprtn[i] * 100);
            SetWindowText(hwndEditTonality[i], szBuf);
            InvalidateRect(hwndStaticTonalityBar[i], NULL, TRUE);

            if(iCurPage == PAGETONALITYANLYZER) {
                ti.uId = i;
                SendMessage(hwndTool, TTM_DELTOOL, 0, (LPARAM)&ti);
                GetWindowRect(hwndStaticTonalityBar[i], &ti.rect);
                MapWindowPoints(HWND_DESKTOP, hwnd, (LPPOINT)&ti.rect, 2);
                ti.lpszText = szBuf;
                SendMessage(hwndTool, TTM_ADDTOOL, 0, (LPARAM)&ti);
            }
        }

        for(i = 0; i < 24; i++) {
            if(fTonalityMax[i]) {
                if(dTonalityProprtnMax == 0) {
                    wsprintf(szBuf, L"Most probable tonality: %s", alpszTonalityLabel[i]);
                    dTonalityProprtnMax = dTonalityProprtn[i];
                } else {
                    wsprintf(szBuf, L"%s, %s", szBuf, alpszTonalityLabel[i]);
                }
            }
        }
        if(dTonalityProprtnMax != 0) {
            swprintf(szBuf, 128, L"%s (%.2f%%)", szBuf, dTonalityProprtnMax * 100);
            SetWindowText(hwndStaticMostProbTonality, szBuf);
        } else {
            SetWindowText(hwndStaticMostProbTonality, L"Most probable tonality: ");
        }
        return 0;

    case WM_DESTROY:
        SendMessage(hwnd, WM_APP_CLOSEFILE, 0, 0);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}


int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int iCmdShow) {
    HWND hwnd;
    MSG msg;
    WNDCLASS wc;
    int iArgc;
    LPWSTR *lplpszArgv;

    HACCEL haccel;

    HWND hwndFocused;
    WCHAR szBuf[128];

    lplpszArgv = CommandLineToArgvW(GetCommandLine(), &iArgc);
    if(iArgc > 1) {
        lstrcpy(szCmdFilePath, lplpszArgv[1]);
    }
    LocalFree(lplpszArgv);

    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIcon(hInstance, L"AnkeMidi");
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = L"AnkeMidi";
    if(!RegisterClass(&wc))
        return -1;
    
    hwnd = CreateWindowEx(WS_EX_ACCEPTFILES, L"AnkeMidi", L"Anke MIDI Reader", 
        WS_OVERLAPPEDWINDOW | WS_VISIBLE, 
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 
        NULL, NULL, hInstance, NULL);
    if(!hwnd)
        return -1;

    haccel = LoadAccelerators(hInstance, L"AnkeMidi");

    hInstanceMain = hInstance;
    hwndMain = hwnd;
    
    while(GetMessage(&msg, NULL, 0, 0)) {
        if(!TranslateAccelerator(hwnd, haccel, &msg)) {
            /* Handle Ctrl+A for edit controls */
            if(msg.message == WM_KEYDOWN && msg.wParam == 'A' && GetKeyState(VK_CONTROL) < 0) {
                hwndFocused = GetFocus();
                GetClassName(hwndFocused, szBuf, 128);
                if(hwndFocused && lstrcmp(szBuf, L"Edit") == 0){
                    SendMessage(hwndFocused, EM_SETSEL, 0, -1);
                    continue;
                }
            }

            /* Pass key messages to the main window if the focus is not on Edit or Tab Control */
            if(msg.message == WM_KEYDOWN || msg.message == WM_KEYUP) {
                hwndFocused = GetFocus();
                GetClassName(hwndFocused, szBuf, 128);
                if(!(hwndFocused && (lstrcmp(szBuf, L"Edit") == 0 || lstrcmp(szBuf, WC_TABCONTROL) == 0)))
                    msg.hwnd = hwndMain;
            }

            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    
    return msg.wParam;
}