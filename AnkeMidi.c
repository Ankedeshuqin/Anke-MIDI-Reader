#define _WIN32_WINNT 0x0400
#include <Windows.h>
#include "cbinfo.h"
#include <CommCtrl.h>
#include <stdio.h> // For float number outputting
#include "StrList.h"
#include "MidiRead.h"
#include "resource.h"

#pragma comment(lib, "comctl32.lib")

#define DLUX(x) MulDiv(cxChar, x, 4) // x in DLUs
#define DLUY(y) MulDiv(cyChar, y, 8) // y in DLUs


#define CPAGE 4

enum ePage {
    PAGEEVTLIST, PAGETEMPOLIST, PAGEPLAYCTL, PAGETONALITYANLYZER
};

enum eIDC {
    /* Controls of common pages */
    IDC_STATUS, IDC_TC, IDC_BTNOPEN, IDC_BTNCLOSE, IDC_EDITFILEPATH, IDC_BTNPLAY, IDC_BTNSTOP, IDC_TBTIME,
    /* Controls of event list page */
    IDC_STATICTYPE, IDC_STATICCTRK, IDC_STATICTB, IDC_STATICCEVT, IDC_CBFILTTRK, IDC_CBFILTCHN, IDC_CBFILTEVTTYPE, IDC_LVEVTLIST,
    /* Controls of tempo list page */
    IDC_STATICINITEMPO, IDC_STATICAVGTEMPO, IDC_STATICDUR, IDC_STATICEVTDENSITY, IDC_LVTEMPOLIST,
    /* Controls of play control page */
    IDC_EDITTRANSP, IDC_UDTRANSP, IDC_BTNTRANSPRESET, IDC_EDITTEMPOR, IDC_UDTEMPOR, IDC_BTNTEMPORRESET, IDC_TVMUT,
    /* Controls of tonality analyzer page */
    IDC_EDITCNOTEFIRST, IDC_EDITCNOTETOTAL = IDC_EDITCNOTEFIRST + 12, IDC_BTNANLYZTONALITY, IDC_STATICMOSTPROBTONALITY, IDC_EDITTONALITYFIRST, IDC_STATICTONALITYBARFIRST = IDC_EDITTONALITYFIRST + 24, IDC_TTTONALITYBARFIRST = IDC_STATICTONALITYBARFIRST + 24, IDC_STATICTONALITYBARLABELFIRST = IDC_STATICTONALITYBARFIRST + 24
};


enum eAppMsg {
    WM_APP_OPENFILE = WM_APP, // lParam: lpszFilePath
    WM_APP_CLOSEFILE,
    WM_APP_PLAYFROMEVT, // lParam: pevt
    WM_APP_STOP,
    WM_APP_FILTCBITEMCLK, // wParam: iIndex; lParam: hwndCheckedCB
    WM_APP_ANLYZTONALITY
};


/* Filter flags */
#define FILT_CHECKED 1
#define FILT_AVAILABLE 0x10000

typedef struct filtstate {
    DWORD dwFiltChn[17];
    DWORD dwFiltEvtType[10];
    DWORD dwFiltTrk[1];
} FILTERSTATES;

int EvtGetFiltTrkIndex(EVENT *pevt) {
    return pevt->wTrk;
}

int EvtGetFiltChnIndex(EVENT *pevt) {
    if(pevt->bStatus < 0xF0)
        return (pevt->bStatus & 0xF) + 1;
    return 0;
}

int EvtGetFiltEvtTypeIndex(EVENT *pevt) {
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

void MakeCBFiltLists(MIDIFILE *pmf, FILTERSTATES *pfiltstate, HWND hwndCBFiltTrk, HWND hwndCBFiltChn, HWND hwndCBFiltEvtType) {
    int i, cCBFiltItem;
    UINT uFiltIndex;
    WCHAR szBuf[128];

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
            wsprintf(szBuf, L"Track #%u", uFiltIndex);
            SendMessage(hwndCBFiltTrk, CB_ADDSTRING, 0, (LPARAM)szBuf);
            SendMessage(hwndCBFiltTrk, CB_SETITEMDATA, i, pfiltstate->dwFiltTrk[uFiltIndex] & FILT_CHECKED);
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
            SendMessage(hwndCBFiltChn, CB_SETITEMDATA, i, pfiltstate->dwFiltChn[uFiltIndex] & FILT_CHECKED);
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
            SendMessage(hwndCBFiltEvtType, CB_SETITEMDATA, i, pfiltstate->dwFiltEvtType[uFiltIndex] & FILT_CHECKED);
            i++;
        }
    }
}

DWORD GetEvtList(HWND hwndLVEvtList, MIDIFILE *pmf, FILTERSTATES *pfiltstate, COLORREF *pcrLVEvtListCD) {
    LVITEM lvitem;

    EVENT *pevtCur;
    DWORD dwRow = 0;
    WCHAR szBuf[128];
    LPWSTR lpszData, lpszDataCmt; // For meta or sys-ex events

    int cSharp = 0; // For key signature events

    int iFiltTrkIndex, iFiltChnIndex, iFiltEvtTypeIndex;
    UINT uFiltIndex, u;

    for(uFiltIndex = 0; uFiltIndex < pmf->cTrk; uFiltIndex++)
        pfiltstate->dwFiltTrk[uFiltIndex] &= ~FILT_AVAILABLE;
    for(uFiltIndex = 0; uFiltIndex < 17; uFiltIndex++)
        pfiltstate->dwFiltChn[uFiltIndex] &= ~FILT_AVAILABLE;
    for(uFiltIndex = 0; uFiltIndex < 10; uFiltIndex++)
        pfiltstate->dwFiltEvtType[uFiltIndex] &= ~FILT_AVAILABLE;

    SendMessage(hwndLVEvtList, WM_SETREDRAW, FALSE, 0);
    ListView_DeleteAllItems(hwndLVEvtList);

    lvitem.mask = LVIF_TEXT;
    pevtCur = pmf->pevtHead;
    while(pevtCur) {
        iFiltTrkIndex = EvtGetFiltTrkIndex(pevtCur);
        iFiltChnIndex = EvtGetFiltChnIndex(pevtCur);
        iFiltEvtTypeIndex = EvtGetFiltEvtTypeIndex(pevtCur);
        if(
            !(pfiltstate->dwFiltTrk[iFiltTrkIndex] & FILT_CHECKED) ||
            !(pfiltstate->dwFiltChn[iFiltChnIndex] & FILT_CHECKED) ||
            !(pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] & FILT_CHECKED)
        ) {
            if(
                !(pfiltstate->dwFiltTrk[iFiltTrkIndex] & FILT_CHECKED) &&
                pfiltstate->dwFiltChn[iFiltChnIndex] & FILT_CHECKED &&
                pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] & FILT_CHECKED
            ) {
                pfiltstate->dwFiltTrk[iFiltTrkIndex] |= FILT_AVAILABLE;
            }
            if(
                pfiltstate->dwFiltTrk[iFiltTrkIndex] & FILT_CHECKED &&
                !(pfiltstate->dwFiltChn[iFiltChnIndex] & FILT_CHECKED) &&
                pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] & FILT_CHECKED
            ) {
                pfiltstate->dwFiltChn[iFiltChnIndex] |= FILT_AVAILABLE;
            }
            if(
                pfiltstate->dwFiltTrk[iFiltTrkIndex] & FILT_CHECKED &&
                pfiltstate->dwFiltChn[iFiltChnIndex] & FILT_CHECKED &&
                !(pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] & FILT_CHECKED)
            ) {
                pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] |= FILT_AVAILABLE;
            }

            pevtCur->fUnfiltered = FALSE;
            pevtCur = pevtCur->pevtNext;
            continue;
        }

        lvitem.iItem = dwRow;
        lvitem.iSubItem = 0;
        wsprintf(szBuf, L"%u", dwRow + 1);
        lvitem.pszText = szBuf;
        ListView_InsertItem(hwndLVEvtList, &lvitem);

        pfiltstate->dwFiltTrk[pevtCur->wTrk] |= FILT_AVAILABLE;
        wsprintf(szBuf, L"%u", pevtCur->wTrk);
        ListView_SetItemText(hwndLVEvtList, dwRow, 1, szBuf);

        wsprintf(szBuf, L"%u", pevtCur->dwTk);
        ListView_SetItemText(hwndLVEvtList, dwRow, 3, szBuf);

        if(pevtCur->bStatus < 0xF0) {
            pfiltstate->dwFiltChn[(pevtCur->bStatus & 0xF) + 1] |= FILT_AVAILABLE;
            wsprintf(szBuf, L"%u", (pevtCur->bStatus & 0xF) + 1);
            ListView_SetItemText(hwndLVEvtList, dwRow, 2, szBuf);

            switch(pevtCur->bStatus >> 4) {
            case 0x8:
                pfiltstate->dwFiltEvtType[0] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow] = RGB(128, 128, 128);
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Note off");
                wsprintf(szBuf, L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszNote[pevtCur->bData1]);
                wsprintf(szBuf, L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);
                break;
            case 0x9:
                if(pevtCur->bData2 == 0) {
                    pfiltstate->dwFiltEvtType[1] |= FILT_AVAILABLE;
                    pcrLVEvtListCD[dwRow] = RGB(128, 128, 128);
                } else {
                    pfiltstate->dwFiltEvtType[2] |= FILT_AVAILABLE;
                    pcrLVEvtListCD[dwRow] = RGB(0, 0, 0);
                }
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Note on");
                wsprintf(szBuf, L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszNote[pevtCur->bData1]);
                wsprintf(szBuf, L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);
                break;
            case 0xA:
                pfiltstate->dwFiltEvtType[3] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow] = RGB(0, 0, 128);
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Note aftertouch");
                wsprintf(szBuf, L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszNote[pevtCur->bData1]);
                wsprintf(szBuf, L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);
                break;
            case 0xB:
                pfiltstate->dwFiltEvtType[4] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow] = RGB(0, 128, 0);
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Controller");
                wsprintf(szBuf, L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszCtl[pevtCur->bData1]);
                wsprintf(szBuf, L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);
                break;
            case 0xC:
                pfiltstate->dwFiltEvtType[5] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow] = RGB(192, 128, 0);
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Program change");
                wsprintf(szBuf, L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszPrg[pevtCur->bData1]);
                break;
            case 0xD:
                pfiltstate->dwFiltEvtType[6] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow] = RGB(0, 0, 255);
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Channel aftertouch");
                wsprintf(szBuf, L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                break;
            case 0xE:
                pfiltstate->dwFiltEvtType[7] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow] = RGB(255, 0, 255);
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Pitch bend");
                wsprintf(szBuf, L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                wsprintf(szBuf, L"%u", pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList, dwRow, 7, szBuf);

                /* Generate comment string for pitch bend events */
                wsprintf(szBuf, L"%d", (int)pevtCur->bData1 | (pevtCur->bData2 << 7) - 8192);
                ListView_SetItemText(hwndLVEvtList, dwRow, 8, szBuf);
                break;
            }
        } else {
            pfiltstate->dwFiltChn[0] |= FILT_AVAILABLE;
            lpszData = (LPWSTR)malloc((pevtCur->cbData * 4 + 1) * sizeof(WCHAR));
            lpszData[0] = '\0';
            for(u = 0; u < pevtCur->cbData; u++) {
                wsprintf(lpszData, L"%s%u", lpszData, pevtCur->abData[u]);
                if(u < pevtCur->cbData - 1)
                    wsprintf(lpszData, L"%s ", lpszData);
            }
            ListView_SetItemText(hwndLVEvtList, dwRow, 7, lpszData);
            free(lpszData);

            switch(pevtCur->bStatus) {
            case 0xF0:
            case 0xF7:
                pfiltstate->dwFiltEvtType[8] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow] = RGB(255, 0, 0);
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"System exclusive");
                break;
            case 0xFF:
                pfiltstate->dwFiltEvtType[9] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow] = RGB(192, 128, 128);
                ListView_SetItemText(hwndLVEvtList, dwRow, 4, L"Meta event");
                wsprintf(szBuf, L"%u", pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList, dwRow, 5, szBuf);
                ListView_SetItemText(hwndLVEvtList, dwRow, 6, alpszMeta[pevtCur->bData1]);

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
        
        pevtCur->fUnfiltered = TRUE;
        pevtCur = pevtCur->pevtNext;
        dwRow++;
    }
    
    SendMessage(hwndLVEvtList, WM_SETREDRAW, TRUE, 0);
    return dwRow;
}

DWORD GetTempoList(HWND hwndLVTempoList, MIDIFILE *pmf) {
    LVITEM lvitem;

    TEMPOEVENT *ptempoevtCur;
    DWORD dwRow = 0;
    WCHAR szBuf[128];

    SendMessage(hwndLVTempoList, WM_SETREDRAW, FALSE, 0);

    lvitem.mask = LVIF_TEXT;
    ptempoevtCur = pmf->ptempoevtHead;
    while(ptempoevtCur) {
        lvitem.iItem = dwRow;
        lvitem.iSubItem = 0;
        wsprintf(szBuf, L"%u", dwRow + 1);
        lvitem.pszText = szBuf;
        ListView_InsertItem(hwndLVTempoList, &lvitem);

        wsprintf(szBuf, L"%u", ptempoevtCur->dwTk);
        ListView_SetItemText(hwndLVTempoList, dwRow, 1, szBuf);

        wsprintf(szBuf, L"%u", ptempoevtCur->cTk);
        ListView_SetItemText(hwndLVTempoList, dwRow, 2, szBuf);

        wsprintf(szBuf, L"%u", ptempoevtCur->dwData);
        ListView_SetItemText(hwndLVTempoList, dwRow, 3, szBuf);

        swprintf(szBuf, 128, L"%f bpm", 60000000. / ptempoevtCur->dwData);
        ListView_SetItemText(hwndLVTempoList, dwRow, 4, szBuf);

        ptempoevtCur = ptempoevtCur->ptempoevtNext;
        dwRow++;
    }

    SendMessage(hwndLVTempoList, WM_SETREDRAW, TRUE, 0);
    return dwRow;
}


void AnalyzeTonality(UINT *pcNote, double *pdTonalityProprtn, BOOL *pfTonalityMax) {
    const int aiMajScaleNote[] = {0, 2, 4, 5, 7, 9, 11};
    const int aiMinScaleNote[] = {0, 2, 3, 5, 7, 8, 10, 11};
    const int aiMajScaleNoteWeight[] = {3, 3, 3, 3, 3, 3, 3};
    const int aiMinScaleNoteWeight[] = {3, 3, 3, 3, 3, 3, 1, 2};

    double adTonality[24];
    double dTonalityTotal = 0, dTonalityMax = 0;

    int iTonalPivot, iFifthPivot, iMajThirdPivot, iMinThirdPivot, iCurScaleNotePivot;
    double dMajThirdProprtn, dMinThirdProprtn;
    UINT cMajThird, cMinThird, cTonalFifth, cMajScaleNote, cMinScaleNote;

    int i;

    for(iTonalPivot = 0; iTonalPivot < 12; iTonalPivot++) {
        if(pcNote[iTonalPivot] > 0) {
            iFifthPivot = (iTonalPivot + 7) % 12;
            cTonalFifth = pcNote[iTonalPivot] + pcNote[iFifthPivot];

            iMajThirdPivot = (iTonalPivot + 4) % 12;
            iMinThirdPivot = (iTonalPivot + 3) % 12;
            cMajThird = pcNote[iMajThirdPivot];
            cMinThird = pcNote[iMinThirdPivot];
            if(cMajThird + cMinThird) {
                dMajThirdProprtn = (double)cMajThird / (cMajThird + cMinThird);
                dMinThirdProprtn = (double)cMinThird / (cMajThird + cMinThird);
            } else {
                dMajThirdProprtn = 0.5;
                dMinThirdProprtn = 0.5;
            }

            cMajScaleNote = 0;
            cMinScaleNote = 0;
            for(i = 0; i < 7; i++) {
                iCurScaleNotePivot = (iTonalPivot + aiMajScaleNote[i]) % 12;
                cMajScaleNote += pcNote[iCurScaleNotePivot] * aiMajScaleNoteWeight[i];
            }
            for(i = 0; i < 8; i++) {
                iCurScaleNotePivot = (iTonalPivot + aiMinScaleNote[i]) % 12;
                cMinScaleNote += pcNote[iCurScaleNotePivot] * aiMinScaleNoteWeight[i];
            }

            adTonality[iTonalPivot * 2] = (double)cTonalFifth * dMajThirdProprtn * cMajScaleNote; // Value for major tonalities
            adTonality[iTonalPivot * 2 + 1] = (double)cTonalFifth * dMinThirdProprtn * cMinScaleNote; // Value for minor tonalities

            dTonalityTotal += adTonality[iTonalPivot * 2] + adTonality[iTonalPivot * 2 + 1];
        } else {
            adTonality[iTonalPivot * 2] = 0;
            adTonality[iTonalPivot * 2 + 1] = 0;
        }
    }

    ZeroMemory(pfTonalityMax, 24 * sizeof(BOOL));
    for(i = 0; i < 24; i++) {
        if(dTonalityTotal != 0) {
            pdTonalityProprtn[i] = adTonality[i] / dTonalityTotal;
            if(adTonality[i] > dTonalityMax) {
                dTonalityMax = adTonality[i];
                ZeroMemory(pfTonalityMax, 24 * sizeof(BOOL));
                pfTonalityMax[i] = TRUE;
            } else if(adTonality[i] == dTonalityMax) {
                pfTonalityMax[i] = TRUE;
            }
        } else {
            pdTonalityProprtn[i] = 0;
        }
    }
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

WNDPROC DefTCProc, DefStaticProc, DefLBProc;

LRESULT CALLBACK TempWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch(uMsg) {
    case WM_COMMAND:
    case WM_NOTIFY:
    case WM_DRAWITEM:
    case WM_HSCROLL:
    case WM_APP_FILTCBITEMCLK:
        return SendMessage(hwndMain, uMsg, wParam, lParam);
    }
    return CallWindowProc(GetWindowLong(hwnd, GWL_ID) == IDC_TC ? DefTCProc : DefStaticProc, hwnd, uMsg, wParam, lParam);
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
    static RECT rcTCDisp;
    static int iCurPage = 0;
    static HWND hwndStaticPage[CPAGE];

    static HWND hwndBtnOpen, hwndBtnClose, hwndEditFilePath, hwndBtnPlay, hwndBtnStop, hwndTBTime,
        hwndStaticType, hwndStaticCTrk, hwndStaticTb, hwndStaticCEvt, hwndCBFiltTrk, hwndCBFiltChn, hwndCBFiltEvtType, hwndLBCBFiltTrk, hwndLBCBFiltChn, hwndLBCBFiltEvtType, hwndLVEvtList,
        hwndStaticIniTempo, hwndStaticAvgTempo, hwndStaticDur, hwndStaticEvtDensity, hwndLVTempoList,
        hwndEditTransp, hwndUdTransp, hwndBtnTranspReset, hwndEditTempoR, hwndUdTempoR, hwndBtnTempoRReset, hwndTVMut,
        hwndEditCNote[12], hwndEditCNoteTotal, hwndBtnAnlyzTonality, hwndStaticMostProbTonality, hwndEditTonality[24], hwndStaticTonalityBar[24], hwndStaticTonalityBarLabel[24];
    HWND hwndTemp;

    COMBOBOXINFO cbi;

    LVCOLUMN lvcol;


    static MIDIFILE mf;
    
    static FILTERSTATES *pfiltstate;
    LPDRAWITEMSTRUCT lpdis;
    RECT rcCheckBox, rcText;
    UINT uState;
    WCHAR szBuf[128];
    int i, cCBFiltItem;
    UINT uFiltIndex;
    static DWORD cEvtListRow = 0, cTempoListRow = 0;
    static COLORREF *pcrLVEvtListCD;
    int iCBFiltTopIndex;
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

    
    switch(uMsg){
    case WM_CREATE:
        InitCommonControls();

        osversioninfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
        GetVersionEx(&osversioninfo);

        ncm.cbSize = sizeof(NONCLIENTMETRICS);
        if(osversioninfo.dwMajorVersion < 6)
            ncm.cbSize -= sizeof(int);
        SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
        hfontGUI = CreateFontIndirect(&ncm.lfMessageFont);
        hdc = GetDC(hwnd);
        SelectObject(hdc, hfontGUI);
        GetTextMetrics(hdc, &tm);
        cxChar = tm.tmAveCharWidth;
        cyChar = tm.tmHeight + tm.tmExternalLeading;
        ReleaseDC(hwnd, hdc);


        /* Controls for common pages */
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

        hwndTC = CreateWindow(WC_TABCONTROL, NULL,
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
            0, 0, 0, 0,
            hwnd, (HMENU)IDC_TC, hInstanceMain, NULL);
        SendMessage(hwndTC, WM_SETFONT, (WPARAM)hfontGUI, TRUE);
        DefTCProc = (WNDPROC)GetWindowLong(hwndTC, GWL_WNDPROC);
        SetWindowLong(hwndTC, GWL_WNDPROC, (LONG)TempWndProc);

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
            SetWindowLong(hwndStaticPage[i], GWL_WNDPROC, (LONG)TempWndProc);
        }
        iCurPage = PAGEEVTLIST;
        ShowWindow(hwndStaticPage[iCurPage], SW_SHOW);

        GetClientRect(hwnd, &rcTCDisp);
        TabCtrl_AdjustRect(hwndTC, FALSE, &rcTCDisp);

        hwndBtnOpen = CreateWindow(L"BUTTON", L"Open",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rcTCDisp.left + DLUX(7), rcTCDisp.top + DLUY(7), DLUX(28), DLUY(14),
            hwndTC, (HMENU)IDC_BTNOPEN, hInstanceMain, NULL);
        SendMessage(hwndBtnOpen, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnClose = CreateWindow(L"BUTTON", L"Close",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rcTCDisp.left + DLUX(39), rcTCDisp.top + DLUY(7), DLUX(28), DLUY(14),
            hwndTC, (HMENU)IDC_BTNCLOSE, hInstanceMain, NULL);
        SendMessage(hwndBtnClose, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndEditFilePath = CreateWindow(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_READONLY | ES_AUTOHSCROLL,
            0, 0, 0, 0,
            hwndTC, (HMENU)IDC_EDITFILEPATH, hInstanceMain, NULL);
        SendMessage(hwndEditFilePath, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnPlay = CreateWindow(L"BUTTON", L"Play",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rcTCDisp.left + DLUX(7), rcTCDisp.top + DLUY(25), DLUX(28), DLUY(14),
            hwndTC, (HMENU)IDC_BTNPLAY, hInstanceMain, NULL);
        SendMessage(hwndBtnPlay, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnStop = CreateWindow(L"BUTTON", L"Stop",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rcTCDisp.left + DLUX(39), rcTCDisp.top + DLUY(25), DLUX(28), DLUY(14),
            hwndTC, (HMENU)IDC_BTNSTOP, hInstanceMain, NULL);
        SendMessage(hwndBtnStop, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndTBTime = CreateWindow(TRACKBAR_CLASS, NULL,
            WS_CHILD | WS_VISIBLE,
            0, 0, 0, 0,
            hwndTC, (HMENU)IDC_TBTIME, hInstanceMain, NULL);
        SendMessage(hwndTBTime, TBM_SETRANGEMIN, TRUE, 0);
        SendMessage(hwndTBTime, TBM_SETRANGEMAX, TRUE, 0);
        SendMessage(hwndTBTime, TBM_SETPOS, TRUE, 0);

        /* Controls for event list page */
        hwndStaticType = CreateWindow(L"STATIC", L"Type: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(7), DLUY(46), DLUX(56), DLUY(8),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_STATICTYPE, hInstanceMain, NULL);
        SendMessage(hwndStaticType, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticCTrk = CreateWindow(L"STATIC", L"Number of tracks: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(67), DLUY(46), DLUX(104), DLUY(8),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_STATICCTRK, hInstanceMain, NULL);
        SendMessage(hwndStaticCTrk, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticTb = CreateWindow(L"STATIC", L"Timebase: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(175), DLUY(46), DLUX(72), DLUY(8),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_STATICTB, hInstanceMain, NULL);
        SendMessage(hwndStaticTb, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticCEvt = CreateWindow(L"STATIC", L"Number of events: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(251), DLUY(46), DLUX(104), DLUY(8),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_STATICCEVT, hInstanceMain, NULL);
        SendMessage(hwndStaticCEvt, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);

        hwndCBFiltTrk = CreateWindow(L"COMBOBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            DLUX(7), DLUY(61), DLUX(122), DLUY(162),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_CBFILTTRK, hInstanceMain, NULL);
        SendMessage(hwndCBFiltTrk, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndCBFiltChn = CreateWindow(L"COMBOBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            DLUX(133), DLUY(61), DLUX(122), DLUY(162),
            hwndStaticPage[PAGEEVTLIST], (HMENU)IDC_CBFILTCHN, hInstanceMain, NULL);
        SendMessage(hwndCBFiltChn, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndCBFiltEvtType = CreateWindow(L"COMBOBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            DLUX(259), DLUY(61), DLUX(122), DLUY(162),
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
        lvcol.pszText = L"Comment";
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
            DLUX(7), DLUY(46), DLUX(136), DLUY(8),
            hwndStaticPage[PAGETEMPOLIST], (HMENU)IDC_STATICINITEMPO, hInstanceMain, NULL);
        SendMessage(hwndStaticIniTempo, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticAvgTempo = CreateWindow(L"STATIC", L"Average tempo: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(147), DLUY(46), DLUX(136), DLUY(8),
            hwndStaticPage[PAGETEMPOLIST], (HMENU)IDC_STATICAVGTEMPO, hInstanceMain, NULL);
        SendMessage(hwndStaticAvgTempo, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticDur = CreateWindow(L"STATIC", L"Duration: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(287), DLUY(46), DLUX(108), DLUY(8),
            hwndStaticPage[PAGETEMPOLIST], (HMENU)IDC_STATICDUR, hInstanceMain, NULL);
        SendMessage(hwndStaticDur, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndStaticEvtDensity = CreateWindow(L"STATIC", L"Event density: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(399), DLUY(46), DLUX(144), DLUY(8),
            hwndStaticPage[PAGETEMPOLIST], (HMENU)IDC_STATICEVTDENSITY, hInstanceMain, NULL);
        SendMessage(hwndStaticEvtDensity, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);

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
            DLUX(7), DLUY(49), DLUX(64), DLUY(8),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndEditTransp = CreateWindow(L"EDIT", NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            DLUX(74), DLUY(46), DLUX(34), DLUY(14),
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
            DLUX(112), DLUY(46), DLUX(28), DLUY(14),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_BTNTRANSPRESET, hInstanceMain, NULL);
        SendMessage(hwndBtnTranspReset, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndTemp = CreateWindow(L"STATIC", L"Tempo ratio:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(159), DLUY(49), DLUX(56), DLUY(8),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndEditTempoR = CreateWindow(L"EDIT", NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            DLUX(218), DLUY(46), DLUX(34), DLUY(14),
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
            DLUX(255), DLUY(49), DLUX(12), DLUY(8),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnTempoRReset = CreateWindow(L"BUTTON", L"Reset",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(267), DLUY(46), DLUX(28), DLUY(14),
            hwndStaticPage[PAGEPLAYCTL], (HMENU)IDC_BTNTEMPORRESET, hInstanceMain, NULL);
        SendMessage(hwndBtnTempoRReset, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndTemp = CreateWindow(L"STATIC", L"Muting:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            DLUX(7), DLUY(67), DLUX(36), DLUY(8),
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
            DLUX(7), DLUY(46), DLUX(68), DLUY(8),
            hwndStaticPage[PAGETONALITYANLYZER], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        for(i = 0; i < 12; i++) {
            hwndTemp = CreateWindow(L"STATIC", alpszNoteLabel[i],
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                DLUX(7 + 36 * i), DLUY(57), DLUX(32), DLUY(8),
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)-1, hInstanceMain, NULL);
            SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
            hwndEditCNote[i] = CreateWindow(L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_AUTOHSCROLL,
                DLUX(7 + 36 * i), DLUY(68), DLUX(32), DLUY(14),
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)(IDC_EDITCNOTEFIRST + i), hInstanceMain, NULL);
            SendMessage(hwndEditCNote[i], WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        }
        hwndTemp = CreateWindow(L"STATIC", L"Total",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            DLUX(439), DLUY(57), DLUX(32), DLUY(8),
            hwndStaticPage[PAGETONALITYANLYZER], (HMENU)-1, hInstanceMain, NULL);
        SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndEditCNoteTotal = CreateWindow(L"EDIT", L"0",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            DLUX(439), DLUY(68), DLUX(32), DLUY(14),
            hwndStaticPage[PAGETONALITYANLYZER], (HMENU)IDC_EDITCNOTETOTAL, hInstanceMain, NULL);
        SendMessage(hwndEditCNoteTotal, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
        hwndBtnAnlyzTonality = CreateWindow(L"BUTTON", L"Analyze tonality",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            DLUX(475), DLUY(68), DLUX(72), DLUY(14),
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
                DLUX(7 + 36 * i), DLUY(104), DLUX(32), DLUY(8),
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)-1, hInstanceMain, NULL);
            SendMessage(hwndTemp, WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
            hwndEditTonality[i] = CreateWindow(L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_READONLY | ES_AUTOHSCROLL,
                DLUX(7 + 36 * i), DLUY(115), DLUX(32), DLUY(14),
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)(IDC_EDITTONALITYFIRST + i), hInstanceMain, NULL);
            SendMessage(hwndEditTonality[i], WM_SETFONT, (WPARAM)hfontGUI, (LPARAM)TRUE);
            hwndStaticTonalityBar[i] = CreateWindow(L"STATIC", L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | SS_OWNERDRAW,
                0, 0, 0, 0,
                hwndStaticPage[PAGETONALITYANLYZER], (HMENU)(IDC_STATICTONALITYBARFIRST + i), hInstanceMain, NULL);
            SetWindowLong(hwndStaticTonalityBar[i], GWL_WNDPROC, (LONG)TempWndProc);
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
        /* Controls for common pages */
        SendMessage(hwndStatus, uMsg, wParam, lParam);
        MoveWindow(hwndTC, 0, 0, LOWORD(lParam), HIWORD(lParam) - cyStatus, TRUE);
        GetClientRect(hwnd, &rcTCDisp);
        rcTCDisp.bottom -= cyStatus;
        TabCtrl_AdjustRect(hwndTC, FALSE, &rcTCDisp);
        for(i = 0; i < CPAGE; i++) {
            MoveWindow(hwndStaticPage[i], rcTCDisp.left, rcTCDisp.top, rcTCDisp.right - rcTCDisp.left, rcTCDisp.bottom - rcTCDisp.top, TRUE);
        }

        MoveWindow(hwndEditFilePath, rcTCDisp.left + DLUX(71), rcTCDisp.top + DLUY(7), rcTCDisp.right - rcTCDisp.left - DLUX(78), DLUY(14), TRUE);
        MoveWindow(hwndTBTime, rcTCDisp.left + DLUX(71), rcTCDisp.top + DLUY(25), rcTCDisp.right - rcTCDisp.left - DLUX(78), DLUY(14), TRUE);
        
        /* Controls for event list page */
        MoveWindow(hwndLVEvtList, DLUX(7), DLUY(osversioninfo.dwMajorVersion >= 6 ? 75 : 79), rcTCDisp.right - rcTCDisp.left - DLUX(14), rcTCDisp.bottom - rcTCDisp.top - DLUY(osversioninfo.dwMajorVersion >= 6 ? 82 : 86), TRUE);

        /* Controls for tempo list page */
        MoveWindow(hwndLVTempoList, DLUX(7), DLUY(61), rcTCDisp.right - rcTCDisp.left - DLUX(14), rcTCDisp.bottom - rcTCDisp.top - DLUY(68), TRUE);

        /* Controls for play control page */
        MoveWindow(hwndTVMut, DLUX(7), DLUY(78), rcTCDisp.right - rcTCDisp.left - DLUX(14), rcTCDisp.bottom - rcTCDisp.top - DLUY(85), TRUE);

        /* Controls for tonality analyzer page */
        MoveWindow(hwndStaticMostProbTonality, DLUX(7), DLUY(89), rcTCDisp.right - rcTCDisp.left - DLUX(7), DLUY(8), TRUE);
        for(i = 0; i < 24; i++) {
            MoveWindow(hwndStaticTonalityBar[i], DLUX(7 + 24 * i), DLUY(133), DLUX(24), rcTCDisp.bottom - rcTCDisp.top - DLUY(151), TRUE);
            MoveWindow(hwndStaticTonalityBarLabel[i], DLUX(7 + 24 * i), rcTCDisp.bottom - rcTCDisp.top - DLUY(15), DLUX(24), DLUY(8), TRUE);

            ti.uId = IDC_TTTONALITYBARFIRST + i;
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
                        ti.uId = IDC_TTTONALITYBARFIRST + i;
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
                    lplvcd->clrText = pcrLVEvtListCD[lplvcd->nmcd.dwItemSpec];
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

            if(lpdis->itemID < 2) {
                rcText = lpdis->rcItem;
                if(wParam == IDC_CBFILTTRK)
                    SendMessage(hwndLBCBFiltTrk, LB_GETTEXT, lpdis->itemID, (LPARAM)szBuf);
                if(wParam == IDC_CBFILTCHN)
                    SendMessage(hwndLBCBFiltChn, LB_GETTEXT, lpdis->itemID, (LPARAM)szBuf);
                if(wParam == IDC_CBFILTEVTTYPE)
                    SendMessage(hwndLBCBFiltEvtType, LB_GETTEXT, lpdis->itemID, (LPARAM)szBuf);
                DrawText(hdc, szBuf, -1, &rcText, DT_VCENTER);
                return 0;
            }

            rcCheckBox = lpdis->rcItem;
            rcText = lpdis->rcItem;
            rcCheckBox.right = rcCheckBox.left + rcCheckBox.bottom - rcCheckBox.top;
            rcText.left = rcCheckBox.right;

            uState = DFCS_BUTTONCHECK;
            if(wParam == IDC_CBFILTTRK) {
                if(SendMessage(hwndCBFiltTrk, CB_GETITEMDATA, lpdis->itemID, 0))
                    uState |= DFCS_CHECKED;
                SendMessage(hwndLBCBFiltTrk, LB_GETTEXT, lpdis->itemID, (LPARAM)szBuf);
            }
            if(wParam == IDC_CBFILTCHN) {
                if(SendMessage(hwndCBFiltChn, CB_GETITEMDATA, lpdis->itemID, 0))
                    uState |= DFCS_CHECKED;
                SendMessage(hwndLBCBFiltChn, LB_GETTEXT, lpdis->itemID, (LPARAM)szBuf);
            }
            if(wParam == IDC_CBFILTEVTTYPE) {
                if(SendMessage(hwndCBFiltEvtType, CB_GETITEMDATA, lpdis->itemID, 0))
                    uState |= DFCS_CHECKED;
                SendMessage(hwndLBCBFiltEvtType, LB_GETTEXT, lpdis->itemID, (LPARAM)szBuf);
            }
            DrawFrameControl(hdc, &rcCheckBox, DFC_BUTTON, uState);
            DrawText(hdc, szBuf, -1, &rcText, DT_VCENTER);
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
                dwStartTime = SendMessage(hwndTBTime, TBM_GETPOS, 0, 0);
                if(!(pevtCurOutput = GetEvtByMs(&mf, dwStartTime, &dwPrevBufEvtTk, &dwCurTempoData)))
                    break;
                SendMessage(hwnd, WM_APP_PLAYFROMEVT, 0, (LPARAM)pevtCurOutput);
                iCurStrmStatus = STRM_PLAY;
                SetWindowText(hwndBtnPlay, L"Pause");
                break;
            case STRM_PLAY: // Pause
                mmt.wType = TIME_MS;
                midiStreamPosition(hms, &mmt, sizeof(MMTIME));
                SendMessage(hwndTBTime, TBM_SETPOS, TRUE, mmt.u.ms * iTempoR / 100 + dwStartTime);

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

        case IDC_BTNANLYZTONALITY:
            SendMessage(hwnd, WM_APP_ANLYZTONALITY, 0, 0);
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
                break;
            case SB_ENDSCROLL:
                fTBIsTracking = FALSE;
                break;
            case SB_THUMBPOSITION:
            case SB_PAGELEFT:
            case SB_PAGERIGHT:
                if(iCurStrmStatus != STRM_PLAY)
                    break;

                dwStartTime = SendMessage(hwndTBTime, TBM_GETPOS, 0, 0);

                iPlayCtlBufRemaining = 2; // See the process of the MM_MOM_DONE message
                midiStreamStop(hms);
                break;
            }
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
        swprintf(szBuf, 128, L"Event density: %f per s", mf.dDur == 0. ? 0. : (double)mf.cEvt / mf.dDur);
        SetWindowText(hwndStaticEvtDensity, szBuf);

        pfiltstate = (FILTERSTATES *)malloc(sizeof(FILTERSTATES) + (mf.cTrk - 1) * sizeof(DWORD));
        for(u = 0; u < mf.cTrk; u++)
            pfiltstate->dwFiltTrk[u] = FILT_CHECKED;
        for(u = 0; u < 17; u++)
            pfiltstate->dwFiltChn[u] = FILT_CHECKED;
        for(u = 0; u < 10; u++)
            pfiltstate->dwFiltEvtType[u] = FILT_CHECKED;

        pcrLVEvtListCD = (COLORREF *)malloc(mf.cEvt * sizeof(COLORREF));
        cEvtListRow = GetEvtList(hwndLVEvtList, &mf, pfiltstate, pcrLVEvtListCD);
        cTempoListRow = GetTempoList(hwndLVTempoList, &mf);

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

        SendMessage(hwndTBTime, TBM_SETRANGEMAX, TRUE, (int)(mf.dDur * 1000));
        SendMessage(hwndTBTime, TBM_SETPOS, TRUE, 0);

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
            wsprintf(szBuf, L"Track #%u", u);
            tvis.item.pszText = szBuf;
            htiTrk = TreeView_InsertItem(hwndTVMut, &tvis);
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

        cCBFiltItem = SendMessage(hwndLBCBFiltTrk, LB_GETCOUNT, 0, 0);
        for(i = 2; i < cCBFiltItem; i++)
            SendMessage(hwndCBFiltTrk, CB_DELETESTRING, 2, 0);
        cCBFiltItem = SendMessage(hwndLBCBFiltChn, LB_GETCOUNT, 0, 0);
        for(i = 2; i < cCBFiltItem; i++)
            SendMessage(hwndCBFiltChn, CB_DELETESTRING, 2, 0);
        cCBFiltItem = SendMessage(hwndLBCBFiltEvtType, LB_GETCOUNT, 0, 0);
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

        SetWindowText(hwndStaticType, L"Type: ");
        SetWindowText(hwndStaticCTrk, L"Number of tracks: ");
        SetWindowText(hwndStaticTb, L"Timebase: ");
        SetWindowText(hwndStaticCEvt, L"Number of events: ");

        SetWindowText(hwndStaticIniTempo, L"Initial tempo: ");
        SetWindowText(hwndStaticAvgTempo, L"Average tempo: ");
        SetWindowText(hwndStaticDur, L"Duration: ");
        SetWindowText(hwndStaticEvtDensity, L"Event density: ");

        free(pcrLVEvtListCD);
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

            ti.uId = IDC_TTTONALITYBARFIRST + i;
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
            SendMessage((HWND)lParam, CB_SETITEMDATA, wParam, 
                !SendMessage((HWND)lParam, CB_GETITEMDATA, wParam, 0));

            // Update filter states
            if((HWND)lParam == hwndCBFiltTrk) {
                for(i = 2, uFiltIndex = 0;; i++, uFiltIndex++) {
                    while(!(pfiltstate->dwFiltTrk[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if(i == wParam)
                        break;
                }
                pfiltstate->dwFiltTrk[uFiltIndex] ^= FILT_CHECKED;
            }
            if((HWND)lParam == hwndCBFiltChn) {
                for(i = 2, uFiltIndex = 0;; i++, uFiltIndex++) {
                    while(!(pfiltstate->dwFiltChn[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if(i == wParam)
                        break;
                }
                pfiltstate->dwFiltChn[uFiltIndex] ^= FILT_CHECKED;
            }
            if((HWND)lParam == hwndCBFiltEvtType) {
                for(i = 2, uFiltIndex = 0;; i++, uFiltIndex++) {
                    while(!(pfiltstate->dwFiltEvtType[uFiltIndex] & FILT_AVAILABLE))
                        uFiltIndex++;
                    if (i == wParam)
                        break;
                }
                pfiltstate->dwFiltEvtType[uFiltIndex] ^= FILT_CHECKED;
            }
        } else { // Check all or check none
            for(i = 2; i < cCBFiltItem; i++) {
                SendMessage((HWND)lParam, CB_SETITEMDATA, i, !wParam);
            }

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

        // Get event list again
        cEvtListRow = GetEvtList(hwndLVEvtList, &mf, pfiltstate, pcrLVEvtListCD);
        wsprintf(szBuf, L"%u event(s) in total.", cEvtListRow);
        SendMessage(hwndStatus, SB_SETTEXT, 0, (LPARAM)szBuf);

        iCBFiltTopIndex = SendMessage((HWND)lParam, LB_GETTOPINDEX, 0, 0);
        SendMessage((HWND)lParam, WM_SETREDRAW, FALSE, 0);

        MakeCBFiltLists(&mf, pfiltstate, hwndCBFiltTrk, hwndCBFiltChn, hwndCBFiltEvtType);

        SendMessage((HWND)lParam, LB_SETTOPINDEX, iCBFiltTopIndex, 0);
        SendMessage((HWND)lParam, WM_SETREDRAW, TRUE, 0);

        if(iCurStrmStatus == STRM_PLAY) {
            pevtTemp = mf.pevtHead;
            uCurEvtListRow = -1;
            while(pevtTemp != pevtCurOutput) {
                if(pevtTemp->fUnfiltered)
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
            if(pevtCurBuf->fUnfiltered)
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
        if(pevtCurOutput->fUnfiltered) {
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
        SendMessage(hwndTBTime, TBM_SETPOS, TRUE, 0);
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
                ti.uId = IDC_TTTONALITYBARFIRST + i;
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
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
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
                if(hwndFocused && lstrcmp(szBuf, L"EDIT")){
                    SendMessage(hwndFocused, EM_SETSEL, 0, -1);
                    continue;
                }
            }

            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    
    return msg.wParam;
}