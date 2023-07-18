#define _WIN32_WINNT 0x0400
#include <Windows.h>
#include "cbinfo.h"
#include <CommCtrl.h>
#include <stdio.h> // For float number outputting
#include "strlist.h"
#include "resource.h"

#pragma comment(lib,"comctl32.lib")
#pragma comment(lib,"winmm.lib")

#define CPAGE 4
#define STRMBUFLEN 15360

/* Filter flags */
#define FILT_CHECKED 1
#define FILT_AVAILABLE 0x10000

enum eIDC{
    IDC_STATUS=1, IDC_TC, IDC_BTNOPEN, IDC_BTNCLOSE, IDC_EDITFILEPATH, IDC_BTNPLAY, IDC_BTNSTOP, IDC_TBTIME,
    IDC_STATICTYPE, IDC_STATICCTRK, IDC_STATICTB, IDC_STATICCEVT, IDC_CBFILTTRK, IDC_CBFILTCHN, IDC_CBFILTEVTTYPE, IDC_LVEVTLIST,
    IDC_STATICINITEMPO, IDC_STATICAVGTEMPO, IDC_STATICDUR, IDC_STATICEVTDENSITY, IDC_LVTEMPOLIST,
    IDC_EDITTRANSP, IDC_UDTRANSP, IDC_BTNTRANSPRESET, IDC_EDITTEMPOR, IDC_UDTEMPOR, IDC_BTNTEMPORRESET, IDC_TVMUT,
    IDC_EDITCNOTEFIRST, IDC_EDITCNOTETOTAL=IDC_EDITCNOTEFIRST+12, IDC_BTNANLYZTONALITY, IDC_STATICMOSTPROBTONALITY, IDC_EDITTONALITYFIRST, IDC_STATICTONALITYBARFIRST, IDC_TTTONALITYBARFIRST=IDC_STATICTONALITYBARFIRST+24, IDC_STATICTONALITYBARLABELFIRST=IDC_STATICTONALITYBARFIRST+24
};

enum ePage{
    PAGEEVTLIST, PAGETEMPOLIST, PAGEPLAYCTL, PAGETONALITYANLYZER
};

enum eAppMsg{
    WM_APP_CLOSEFILE=WM_APP,
    WM_APP_ANLYZTONALITY,
    WM_APP_STOP,
    WM_APP_PLAYFROMEVT, // lParam: pevt
    WM_APP_OPENFILE, // lParam: szFilePath
    WM_APP_FILTCHECKCHANGE, // wParam: iIndex; lParam: hwndCB
    WM_APP_FILTCHECKALL // wParam: TRUE (all check) or FALSE (all uncheck); lParam: hwndCB
};

enum eStrmStatus{
    STRM_STOP, STRM_PLAY
};

typedef struct evt{
    struct evt *pevtNext;
    DWORD dwTk;
    WORD wTrk;
    DWORD cbData; // For meta or sys-ex events
    BOOL fUnfiltered;
    BYTE bStatus;
    BYTE bData1;
    BYTE bData2;
    BYTE bData[1]; // For meta or sys-ex events
}EVENT;

typedef struct tempoevt{
    struct tempoevt *ptempoevtNext;
    DWORD dwTk;
    DWORD cTk;
    DWORD dwData;
}TEMPOEVENT;

typedef struct mf{
    BOOL fOpened;

    BYTE *pb;
    DWORD dwLoc;
    EVENT *pevtHead;
    TEMPOEVENT *ptempoevtHead;

    DWORD cEvt;
    DWORD cTempoEvt;

    WORD wType;
    WORD cTrk;
    WORD wTb;
    
    UINT cNote[12]; // For tonality analysis

    DWORD cTk;
    DWORD cPlayableTk;

    double dDur;
    double dIniTempo;
    double dAvgTempo;

    WORD *pwTrkChnUsage; // Channel usage of tracks, represented by bits
}MIDIFILE;

typedef struct filtstate{
    DWORD dwFiltChn[17];
    DWORD dwFiltEvtType[10];
    DWORD dwFiltTrk[1];
}FILTERSTATES;

LPWSTR szCmdFilePath=NULL;
HINSTANCE hInstanceMain;
HWND hwndMain;

WNDPROC DefTCProc,DefStaticProc,DefLBProc;

LPWSTR szNoteLabel[]={L"C",L"C#",L"D",L"D#",L"E",L"F",L"F#",L"G",L"G#",L"A",L"A#",L"B"};
LPWSTR szTonalityLabel[]={L"C",L"Cm",L"Db",L"C#m",L"D",L"Dm",L"Eb",L"Ebm",L"E",L"Em",L"F",L"Fm",L"F#",L"F#m",L"G",L"Gm",L"Ab",L"G#m",L"A",L"Am",L"Bb",L"Bbm",L"B",L"Bm"};

BYTE ReadByte(MIDIFILE *pmf){
    return pmf->pb[pmf->dwLoc++];
}

/* Read a multi byte big-endian integer */
DWORD ReadBEInt(MIDIFILE *pmf, UINT cb){
    DWORD dwRet=0;
    UINT u;
    for(u=0;u<cb;u++){
        dwRet=dwRet << 8 | ReadByte(pmf);
    }
    return dwRet;
}

/* Read a variable-length integer */
DWORD ReadVarLenInt(MIDIFILE *pmf){
    DWORD dwRet=0;
    BYTE b;
    do{
        b=ReadByte(pmf);
        dwRet=dwRet << 7 | (b & 127);
    }while(b & 128);
    return dwRet;
}

/* Read MIDI File */
BOOL ReadMidi(LPWSTR szPath, MIDIFILE *pmf){
    HANDLE hFile;
    DWORD dwFileLen;
    EVENT *pevtCur=NULL, *pevtPrev=NULL;
    TEMPOEVENT *ptempoevtCur=NULL, *ptempoevtPrev=NULL;

    DWORD dwCurChkName; // Only used when determining whether it is an RIFF midi file
    DWORD dwCurChkLen;
    DWORD dwCurChkEndPos;

    DWORD *pdwTrkLoc=NULL;
    BYTE *pbTrkCurStatus=NULL;
    DWORD *pdwTrkCurTk=NULL;
    DWORD *pdwTrkEndPos=NULL;
    BOOL *pfTrkIsEnd=NULL;
    BOOL fMidiIsEnd;

    DWORD dwCurTk=0, dwCurPlayableTk=0;
    WORD wCurEvtTrk;
    BYTE bCurEvtStatus;
    BYTE bCurEvtChn;
    BYTE bCurEvtData1;
    BYTE bCurEvtData2;
    DWORD cbCurEvtData;
    short iChnCurPrg[16]; // For excluding notes of non-chromatic instruments when counting notes

    DWORD dwCurTempoTk,dwPrevTempoTk,cPrevTempoTk,cLastTempoTk;
    DWORD dwCurTempoData,dwPrevTempoData;
    DWORDLONG qwMusTb; // An intermediate variable when calculating duration (equals to the microsecond count multiplies the timebase)

    UINT u;
    BYTE b;

    hFile=CreateFile(szPath,GENERIC_READ,0,NULL,
        OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
    if(hFile==INVALID_HANDLE_VALUE)
        return FALSE;
    pmf->pb=(BYTE *)malloc(GetFileSize(hFile,NULL));
    ReadFile(hFile,pmf->pb,GetFileSize(hFile,NULL),&dwFileLen,NULL);
    CloseHandle(hFile);

    pmf->dwLoc=0;

    /* Read the header chunk */
    if(pmf->dwLoc+8>dwFileLen) goto err;
    dwCurChkName=ReadBEInt(pmf,4);
    if(dwCurChkName==0x52494646){ // RIFF file
        if(pmf->dwLoc+16>dwFileLen) goto err;
        pmf->dwLoc+=16;
        if(pmf->dwLoc+8>dwFileLen) goto err;
        dwCurChkName=ReadBEInt(pmf,4);
    }
    if(dwCurChkName!=0x4D546864) goto err;
    dwCurChkLen=ReadBEInt(pmf,4);
    dwCurChkEndPos=pmf->dwLoc+dwCurChkLen;
    if(dwCurChkEndPos>dwFileLen) goto err;
    pmf->wType=ReadBEInt(pmf,2);
    pmf->cTrk=ReadBEInt(pmf,2);
    pmf->wTb=ReadBEInt(pmf,2);
    pmf->dwLoc=dwCurChkEndPos;
    
    pdwTrkLoc=(DWORD *)malloc(pmf->cTrk*sizeof(DWORD));
    pdwTrkEndPos=(DWORD *)malloc(pmf->cTrk*sizeof(DWORD));
    pbTrkCurStatus=(BYTE *)malloc(pmf->cTrk*sizeof(BYTE));
    pdwTrkCurTk=(DWORD *)malloc(pmf->cTrk*sizeof(DWORD));
    pfTrkIsEnd=(BOOL *)malloc(pmf->cTrk*sizeof(BOOL));

    pmf->pwTrkChnUsage=(WORD *)malloc(pmf->cTrk*sizeof(WORD));

    /* Locate each track chunk and prepare event reading */
    u=0;
    while(u<pmf->cTrk){
        if(pmf->dwLoc+8>dwFileLen) goto err;

        if(ReadBEInt(pmf,4)!=0x4D54726B){ // Skip non-standard MIDI track chunks
            dwCurChkLen=ReadBEInt(pmf,4);
            dwCurChkEndPos=pmf->dwLoc+dwCurChkLen;
            if(dwCurChkEndPos>dwFileLen) goto err;
            pmf->dwLoc=dwCurChkEndPos;
            continue;
        }

        dwCurChkLen=ReadBEInt(pmf,4);
        dwCurChkEndPos=pmf->dwLoc+dwCurChkLen;
        if(dwCurChkEndPos>dwFileLen) goto err;
        pdwTrkEndPos[u]=dwCurChkEndPos;
        pbTrkCurStatus[u]=0;
        pdwTrkCurTk[u]=ReadVarLenInt(pmf);
        pdwTrkLoc[u]=pmf->dwLoc;
        pfTrkIsEnd[u]=FALSE;

        pmf->pwTrkChnUsage[u]=0;

        pmf->dwLoc=dwCurChkEndPos;
        u++;
    }

    pmf->cEvt=0;
    for(u=0;u<12;u++)
        pmf->cNote[u]=0;
    for(u=0;u<16;u++)
        iChnCurPrg[u]=-1;

    pmf->cTempoEvt=0;
    qwMusTb=0;

    /* Begin to read a new event */
    do{
        if(!pmf->cTrk)
            break;

        /* Find the track with the earliest event to be read */
        u=0;
        while(pfTrkIsEnd[u]){
            u++;
        }
        dwCurTk=pdwTrkCurTk[u];
        wCurEvtTrk=u;
        for(;u<pmf->cTrk;u++){
            if(!pfTrkIsEnd[u] && pdwTrkCurTk[u]<dwCurTk){
                dwCurTk=pdwTrkCurTk[u];
                wCurEvtTrk=u;
            }
        }

        pmf->dwLoc=pdwTrkLoc[wCurEvtTrk];

        if(pmf->pb[pmf->dwLoc] & 0x80){
            bCurEvtStatus=ReadByte(pmf);
            if(bCurEvtStatus<0xF0) pbTrkCurStatus[wCurEvtTrk]=bCurEvtStatus;
        }else{
            bCurEvtStatus=pbTrkCurStatus[wCurEvtTrk];
        }

        if(bCurEvtStatus<0xF0){
            pmf->pwTrkChnUsage[wCurEvtTrk] |= 1 << (bCurEvtStatus & 0xF);

            bCurEvtChn=(bCurEvtStatus & 0xF)+1;
            bCurEvtData1=ReadByte(pmf);

            pevtCur=(EVENT *)malloc(sizeof(EVENT));
            pevtCur->dwTk=dwCurTk;
            pevtCur->wTrk=wCurEvtTrk;
            pevtCur->bStatus=bCurEvtStatus;
            pevtCur->bData1=bCurEvtData1;

            switch(bCurEvtStatus >> 4){
            case 0x8:
            case 0xA:
            case 0xB:
            case 0xE:
                bCurEvtData2=ReadByte(pmf);
                pevtCur->bData2=bCurEvtData2;
                break;

            case 0x9:
                bCurEvtData2=ReadByte(pmf);
                pevtCur->bData2=bCurEvtData2;

                /* Count each note, of chromatic instruments only */
                if(bCurEvtData2!=0 && bCurEvtChn!=10 && iChnCurPrg[bCurEvtChn-1]<115)
                    pmf->cNote[bCurEvtData1%12]++;
                break;

            case 0xC:
                iChnCurPrg[bCurEvtChn-1]=bCurEvtData1;
            case 0xD:
                pevtCur->bData2=0;
                break;
            }

            dwCurPlayableTk=dwCurTk;
        }else{
            switch(bCurEvtStatus){
            case 0xF0:
            case 0xF7:
                cbCurEvtData=ReadVarLenInt(pmf);
                if(bCurEvtStatus==0xF0)
                    cbCurEvtData++;

                pevtCur=(EVENT *)malloc(sizeof(EVENT)+cbCurEvtData);
                pevtCur->pevtNext=NULL;
                pevtCur->dwTk=dwCurTk;
                pevtCur->wTrk=wCurEvtTrk;
                pevtCur->bStatus=bCurEvtStatus;
                pevtCur->cbData=cbCurEvtData;
                if(bCurEvtStatus==0xF0){
                    u=1;
                    pevtCur->bData[0]=0xF0;
                }else{
                    u=0;
                }
                for(;u<cbCurEvtData;u++)
                    pevtCur->bData[u]=ReadByte(pmf);

                dwCurPlayableTk=dwCurTk;
                break;

            case 0xFF:
                bCurEvtData1=ReadByte(pmf);
                cbCurEvtData=ReadVarLenInt(pmf);

                pevtCur=(EVENT *)malloc(sizeof(EVENT)+cbCurEvtData);
                pevtCur->pevtNext=NULL;
                pevtCur->dwTk=dwCurTk;
                pevtCur->wTrk=wCurEvtTrk;
                pevtCur->bStatus=bCurEvtStatus;
                pevtCur->bData1=bCurEvtData1;
                pevtCur->cbData=cbCurEvtData;

                if(bCurEvtData1==0x2F){
                    pfTrkIsEnd[wCurEvtTrk]=TRUE;
                }

                if(bCurEvtData1==0x51){
                    if(cbCurEvtData!=3) goto err;

                    dwCurTempoTk=dwCurTk;
                    dwCurTempoData=0;
                    for(u=0;u<3;u++){
                        b=ReadByte(pmf);
                        pevtCur->bData[u]=b;
                        dwCurTempoData=dwCurTempoData << 8 | b;
                    }

                    ptempoevtCur=(TEMPOEVENT *)malloc(sizeof(TEMPOEVENT));
                    ptempoevtCur->dwTk=dwCurTempoTk;
                    ptempoevtCur->dwData=dwCurTempoData;

                    if(pmf->cTempoEvt>0){ // Not the first tempo event
                        /* Calculate the tick count of the previous tempo event */
                        cPrevTempoTk=dwCurTempoTk-dwPrevTempoTk;
                        qwMusTb+=(DWORDLONG)dwPrevTempoData*cPrevTempoTk;

                        ptempoevtPrev->cTk=cPrevTempoTk;
                        ptempoevtPrev->ptempoevtNext=ptempoevtCur;
                    }else{ // The first tempo event
                        qwMusTb=500000*dwCurTempoTk;
                        pmf->dIniTempo=60000000./dwCurTempoData;

                        pmf->ptempoevtHead=ptempoevtCur;
                    }

                    dwPrevTempoTk=dwCurTempoTk;
                    dwPrevTempoData=dwCurTempoData;
                    ptempoevtPrev=ptempoevtCur;

                    dwCurPlayableTk=dwCurTk;

                    pmf->cTempoEvt++;
                }else{
                    for(u=0;u<cbCurEvtData;u++)
                        pevtCur->bData[u]=ReadByte(pmf);
                }
                break;

            default:
                goto err;
            }
        }

        if(pmf->cEvt>0){ // Not the first event
            pevtPrev->pevtNext=pevtCur;
        }else{ // The first event
            pmf->pevtHead=pevtCur;
        }

        pevtPrev=pevtCur;

        pmf->cEvt++;

        pdwTrkLoc[wCurEvtTrk]=pmf->dwLoc;
        pfTrkIsEnd[wCurEvtTrk]=pfTrkIsEnd[wCurEvtTrk] || pdwTrkLoc[wCurEvtTrk]>=pdwTrkEndPos[wCurEvtTrk];

        /* Find whether all tracks have already ended */
        fMidiIsEnd=TRUE;
        for(u=0;u<pmf->cTrk && fMidiIsEnd;u++)
            fMidiIsEnd=fMidiIsEnd && pfTrkIsEnd[u];

        if(!pfTrkIsEnd[wCurEvtTrk]){
            pdwTrkCurTk[wCurEvtTrk]+=ReadVarLenInt(pmf);
            pdwTrkLoc[wCurEvtTrk]=pmf->dwLoc;
        }
    }while(!fMidiIsEnd);

    /* Complete the event link */
    if(pmf->cEvt!=0){
        pevtPrev->pevtNext=NULL;
    }

    /* Complete the tempo event link */
    if(pmf->cTempoEvt==0){ // In the extreme case of midi with no tempo events
        qwMusTb=(DWORDLONG)500000*dwCurPlayableTk;
        pmf->dIniTempo=0;
    }else{
        /* Calculate the tick count of the last tempo event */
        cLastTempoTk=dwCurPlayableTk-dwCurTempoTk;
        qwMusTb+=(DWORDLONG)dwCurTempoData*cLastTempoTk;

        ptempoevtPrev->cTk=cLastTempoTk;
        ptempoevtPrev->ptempoevtNext=NULL;
    }

    /* Calculate duration and average tempo */
    pmf->dDur=(double)qwMusTb/pmf->wTb/1000000;
    if(qwMusTb==0){ // In the extreme case of zero-duration midi
        pmf->dAvgTempo=pmf->dIniTempo;
    }else{
        pmf->dAvgTempo=60000000.*dwCurPlayableTk/qwMusTb;
    }
    
    pmf->cTk=dwCurTk;
    pmf->cPlayableTk=dwCurPlayableTk;
    
    pmf->fOpened=TRUE;

    free(pmf->pb);
    free(pdwTrkLoc);
    free(pbTrkCurStatus);
    free(pdwTrkCurTk);
    free(pdwTrkEndPos);
    free(pfTrkIsEnd);
    return TRUE;

err:
    free(pmf->pb);
    free(pdwTrkLoc);
    free(pbTrkCurStatus);
    free(pdwTrkCurTk);
    free(pdwTrkEndPos);
    free(pfTrkIsEnd);
    return FALSE;
}

void FreeMidi(MIDIFILE *pmf){
    EVENT *pevtCur, *pevtNext;
    TEMPOEVENT *ptempoevtCur, *ptempoevtNext;

    free(pmf->pwTrkChnUsage);
    
    pevtCur=pmf->pevtHead;
    while(pevtCur){
        pevtNext=pevtCur->pevtNext;
        free(pevtCur);
        pevtCur=pevtNext;
    }

    ptempoevtCur=pmf->ptempoevtHead;
    while(ptempoevtCur){
        ptempoevtNext=ptempoevtCur->ptempoevtNext;
        free(ptempoevtCur);
        ptempoevtCur=ptempoevtNext;
    }

    ZeroMemory(pmf,sizeof(MIDIFILE));
}

EVENT *GetEvtByMs(MIDIFILE *pmf, DWORD dwMs, DWORD *pdwTk, DWORD *pdwCurTempoData){
    DWORDLONG qwMusTb=(DWORDLONG)dwMs*pmf->wTb*1000;
    
    EVENT *pevtCur;
    DWORD dwCurTempoData=500000;
    DWORDLONG qwCurMusTb=0,qwPrevMusTb=0;
    DWORD dwPrevEvtTk=0;

    pevtCur=pmf->pevtHead;
    while(qwCurMusTb<qwMusTb && pevtCur){
        qwPrevMusTb=qwCurMusTb;
        qwCurMusTb+=(DWORDLONG)dwCurTempoData*(pevtCur->dwTk-dwPrevEvtTk);

        if(qwCurMusTb>=qwMusTb)
            break;
        dwPrevEvtTk=pevtCur->dwTk;
        if(pevtCur->bStatus==0xFF && pevtCur->bData1==0x51)
            dwCurTempoData=pevtCur->bData[0]<<16 | pevtCur->bData[1]<<8 | pevtCur->bData[2];
        pevtCur=pevtCur->pevtNext;
    }

    *pdwTk=dwPrevEvtTk+(qwMusTb-qwPrevMusTb)/dwCurTempoData;
    *pdwCurTempoData=dwCurTempoData;
    return pevtCur;
}

int EvtGetFiltTrkIndex(EVENT *pevt){
    return pevt->wTrk;
}

int EvtGetFiltChnIndex(EVENT *pevt){
    if(pevt->bStatus<0xF0)
        return (pevt->bStatus & 0xF)+1;
    return 0;
}

int EvtGetFiltEvtTypeIndex(EVENT *pevt){
    switch(pevt->bStatus >> 4){
    case 0x8:
        return 0;
    case 0x9:
        if(pevt->bData2==0)
            return 1;
        return 2;
    case 0xA:
    case 0xB:
    case 0xC:
    case 0xD:
    case 0xE:
        return (pevt->bStatus >> 4)-7;
    }
    if(pevt->bStatus==0xF0 || pevt->bStatus==0xF7)
        return 8;
    return 9;
}

DWORD GetEvtList(HWND hwndLVEvtList, MIDIFILE *pmf, FILTERSTATES *pfiltstate, COLORREF *pcrLVEvtListCD){
    LVITEM lvitem;

    EVENT *pevtCur;
    DWORD dwRow=0;
    WCHAR szBuf[128];
    LPWSTR szData, szDataCmt; // For meta or sys-ex events

    int cSharp=0; // For key signature events

    int iFiltTrkIndex, iFiltChnIndex, iFiltEvtTypeIndex;

    UINT u;
    
    SendMessage(hwndLVEvtList,WM_SETREDRAW,FALSE,0);
    ListView_DeleteAllItems(hwndLVEvtList);

    lvitem.mask=LVIF_TEXT;
    pevtCur=pmf->pevtHead;
    while(pevtCur){
        iFiltTrkIndex=EvtGetFiltTrkIndex(pevtCur);
        iFiltChnIndex=EvtGetFiltChnIndex(pevtCur);
        iFiltEvtTypeIndex=EvtGetFiltEvtTypeIndex(pevtCur);
        if(
            !(pfiltstate->dwFiltTrk[iFiltTrkIndex] & FILT_CHECKED) ||
            !(pfiltstate->dwFiltChn[iFiltChnIndex] & FILT_CHECKED) ||
            !(pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] & FILT_CHECKED)
        ){
            if(
                !(pfiltstate->dwFiltTrk[iFiltTrkIndex] & FILT_CHECKED) &&
                pfiltstate->dwFiltChn[iFiltChnIndex] & FILT_CHECKED &&
                pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] & FILT_CHECKED
            ){
                pfiltstate->dwFiltTrk[iFiltTrkIndex] |= FILT_AVAILABLE;
            }
            if(
                pfiltstate->dwFiltTrk[iFiltTrkIndex] & FILT_CHECKED &&
                !(pfiltstate->dwFiltChn[iFiltChnIndex] & FILT_CHECKED) &&
                pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] & FILT_CHECKED
            ){
                pfiltstate->dwFiltChn[iFiltChnIndex] |= FILT_AVAILABLE;
            }
            if(
                pfiltstate->dwFiltTrk[iFiltTrkIndex] & FILT_CHECKED &&
                pfiltstate->dwFiltChn[iFiltChnIndex] & FILT_CHECKED &&
                !(pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] & FILT_CHECKED)
            ){
                pfiltstate->dwFiltEvtType[iFiltEvtTypeIndex] |= FILT_AVAILABLE;
            }

            pevtCur->fUnfiltered=FALSE;
            pevtCur=pevtCur->pevtNext;
            continue;
        }

        lvitem.iItem=dwRow;
        lvitem.iSubItem=0;
        wsprintf(szBuf,L"%u",dwRow+1);
        lvitem.pszText=szBuf;
        lvitem.cchTextMax=128;
        ListView_InsertItem(hwndLVEvtList,&lvitem);

        pfiltstate->dwFiltTrk[pevtCur->wTrk] |= FILT_AVAILABLE;
        wsprintf(szBuf,L"%u",pevtCur->wTrk);
        ListView_SetItemText(hwndLVEvtList,dwRow,1,szBuf);

        wsprintf(szBuf,L"%u",pevtCur->dwTk);
        ListView_SetItemText(hwndLVEvtList,dwRow,3,szBuf);

        if(pevtCur->bStatus < 0xF0){
            pfiltstate->dwFiltChn[(pevtCur->bStatus & 0xF)+1] |= FILT_AVAILABLE;
            wsprintf(szBuf,L"%u",(pevtCur->bStatus & 0xF)+1);
            ListView_SetItemText(hwndLVEvtList,dwRow,2,szBuf);

            switch(pevtCur->bStatus >> 4){
            case 0x8:
                pfiltstate->dwFiltEvtType[0] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow]=RGB(128,128,128);
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"Note off");
                wsprintf(szBuf,L"%u",pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList,dwRow,5,szBuf);
                ListView_SetItemText(hwndLVEvtList,dwRow,6,szNote[pevtCur->bData1]);
                wsprintf(szBuf,L"%u",pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList,dwRow,7,szBuf);
                break;
            case 0x9:
                if(pevtCur->bData2==0){
                    pfiltstate->dwFiltEvtType[1] |= FILT_AVAILABLE;
                    pcrLVEvtListCD[dwRow]=RGB(128,128,128);
                }else{
                    pfiltstate->dwFiltEvtType[2] |= FILT_AVAILABLE;
                    pcrLVEvtListCD[dwRow]=RGB(0,0,0);
                }
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"Note on");
                wsprintf(szBuf,L"%u",pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList,dwRow,5,szBuf);
                ListView_SetItemText(hwndLVEvtList,dwRow,6,szNote[pevtCur->bData1]);
                wsprintf(szBuf,L"%u",pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList,dwRow,7,szBuf);
                break;
            case 0xA:
                pfiltstate->dwFiltEvtType[3] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow]=RGB(0,0,128);
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"Note aftertouch");
                wsprintf(szBuf,L"%u",pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList,dwRow,5,szBuf);
                ListView_SetItemText(hwndLVEvtList,dwRow,6,szNote[pevtCur->bData1]);
                wsprintf(szBuf,L"%u",pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList,dwRow,7,szBuf);
                break;
            case 0xB:
                pfiltstate->dwFiltEvtType[4] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow]=RGB(0,128,0);
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"Controller");
                wsprintf(szBuf,L"%u",pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList,dwRow,5,szBuf);
                ListView_SetItemText(hwndLVEvtList,dwRow,6,szCtl[pevtCur->bData1]);
                wsprintf(szBuf,L"%u",pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList,dwRow,7,szBuf);
                break;
            case 0xC:
                pfiltstate->dwFiltEvtType[5] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow]=RGB(192,128,0);
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"Program change");
                wsprintf(szBuf,L"%u",pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList,dwRow,5,szBuf);
                ListView_SetItemText(hwndLVEvtList,dwRow,6,szPrg[pevtCur->bData1]);
                break;
            case 0xD:
                pfiltstate->dwFiltEvtType[6] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow]=RGB(0,0,255);
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"Channel aftertouch");
                wsprintf(szBuf,L"%u",pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList,dwRow,5,szBuf);
                break;
            case 0xE:
                pfiltstate->dwFiltEvtType[7] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow]=RGB(255,0,255);
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"Pitch bend");
                wsprintf(szBuf,L"%u",pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList,dwRow,5,szBuf);
                wsprintf(szBuf,L"%u",pevtCur->bData2);
                ListView_SetItemText(hwndLVEvtList,dwRow,7,szBuf);
                /* Generate comment string for pitch bend events */
                wsprintf(szBuf,L"%d",(int)pevtCur->bData1 | (pevtCur->bData2 << 7) - 8192);
                ListView_SetItemText(hwndLVEvtList,dwRow,8,szBuf);
                break;
            }
        }else{
            pfiltstate->dwFiltChn[0] |= FILT_AVAILABLE;
            szData=(LPWSTR)malloc((pevtCur->cbData*4+1)*sizeof(WCHAR));
            szData[0]='\0';
            for(u=0;u<pevtCur->cbData;u++){
                wsprintf(szData,L"%s\\x%02X",szData,pevtCur->bData[u]);
            }
            ListView_SetItemText(hwndLVEvtList,dwRow,7,szData);
            free(szData);

            switch(pevtCur->bStatus){
            case 0xF0:
            case 0xF7:
                pfiltstate->dwFiltEvtType[8] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow]=RGB(255,0,0);
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"System exclusive");
                break;
            case 0xFF:
                pfiltstate->dwFiltEvtType[9] |= FILT_AVAILABLE;
                pcrLVEvtListCD[dwRow]=RGB(192,128,128);
                ListView_SetItemText(hwndLVEvtList,dwRow,4,L"Meta event");
                wsprintf(szBuf,L"%u",pevtCur->bData1);
                ListView_SetItemText(hwndLVEvtList,dwRow,5,szBuf);
                ListView_SetItemText(hwndLVEvtList,dwRow,6,szMeta[pevtCur->bData1]);

                /* Generate comment string for text-typed meta events */
                if(pevtCur->bData1>=0x1 && pevtCur->bData1<=0x9){
                    szDataCmt=(LPWSTR)malloc((pevtCur->cbData+3)*sizeof(WCHAR));
                    ZeroMemory(szDataCmt,(pevtCur->cbData+3)*sizeof(WCHAR));
                    szDataCmt[0]='\"';
                    MultiByteToWideChar(CP_ACP,MB_COMPOSITE,(LPCSTR)pevtCur->bData,pevtCur->cbData,szDataCmt+1,(pevtCur->cbData+2)*sizeof(WCHAR));
                    lstrcat(szDataCmt,L"\"");
                    ListView_SetItemText(hwndLVEvtList,dwRow,8,szDataCmt);
                    free(szDataCmt);
                }

                /* Generate comment string for tempo events */
                if(pevtCur->bData1==0x51){
                    swprintf(szBuf,128,L"%f bpm",60000000./(pevtCur->bData[0]<<16 | pevtCur->bData[1]<<8 | pevtCur->bData[2]));
                    ListView_SetItemText(hwndLVEvtList,dwRow,8,szBuf);
                }

                /* Generate comment string for time signature events */
                if(pevtCur->bData1==0x58){
                    if(pevtCur->cbData>=2){
                        wsprintf(szBuf,L"%d/%d",pevtCur->bData[0],1<<pevtCur->bData[1]);
                        ListView_SetItemText(hwndLVEvtList,dwRow,8,szBuf);
                    }
                }

                /* Generate comment string for key signature events */
                if(pevtCur->bData1==0x59){
                    if(pevtCur->cbData>=2 && (pevtCur->bData[0]<=7 || pevtCur->bData[0]>=121) && pevtCur->bData[1]<=1){
                        cSharp=(pevtCur->bData[0]+64)%128-64;
                        
                        if(!pevtCur->bData[1]){ // Major
                            wsprintf(szBuf,L"%c",'A'+(4*cSharp+30)%7);
                            if(cSharp>=6)
                                lstrcat(szBuf,L"-sharp");
                            if(cSharp<=-2)
                                lstrcat(szBuf,L"-flat");
                            lstrcat(szBuf,L" major");
                        }else{ // Minor
                            wsprintf(szBuf,L"%c",'A'+(4*cSharp+28)%7);
                            if(cSharp>=3)
                                lstrcat(szBuf,L"-sharp");
                            if(cSharp<=-5)
                                lstrcat(szBuf,L"-flat");
                            lstrcat(szBuf,L" minor");
                        }
                        ListView_SetItemText(hwndLVEvtList,dwRow,8,szBuf);
                    }
                }
                break;
            }
        }
        
        pevtCur->fUnfiltered=TRUE;
        pevtCur=pevtCur->pevtNext;
        dwRow++;
    }
    
    SendMessage(hwndLVEvtList,WM_SETREDRAW,TRUE,0);
    return dwRow;
}

DWORD GetTempoList(HWND hwndLVTempoList,MIDIFILE *pmf){
    LVITEM lvitem;

    TEMPOEVENT *ptempoevtCur;
    DWORD dwRow=0;
    WCHAR szBuf[128];

    SendMessage(hwndLVTempoList,WM_SETREDRAW,FALSE,0);

    lvitem.mask=LVIF_TEXT;
    ptempoevtCur=pmf->ptempoevtHead;
    while(ptempoevtCur){
        lvitem.iItem=dwRow;
        lvitem.iSubItem=0;
        wsprintf(szBuf,L"%u",dwRow+1);
        lvitem.pszText=szBuf;
        lvitem.cchTextMax=128;
        ListView_InsertItem(hwndLVTempoList,&lvitem);

        wsprintf(szBuf,L"%u",ptempoevtCur->dwTk);
        ListView_SetItemText(hwndLVTempoList,dwRow,1,szBuf);
        
        wsprintf(szBuf,L"%u",ptempoevtCur->cTk);
        ListView_SetItemText(hwndLVTempoList,dwRow,2,szBuf);
        
        wsprintf(szBuf,L"%u",ptempoevtCur->dwData);
        ListView_SetItemText(hwndLVTempoList,dwRow,3,szBuf);

        swprintf(szBuf,128,L"%f bpm",60000000./ptempoevtCur->dwData);
        ListView_SetItemText(hwndLVTempoList,dwRow,4,szBuf);

        ptempoevtCur=ptempoevtCur->ptempoevtNext;
        dwRow++;
    }

    SendMessage(hwndLVTempoList,WM_SETREDRAW,TRUE,0);
    return dwRow;
}

void AnalyzeTonality(UINT *pcNote,double *pdTonalityProprtn,BOOL *pfTonalityMax){
    UINT uMajScaleNote[]={0,2,4,5,7,9,11};
    UINT uMinScaleNote[]={0,2,3,5,7,8,10,11};
    UINT uMajScaleNoteWeight[]={3,3,3,3,3,3,3};
    UINT uMinScaleNoteWeight[]={3,3,3,3,3,3,1,2};

    double dTonality[24];
    double dTonalityTotal=0, dTonalityMax=0;

    UINT uTonalPivot, uFifthPivot, uMajThirdPivot, uMinThirdPivot, uCurScaleNotePivot;
    double dMajThirdProprtn, dMinThirdProprtn;
    UINT cMajThird, cMinThird, cTonalFifth, cMajScaleNote, cMinScaleNote;

    UINT u;

    for(uTonalPivot=0;uTonalPivot<12;uTonalPivot++){
        if(pcNote[uTonalPivot]>0){
            uFifthPivot=(uTonalPivot+7)%12;
            cTonalFifth=pcNote[uTonalPivot]+pcNote[uFifthPivot];

            uMajThirdPivot=(uTonalPivot+4)%12;
            uMinThirdPivot=(uTonalPivot+3)%12;
            cMajThird=pcNote[uMajThirdPivot];
            cMinThird=pcNote[uMinThirdPivot];
            if(cMajThird+cMinThird){
                dMajThirdProprtn=(double)cMajThird/(cMajThird+cMinThird);
                dMinThirdProprtn=(double)cMinThird/(cMajThird+cMinThird);
            }else{
                dMajThirdProprtn=0.5;
                dMinThirdProprtn=0.5;
            }

            cMajScaleNote=0;
            cMinScaleNote=0;
            for(u=0;u<7;u++){
                uCurScaleNotePivot=(uTonalPivot+uMajScaleNote[u])%12;
                cMajScaleNote+=pcNote[uCurScaleNotePivot]*uMajScaleNoteWeight[u];
            }
            for(u=0;u<8;u++){
                uCurScaleNotePivot=(uTonalPivot+uMinScaleNote[u])%12;
                cMinScaleNote+=pcNote[uCurScaleNotePivot]*uMinScaleNoteWeight[u];
            }

            dTonality[uTonalPivot*2]=(double)cTonalFifth*dMajThirdProprtn*cMajScaleNote; // Value for major tonalities
            dTonality[uTonalPivot*2+1]=(double)cTonalFifth*dMinThirdProprtn*cMinScaleNote; // Value for minor tonalities

            dTonalityTotal+=dTonality[uTonalPivot*2]+dTonality[uTonalPivot*2+1];
        }else{
            dTonality[uTonalPivot*2]=0;
            dTonality[uTonalPivot*2+1]=0;
        }
    }

    ZeroMemory(pfTonalityMax,24*sizeof(BOOL));
    for(u=0;u<24;u++){
        if(dTonalityTotal!=0){
            pdTonalityProprtn[u]=dTonality[u]/dTonalityTotal;
            if(dTonality[u]>dTonalityMax){
                dTonalityMax=dTonality[u];
                ZeroMemory(pfTonalityMax,24*sizeof(BOOL));
                pfTonalityMax[u]=TRUE;
            }else if(dTonality[u]==dTonalityMax){
                pfTonalityMax[u]=TRUE;
            }
        }else{
            pdTonalityProprtn[u]=0;
        }
    }
}

LRESULT CALLBACK TempWndProc(HWND hwnd,UINT uMsg,WPARAM wParam,LPARAM lParam){
    switch(uMsg){
    case WM_COMMAND:
    case WM_NOTIFY:
    case WM_DRAWITEM:
    case WM_HSCROLL:
    case WM_APP_FILTCHECKCHANGE:
    case WM_APP_FILTCHECKALL:
        return SendMessage(hwndMain,uMsg,wParam,lParam);
    }
    return CallWindowProc(GetWindowLong(hwnd,GWL_ID)==IDC_TC ? DefTCProc : DefStaticProc,hwnd,uMsg,wParam,lParam);
}

LRESULT CALLBACK LBCheckedCBProc(HWND hwnd,UINT uMsg,WPARAM wParam,LPARAM lParam){
    int iIndex, cIndex;
    BOOL fNonClient;
    HWND hwndParent;
    int i;

    switch(uMsg){
    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK:
        fNonClient=HIWORD(SendMessage(hwnd,LB_ITEMFROMPOINT,0,lParam));
        if(fNonClient)
            break;

        iIndex=LOWORD(SendMessage(hwnd,LB_ITEMFROMPOINT,0,lParam));
        hwndParent=GetParent((HWND)GetWindowLong(hwnd,GWL_USERDATA)); // The user data of list boxes from combo boxes are set to the combo boxes' window handle

        if(iIndex<2){ // Check all or check none
            cIndex=SendMessage(hwnd,LB_GETCOUNT,0,0);
            for(i=2;i<cIndex;i++)
                SendMessage(hwnd,LB_SETITEMDATA,i,!iIndex);
            SendMessage(hwndParent,WM_APP_FILTCHECKALL,!iIndex,(LPARAM)hwnd);
        }else{
            SendMessage(hwnd,LB_SETITEMDATA,iIndex,!SendMessage(hwnd,LB_GETITEMDATA,iIndex,0));
            SendMessage(hwndParent,WM_APP_FILTCHECKCHANGE,iIndex-2,(LPARAM)hwnd);
        }
        InvalidateRect(hwnd,NULL,FALSE);
        return 0;
    }
    return CallWindowProc(DefLBProc,hwnd,uMsg,wParam,lParam);
}

LRESULT CALLBACK WndProc(HWND hwnd,UINT uMsg,WPARAM wParam,LPARAM lParam){
    static HFONT hfontGUI;
    static TEXTMETRIC tm;
    static long cxChar,cyChar;
    NONCLIENTMETRICS ncm;
    HDC hdc;

    static OSVERSIONINFO osversioninfo;

    HDROP hdrop;
    WCHAR szDragFilePath[MAX_PATH+1];

    static HWND hwndStatus;
    static long cyStatus;
    
    static HWND hwndTool;
    static TOOLINFO ti;

    static HWND hwndTC;
    static TCITEM tci;
    static RECT rcTCDisp;
    static int iCurPage=0;
    static HWND hwndStaticPage[CPAGE];
    static BOOL fTBIsTracking=FALSE;

    static HWND hwndBtnOpen,hwndBtnClose,hwndEditFilePath, hwndBtnPlay,hwndBtnStop,hwndTBTime,
        hwndStaticType, hwndStaticCTrk, hwndStaticTb, hwndStaticCEvt, hwndCBFiltTrk, hwndCBFiltChn, hwndCBFiltEvtType, hwndLBCBFiltTrk, hwndLBCBFiltChn, hwndLBCBFiltEvtType, hwndLVEvtList,
        hwndStaticIniTempo, hwndStaticAvgTempo, hwndStaticDur, hwndStaticEvtDensity, hwndLVTempoList,
        hwndEditTransp, hwndUdTransp, hwndBtnTranspReset, hwndEditTempoR, hwndUdTempoR, hwndBtnTempoRReset, hwndTVMut,
        hwndEditCNote[12], hwndEditCNoteTotal, hwndBtnAnlyzTonality, hwndStaticMostProbTonality, hwndEditTonality[24], hwndStaticTonalityBar[24], hwndStaticTonalityBarLabel[24];
    HWND hwndTemp;

    COMBOBOXINFO cbi;
    LPDRAWITEMSTRUCT lpdis;
    RECT rcCheckBox,rcText;
    UINT uState;
    static FILTERSTATES *pfiltstate=NULL;
    int iLBTopIndex;

    static OPENFILENAME ofn;
    static WCHAR szFilePath[MAX_PATH+1], szFileName[MAX_PATH+1];

    LVCOLUMN lvcol;
    LPNMLVCUSTOMDRAW lplvcd;
    static COLORREF *pcrLVEvtListCD=NULL;

    static MIDIFILE mf;
    static BOOL fAnlyzTonality=FALSE;

    static DWORD cEvtListRow=0, cTempoListRow=0;

    UINT cNoteTotal=0;
    static UINT cNote[12];
    static double dTonalityProprtn[24];
    static BOOL fTonalityMax[24];
    double dTonalityProprtnMax;
    HBRUSH hbr;

    static HMIDISTRM hms;
    static MIDIHDR mhdr[2];
    static UINT uDevID=MIDI_MAPPER;
    static UINT uStrmDataOffset;
    static BYTE *pbStrmData;
    static UINT uCurBuf;
    static UINT uCurOutputEvt, uCurEvtListRow, uCurTempoListRow;
    MIDIEVENT mevtTemp;
    MIDIPROPTIMEDIV mptd;
    MIDIPROPTEMPO mptempo;
    static EVENT *pevtCurBuf, *pevtCurOutput;
    static DWORD dwPrevEvtTk;
    static BOOL fDone;
    UINT cbLongEvtStrmData;
    static DWORD dwCurTk;
    MMTIME mmt;
    static int iCurStrmStatus=STRM_STOP;
    static DWORD dwStartTime;
    static DWORD dwCurTempoData;
    EVENT *pevtTemp;

    static int iTransp;
    static int iTempoR;
    TVINSERTSTRUCT tvis;
    static HTREEITEM htiRoot;
    HTREEITEM htiTrk, htiChn, htiTemp;
    TVHITTESTINFO tvhti;
    int iTVMutIndex, iMutTrk, iMutChn;
    static WORD *pwTrkChnUnmuted;
    static BOOL fPlayControlling; // See the process of the MM_MOM_DONE message

    RECT rc;
    WCHAR szBuf[128];
    int i, iFiltIndex, cLBFiltItem;
    UINT u;
    BOOL f;
    
    switch(uMsg){
    case WM_CREATE:
        InitCommonControls();

        ncm.cbSize=sizeof(NONCLIENTMETRICS);
        SystemParametersInfo(SPI_GETNONCLIENTMETRICS,sizeof(NONCLIENTMETRICS),&ncm,0);
        hfontGUI=CreateFontIndirect(&ncm.lfMessageFont);
        hdc=GetDC(hwnd);
        SelectObject(hdc,hfontGUI);
        GetTextMetrics(hdc,&tm);
        cxChar=tm.tmAveCharWidth;
        cyChar=tm.tmHeight;
        ReleaseDC(hwnd,hdc);

        osversioninfo.dwOSVersionInfoSize=sizeof(OSVERSIONINFO);
        GetVersionEx(&osversioninfo);

        hwndStatus=CreateStatusWindow(WS_CHILD | WS_VISIBLE | CCS_BOTTOM | SBARS_SIZEGRIP,
            L"",hwnd,IDC_STATUS);
        GetWindowRect(hwndStatus,&rc);
        cyStatus=rc.bottom-rc.top;

        hwndTool=CreateWindow(TOOLTIPS_CLASS,NULL,
            TTS_ALWAYSTIP,
            CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,
            hwnd,NULL,hInstanceMain,NULL);
        ti.cbSize=sizeof(TOOLINFO);
        ti.uFlags=TTF_SUBCLASS;
        ti.hwnd=hwnd;

        hwndTC=CreateWindow(WC_TABCONTROL,NULL,
            WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
            0,0,0,0,
            hwnd,(HMENU)IDC_TC,hInstanceMain,NULL);
        SendMessage(hwndTC,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        DefTCProc=(WNDPROC)GetWindowLong(hwndTC,GWL_WNDPROC);
        SetWindowLong(hwndTC,GWL_WNDPROC,(LONG)TempWndProc);
        
        tci.mask=TCIF_TEXT;
        tci.pszText=L"Event list";
        TabCtrl_InsertItem(hwndTC,PAGEEVTLIST,&tci);
        tci.pszText=L"Tempo list";
        TabCtrl_InsertItem(hwndTC,PAGETEMPOLIST,&tci);
        tci.pszText=L"Play control";
        TabCtrl_InsertItem(hwndTC,PAGEPLAYCTL,&tci);
        tci.pszText=L"Tonality analyzer";
        TabCtrl_InsertItem(hwndTC,PAGETONALITYANLYZER,&tci);
        for(i=0;i<CPAGE;i++){
            hwndStaticPage[i]=CreateWindow(L"STATIC",L"",
                WS_CHILD,
                0,0,0,0,
                hwndTC,(HMENU)-1,hInstanceMain,NULL);
            if(i==0)
                DefStaticProc=(WNDPROC)GetWindowLong(hwndStaticPage[i],GWL_WNDPROC);
            SetWindowLong(hwndStaticPage[i],GWL_WNDPROC,(LONG)TempWndProc);
        }
        
        GetClientRect(hwnd,&rcTCDisp);
        TabCtrl_AdjustRect(hwndTC,FALSE,&rcTCDisp);

        /* Controls for common pages */
        hwndBtnOpen=CreateWindow(L"BUTTON",L"Open",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rcTCDisp.left+MulDiv(cxChar,7,4),rcTCDisp.top+MulDiv(cyChar,7,8),MulDiv(cxChar,28,4),MulDiv(cyChar,14,8),
            hwndTC,(HMENU)IDC_BTNOPEN,hInstanceMain,NULL);
        SendMessage(hwndBtnOpen,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndBtnClose=CreateWindow(L"BUTTON",L"Close",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rcTCDisp.left+MulDiv(cxChar,39,4),rcTCDisp.top+MulDiv(cyChar,7,8),MulDiv(cxChar,28,4),MulDiv(cyChar,14,8),
            hwndTC,(HMENU)IDC_BTNCLOSE,hInstanceMain,NULL);
        SendMessage(hwndBtnClose,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndEditFilePath=CreateWindow(L"EDIT",L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_READONLY | ES_AUTOHSCROLL,
            0,0,0,0,
            hwndTC,(HMENU)IDC_EDITFILEPATH,hInstanceMain,NULL);
        SendMessage(hwndEditFilePath,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndBtnPlay=CreateWindow(L"BUTTON",L"Play",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rcTCDisp.left+MulDiv(cxChar,7,4),rcTCDisp.top+MulDiv(cyChar,25,8),MulDiv(cxChar,28,4),MulDiv(cyChar,14,8),
            hwndTC,(HMENU)IDC_BTNPLAY,hInstanceMain,NULL);
        SendMessage(hwndBtnPlay,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndBtnStop=CreateWindow(L"BUTTON",L"Stop",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rcTCDisp.left+MulDiv(cxChar,39,4),rcTCDisp.top+MulDiv(cyChar,25,8),MulDiv(cxChar,28,4),MulDiv(cyChar,14,8),
            hwndTC,(HMENU)IDC_BTNSTOP,hInstanceMain,NULL);
        SendMessage(hwndBtnStop,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndTBTime=CreateWindow(TRACKBAR_CLASS,NULL,
            WS_CHILD | WS_VISIBLE,
            0,0,0,0,
            hwndTC,(HMENU)IDC_TBTIME,hInstanceMain,NULL);
        SendMessage(hwndTBTime,TBM_SETRANGEMIN,TRUE,0);
        SendMessage(hwndTBTime,TBM_SETRANGEMAX,TRUE,0);
        SendMessage(hwndTBTime,TBM_SETPOS,TRUE,0);

        /* Controls for event list page */
        hwndStaticType=CreateWindow(L"STATIC",L"Type: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,7,4),MulDiv(cyChar,46,8),MulDiv(cxChar,56,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGEEVTLIST],(HMENU)IDC_STATICTYPE,hInstanceMain,NULL);
        SendMessage(hwndStaticType,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndStaticCTrk=CreateWindow(L"STATIC",L"Number of tracks: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,67,4),MulDiv(cyChar,46,8),MulDiv(cxChar,104,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGEEVTLIST],(HMENU)IDC_STATICCTRK,hInstanceMain,NULL);
        SendMessage(hwndStaticCTrk,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndStaticTb=CreateWindow(L"STATIC",L"Timebase: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,175,4),MulDiv(cyChar,46,8),MulDiv(cxChar,72,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGEEVTLIST],(HMENU)IDC_STATICTB,hInstanceMain,NULL);
        SendMessage(hwndStaticTb,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndStaticCEvt=CreateWindow(L"STATIC",L"Number of events: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,251,4),MulDiv(cyChar,46,8),MulDiv(cxChar,104,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGEEVTLIST],(HMENU)IDC_STATICCEVT,hInstanceMain,NULL);
        SendMessage(hwndStaticCEvt,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndCBFiltTrk=CreateWindow(L"COMBOBOX",NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            MulDiv(cxChar,7,4),MulDiv(cyChar,61,8),MulDiv(cxChar,122,4),MulDiv(cyChar,162,8),
            hwndStaticPage[PAGEEVTLIST],(HMENU)IDC_CBFILTTRK,hInstanceMain,NULL);
        SendMessage(hwndCBFiltTrk,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndCBFiltChn=CreateWindow(L"COMBOBOX",NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            MulDiv(cxChar,133,4),MulDiv(cyChar,61,8),MulDiv(cxChar,122,4),MulDiv(cyChar,162,8),
            hwndStaticPage[PAGEEVTLIST],(HMENU)IDC_CBFILTCHN,hInstanceMain,NULL);
        SendMessage(hwndCBFiltChn,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndCBFiltEvtType=CreateWindow(L"COMBOBOX",NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            MulDiv(cxChar,259,4),MulDiv(cyChar,61,8),MulDiv(cxChar,122,4),MulDiv(cyChar,162,8),
            hwndStaticPage[PAGEEVTLIST],(HMENU)IDC_CBFILTEVTTYPE,hInstanceMain,NULL);
        SendMessage(hwndCBFiltEvtType,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndLVEvtList=CreateWindow(WC_LISTVIEW,NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            0,0,0,0,
            hwndStaticPage[PAGEEVTLIST],(HMENU)IDC_LVEVTLIST,hInstanceMain,NULL);
        ListView_SetExtendedListViewStyle(hwndLVEvtList,LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);

        /* Controls for tempo list page */
        hwndStaticIniTempo=CreateWindow(L"STATIC",L"Initial tempo: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,7,4),MulDiv(cyChar,46,8),MulDiv(cxChar,136,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGETEMPOLIST],(HMENU)IDC_STATICINITEMPO,hInstanceMain,NULL);
        SendMessage(hwndStaticIniTempo,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndStaticAvgTempo=CreateWindow(L"STATIC",L"Average tempo: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,147,4),MulDiv(cyChar,46,8),MulDiv(cxChar,136,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGETEMPOLIST],(HMENU)IDC_STATICAVGTEMPO,hInstanceMain,NULL);
        SendMessage(hwndStaticAvgTempo,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndStaticDur=CreateWindow(L"STATIC",L"Duration: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,287,4),MulDiv(cyChar,46,8),MulDiv(cxChar,108,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGETEMPOLIST],(HMENU)IDC_STATICDUR,hInstanceMain,NULL);
        SendMessage(hwndStaticDur,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndStaticEvtDensity=CreateWindow(L"STATIC",L"Event density: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,399,4),MulDiv(cyChar,46,8),MulDiv(cxChar,144,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGETEMPOLIST],(HMENU)IDC_STATICEVTDENSITY,hInstanceMain,NULL);
        SendMessage(hwndStaticEvtDensity,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndLVTempoList=CreateWindow(WC_LISTVIEW,NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            0,0,0,0,
            hwndStaticPage[PAGETEMPOLIST],(HMENU)IDC_LVTEMPOLIST,hInstanceMain,NULL);
        ListView_SetExtendedListViewStyle(hwndLVTempoList,LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT);

        /* Controls for play control page */
        hwndTemp=CreateWindow(L"STATIC",L"Transposition:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,7,4),MulDiv(cyChar,49,8),MulDiv(cxChar,64,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGEPLAYCTL],(HMENU)-1,hInstanceMain,NULL);
        SendMessage(hwndTemp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndEditTransp=CreateWindow(L"EDIT",NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            MulDiv(cxChar,74,4),MulDiv(cyChar,46,8),MulDiv(cxChar,34,4),MulDiv(cyChar,14,8),
            hwndStaticPage[PAGEPLAYCTL],(HMENU)IDC_EDITTRANSP,hInstanceMain,NULL);
        SendMessage(hwndEditTransp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndUdTransp=CreateWindow(UPDOWN_CLASS,NULL,
            WS_CHILD | WS_VISIBLE | UDS_ALIGNRIGHT | UDS_ARROWKEYS | UDS_SETBUDDYINT,
            0,0,0,0,
            hwndStaticPage[PAGEPLAYCTL],(HMENU)IDC_UDTRANSP,hInstanceMain,NULL);
        SendMessage(hwndUdTransp,UDM_SETBUDDY,(WPARAM)hwndEditTransp,0);
        SendMessage(hwndUdTransp,UDM_SETRANGE,0,MAKELPARAM(36,-36));
        SendMessage(hwndUdTransp,UDM_SETPOS,0,0);
        hwndBtnTranspReset=CreateWindow(L"BUTTON",L"Reset",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            MulDiv(cxChar,112,4),MulDiv(cyChar,46,8),MulDiv(cxChar,28,4),MulDiv(cyChar,14,8),
            hwndStaticPage[PAGEPLAYCTL],(HMENU)IDC_BTNTRANSPRESET,hInstanceMain,NULL);
        SendMessage(hwndBtnTranspReset,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndTemp=CreateWindow(L"STATIC",L"Tempo ratio:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,159,4),MulDiv(cyChar,49,8),MulDiv(cxChar,56,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGEPLAYCTL],(HMENU)-1,hInstanceMain,NULL);
        SendMessage(hwndTemp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndEditTempoR=CreateWindow(L"EDIT",NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            MulDiv(cxChar,218,4),MulDiv(cyChar,46,8),MulDiv(cxChar,34,4),MulDiv(cyChar,14,8),
            hwndStaticPage[PAGEPLAYCTL],(HMENU)IDC_EDITTEMPOR,hInstanceMain,NULL);
        SendMessage(hwndEditTempoR,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndUdTempoR=CreateWindow(UPDOWN_CLASS,NULL,
            WS_CHILD | WS_VISIBLE | UDS_ALIGNRIGHT | UDS_ARROWKEYS | UDS_SETBUDDYINT,
            0,0,0,0,
            hwndStaticPage[PAGEPLAYCTL],(HMENU)IDC_UDTEMPOR,hInstanceMain,NULL);
        SendMessage(hwndUdTempoR,UDM_SETBUDDY,(WPARAM)hwndEditTempoR,0);
        SendMessage(hwndUdTempoR,UDM_SETRANGE,0,MAKELPARAM(500,20));
        SendMessage(hwndUdTempoR,UDM_SETPOS,0,100);
        hwndTemp=CreateWindow(L"STATIC",L"%",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,255,4),MulDiv(cyChar,49,8),MulDiv(cxChar,12,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGEPLAYCTL],(HMENU)-1,hInstanceMain,NULL);
        SendMessage(hwndTemp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndBtnTempoRReset=CreateWindow(L"BUTTON",L"Reset",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            MulDiv(cxChar,267,4),MulDiv(cyChar,46,8),MulDiv(cxChar,28,4),MulDiv(cyChar,14,8),
            hwndStaticPage[PAGEPLAYCTL],(HMENU)IDC_BTNTEMPORRESET,hInstanceMain,NULL);
        SendMessage(hwndBtnTempoRReset,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndTemp=CreateWindow(L"STATIC",L"Muting:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,7,4),MulDiv(cyChar,67,8),MulDiv(cxChar,36,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGEPLAYCTL],(HMENU)-1,hInstanceMain,NULL);
        SendMessage(hwndTemp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndTVMut=CreateWindow(WC_TREEVIEW,NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | TVS_HASBUTTONS | TVS_LINESATROOT | TVS_HASLINES | TVS_NOHSCROLL,
            0,0,0,0,
            hwndStaticPage[PAGEPLAYCTL],(HMENU)IDC_TVMUT,hInstanceMain,NULL);
        SetWindowLong(hwndTVMut,GWL_STYLE,GetWindowLong(hwndTVMut,GWL_STYLE) | TVS_CHECKBOXES); // According to https://devblogs.microsoft.com/oldnewthing/20171127-00/?p=97465, we have to do so
        
        /* Controls for tonality analysis page */
        hwndTemp=CreateWindow(L"STATIC",L"Note count:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            MulDiv(cxChar,7,4),MulDiv(cyChar,46,8),MulDiv(cxChar,68,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGETONALITYANLYZER],(HMENU)-1,hInstanceMain,NULL);
        SendMessage(hwndTemp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        for(u=0;u<12;u++){
            hwndTemp=CreateWindow(L"STATIC",szNoteLabel[u],
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                MulDiv(cxChar,7+36*u,4),MulDiv(cyChar,57,8),MulDiv(cxChar,32,4),MulDiv(cyChar,8,8),
                hwndStaticPage[PAGETONALITYANLYZER],(HMENU)-1,hInstanceMain,NULL);
            SendMessage(hwndTemp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
            hwndEditCNote[u]=CreateWindow(L"EDIT",L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_AUTOHSCROLL,
                MulDiv(cxChar,7+36*u,4),MulDiv(cyChar,68,8),MulDiv(cxChar,32,4),MulDiv(cyChar,14,8),
                hwndStaticPage[PAGETONALITYANLYZER],(HMENU)(IDC_EDITCNOTEFIRST+u),hInstanceMain,NULL);
            SendMessage(hwndEditCNote[u],WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        }
        hwndTemp=CreateWindow(L"STATIC",L"Total",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            MulDiv(cxChar,439,4),MulDiv(cyChar,57,8),MulDiv(cxChar,32,4),MulDiv(cyChar,8,8),
            hwndStaticPage[PAGETONALITYANLYZER],(HMENU)-1,hInstanceMain,NULL);
        SendMessage(hwndTemp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndEditCNoteTotal=CreateWindow(L"EDIT",L"0",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_READONLY | ES_AUTOHSCROLL,
            MulDiv(cxChar,439,4),MulDiv(cyChar,68,8),MulDiv(cxChar,32,4),MulDiv(cyChar,14,8),
            hwndStaticPage[PAGETONALITYANLYZER],(HMENU)IDC_EDITCNOTETOTAL,hInstanceMain,NULL);
        SendMessage(hwndEditCNoteTotal,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndBtnAnlyzTonality=CreateWindow(L"BUTTON",L"Analyze tonality",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            MulDiv(cxChar,475,4),MulDiv(cyChar,68,8),MulDiv(cxChar,72,4),MulDiv(cyChar,14,8),
            hwndStaticPage[PAGETONALITYANLYZER],(HMENU)IDC_BTNANLYZTONALITY,hInstanceMain,NULL);
        SendMessage(hwndBtnAnlyzTonality,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        hwndStaticMostProbTonality=CreateWindow(L"STATIC",L"Most probable tonality: ",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            0,0,0,0,
            hwndStaticPage[PAGETONALITYANLYZER],(HMENU)IDC_STATICMOSTPROBTONALITY,hInstanceMain,NULL);
        SendMessage(hwndStaticMostProbTonality,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        for(u=0;u<24;u++){
            hwndTemp=CreateWindow(L"STATIC",szTonalityLabel[u],
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                MulDiv(cxChar,7+36*u,4),MulDiv(cyChar,104,8),MulDiv(cxChar,32,4),MulDiv(cyChar,8,8),
                hwndStaticPage[PAGETONALITYANLYZER],(HMENU)-1,hInstanceMain,NULL);
            SendMessage(hwndTemp,WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
            hwndEditTonality[u]=CreateWindow(L"EDIT",L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_READONLY | ES_AUTOHSCROLL,
                MulDiv(cxChar,7+36*u,4),MulDiv(cyChar,115,8),MulDiv(cxChar,32,4),MulDiv(cyChar,14,8),
                hwndStaticPage[PAGETONALITYANLYZER],(HMENU)(IDC_EDITTONALITYFIRST+u),hInstanceMain,NULL);
            SendMessage(hwndEditTonality[u],WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
            hwndStaticTonalityBar[u]=CreateWindow(L"STATIC",L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | SS_OWNERDRAW,
                0,0,0,0,
                hwndStaticPage[PAGETONALITYANLYZER],(HMENU)(IDC_STATICTONALITYBARFIRST+u),hInstanceMain,NULL);
            SetWindowLong(hwndStaticTonalityBar[u],GWL_WNDPROC,(LONG)TempWndProc);
            hwndStaticTonalityBarLabel[u]=CreateWindow(L"STATIC",szTonalityLabel[u],
                WS_CHILD | WS_VISIBLE | SS_CENTER,
                0,0,0,0,
                hwndStaticPage[PAGETONALITYANLYZER],(HMENU)(IDC_STATICTONALITYBARLABELFIRST+u),hInstanceMain,NULL);
            SendMessage(hwndStaticTonalityBarLabel[u],WM_SETFONT,(WPARAM)hfontGUI,(LPARAM)TRUE);
        }

        ofn.lStructSize=sizeof(OPENFILENAME);
        ofn.hwndOwner=hwnd;
        ofn.lpstrFilter=L"MIDI Files (*.mid;*.midi;*.kar;*.rmi)\0*.mid;*.midi;*.kar;*.rmi\0"
            L"All Files (*.*)\0*.*\0\0";
        ofn.nFilterIndex=0;
        ofn.lpstrFile=szFilePath;
        ofn.nMaxFile=MAX_PATH+1;
        ofn.lpstrFileTitle=szFileName;
        ofn.nMaxFileTitle=MAX_PATH+1;
        ofn.Flags=OFN_FILEMUSTEXIST;

        lvcol.mask=LVCF_FMT | LVCF_WIDTH | LVCF_TEXT;
        lvcol.fmt=LVCFMT_LEFT;
        lvcol.cx=cxChar*10;
        lvcol.pszText=L"#";
        ListView_InsertColumn(hwndLVEvtList,0,&lvcol);
        lvcol.pszText=L"Track";
        ListView_InsertColumn(hwndLVEvtList,1,&lvcol);
        lvcol.pszText=L"Channel";
        ListView_InsertColumn(hwndLVEvtList,2,&lvcol);
        lvcol.cx=cxChar*15;
        lvcol.pszText=L"Start tick";
        ListView_InsertColumn(hwndLVEvtList,3,&lvcol);
        lvcol.cx=cxChar*25;
        lvcol.pszText=L"Event type";
        ListView_InsertColumn(hwndLVEvtList,4,&lvcol);
        lvcol.cx=cxChar*10;
        lvcol.pszText=L"Data 1";
        ListView_InsertColumn(hwndLVEvtList,5,&lvcol);
        lvcol.cx=cxChar*25;
        lvcol.pszText=L"Comment";
        ListView_InsertColumn(hwndLVEvtList,6,&lvcol);
        lvcol.cx=cxChar*10;
        lvcol.pszText=L"Data 2";
        ListView_InsertColumn(hwndLVEvtList,7,&lvcol);
        lvcol.cx=cxChar*50;
        lvcol.pszText=L"Comment";
        ListView_InsertColumn(hwndLVEvtList,8,&lvcol);

        lvcol.cx=cxChar*10;
        lvcol.pszText=L"#";
        ListView_InsertColumn(hwndLVTempoList,0,&lvcol);
        lvcol.cx=cxChar*15;
        lvcol.pszText=L"Start tick";
        ListView_InsertColumn(hwndLVTempoList,1,&lvcol);
        lvcol.pszText=L"Lasting tick";
        ListView_InsertColumn(hwndLVTempoList,2,&lvcol);
        lvcol.cx=cxChar*20;
        lvcol.pszText=L"Tempo data";
        ListView_InsertColumn(hwndLVTempoList,3,&lvcol);
        lvcol.cx=cxChar*30;
        lvcol.pszText=L"Tempo";
        ListView_InsertColumn(hwndLVTempoList,4,&lvcol);

        cbi.cbSize=sizeof(COMBOBOXINFO);
        SendMessage(hwndCBFiltTrk,CB_GETCOMBOBOXINFO,0,(LPARAM)&cbi);
        hwndLBCBFiltTrk=cbi.hwndList;
        DefLBProc=(WNDPROC)GetWindowLong(hwndLBCBFiltTrk,GWL_WNDPROC);
        SendMessage(hwndCBFiltChn,CB_GETCOMBOBOXINFO,0,(LPARAM)&cbi);
        hwndLBCBFiltChn=cbi.hwndList;
        SendMessage(hwndCBFiltEvtType,CB_GETCOMBOBOXINFO,0,(LPARAM)&cbi);
        hwndLBCBFiltEvtType=cbi.hwndList;
        SetWindowLong(hwndLBCBFiltTrk,GWL_WNDPROC,(LONG)LBCheckedCBProc);
        SetWindowLong(hwndLBCBFiltChn,GWL_WNDPROC,(LONG)LBCheckedCBProc);
        SetWindowLong(hwndLBCBFiltEvtType,GWL_WNDPROC,(LONG)LBCheckedCBProc);
        SetWindowLong(hwndLBCBFiltTrk,GWL_USERDATA,(LONG)hwndCBFiltTrk);
        SetWindowLong(hwndLBCBFiltChn,GWL_USERDATA,(LONG)hwndCBFiltChn);
        SetWindowLong(hwndLBCBFiltEvtType,GWL_USERDATA,(LONG)hwndCBFiltEvtType);

        ZeroMemory(&mf,sizeof(MIDIFILE));
        for(u=0;u<24;u++)
            dTonalityProprtn[u]=0;
        
        ShowWindow(hwndStaticPage[PAGEEVTLIST],SW_SHOW);

        if(szCmdFilePath){
            PostMessage(hwnd,WM_APP_OPENFILE,0,(LPARAM)szCmdFilePath);
        }
        
        SendMessage(hwndCBFiltTrk,CB_ADDSTRING,0,(LPARAM)L"Select all");
        SendMessage(hwndCBFiltTrk,CB_ADDSTRING,0,(LPARAM)L"Select none");
        SendMessage(hwndCBFiltChn,CB_ADDSTRING,0,(LPARAM)L"Select all");
        SendMessage(hwndCBFiltChn,CB_ADDSTRING,0,(LPARAM)L"Select none");
        SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Select all");
        SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Select none");

        iTransp=0;
        iTempoR=100;
        return 0;

    case WM_SIZE:
        SendMessage(hwndStatus,uMsg,wParam,lParam);
        MoveWindow(hwndTC,0,0,LOWORD(lParam),HIWORD(lParam)-cyStatus,TRUE);
        GetClientRect(hwnd,&rcTCDisp);
        rcTCDisp.bottom-=cyStatus;
        TabCtrl_AdjustRect(hwndTC,FALSE,&rcTCDisp);
        for(i=0;i<CPAGE;i++){
            MoveWindow(hwndStaticPage[i],rcTCDisp.left,rcTCDisp.top,rcTCDisp.right-rcTCDisp.left,rcTCDisp.bottom-rcTCDisp.top,TRUE);
        }

        MoveWindow(hwndEditFilePath,rcTCDisp.left+MulDiv(cxChar,71,4),rcTCDisp.top+MulDiv(cyChar,7,8),rcTCDisp.right-rcTCDisp.left-MulDiv(cxChar,78,4),MulDiv(cyChar,14,8),TRUE);
        MoveWindow(hwndTBTime,rcTCDisp.left+MulDiv(cxChar,71,4),rcTCDisp.top+MulDiv(cyChar,25,8),rcTCDisp.right-rcTCDisp.left-MulDiv(cxChar,78,4),MulDiv(cyChar,14,8),TRUE);
        
        MoveWindow(hwndLVEvtList,MulDiv(cxChar,7,4),MulDiv(cyChar,osversioninfo.dwMajorVersion>=6?75:79,8),rcTCDisp.right-rcTCDisp.left-MulDiv(cxChar,14,4),rcTCDisp.bottom-rcTCDisp.top-MulDiv(cyChar,osversioninfo.dwMajorVersion>=6?82:86,8),TRUE);

        MoveWindow(hwndLVTempoList,MulDiv(cxChar,7,4),MulDiv(cyChar,61,8),rcTCDisp.right-rcTCDisp.left-MulDiv(cxChar,14,4),rcTCDisp.bottom-rcTCDisp.top-MulDiv(cyChar,68,8),TRUE);

        MoveWindow(hwndTVMut,MulDiv(cxChar,7,4),MulDiv(cyChar,78,8),rcTCDisp.right-rcTCDisp.left-MulDiv(cxChar,14,4),rcTCDisp.bottom-rcTCDisp.top-MulDiv(cyChar,85,8),TRUE);

        MoveWindow(hwndStaticMostProbTonality,MulDiv(cxChar,7,4),MulDiv(cyChar,89,8),rcTCDisp.right-rcTCDisp.left-MulDiv(cxChar,7,4),MulDiv(cyChar,8,8),TRUE);
        for(u=0;u<24;u++){
            MoveWindow(hwndStaticTonalityBar[u],MulDiv(cxChar,7+24*u,4),MulDiv(cyChar,133,8),MulDiv(cxChar,24,4),rcTCDisp.bottom-rcTCDisp.top-MulDiv(cyChar,151,8),TRUE);
            MoveWindow(hwndStaticTonalityBarLabel[u],MulDiv(cxChar,7+24*u,4),rcTCDisp.bottom-rcTCDisp.top-MulDiv(cyChar,15,8),MulDiv(cxChar,24,4),MulDiv(cyChar,8,8),TRUE);
            
            ti.uId=IDC_TTTONALITYBARFIRST+u;
            GetWindowRect(hwndStaticTonalityBar[u],&ti.rect);
            MapWindowPoints(HWND_DESKTOP,hwnd,(LPPOINT)&ti.rect,2);
            SendMessage(hwndTool,TTM_NEWTOOLRECT,0,(LPARAM)&ti);
        }
        return 0;

    case WM_DROPFILES:
        hdrop=(HDROP)wParam;
        DragQueryFile(hdrop,0,szDragFilePath,MAX_PATH+1);
        DragFinish(hdrop);
        SendMessage(hwnd,WM_APP_OPENFILE,0,(LPARAM)szDragFilePath);
        return 0;

    case WM_COMMAND:
        if(LOWORD(wParam)>=IDC_EDITCNOTEFIRST && LOWORD(wParam)<IDC_EDITCNOTEFIRST+12 && HIWORD(wParam)==EN_UPDATE){
            cNoteTotal=0;
            for(u=0;u<12;u++){
                cNoteTotal+=GetDlgItemInt(hwndStaticPage[PAGETONALITYANLYZER],IDC_EDITCNOTEFIRST+u,NULL,FALSE);
            }
            SetDlgItemInt(hwndStaticPage[PAGETONALITYANLYZER],IDC_EDITCNOTETOTAL,cNoteTotal,FALSE);
        }

        switch(LOWORD(wParam)){
        case IDC_BTNOPEN:
            if(!GetOpenFileName(&ofn))
                break;

            SendMessage(hwnd,WM_APP_OPENFILE,0,(LPARAM)szFilePath);
            break;

        case IDC_BTNCLOSE:
            SendMessage(hwnd,WM_APP_CLOSEFILE,0,0);
            break;

        case IDC_BTNPLAY:
            if(!mf.fOpened)
                break;

            switch(iCurStrmStatus){
            case STRM_STOP: // Play
                dwStartTime=SendMessage(hwndTBTime,TBM_GETPOS,0,0);
                if(!(pevtCurOutput=GetEvtByMs(&mf,dwStartTime,&dwPrevEvtTk,&dwCurTempoData)))
                    break;
                SendMessage(hwnd,WM_APP_PLAYFROMEVT,0,(LPARAM)pevtCurOutput);
                iCurStrmStatus=STRM_PLAY;
                SetWindowText(hwndBtnPlay,L"Pause");
                break;
            case STRM_PLAY: // Pause
                mmt.wType=TIME_MS;
                midiStreamPosition(hms,&mmt,sizeof(MMTIME));
                SendMessage(hwndTBTime,TBM_SETPOS,TRUE,mmt.u.ms*iTempoR/100+dwStartTime);

                pevtCurOutput=NULL; // See the process of the MM_MOM_POSITIONCB message
                fDone=TRUE;
                midiStreamStop(hms);
                iCurStrmStatus=STRM_STOP;
                SetWindowText(hwndBtnPlay,L"Play");
                break;
            }
            break;

        case IDC_BTNSTOP:
            SendMessage(hwnd,WM_APP_STOP,0,0);
            break;

        case IDC_BTNANLYZTONALITY:
            SendMessage(hwnd,WM_APP_ANLYZTONALITY,0,0);
            break;

        case IDC_EDITTRANSP:
        case IDC_EDITTEMPOR:
            if(HIWORD(wParam)!=EN_UPDATE)
                break;

            if(iCurStrmStatus==STRM_PLAY){
                mmt.wType=TIME_MS;
                midiStreamPosition(hms,&mmt,sizeof(MMTIME));
                dwStartTime+=mmt.u.ms*iTempoR/100;
                SendMessage(hwndTBTime,TBM_SETPOS,TRUE,dwStartTime);
            }

            if(LOWORD(wParam)==IDC_EDITTRANSP){
                iTransp=GetDlgItemInt(hwndStaticPage[PAGEPLAYCTL],IDC_EDITTRANSP,NULL,TRUE);
            }else{
                iTempoR=GetDlgItemInt(hwndStaticPage[PAGEPLAYCTL],IDC_EDITTEMPOR,NULL,FALSE);
            }

            if(iCurStrmStatus!=STRM_PLAY)
                break;

            fDone=TRUE;
            fPlayControlling=TRUE; // See the process of the MM_MOM_DONE message
            midiStreamStop(hms);
            break;

        case IDC_BTNTRANSPRESET:
            SendMessage(hwndUdTransp,UDM_SETPOS,0,0);
            break;

        case IDC_BTNTEMPORRESET:
            SendMessage(hwndUdTempoR,UDM_SETPOS,0,100);
            break;

        case ID_PLAY:
            SendMessage(hwndBtnPlay,BM_CLICK,0,0);
            break;
        }
        return 0;

    case WM_APP_OPENFILE:
        SendMessage(hwnd,WM_APP_CLOSEFILE,0,0);

        if(!ReadMidi((LPWSTR)lParam,&mf)){
            wsprintf(szBuf,L"Cannot read file \"%s\"!",(LPWSTR)lParam);
            MessageBox(hwnd,szBuf,L"Error",MB_ICONEXCLAMATION);
            break;
        }

        if(szFileName[0]=='\0'){
            i=lstrlen((LPWSTR)lParam);
            while(((LPWSTR)lParam)[i-1]!='\\')
                i--;
            lstrcpy(szFileName,((LPWSTR)lParam)+i);
        }

        SetWindowText(hwndEditFilePath,(LPWSTR)lParam);
            
        wsprintf(szBuf,L"Type: %u",mf.wType);
        SetWindowText(hwndStaticType,szBuf);
        wsprintf(szBuf,L"Number of tracks: %u",mf.cTrk);
        SetWindowText(hwndStaticCTrk,szBuf);
        wsprintf(szBuf,L"Timebase: %u",mf.wTb);
        SetWindowText(hwndStaticTb,szBuf);
        wsprintf(szBuf,L"Number of events: %u",mf.cEvt);
        SetWindowText(hwndStaticCEvt,szBuf);

        swprintf(szBuf,128,L"Initial tempo: %f bpm",mf.dIniTempo);
        SetWindowText(hwndStaticIniTempo,szBuf);
        swprintf(szBuf,128,L"Average tempo: %f bpm",mf.dAvgTempo);
        SetWindowText(hwndStaticAvgTempo,szBuf);
        swprintf(szBuf,128,L"Duration: %f s",mf.dDur);
        SetWindowText(hwndStaticDur,szBuf);
        swprintf(szBuf,128,L"Event density: %f per s",mf.dDur==0. ? 0. : (double)mf.cEvt/mf.dDur);
        SetWindowText(hwndStaticEvtDensity,szBuf);

        pfiltstate=(FILTERSTATES *)malloc(sizeof(FILTERSTATES)+(mf.cTrk-1)*sizeof(DWORD));
        for(u=0;u<mf.cTrk;u++){
            pfiltstate->dwFiltTrk[u]=FILT_CHECKED;
        }
        for(u=0;u<17;u++){
            pfiltstate->dwFiltChn[u]=FILT_CHECKED;
        }
        for(u=0;u<10;u++){
            pfiltstate->dwFiltEvtType[u]=FILT_CHECKED;
        }

        pcrLVEvtListCD=(COLORREF *)malloc(mf.cEvt*sizeof(COLORREF));
        cEvtListRow=GetEvtList(hwndLVEvtList,&mf,pfiltstate,pcrLVEvtListCD);
        cTempoListRow=GetTempoList(hwndLVTempoList,&mf);

        for(u=0,i=2;u<mf.cTrk;u++){
            if(pfiltstate->dwFiltTrk[u] & FILT_AVAILABLE){
                wsprintf(szBuf,L"Track #%u",u);
                SendMessage(hwndCBFiltTrk,CB_ADDSTRING,0,(LPARAM)szBuf);
                SendMessage(hwndLBCBFiltTrk,LB_SETITEMDATA,i,TRUE);
                i++;
            }
        }
        for(u=0,i=2;u<17;u++){
            if(pfiltstate->dwFiltChn[u] & FILT_AVAILABLE){
                if(u==0){
                    SendMessage(hwndCBFiltChn,CB_ADDSTRING,0,(LPARAM)L"Non-channel");
                }else{
                    wsprintf(szBuf,L"Channel #%u",u);
                    SendMessage(hwndCBFiltChn,CB_ADDSTRING,0,(LPARAM)szBuf);
                }
                SendMessage(hwndLBCBFiltChn,LB_SETITEMDATA,i,TRUE);
                i++;
            }
        }
        for(u=0,i=2;u<10;u++){
            if(pfiltstate->dwFiltEvtType[u] & FILT_AVAILABLE){
                switch(u){
                case 0: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Note off"); break;
                case 1: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Note on (with no velocity)"); break;
                case 2: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Note on (with velocity)"); break;
                case 3: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Note aftertouch"); break;
                case 4: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Controller"); break;
                case 5: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Program change"); break;
                case 6: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Channel aftertouch"); break;
                case 7: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Pitch bend"); break;
                case 8: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"System exclusive"); break;
                case 9: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Meta event"); break;
                }
                SendMessage(hwndLBCBFiltEvtType,LB_SETITEMDATA,i,TRUE);
                i++;
            }
        }

        switch(iCurPage){
        case PAGEEVTLIST:
            wsprintf(szBuf,L"%u event(s) in total.",cEvtListRow);
            SendMessage(hwndStatus,SB_SETTEXT,0,(LPARAM)szBuf);
            break;
        case PAGETEMPOLIST:
            wsprintf(szBuf,L"%u tempo event(s) in total.",cTempoListRow);
            SendMessage(hwndStatus,SB_SETTEXT,0,(LPARAM)szBuf);
            break;
        default:
            SendMessage(hwndStatus,SB_SETTEXT,0,(LPARAM)L"");
            break;
        }

        for(u=0;u<12;u++)
            SetDlgItemInt(hwndStaticPage[PAGETONALITYANLYZER],IDC_EDITCNOTEFIRST+u,mf.cNote[u],FALSE);
        SendMessage(hwnd,WM_APP_ANLYZTONALITY,0,0);

        SendMessage(hwndTBTime,TBM_SETRANGEMAX,TRUE,(int)(mf.dDur*1000));
        SendMessage(hwndTBTime,TBM_SETPOS,TRUE,0);

        pwTrkChnUnmuted=(WORD *)malloc(mf.cTrk*sizeof(WORD));
        tvis.hInsertAfter=TVI_LAST;
        tvis.item.mask=TVIF_STATE | TVIF_TEXT;
        tvis.item.state=TVIS_EXPANDED;
        tvis.item.stateMask=TVIS_EXPANDED;
        tvis.hParent=TVI_ROOT;
        tvis.item.pszText=szFileName;
        htiRoot=TreeView_InsertItem(hwndTVMut,&tvis);
        TreeView_SetCheckState(hwndTVMut,htiRoot,TRUE);
        for(u=0;u<mf.cTrk;u++){
            pwTrkChnUnmuted[u]=0xFFFF;

            if(!mf.pwTrkChnUsage[u])
                continue;
            tvis.hParent=htiRoot;
            wsprintf(szBuf,L"Track #%u",u);
            tvis.item.pszText=szBuf;
            htiTrk=TreeView_InsertItem(hwndTVMut,&tvis);
            TreeView_SetCheckState(hwndTVMut,htiTrk,TRUE);
            
            tvis.hParent=htiTrk;
            for(i=0;i<16;i++){
                if(mf.pwTrkChnUsage[u] & (1 << i)){
                    wsprintf(szBuf,L"Channel #%d",i+1);
                    tvis.item.pszText=szBuf;
                    htiChn=TreeView_InsertItem(hwndTVMut,&tvis);
                    TreeView_SetCheckState(hwndTVMut,htiChn,TRUE);
                }
            }
        }

        midiStreamOpen(&hms,&uDevID,1,(DWORD)hwnd,0,CALLBACK_WINDOW);
        dwPrevEvtTk=0;

        fPlayControlling=FALSE;
        return 0;

    case WM_APP_CLOSEFILE:
        if(!mf.fOpened)
            break;

        szFileName[0]='\0';

        cLBFiltItem=SendMessage(hwndLBCBFiltTrk,LB_GETCOUNT,0,0);
        for(i=2;i<cLBFiltItem;i++){
            SendMessage(hwndCBFiltTrk,CB_DELETESTRING,2,0);
        }
        cLBFiltItem=SendMessage(hwndLBCBFiltChn,LB_GETCOUNT,0,0);
        for(i=2;i<cLBFiltItem;i++){
            SendMessage(hwndCBFiltChn,CB_DELETESTRING,2,0);
        }
        cLBFiltItem=SendMessage(hwndLBCBFiltEvtType,LB_GETCOUNT,0,0);
        for(i=2;i<cLBFiltItem;i++){
            SendMessage(hwndCBFiltEvtType,CB_DELETESTRING,2,0);
        }

        SendMessage(hwnd,WM_APP_STOP,0,0);

        midiOutReset((HMIDIOUT)hms);
        for(u=0;u<2;u++){
            midiOutUnprepareHeader((HMIDIOUT)hms,&mhdr[u],sizeof(MIDIHDR));
        }
        midiStreamClose(hms);

        free(pfiltstate);
        pfiltstate=NULL;

        SetWindowText(hwndEditFilePath,L"");
        FreeMidi(&mf);
        SendMessage(hwndTBTime,TBM_SETRANGEMAX,TRUE,0);

        SetWindowText(hwndStaticType,L"Type: ");
        SetWindowText(hwndStaticCTrk,L"Number of tracks: ");
        SetWindowText(hwndStaticTb,L"Timebase: ");
        SetWindowText(hwndStaticCEvt,L"Number of events: ");

        SetWindowText(hwndStaticIniTempo,L"Initial tempo: ");
        SetWindowText(hwndStaticAvgTempo,L"Average tempo: ");
        SetWindowText(hwndStaticDur,L"Duration: ");
        SetWindowText(hwndStaticEvtDensity,L"Event density: ");

        if(pcrLVEvtListCD){
            free(pcrLVEvtListCD);
            pcrLVEvtListCD=NULL;
        }
        ListView_DeleteAllItems(hwndLVEvtList);
        ListView_DeleteAllItems(hwndLVTempoList);

        SendMessage(hwndStatus,SB_SETTEXT,0,(LPARAM)L"");

        fAnlyzTonality=FALSE;

        for(u=0;u<12;u++)
            SetWindowText(hwndEditCNote[u],L"");

        for(u=0;u<24;u++){
            SetWindowText(hwndEditTonality[u],L"");
            dTonalityProprtn[u]=0;
            InvalidateRect(hwndStaticTonalityBar[u],NULL,TRUE);

            ti.uId=IDC_TTTONALITYBARFIRST+u;
            SendMessage(hwndTool,TTM_DELTOOL,0,(LPARAM)&ti);
        }

        SetWindowText(hwndStaticMostProbTonality,L"Most probable tonality: ");

        free(pwTrkChnUnmuted);
        TreeView_DeleteAllItems(hwndTVMut);
        return 0;

    case WM_HSCROLL:
        if((HWND)lParam==hwndTBTime){
            switch(LOWORD(wParam)){
            case SB_THUMBTRACK:
                fTBIsTracking=TRUE;
                break;
            case SB_ENDSCROLL:
                fTBIsTracking=FALSE;
                break;
            case SB_THUMBPOSITION:
            case SB_PAGELEFT:
            case SB_PAGERIGHT:
                if(iCurStrmStatus!=STRM_PLAY)
                    break;
                
                dwStartTime=SendMessage(hwndTBTime,TBM_GETPOS,0,0);

                fDone=TRUE;
                fPlayControlling=TRUE; // See the process of the MM_MOM_DONE message
                midiStreamStop(hms);
                break;
            }
        }
        return 0;

    case MM_MOM_OPEN:
        for(u=0;u<2;u++){
            mhdr[u].lpData=(LPSTR)malloc(STRMBUFLEN);
            mhdr[u].dwBufferLength=STRMBUFLEN;
        }

        mptd.cbStruct=sizeof(MIDIPROPTIMEDIV);
        mptd.dwTimeDiv=mf.wTb;
        midiStreamProperty(hms,(LPBYTE)&mptd,MIDIPROP_SET | MIDIPROP_TIMEDIV);
        return 0;

    case WM_APP_PLAYFROMEVT:
        mptempo.cbStruct=sizeof(MIDIPROPTEMPO);
        mptempo.dwTempo=max(min(dwCurTempoData*100/iTempoR,0xFFFFFF),1);
        midiStreamProperty(hms,(LPBYTE)&mptempo,MIDIPROP_SET | MIDIPROP_TEMPO);

        uCurBuf=0;
        pbStrmData=(BYTE *)mhdr[0].lpData;
        uStrmDataOffset=0;

        pevtCurBuf=mf.pevtHead;
        uCurOutputEvt=0;
        uCurEvtListRow=-1;
        uCurTempoListRow=-1;
        while(pevtCurBuf!=(EVENT *)lParam){
            uCurOutputEvt++;
            if(pevtCurBuf->fUnfiltered)
                uCurEvtListRow++;
            if(pevtCurBuf->bStatus==0xFF && pevtCurBuf->bData1==0x51)
                uCurTempoListRow++;
            pevtCurBuf=pevtCurBuf->pevtNext;
        }
        ListView_SetItemState(hwndLVEvtList,uCurEvtListRow,LVIS_SELECTED,LVIS_SELECTED);
        ListView_EnsureVisible(hwndLVEvtList,uCurEvtListRow,FALSE);
        ListView_SetItemState(hwndLVTempoList,uCurTempoListRow,LVIS_SELECTED,LVIS_SELECTED);
        ListView_EnsureVisible(hwndLVTempoList,uCurTempoListRow,FALSE);
        fDone=FALSE;

        /* Firstly fill the two buffers at once */
        while(pevtCurBuf){
            if(pevtCurBuf->bStatus==0xF0 || pevtCurBuf->bStatus==0xF7)
                cbLongEvtStrmData=(pevtCurBuf->cbData+3)/4*4;
            else
                cbLongEvtStrmData=0;

            if(uStrmDataOffset+12+cbLongEvtStrmData>STRMBUFLEN){
                if(uCurBuf==0){ // The first buffer has been finished filling
                    mhdr[0].dwBytesRecorded=uStrmDataOffset;
                    mhdr[0].dwFlags=0;
                    uCurBuf=1;
                    pbStrmData=(BYTE *)mhdr[1].lpData;
                    uStrmDataOffset=0;
                }else{ // The second buffer has also been finished filling
                    break;
                }
            }
            mevtTemp.dwDeltaTime=pevtCurBuf->dwTk-dwPrevEvtTk;
            mevtTemp.dwStreamID=0;
            if(cbLongEvtStrmData){
                mevtTemp.dwEvent=pevtCurBuf->cbData | MEVT_LONGMSG<<24;
            }else if(pevtCurBuf->bStatus==0xFF){
                if(pevtCurBuf->bData1==0x51){
                    mevtTemp.dwEvent=max(min((pevtCurBuf->bData[0] << 16 | pevtCurBuf->bData[1] << 8 | pevtCurBuf->bData[2])*100/iTempoR,0xFFFFFF),1) | MEVT_TEMPO<<24;
                }else{
                    mevtTemp.dwEvent=MEVT_NOP<<24;
                }
            }else if(pwTrkChnUnmuted[pevtCurBuf->wTrk] & (1 << (pevtCurBuf->bStatus & 0xF))){
                if((pevtCurBuf->bStatus&0xF)!=0x9 && pevtCurBuf->bStatus>=0x80 && pevtCurBuf->bStatus<0xB0){
                    if((int)pevtCurBuf->bData1+iTransp>=0 && (int)pevtCurBuf->bData1+iTransp<128){
                        mevtTemp.dwEvent=MAKEFOURCC(pevtCurBuf->bStatus, pevtCurBuf->bData1+iTransp, pevtCurBuf->bData2, MEVT_SHORTMSG);
                    }else{
                        mevtTemp.dwEvent=MEVT_NOP<<24;
                    }
                }else{
                    mevtTemp.dwEvent=MAKEFOURCC(pevtCurBuf->bStatus, pevtCurBuf->bData1, pevtCurBuf->bData2, MEVT_SHORTMSG);
                }
            }else{
                mevtTemp.dwEvent=MEVT_NOP<<24;
            }
            mevtTemp.dwEvent |= MEVT_F_CALLBACK;
            memcpy(pbStrmData+uStrmDataOffset,&mevtTemp,12);
            /* Write the data of long events */
            if(cbLongEvtStrmData)
                memcpy(pbStrmData+uStrmDataOffset+12,pevtCurBuf->bData,pevtCurBuf->cbData);
            uStrmDataOffset+=12+cbLongEvtStrmData;

            dwPrevEvtTk=pevtCurBuf->dwTk;
            pevtCurBuf=pevtCurBuf->pevtNext;
        }
        mhdr[uCurBuf].dwBytesRecorded=uStrmDataOffset;
        mhdr[uCurBuf].dwFlags=0;

        /* Begin outputting with the first buffer */
        midiOutPrepareHeader((HMIDIOUT)hms,&mhdr[0],sizeof(MIDIHDR));
        midiStreamOut(hms,&mhdr[0],sizeof(MIDIHDR));

        uCurBuf=0;

        midiStreamRestart(hms);
        return 0;

    case MM_MOM_DONE: // When one buffer has been finished outputting
        if(fDone){
            /* Since the MM_MOM_DONE message will always be processed after the WM_APP_PLAYFROMEVENT message due to unknown reasons, even if the former is sent before the latter, the WM_APP_PLAYFROMEVENT message sending when play controlling was moved here */
            if(fPlayControlling){
                if(!(pevtCurOutput=GetEvtByMs(&mf,dwStartTime,&dwPrevEvtTk,&dwCurTempoData)))
                    break;
                SendMessage(hwnd,WM_APP_PLAYFROMEVT,0,(LPARAM)pevtCurOutput);
                fPlayControlling=FALSE;
            }
            return 0;
        }

        /* Output the other buffer, while refill the buffer just been finished outputting */
        midiOutPrepareHeader((HMIDIOUT)hms,&mhdr[(uCurBuf+1)%2],sizeof(MIDIHDR));
        midiStreamOut(hms,&mhdr[(uCurBuf+1)%2],sizeof(MIDIHDR));

        pbStrmData=(BYTE *)mhdr[uCurBuf].lpData;
        uStrmDataOffset=0;

        while(pevtCurBuf){
            if(pevtCurBuf->bStatus==0xF0 || pevtCurBuf->bStatus==0xF7)
                cbLongEvtStrmData=(pevtCurBuf->cbData+3)/4*4;
            else
                cbLongEvtStrmData=0;

            if(uStrmDataOffset+12+cbLongEvtStrmData>STRMBUFLEN){ // The buffer just been finished has been finished refilling
                break;
            }
            mevtTemp.dwDeltaTime=pevtCurBuf->dwTk-dwPrevEvtTk;
            mevtTemp.dwStreamID=0;
            if(cbLongEvtStrmData){
                mevtTemp.dwEvent=pevtCurBuf->cbData | MEVT_LONGMSG<<24;
            }else if(pevtCurBuf->bStatus==0xFF){
                if(pevtCurBuf->bData1==0x51){
                    mevtTemp.dwEvent=max(min((pevtCurBuf->bData[0] << 16 | pevtCurBuf->bData[1] << 8 | pevtCurBuf->bData[2])*100/iTempoR,0xFFFFFF),1) | MEVT_TEMPO<<24;
                }else{
                    mevtTemp.dwEvent=MEVT_NOP<<24;
                }
            }else if(pwTrkChnUnmuted[pevtCurBuf->wTrk] & (1 << (pevtCurBuf->bStatus & 0xF))){
                if((pevtCurBuf->bStatus&0xF)!=0x9 && pevtCurBuf->bStatus>=0x80 && pevtCurBuf->bStatus<0xB0){
                    if((int)pevtCurBuf->bData1+iTransp>=0 && (int)pevtCurBuf->bData1+iTransp<128){
                        mevtTemp.dwEvent=MAKEFOURCC(pevtCurBuf->bStatus, pevtCurBuf->bData1+iTransp, pevtCurBuf->bData2, MEVT_SHORTMSG);
                    }else{
                        mevtTemp.dwEvent=MEVT_NOP<<24;
                    }
                }else{
                    mevtTemp.dwEvent=MAKEFOURCC(pevtCurBuf->bStatus, pevtCurBuf->bData1, pevtCurBuf->bData2, MEVT_SHORTMSG);
                }
            }else{
                mevtTemp.dwEvent=MEVT_NOP<<24;
            }
            mevtTemp.dwEvent |= MEVT_F_CALLBACK;
            memcpy(pbStrmData+uStrmDataOffset,&mevtTemp,12);
            /* Write the data of long events */
            if(cbLongEvtStrmData)
                memcpy(pbStrmData+uStrmDataOffset+12,pevtCurBuf->bData,pevtCurBuf->cbData);
            uStrmDataOffset+=12+cbLongEvtStrmData;

            dwPrevEvtTk=pevtCurBuf->dwTk;
            pevtCurBuf=pevtCurBuf->pevtNext;
        }
        mhdr[uCurBuf].dwBytesRecorded=uStrmDataOffset;
        mhdr[uCurBuf].dwFlags=0;

        uCurBuf=(uCurBuf+1)%2;
        return 0;

    case MM_MOM_POSITIONCB:
        if(!pevtCurOutput){
            /* To prevent potentially continue outputting after pausing or stopping, which may cause bugs */
            break;
        }

        uCurOutputEvt++;
        if(pevtCurOutput->fUnfiltered){
            uCurEvtListRow++;
            ListView_SetItemState(hwndLVEvtList,uCurEvtListRow,LVIS_SELECTED,LVIS_SELECTED);
            ListView_EnsureVisible(hwndLVEvtList,uCurEvtListRow,FALSE);
        }
        if(pevtCurOutput->bStatus==0xFF && pevtCurOutput->bData1==0x51){
            uCurTempoListRow++;
            ListView_SetItemState(hwndLVTempoList,uCurTempoListRow,LVIS_SELECTED,LVIS_SELECTED);
            ListView_EnsureVisible(hwndLVTempoList,uCurTempoListRow,FALSE);
        }
        if(!fTBIsTracking){
            mmt.wType=TIME_MS;
            midiStreamPosition(hms,&mmt,sizeof(MMTIME));
            SendMessage(hwndTBTime,TBM_SETPOS,TRUE,mmt.u.ms*iTempoR/100+dwStartTime);
        }

        pevtCurOutput=pevtCurOutput->pevtNext;

        if(uCurOutputEvt==mf.cEvt){
            SendMessage(hwnd,WM_APP_STOP,0,0);
        }
        return 0;

    case WM_APP_STOP:
        if(!mf.fOpened)
            return 0;

        pevtCurOutput=NULL; // See the process of the MM_MOM_POSITIONCB message
        fDone=TRUE;
        midiStreamStop(hms);
        if(iCurStrmStatus!=STRM_STOP)
            SetWindowText(hwndBtnPlay,L"Play");
        iCurStrmStatus=STRM_STOP;
        SendMessage(hwndTBTime,TBM_SETPOS,TRUE,0);
        return 0;

    case WM_APP_ANLYZTONALITY:
        fAnlyzTonality=TRUE;
        dTonalityProprtnMax=0;

        for(u=0;u<12;u++)
            cNote[u]=GetDlgItemInt(hwndStaticPage[PAGETONALITYANLYZER],IDC_EDITCNOTEFIRST+u,NULL,FALSE);

        AnalyzeTonality(cNote,dTonalityProprtn,fTonalityMax);
        for(u=0;u<24;u++){
            swprintf(szBuf,128,L"%.2f%%",dTonalityProprtn[u]*100);
            SetWindowText(hwndEditTonality[u],szBuf);
            InvalidateRect(hwndStaticTonalityBar[u],NULL,TRUE);

            if(iCurPage==PAGETONALITYANLYZER){
                ti.uId=IDC_TTTONALITYBARFIRST+u;
                SendMessage(hwndTool,TTM_DELTOOL,0,(LPARAM)&ti);
                GetWindowRect(hwndStaticTonalityBar[u],&ti.rect);
                MapWindowPoints(HWND_DESKTOP,hwnd,(LPPOINT)&ti.rect,2);
                ti.lpszText=szBuf;
                SendMessage(hwndTool,TTM_ADDTOOL,0,(LPARAM)&ti);
            }
        }

        for(u=0;u<24;u++){
            if(fTonalityMax[u]){
                if(dTonalityProprtnMax==0){
                    wsprintf(szBuf,L"Most probable tonality: %s",szTonalityLabel[u]);
                    dTonalityProprtnMax=dTonalityProprtn[u];
                }else{
                    wsprintf(szBuf,L"%s, %s",szBuf,szTonalityLabel[u]);
                }
            }
        }
        if(dTonalityProprtnMax!=0){
            swprintf(szBuf,128,L"%s (%.2f%%)",szBuf,dTonalityProprtnMax*100);
            SetWindowText(hwndStaticMostProbTonality,szBuf);
        }else{
            SetWindowText(hwndStaticMostProbTonality,L"Most probable tonality: ");
        }
        return 0;

    case WM_NOTIFY:
        switch(((NMHDR *)lParam)->idFrom){
        case IDC_TC:
            switch(((NMHDR *)lParam)->code){
            case TCN_SELCHANGE:
                ShowWindow(hwndStaticPage[iCurPage],SW_HIDE);
                iCurPage=TabCtrl_GetCurSel(hwndTC);
                ShowWindow(hwndStaticPage[iCurPage],SW_SHOW);

                if(mf.fOpened){
                    switch(iCurPage){
                    case PAGEEVTLIST:
                        wsprintf(szBuf,L"%u event(s) in total.",cEvtListRow);
                        SendMessage(hwndStatus,SB_SETTEXT,0,(LPARAM)szBuf);
                        break;
                    case PAGETEMPOLIST:
                        wsprintf(szBuf,L"%u tempo event(s) in total.",cTempoListRow);
                        SendMessage(hwndStatus,SB_SETTEXT,0,(LPARAM)szBuf);
                        break;
                    default:
                        SendMessage(hwndStatus,SB_SETTEXT,0,(LPARAM)L"");
                        break;
                    }
                }

                if(fAnlyzTonality){
                    for(u=0;u<24;u++){
                        ti.uId=IDC_TTTONALITYBARFIRST+u;
                        if(iCurPage==PAGETONALITYANLYZER){
                            swprintf(szBuf,128,L"%.2f%%",dTonalityProprtn[u]*100);
                            GetWindowRect(hwndStaticTonalityBar[u],&ti.rect);
                            MapWindowPoints(HWND_DESKTOP,hwnd,(LPPOINT)&ti.rect,2);
                            ti.lpszText=szBuf;
                            SendMessage(hwndTool,TTM_ADDTOOL,0,(LPARAM)&ti);
                        }else{
                            SendMessage(hwndTool,TTM_DELTOOL,0,(LPARAM)&ti);
                        }
                    }
                }
                break;
            }
            break;

        case IDC_LVEVTLIST:
            switch(((NMHDR *)lParam)->code){
            case NM_CUSTOMDRAW:
                lplvcd=(LPNMLVCUSTOMDRAW)lParam;
                switch(lplvcd->nmcd.dwDrawStage){
                case CDDS_PREPAINT:
                    return CDRF_NOTIFYITEMDRAW;
                case CDDS_ITEMPREPAINT:
                    lplvcd->clrText=pcrLVEvtListCD[lplvcd->nmcd.dwItemSpec];
                    return CDRF_NEWFONT;
                }
                break;
            }
            break;

        case IDC_TVMUT:
            switch(((NMHDR *)lParam)->code){
            case NM_CLICK:
            case TVN_KEYDOWN:
                if(((NMHDR *)lParam)->code==NM_CLICK){
                    GetCursorPos(&tvhti.pt);
                    MapWindowPoints(NULL,hwndTVMut,&tvhti.pt,1);
                    TreeView_HitTest(hwndTVMut,&tvhti);
                    if(tvhti.flags!=TVHT_ONITEMSTATEICON)
                        break;
                    htiTemp=tvhti.hItem;
                }else{
                    if(((NMTVKEYDOWN *)lParam)->wVKey!=VK_SPACE)
                        break;
                    if(!(htiTemp=TreeView_GetSelection(hwndTVMut)))
                        break;
                }

                if(!TreeView_GetParent(hwndTVMut,htiTemp)){ // Mute or unmute all
                    f=!TreeView_GetCheckState(hwndTVMut,htiRoot);
                    for(htiTrk=TreeView_GetChild(hwndTVMut,htiRoot),iMutTrk=0;htiTrk;htiTrk=TreeView_GetNextSibling(hwndTVMut,htiTrk),iMutTrk++){
                        while(!mf.pwTrkChnUsage[iMutTrk])
                            iMutTrk++;
                        TreeView_SetCheckState(hwndTVMut,htiTrk,f);

                        for(htiChn=TreeView_GetChild(hwndTVMut,htiTrk),iMutChn=0;htiChn;htiChn=TreeView_GetNextSibling(hwndTVMut,htiChn),iMutChn++){
                            while(!(mf.pwTrkChnUsage[iMutTrk] & (1 << iMutChn)))
                                iMutChn++;
                            if(f)
                                pwTrkChnUnmuted[iMutTrk] |= (1 << iMutChn);
                            else
                                pwTrkChnUnmuted[iMutTrk] &= ~(1 << iMutChn);
                            TreeView_SetCheckState(hwndTVMut,htiChn,f);
                        }
                    }
                }else if(TreeView_GetParent(hwndTVMut,htiTemp)==htiRoot){ // Mute or unmute a track
                    htiTrk=htiTemp;

                    iTVMutIndex=0;
                    while(htiTemp=TreeView_GetPrevSibling(hwndTVMut,htiTemp))
                        iTVMutIndex++;
                    for(iMutTrk=0,i=0;;iMutTrk++,i++){
                        while(!mf.pwTrkChnUsage[iMutTrk])
                            iMutTrk++;
                        if(i==iTVMutIndex) break;
                    }

                    f=!TreeView_GetCheckState(hwndTVMut,htiTrk);
                    for(htiChn=TreeView_GetChild(hwndTVMut,htiTrk),iMutChn=0;htiChn;htiChn=TreeView_GetNextSibling(hwndTVMut,htiChn),iMutChn++){
                        while(!(mf.pwTrkChnUsage[iMutTrk] & (1 << iMutChn)))
                            iMutChn++;
                        if(f)
                            pwTrkChnUnmuted[iMutTrk] |= (1 << iMutChn);
                        else
                            pwTrkChnUnmuted[iMutTrk] &= ~(1 << iMutChn);
                        TreeView_SetCheckState(hwndTVMut,htiChn,f);
                    }

                    f=TRUE;
                    for(u=0;u<mf.cTrk;u++){
                        if(pwTrkChnUnmuted[u]!=0xFFFF){
                            f=FALSE;
                            break;
                        }
                    }
                    TreeView_SetCheckState(hwndTVMut,htiRoot,f);
                }else{ // Mute or unmute a channel within a track
                    htiTrk=TreeView_GetParent(hwndTVMut,htiTemp);
                    htiChn=htiTemp;

                    iTVMutIndex=0;
                    htiTemp=htiTrk;
                    while(htiTemp=TreeView_GetPrevSibling(hwndTVMut,htiTemp))
                        iTVMutIndex++;
                    for(iMutTrk=0,i=0;;iMutTrk++,i++){
                        while(!mf.pwTrkChnUsage[iMutTrk])
                            iMutTrk++;
                        if(i==iTVMutIndex) break;
                    }

                    iTVMutIndex=0;
                    htiTemp=htiChn;
                    while(htiTemp=TreeView_GetPrevSibling(hwndTVMut,htiTemp))
                        iTVMutIndex++;
                    for(iMutChn=0,i=0;;iMutChn++,i++){
                        while(!(mf.pwTrkChnUsage[iMutTrk] & (1 << iMutChn)))
                            iMutChn++;
                        if(i==iTVMutIndex) break;
                    }
                    pwTrkChnUnmuted[iMutTrk] ^= (1 << iMutChn);

                    TreeView_SetCheckState(hwndTVMut,htiTrk,pwTrkChnUnmuted[iMutTrk]==0xFFFF);

                    f=TRUE;
                    for(u=0;u<mf.cTrk;u++){
                        if(pwTrkChnUnmuted[u]!=0xFFFF){
                            f=FALSE;
                            break;
                        }
                    }
                    TreeView_SetCheckState(hwndTVMut,htiRoot,f);
                }

                if(iCurStrmStatus!=STRM_PLAY)
                    break;
                
                mmt.wType=TIME_MS;
                midiStreamPosition(hms,&mmt,sizeof(MMTIME));
                dwStartTime+=mmt.u.ms*iTempoR/100;
                SendMessage(hwndTBTime,TBM_SETPOS,TRUE,dwStartTime);
            
                fDone=TRUE;
                fPlayControlling=TRUE; // See the process of the MM_MOM_DONE message
                midiStreamStop(hms);
                break;
            }
            break;
        }
        return 0;

    case WM_APP_FILTCHECKCHANGE:
    case WM_APP_FILTCHECKALL:
        if(!mf.fOpened)
            return 0;

        cLBFiltItem=SendMessage((HWND)lParam,LB_GETCOUNT,0,0);

        if(uMsg==WM_APP_FILTCHECKCHANGE){
            if((HWND)lParam==hwndLBCBFiltTrk){
                for(i=0,iFiltIndex=0;;i++,iFiltIndex++){
                    while(!(pfiltstate->dwFiltTrk[iFiltIndex] & FILT_AVAILABLE))
                        iFiltIndex++;
                    if(i==wParam) break;
                }
                pfiltstate->dwFiltTrk[iFiltIndex] ^= FILT_CHECKED;
            }
            if((HWND)lParam==hwndLBCBFiltChn){
                for(i=0,iFiltIndex=0;;i++,iFiltIndex++){
                    while(!(pfiltstate->dwFiltChn[iFiltIndex] & FILT_AVAILABLE))
                        iFiltIndex++;
                    if(i==wParam) break;
                }
                pfiltstate->dwFiltChn[iFiltIndex] ^= FILT_CHECKED;
            }
            if((HWND)lParam==hwndLBCBFiltEvtType){
                for(i=0,iFiltIndex=0;;i++,iFiltIndex++){
                    while(!(pfiltstate->dwFiltEvtType[iFiltIndex] & FILT_AVAILABLE))
                        iFiltIndex++;
                    if(i==wParam) break;
                }
                pfiltstate->dwFiltEvtType[iFiltIndex] ^= FILT_CHECKED;
            }
        }else{
            if((HWND)lParam==hwndLBCBFiltTrk){
                for(iFiltIndex=0;;iFiltIndex++){
                    while(iFiltIndex<mf.cTrk && !(pfiltstate->dwFiltTrk[iFiltIndex] & FILT_AVAILABLE))
                        iFiltIndex++;
                    if(iFiltIndex==mf.cTrk)
                        break;
                    if(wParam)
                        pfiltstate->dwFiltTrk[iFiltIndex] |= FILT_CHECKED;
                    else
                        pfiltstate->dwFiltTrk[iFiltIndex] &= ~FILT_CHECKED;
                }
            }
            if((HWND)lParam==hwndLBCBFiltChn){
                for(iFiltIndex=0;;iFiltIndex++){
                    while(iFiltIndex<17 && !(pfiltstate->dwFiltChn[iFiltIndex] & FILT_AVAILABLE))
                        iFiltIndex++;
                    if(iFiltIndex==17)
                        break;
                    if(wParam)
                        pfiltstate->dwFiltChn[iFiltIndex] |= FILT_CHECKED;
                    else
                        pfiltstate->dwFiltChn[iFiltIndex] &= ~FILT_CHECKED;
                }
            }
            if((HWND)lParam==hwndLBCBFiltEvtType){
                for(iFiltIndex=0;;iFiltIndex++){
                    while(iFiltIndex<10 && !(pfiltstate->dwFiltEvtType[iFiltIndex] & FILT_AVAILABLE))
                        iFiltIndex++;
                    if(iFiltIndex==10)
                        break;
                    if(wParam)
                        pfiltstate->dwFiltEvtType[iFiltIndex] |= FILT_CHECKED;
                    else
                        pfiltstate->dwFiltEvtType[iFiltIndex] &= ~FILT_CHECKED;
                }
            }
        }

        for(u=0;u<mf.cTrk;u++){
            pfiltstate->dwFiltTrk[u] &= ~FILT_AVAILABLE;
        }
        for(u=0;u<17;u++){
            pfiltstate->dwFiltChn[u] &= ~FILT_AVAILABLE;
        }
        for(u=0;u<10;u++){
            pfiltstate->dwFiltEvtType[u] &= ~FILT_AVAILABLE;
        }
    
        cEvtListRow=GetEvtList(hwndLVEvtList,&mf,pfiltstate,pcrLVEvtListCD);
        wsprintf(szBuf,L"%u event(s) in total.",cEvtListRow);
        SendMessage(hwndStatus,SB_SETTEXT,0,(LPARAM)szBuf);

        iLBTopIndex=SendMessage((HWND)lParam,LB_GETTOPINDEX,0,0);
        SendMessage((HWND)lParam,WM_SETREDRAW,FALSE,0);

        cLBFiltItem=SendMessage(hwndLBCBFiltTrk,LB_GETCOUNT,0,0);
        for(i=2;i<cLBFiltItem;i++){
            SendMessage(hwndCBFiltTrk,CB_DELETESTRING,2,0);
        }
        cLBFiltItem=SendMessage(hwndLBCBFiltChn,LB_GETCOUNT,0,0);
        for(i=2;i<cLBFiltItem;i++){
            SendMessage(hwndCBFiltChn,CB_DELETESTRING,2,0);
        }
        cLBFiltItem=SendMessage(hwndLBCBFiltEvtType,LB_GETCOUNT,0,0);
        for(i=2;i<cLBFiltItem;i++){
            SendMessage(hwndCBFiltEvtType,CB_DELETESTRING,2,0);
        }

        for(u=0,i=2;u<mf.cTrk;u++){
            if(pfiltstate->dwFiltTrk[u] & FILT_AVAILABLE){
                wsprintf(szBuf,L"Track #%u",u);
                SendMessage(hwndCBFiltTrk,CB_ADDSTRING,0,(LPARAM)szBuf);
                SendMessage(hwndLBCBFiltTrk,LB_SETITEMDATA,i,pfiltstate->dwFiltTrk[u] & FILT_CHECKED);
                i++;
            }
        }
        for(u=0,i=2;u<17;u++){
            if(pfiltstate->dwFiltChn[u] & FILT_AVAILABLE){
                if(u==0){
                    SendMessage(hwndCBFiltChn,CB_ADDSTRING,0,(LPARAM)L"Non-channel");
                }else{
                    wsprintf(szBuf,L"Channel #%u",u);
                    SendMessage(hwndCBFiltChn,CB_ADDSTRING,0,(LPARAM)szBuf);
                }
                SendMessage(hwndLBCBFiltChn,LB_SETITEMDATA,i,pfiltstate->dwFiltChn[u] & FILT_CHECKED);
                i++;
            }
        }
        for(u=0,i=2;u<10;u++){
            if(pfiltstate->dwFiltEvtType[u] & FILT_AVAILABLE){
                switch(u){
                case 0: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Note off"); break;
                case 1: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Note on (with no velocity)"); break;
                case 2: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Note on (with velocity)"); break;
                case 3: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Note aftertouch"); break;
                case 4: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Controller"); break;
                case 5: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Program change"); break;
                case 6: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Channel aftertouch"); break;
                case 7: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Pitch bend"); break;
                case 8: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"System exclusive"); break;
                case 9: SendMessage(hwndCBFiltEvtType,CB_ADDSTRING,0,(LPARAM)L"Meta event"); break;
                }
                SendMessage(hwndLBCBFiltEvtType,LB_SETITEMDATA,i,pfiltstate->dwFiltEvtType[u] & FILT_CHECKED);
                i++;
            }
        }

        SendMessage((HWND)lParam,LB_SETTOPINDEX,iLBTopIndex,0);
        SendMessage((HWND)lParam,WM_SETREDRAW,TRUE,0);

        if(iCurStrmStatus==STRM_PLAY){
            pevtTemp=mf.pevtHead;
            uCurEvtListRow=-1;
            while(pevtTemp!=pevtCurOutput){
                if(pevtTemp->fUnfiltered)
                    uCurEvtListRow++;
                pevtTemp=pevtTemp->pevtNext;
            }
            ListView_SetItemState(hwndLVEvtList,uCurEvtListRow,LVIS_SELECTED,LVIS_SELECTED);
            ListView_EnsureVisible(hwndLVEvtList,uCurEvtListRow,FALSE);
        }
        return 0;

    case WM_DRAWITEM:
        lpdis=(LPDRAWITEMSTRUCT)lParam;

        if(wParam>=IDC_STATICTONALITYBARFIRST && wParam<IDC_STATICTONALITYBARFIRST+24){
            FillRect(lpdis->hDC,&lpdis->rcItem,(HBRUSH)GetStockObject(WHITE_BRUSH));
            hbr=CreateSolidBrush((wParam-IDC_STATICTONALITYBARFIRST)%2==0 ? RGB(255,0,255) : RGB(0,0,255));
            lpdis->rcItem.top+=(int)((double)(lpdis->rcItem.bottom-lpdis->rcItem.top)*(1.-dTonalityProprtn[wParam-IDC_STATICTONALITYBARFIRST]));
            FillRect(lpdis->hDC,&lpdis->rcItem,hbr);
            DeleteObject(hbr);
        }

        if(wParam>=IDC_CBFILTTRK && wParam<=IDC_CBFILTEVTTYPE){
            if(lpdis->itemState & ODS_COMBOBOXEDIT){
                hdc=lpdis->hDC;
                rcText=lpdis->rcItem;
                if(wParam==IDC_CBFILTTRK)
                    DrawText(hdc,L"Filter by track",-1,&rcText,DT_VCENTER);
                if(wParam==IDC_CBFILTCHN)
                    DrawText(hdc,L"Filter by channel",-1,&rcText,DT_VCENTER);
                if(wParam==IDC_CBFILTEVTTYPE)
                    DrawText(hdc,L"Filter by event type",-1,&rcText,DT_VCENTER);
                return 0;
            }

            if(lpdis->itemID<2){
                hdc=lpdis->hDC;
                rcText=lpdis->rcItem;
                if(wParam==IDC_CBFILTTRK)
                    SendMessage(hwndLBCBFiltTrk,LB_GETTEXT,lpdis->itemID,(LPARAM)szBuf);
                if(wParam==IDC_CBFILTCHN)
                    SendMessage(hwndLBCBFiltChn,LB_GETTEXT,lpdis->itemID,(LPARAM)szBuf);
                if(wParam==IDC_CBFILTEVTTYPE)
                    SendMessage(hwndLBCBFiltEvtType,LB_GETTEXT,lpdis->itemID,(LPARAM)szBuf);
                DrawText(hdc,szBuf,-1,&rcText,DT_VCENTER);
                return 0;
            }

            hdc=lpdis->hDC;
            rcCheckBox=lpdis->rcItem;
            rcText=lpdis->rcItem;

            rcCheckBox.right=rcCheckBox.left+rcCheckBox.bottom-rcCheckBox.top;
            rcText.left=rcCheckBox.right;

            uState=DFCS_BUTTONCHECK;
            if(wParam==IDC_CBFILTTRK){
                if(SendMessage(hwndLBCBFiltTrk,LB_GETITEMDATA,lpdis->itemID,0))
                    uState |= DFCS_CHECKED;
                SendMessage(hwndLBCBFiltTrk,LB_GETTEXT,lpdis->itemID,(LPARAM)szBuf);
            }
            if(wParam==IDC_CBFILTCHN){
                if(SendMessage(hwndLBCBFiltChn,LB_GETITEMDATA,lpdis->itemID,0))
                    uState |= DFCS_CHECKED;
                SendMessage(hwndLBCBFiltChn,LB_GETTEXT,lpdis->itemID,(LPARAM)szBuf);
            }
            if(wParam==IDC_CBFILTEVTTYPE){
                if(SendMessage(hwndLBCBFiltEvtType,LB_GETITEMDATA,lpdis->itemID,0))
                    uState |= DFCS_CHECKED;
                SendMessage(hwndLBCBFiltEvtType,LB_GETTEXT,lpdis->itemID,(LPARAM)szBuf);
            }
            DrawFrameControl(hdc,&rcCheckBox,DFC_BUTTON,uState);
            DrawText(hdc,szBuf,-1,&rcText,DT_VCENTER);
        }
        return TRUE;

    case WM_DESTROY:
        SendMessage(hwnd,WM_APP_CLOSEFILE,0,0);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hwnd,uMsg,wParam,lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance,LPWSTR szCmdLine,int iCmdShow){
    HWND hwnd;
    MSG msg;
    WNDCLASS wc;

    HACCEL haccel;

    HWND hwndFocused;
    WCHAR szBuf[128];

    if(lstrlen(szCmdLine)){
        szCmdFilePath=szCmdLine;
        if(szCmdFilePath[0]=='\"'){
            szCmdFilePath++;
            szCmdFilePath[lstrlen(szCmdFilePath)-1]='\0';
        }
    }

    wc.style=CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc=WndProc;
    wc.cbClsExtra=0;
    wc.cbWndExtra=0;
    wc.hInstance=hInstance;
    wc.hIcon=LoadIcon(NULL,IDI_APPLICATION);
    wc.hCursor=LoadCursor(NULL,IDC_ARROW);
    wc.hbrBackground=(HBRUSH)(COLOR_BTNFACE+1);
    wc.lpszMenuName=NULL;
    wc.lpszClassName=L"ANKE";
    if(!RegisterClass(&wc))
        return -1;
    
    hwnd=CreateWindowEx(WS_EX_ACCEPTFILES,L"ANKE",L"Anke Midi Reader",
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,
        NULL,NULL,hInstance,NULL);
    if(!hwnd)
        return -1;

    haccel=LoadAccelerators(hInstance,L"ANKE");

    hInstanceMain=hInstance;
    hwndMain=hwnd;
    
    while(GetMessage(&msg,NULL,0,0)){
        if(!TranslateAccelerator(hwnd,haccel,&msg)){
            /* Handle Ctrl+A for edit controls */
            if(msg.message==WM_KEYDOWN && msg.wParam=='A' && GetKeyState(VK_CONTROL)<0){
                hwndFocused=GetFocus();
                GetClassName(hwndFocused,szBuf,128);
                if(hwndFocused && lstrcmp(szBuf,L"EDIT")){
                    SendMessage(hwndFocused,EM_SETSEL,0,-1);
                    continue;
                }
            }

            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    
    return msg.wParam;
}