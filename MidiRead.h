typedef struct evt {
    struct evt *pevtNext;
    DWORD dwTk;
    WORD wTrk;
    DWORD cbData; // For meta or sys-ex events
    BYTE bStatus;
    BYTE bData1;
    BYTE bData2;
    BYTE abData[1]; // For meta or sys-ex events
} EVENT;

typedef struct tempoevt {
    struct tempoevt *ptempoevtNext;
    DWORD dwTk;
    DWORD cTk;
    DWORD dwData;
} TEMPOEVENT;

typedef struct mf {
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

    UINT acNote[12]; // For tonality analysis

    DWORD cTk;
    DWORD cPlayableTk;

    double dDur;
    double dIniTempo;
    double dAvgTempo;

    WORD *pwTrkChnUsage; // Channel usage for each track, represented by bits
    EVENT **ppevtTrkName; // Track name event for each track
} MIDIFILE;

BYTE ReadByte(MIDIFILE *pmf) {
    return pmf->pb[pmf->dwLoc++];
}

/* Read a multi-byte big-endian integer */
DWORD ReadBEInt(MIDIFILE *pmf, int cb) {
    DWORD dwRet = 0;
    int u;
    for(u = 0; u < cb; u++) {
        dwRet = dwRet << 8 | ReadByte(pmf);
    }
    return dwRet;
}

/* Read a variable-length integer */
DWORD ReadVarLenInt(MIDIFILE *pmf) {
    DWORD dwRet = 0;
    BYTE b;
    do {
        b = ReadByte(pmf);
        dwRet = dwRet << 7 | (b & 127);
    } while(b & 128);
    return dwRet;
}

/* Read MIDI file */
BOOL ReadMidi(LPCWSTR lpszPath, MIDIFILE *pmf) {
    HANDLE hFile;
    DWORD dwFileSize;

    EVENT *pevtCur = NULL, *pevtPrev = NULL;
    TEMPOEVENT *ptempoevtCur = NULL, *ptempoevtPrev = NULL;

    DWORD dwCurChkName; // Only used when determining whether it is an RIFF MIDI file
    DWORD dwCurChkSize;
    DWORD dwCurChkEndPos;

    DWORD *pdwTrkLoc = NULL;
    BYTE *pbTrkCurStatus = NULL;
    DWORD *pdwTrkCurTk = NULL;
    DWORD *pdwTrkEndPos = NULL;
    BOOL *pfTrkIsEnd = NULL;
    BOOL fMidiIsEnd;

    DWORD dwCurTk = 0, dwCurPlayableTk = 0;
    WORD wCurEvtTrk;
    BYTE bCurEvtStatus;
    BYTE bCurEvtChn;
    BYTE bCurEvtData1;
    BYTE bCurEvtData2;
    DWORD cbCurEvtData;
    int aiChnCurPrg[16]; // For excluding notes of non-chromatic instruments when counting notes

    DWORD dwCurTempoTk, dwPrevTempoTk, cPrevTempoTk, cLastTempoTk;
    DWORD dwCurTempoData, dwPrevTempoData;
    DWORDLONG qwMusTb; // An intermediate value when calculating duration (equals to microsecond count * timebase)

    UINT u;
    
    
    hFile = CreateFile(lpszPath, GENERIC_READ, 0, NULL,
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if(hFile == INVALID_HANDLE_VALUE)
        return FALSE;
    pmf->pb = (BYTE *)malloc(GetFileSize(hFile, NULL));
    ReadFile(hFile, pmf->pb, GetFileSize(hFile, NULL), &dwFileSize, NULL);
    CloseHandle(hFile);

    pmf->dwLoc = 0;

    /* Read the header chunk */
    if(pmf->dwLoc + 8 > dwFileSize) goto err;
    dwCurChkName = ReadBEInt(pmf, 4);
    if(dwCurChkName == 0x52494646) { // RIFF MIDI file
        if(pmf->dwLoc + 16 > dwFileSize) goto err;
        pmf->dwLoc += 16;
        if(pmf->dwLoc + 8 > dwFileSize) goto err;
        dwCurChkName = ReadBEInt(pmf, 4);
    }
    if(dwCurChkName != 0x4D546864) goto err;
    dwCurChkSize = ReadBEInt(pmf, 4);
    dwCurChkEndPos = pmf->dwLoc + dwCurChkSize;
    if(dwCurChkEndPos > dwFileSize) goto err;
    pmf->wType = ReadBEInt(pmf, 2);
    pmf->cTrk = ReadBEInt(pmf, 2);
    pmf->wTb = ReadBEInt(pmf, 2);
    pmf->dwLoc = dwCurChkEndPos;

    pdwTrkLoc = (DWORD *)malloc(pmf->cTrk * sizeof(DWORD));
    pbTrkCurStatus = (BYTE *)malloc(pmf->cTrk * sizeof(BYTE));
    pdwTrkCurTk = (DWORD *)malloc(pmf->cTrk * sizeof(DWORD));
    pdwTrkEndPos = (DWORD *)malloc(pmf->cTrk * sizeof(DWORD));
    pfTrkIsEnd = (BOOL *)malloc(pmf->cTrk * sizeof(BOOL));

    pmf->pwTrkChnUsage = (WORD *)malloc(pmf->cTrk * sizeof(WORD));
    pmf->ppevtTrkName = (EVENT **)malloc(pmf->cTrk * sizeof(EVENT *));

    /* Locate each track chunk and prepare for event reading */
    u = 0;
    while(u < pmf->cTrk) {
        if(pmf->dwLoc + 8 > dwFileSize) goto err;

        if(ReadBEInt(pmf, 4) != 0x4D54726B) { // Skip non-standard MIDI track chunks
            dwCurChkSize = ReadBEInt(pmf, 4);
            dwCurChkEndPos = pmf->dwLoc + dwCurChkSize;
            if(dwCurChkEndPos > dwFileSize) goto err;
            pmf->dwLoc = dwCurChkEndPos;
            continue;
        }

        dwCurChkSize = ReadBEInt(pmf, 4);
        dwCurChkEndPos = pmf->dwLoc + dwCurChkSize;
        if(dwCurChkEndPos > dwFileSize) goto err;
        pdwTrkEndPos[u] = dwCurChkEndPos;

        pbTrkCurStatus[u] = 0;
        pdwTrkCurTk[u] = ReadVarLenInt(pmf);
        pdwTrkLoc[u] = pmf->dwLoc;
        pfTrkIsEnd[u] = FALSE;

        pmf->pwTrkChnUsage[u] = 0;
        pmf->ppevtTrkName[u] = NULL;

        pmf->dwLoc = dwCurChkEndPos;
        u++;
    }
    
    for(u = 0; u < 16; u++)
        aiChnCurPrg[u] = -1;
    qwMusTb = 0;

    /* Begin to read a new event */
    do {
        if(!pmf->cTrk)
            break;

        /* Find the track with the earliest event to be read */
        u = 0;
        while(pfTrkIsEnd[u]) {
            u++;
        }
        dwCurTk = pdwTrkCurTk[u];
        wCurEvtTrk = u;
        for(; u < pmf->cTrk; u++) {
            if(!pfTrkIsEnd[u] && pdwTrkCurTk[u] < dwCurTk) {
                dwCurTk = pdwTrkCurTk[u];
                wCurEvtTrk = u;
            }
        }
        pmf->dwLoc = pdwTrkLoc[wCurEvtTrk];

        if(pmf->pb[pmf->dwLoc] & 0x80) {
            bCurEvtStatus = ReadByte(pmf);
            if(bCurEvtStatus < 0xF0)
                pbTrkCurStatus[wCurEvtTrk] = bCurEvtStatus;
        } else {
            bCurEvtStatus = pbTrkCurStatus[wCurEvtTrk];
        }

        if(bCurEvtStatus < 0xF0) {
            pmf->pwTrkChnUsage[wCurEvtTrk] |= 1 << (bCurEvtStatus & 0xF);

            bCurEvtChn = (bCurEvtStatus & 0xF) + 1;
            bCurEvtData1 = ReadByte(pmf);

            pevtCur = (EVENT *)malloc(sizeof(EVENT));
            pevtCur->dwTk = dwCurTk;
            pevtCur->wTrk = wCurEvtTrk;
            pevtCur->bStatus = bCurEvtStatus;
            pevtCur->bData1 = bCurEvtData1;

            switch(bCurEvtStatus >> 4) {
            case 0x8:
            case 0x9:
            case 0xA:
            case 0xB:
            case 0xE:
                bCurEvtData2 = ReadByte(pmf);
                pevtCur->bData2 = bCurEvtData2;

                if(bCurEvtStatus >> 4 == 0x9) {
                    /* Count each note, of chromatic instruments only */
                    if(bCurEvtData2 != 0 && bCurEvtChn != 10 && aiChnCurPrg[bCurEvtChn - 1] < 115)
                        pmf->acNote[bCurEvtData1 % 12]++;
                }
                break;

            case 0xC:
                aiChnCurPrg[bCurEvtChn - 1] = bCurEvtData1;
            case 0xD:
                pevtCur->bData2 = 0;
                break;
            }

            dwCurPlayableTk = dwCurTk;
        } else {
            switch(bCurEvtStatus) {
            case 0xF0:
            case 0xF7:
                cbCurEvtData = ReadVarLenInt(pmf);
                if(bCurEvtStatus == 0xF0)
                    cbCurEvtData++;

                pevtCur = (EVENT *)malloc(sizeof(EVENT) + cbCurEvtData - 1);
                pevtCur->dwTk = dwCurTk;
                pevtCur->wTrk = wCurEvtTrk;
                pevtCur->bStatus = bCurEvtStatus;
                pevtCur->cbData = cbCurEvtData;
                if(bCurEvtStatus == 0xF0) {
                    u = 1;
                    pevtCur->abData[0] = 0xF0;
                } else {
                    u = 0;
                }
                for(; u < cbCurEvtData; u++)
                    pevtCur->abData[u] = ReadByte(pmf);
                
                dwCurPlayableTk = dwCurTk;
                break;

            case 0xFF:
                bCurEvtData1 = ReadByte(pmf);
                cbCurEvtData = ReadVarLenInt(pmf);

                pevtCur = (EVENT *)malloc(sizeof(EVENT) + cbCurEvtData - 1);
                pevtCur->dwTk = dwCurTk;
                pevtCur->wTrk = wCurEvtTrk;
                pevtCur->bStatus = bCurEvtStatus;
                pevtCur->bData1 = bCurEvtData1;
                pevtCur->cbData = cbCurEvtData;

                if(bCurEvtData1 == 0x2F) {
                    pfTrkIsEnd[wCurEvtTrk] = TRUE;
                }

                if(bCurEvtData1 == 0x51) {
                    if(cbCurEvtData != 3) goto err;

                    dwCurTempoTk = dwCurTk;
                    dwCurTempoData = 0;
                    for(u = 0; u < 3; u++) {
                        dwCurTempoData = dwCurTempoData << 8 | pmf->pb[pmf->dwLoc + u];
                    }

                    ptempoevtCur = (TEMPOEVENT *)malloc(sizeof(TEMPOEVENT));
                    ptempoevtCur->dwTk = dwCurTempoTk;
                    ptempoevtCur->dwData = dwCurTempoData;

                    if(pmf->cTempoEvt > 0) { // Not the first tempo event
                        /* Calculate the tick count of the previous tempo event */
                        cPrevTempoTk = dwCurTempoTk - dwPrevTempoTk;
                        qwMusTb += (DWORDLONG)dwPrevTempoData * cPrevTempoTk;

                        ptempoevtPrev->cTk = cPrevTempoTk;
                        ptempoevtPrev->ptempoevtNext = ptempoevtCur;
                    } else { // The first tempo event
                        qwMusTb = 500000 * dwCurTempoTk;
                        pmf->dIniTempo = 60000000. / dwCurTempoData;

                        pmf->ptempoevtHead = ptempoevtCur;
                    }

                    dwPrevTempoTk = dwCurTempoTk;
                    dwPrevTempoData = dwCurTempoData;
                    ptempoevtPrev = ptempoevtCur;

                    dwCurPlayableTk = dwCurTk;

                    pmf->cTempoEvt++;
                }
                
                for(u = 0; u < cbCurEvtData; u++)
                    pevtCur->abData[u] = ReadByte(pmf);

                if(bCurEvtData1 == 0x03 && pmf->ppevtTrkName[wCurEvtTrk] == NULL) {
                    pmf->ppevtTrkName[wCurEvtTrk] = pevtCur;
                }
                break;

            default:
                goto err;
            }
        }

        if(pmf->cEvt > 0) { // Not the first event
            pevtPrev->pevtNext = pevtCur;
        } else { // The first event
            pmf->pevtHead = pevtCur;
        }

        pevtPrev = pevtCur;

        pmf->cEvt++;

        pdwTrkLoc[wCurEvtTrk] = pmf->dwLoc;
        pfTrkIsEnd[wCurEvtTrk] = pfTrkIsEnd[wCurEvtTrk] || pdwTrkLoc[wCurEvtTrk] >= pdwTrkEndPos[wCurEvtTrk];

        /* Find whether all tracks have already ended */
        fMidiIsEnd = TRUE;
        for(u = 0; u < pmf->cTrk && fMidiIsEnd; u++)
            fMidiIsEnd = fMidiIsEnd && pfTrkIsEnd[u];

        if(!pfTrkIsEnd[wCurEvtTrk]) {
            pdwTrkCurTk[wCurEvtTrk] += ReadVarLenInt(pmf);
            pdwTrkLoc[wCurEvtTrk] = pmf->dwLoc;
        }
    } while(!fMidiIsEnd);

    /* Complete the event link */
    if(pmf->cEvt != 0) {
        pevtPrev->pevtNext = NULL;
    }

    /* Complete the tempo event link */
    if(pmf->cTempoEvt == 0) { // In the extreme case of MIDI with no tempo events
        qwMusTb = (DWORDLONG)500000 * dwCurPlayableTk;
        pmf->dIniTempo = 0;
    } else {
        /* Calculate the tick count of the last tempo event */
        cLastTempoTk = dwCurPlayableTk - dwPrevTempoTk;
        qwMusTb += (DWORDLONG)dwPrevTempoData * cLastTempoTk;

        ptempoevtPrev->cTk = cLastTempoTk;
        ptempoevtPrev->ptempoevtNext = NULL;
    }

    /* Calculate duration and average tempo */
    pmf->dDur = (double)qwMusTb / pmf->wTb / 1000000;
    if(qwMusTb == 0) { // In the extreme case of zero-duration MIDI
        pmf->dAvgTempo = pmf->dIniTempo;
    } else {
        pmf->dAvgTempo = 60000000. * dwCurPlayableTk / qwMusTb;
    }

    pmf->cTk = dwCurTk;
    pmf->cPlayableTk = dwCurPlayableTk;

    pmf->fOpened = TRUE;

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

void FreeMidi(MIDIFILE *pmf) {
    EVENT *pevtCur, *pevtNext;
    TEMPOEVENT *ptempoevtCur, *ptempoevtNext;
    
    free(pmf->pwTrkChnUsage);
    free(pmf->ppevtTrkName);

    pevtCur = pmf->pevtHead;
    while(pevtCur) {
        pevtNext = pevtCur->pevtNext;
        free(pevtCur);
        pevtCur = pevtNext;
    }

    ptempoevtCur = pmf->ptempoevtHead;
    while(ptempoevtCur) {
        ptempoevtNext = ptempoevtCur->ptempoevtNext;
        free(ptempoevtCur);
        ptempoevtCur = ptempoevtNext;
    }

    ZeroMemory(pmf, sizeof(MIDIFILE));
}

EVENT *GetEvtByMs(MIDIFILE *pmf, DWORD dwMs, DWORD *pdwTk, DWORD *pdwCurTempoData) {
    DWORDLONG qwMusTb = (DWORDLONG)dwMs * pmf->wTb * 1000;

    EVENT *pevtCur;
    DWORD dwCurTempoData = 500000;
    DWORDLONG qwCurMusTb = 0, qwPrevMusTb = 0;
    DWORD dwPrevEvtTk = 0;

    pevtCur = pmf->pevtHead;
    while(qwCurMusTb < qwMusTb && pevtCur) {
        qwPrevMusTb = qwCurMusTb;
        qwCurMusTb += (DWORDLONG)dwCurTempoData * (pevtCur->dwTk - dwPrevEvtTk);

        if(qwCurMusTb >= qwMusTb)
            break;
        dwPrevEvtTk = pevtCur->dwTk;
        if(pevtCur->bStatus == 0xFF && pevtCur->bData1 == 0x51)
            dwCurTempoData = pevtCur->abData[0] << 16 | pevtCur->abData[1] << 8 | pevtCur->abData[2];
        pevtCur = pevtCur->pevtNext;
    }

    *pdwTk = dwPrevEvtTk + (qwMusTb - qwPrevMusTb) / dwCurTempoData;
    *pdwCurTempoData = dwCurTempoData;
    return pevtCur;
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

    if(dTonalityTotal != 0) {
        for(i = 0; i < 24; i++) {
            pdTonalityProprtn[i] = adTonality[i] / dTonalityTotal;
            if(adTonality[i] > dTonalityMax)
                dTonalityMax = adTonality[i];
        }
        for(i = 0; i < 24; i++)
            pfTonalityMax[i] = adTonality[i] == dTonalityMax;
    } else {
        for(i = 0; i < 24; i++) {
            pdTonalityProprtn[i] = 0;
            pfTonalityMax[i] = FALSE;
        }
    }
}