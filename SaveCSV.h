void SaveEvtListAsCSV(MIDIFILE *pmf, FILTERSTATES *pfiltstate, BOOL fHex, LPCWSTR lpszFilePath) {
    FILE *pf;

    EVENT *pevtCur;
    DWORD dwRow = 0;
    LPWSTR lpszDataCmt; // For text-typed meta events
    WCHAR *ch;

    int cSharp = 0; // For key signature events

    UINT u;

    pf = _wfopen(lpszFilePath, L"w, ccs=UTF-8");
    if(!pf)
        return;

    fwprintf(pf, L"#,Track,Channel,Start tick,Event type,Data 1,Data 1 type,Data 2,Comment");

    pevtCur = pmf->pevtHead;
    while(pevtCur) {
        if(!IsEvtUnfiltered(pevtCur, pfiltstate)) {
            pevtCur = pevtCur->pevtNext;
            continue;
        }

        fwprintf(pf, L"\n%u,%u,", dwRow + 1, pevtCur->wTrk);

        if(pevtCur->bStatus < 0xF0) {
            fwprintf(pf, L"%u,%u,", (pevtCur->bStatus & 0xF) + 1, pevtCur->dwTk);

            switch(pevtCur->bStatus >> 4) {
            case 0x8:
                fwprintf(pf, L"Note off,");
                fwprintf(pf, fHex ? L"%02X," : L"%u,", pevtCur->bData1);
                fwprintf(pf, alpszNote[pevtCur->bData1]);
                fwprintf(pf, fHex ? L",%02X," : L",%u,", pevtCur->bData2);
                break;
            case 0x9:
                fwprintf(pf, L"Note on,");
                fwprintf(pf, fHex ? L"%02X," : L"%u,", pevtCur->bData1);
                fwprintf(pf, alpszNote[pevtCur->bData1]);
                fwprintf(pf, fHex ? L",%02X," : L",%u,", pevtCur->bData2);
                break;
            case 0xA:
                fwprintf(pf, L"Note aftertouch,");
                fwprintf(pf, fHex ? L"%02X," : L"%u,", pevtCur->bData1);
                fwprintf(pf, alpszNote[pevtCur->bData1]);
                fwprintf(pf, fHex ? L",%02X," : L",%u,", pevtCur->bData2);
                break;
            case 0xB:
                fwprintf(pf, L"Controller,");
                fwprintf(pf, fHex ? L"%02X," : L"%u,", pevtCur->bData1);
                fwprintf(pf, alpszCtl[pevtCur->bData1]);
                fwprintf(pf, fHex ? L",%02X," : L",%u,", pevtCur->bData2);
                break;
            case 0xC:
                fwprintf(pf, L"Program change,");
                fwprintf(pf, fHex ? L"%02X," : L"%u,", pevtCur->bData1);
                fwprintf(pf, alpszPrg[pevtCur->bData1]);
                fwprintf(pf, L",,");
                break;
            case 0xD:
                fwprintf(pf, L"Channel aftertouch,");
                fwprintf(pf, fHex ? L"%02X,,," : L"%u,", pevtCur->bData1);
                break;
            case 0xE:
                fwprintf(pf, L"Pitch bend,");
                fwprintf(pf, fHex ? L"%02X,," : L"%u,,%u,", pevtCur->bData1, pevtCur->bData2);

                /* Write comment string for pitch bend events */
                fwprintf(pf, L"%d,", (int)pevtCur->bData1 | (pevtCur->bData2 << 7) - 8192);
                break;
            }
        } else {
            fwprintf(pf, L",%u,", pevtCur->dwTk);

            switch(pevtCur->bStatus) {
            case 0xF0:
            case 0xF7:
                fwprintf(pf, L"System exclusive,");
                
                for(u = 0; u < pevtCur->cbData; u++) {
                    fwprintf(pf, fHex ? L"%02X" : L"%u", pevtCur->abData[u]);
                    if(u < pevtCur->cbData - 1)
                        fwprintf(pf, L" ");
                }
                fwprintf(pf, L",,,");
                break;
            case 0xFF:
                fwprintf(pf, L"Meta event,");
                fwprintf(pf, fHex ? L"%02X," : L"%u,", pevtCur->bData1);
                fwprintf(pf, alpszMeta[pevtCur->bData1]);
                fwprintf(pf, L",");
                
                for(u = 0; u < pevtCur->cbData; u++) {
                    fwprintf(pf, fHex ? L"%02X" : L"%u", pevtCur->abData[u]);
                    if(u < pevtCur->cbData - 1)
                        fwprintf(pf, L" ");
                }
                fwprintf(pf, L",");

                /* Write comment string for text-typed meta events */
                if(pevtCur->bData1 >= 0x1 && pevtCur->bData1 <= 0x9) {
                    lpszDataCmt = (LPWSTR)malloc((pevtCur->cbData + 1) * sizeof(WCHAR));
                    ZeroMemory(lpszDataCmt, (pevtCur->cbData + 1) * sizeof(WCHAR));
                    MultiByteToWideChar(CP_ACP, MB_COMPOSITE, (LPCSTR)pevtCur->abData, pevtCur->cbData, lpszDataCmt, pevtCur->cbData * sizeof(WCHAR));

                    fwprintf(pf, L"\"\"\"");
                    for(ch = lpszDataCmt; *ch; ch++) {
                        if(*ch == '"')
                            fwprintf(pf, L"\"\"");
                        else
                            fwprintf(pf, L"%c", *ch);
                    }
                    fwprintf(pf, L"\"\"\"");

                    free(lpszDataCmt);
                }

                /* Write comment string for tempo events */
                if(pevtCur->bData1 == 0x51) {
                    fwprintf(pf, L"%f bpm", 60000000. / (pevtCur->abData[0] << 16 | pevtCur->abData[1] << 8 | pevtCur->abData[2]));
                }

                /* Write comment string for time signature events */
                if(pevtCur->bData1 == 0x58) {
                    if(pevtCur->cbData >= 2) {
                        fwprintf(pf, L"%d/%d", pevtCur->abData[0], 1 << pevtCur->abData[1]);
                    }
                }

                /* Write comment string for key signature events */
                if(pevtCur->bData1 == 0x59) {
                    if(pevtCur->cbData >= 2 && (pevtCur->abData[0] <= 7 || pevtCur->abData[0] >= 121) && pevtCur->abData[1] <= 1){
                        cSharp = (pevtCur->abData[0] + 64) % 128 - 64;

                        if(!pevtCur->abData[1]) { // Major
                            fwprintf(pf, L"%c", 'A' + (4 * cSharp + 30) % 7);
                            if (cSharp >= 6)
                                fwprintf(pf, L"-sharp");
                            if (cSharp <= -2)
                                fwprintf(pf, L"-flat");
                            fwprintf(pf, L" major");
                        } else { // Minor
                            fwprintf(pf, L"%c", 'A' + (4 * cSharp + 28) % 7);
                            if (cSharp >= 3)
                                fwprintf(pf, L"-sharp");
                            if (cSharp <= -5)
                                fwprintf(pf, L"-flat");
                            fwprintf(pf, L" minor");
                        }
                    }
                }
                break;
            }
            
        }
        
        pevtCur = pevtCur->pevtNext;
        dwRow++;
    }

    fclose(pf);
}

void SaveTempoListAsCSV(MIDIFILE *pmf, BOOL fHex, LPCWSTR lpszFilePath) {
    FILE *pf;
    DWORD dwRow = 0;
    TEMPOEVENT *ptempoevtCur;

    pf = _wfopen(lpszFilePath, L"w, ccs=UTF-8");
    if(!pf)
        return;
    
    fwprintf(pf, L"#,Start tick,Lasting ticks,Tempo data,Tempo");
    
    ptempoevtCur = pmf->ptempoevtHead;
    while(ptempoevtCur) {
        fwprintf(pf, L"\n%u,%u,%u,", dwRow + 1, ptempoevtCur->dwTk, ptempoevtCur->cTk);

        fwprintf(pf, fHex ? L"%06X," : L"%u,", ptempoevtCur->dwData);

        fwprintf(pf, L"%f bpm", 60000000. / ptempoevtCur->dwData);

        ptempoevtCur = ptempoevtCur->ptempoevtNext;
        dwRow++;
    }

    fclose(pf);
}