#include <stdio.h>
#include <stdlib.h>
#include <windows.h>
#include <Shlwapi.h>
#include <wchar.h>
#include <iostream>
#include <queue>

#include <mapix.h>
#include <mapiutil.h>
#include <mapi.h>
#include <MSPST.h>
#include "PSTARMime.h"
#include "Source.h"

// define interface GUID
#define USES_IID_IMsgServiceAdmin2
#define USES_IID_IMessage
#include <initguid.h>
#include <MAPIguid.h>
#include <mapiaux.h>
#include <imessage.h>
// Class Identifiers
// {4e3a7680-b77a-11d0-9da5-00c04fd65685}
DEFINE_GUID(CLSID_IConverterSession, 0x4e3a7680, 0xb77a, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);
// Interface Identifiers
// {4b401570-b77b-11d0-9da5-00c04fd65685}
DEFINE_GUID(IID_IConverterSession, 0x4b401570, 0xb77b, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);


using namespace std;

typedef struct folder_struct{
    wchar_t *folder_name;
    LPMAPIFOLDER folderP;
}folder_s;


const char mProfileName[] = "profile_PSTAR";
bool mVerbose = true;
IConverterSession *mlpConverter = NULL;

void logdebug(const wchar_t* format, ...)
{
    wchar_t fmt[256] = { 0 };

    if (mVerbose)
    {
        va_list argList;
        va_start(argList, format);
        wvsprintf(fmt, format, argList);
        va_end(argList);

        wprintf(fmt);
    }
}

IConverterSession* CreateConverter()
{
    if (!mlpConverter)
        CoCreateInstance(CLSID_IConverterSession,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_IConverterSession,
        (LPVOID*)&mlpConverter);

    return mlpConverter;
}

void DeleteConverter()
{
    if (mlpConverter)
        mlpConverter->Release();
}

HRESULT OpenDefaultMsgStore(LPMAPISESSION session, LPMDB *lppDefaultMDB)
{
    HRESULT hRes = S_OK;
    LPMAPITABLE pStoresTbl = NULL;
    SRestriction sres;
    SPropValue spv;
    LPSRowSet  pRow = NULL;

    // define sptEIDCol
    enum
    {
        EID,
        NUM_COLS
    };
    const SizedSPropTagArray(NUM_COLS, sptEIDCol) =
    {
        NUM_COLS,
        PR_ENTRYID,
    };

    hRes = session->GetMsgStoresTable(0, &pStoresTbl);

    if (SUCCEEDED(hRes))
    {
        sres.rt = RES_PROPERTY; // gonna compare a property
        sres.res.resProperty.relop = RELOP_EQ; // gonna test equality
        sres.res.resProperty.ulPropTag = PR_DEFAULT_STORE; // tag to compare
        sres.res.resProperty.lpProp = &spv; // prop tag to compare against

        spv.ulPropTag = PR_DEFAULT_STORE; // tag type
        spv.Value.b = true; // tag value

        hRes = HrQueryAllRows(
            pStoresTbl,                     // table to query
            (LPSPropTagArray)&sptEIDCol,    // columns to get
            &sres,                          // restriction to use
            NULL,                           // sort order
            0,                              // max number of rows - 0 means ALL
            &pRow);
        if (pRow && pRow->cRows)
        {
            SBinary lpEID = pRow->aRow[0].lpProps[EID].Value.bin;
            hRes = session->OpenMsgStore(
                NULL,
                lpEID.cb,
                (LPENTRYID)lpEID.lpb,
                NULL,
                MDB_WRITE,
                lppDefaultMDB);
        }
        else hRes = MAPI_E_NOT_FOUND;
    }

    if (pRow)
        FreeProws(pRow);
    if (pStoresTbl)
        pStoresTbl->Release();
    return hRes;
}

HRESULT GetEntryIDFromMDB(LPMDB lpMDB, ULONG ulPropTag, ULONG* lpcbeid, LPENTRYID* lppeid)
{
    HRESULT hRes = S_OK;
    LPSPropValue lpEIDProp = NULL;

    if (!lpMDB || !lpcbeid || !lppeid)
        return MAPI_E_INVALID_PARAMETER;

    hRes = HrGetOneProp(lpMDB, ulPropTag, &lpEIDProp);

    if (SUCCEEDED(hRes) && lpEIDProp)
    {
        hRes = MAPIAllocateBuffer(lpEIDProp->Value.bin.cb, (LPVOID*)lppeid);

        if (SUCCEEDED(hRes))
        {
            *lpcbeid = lpEIDProp->Value.bin.cb;
            memcpy((void*)*lppeid, lpEIDProp->Value.bin.lpb, *lpcbeid);
        }
    }

    if (lpEIDProp)
        MAPIFreeBuffer(lpEIDProp);

    return hRes;
}

HRESULT OpenDefaultMAPIFolder(LPMDB lpMDB, ULONG ulPropTag, LPMAPIFOLDER* folderP)
{
    HRESULT hRes = S_OK;
    ULONG cb = NULL;
    LPENTRYID lpeid = NULL;
    ULONG ulObjType = NULL;


    if (!lpMDB)
        return MAPI_E_INVALID_PARAMETER;

    if (ulPropTag == PR_IPM_SUBTREE_ENTRYID)
    {
        hRes = GetEntryIDFromMDB(lpMDB, PR_IPM_SUBTREE_ENTRYID, &cb, &lpeid);
    }
    else
        return MAPI_E_INVALID_PARAMETER;

    LPUNKNOWN lpUnk = NULL;
    hRes = lpMDB->OpenEntry(
        cb,
        lpeid,
        NULL,
        MAPI_BEST_ACCESS,
        &ulObjType,
        &lpUnk);

    if (FAILED(hRes) && lpUnk)
    {
        lpUnk->Release();
        lpUnk = NULL;
    }
    else
        *folderP = (LPMAPIFOLDER)lpUnk;

    MAPIFreeBuffer(lpeid);

    return hRes;
}


int DeleteProfile()
{
    int rc = 0;
    LPPROFADMIN     lpProfAdmin = NULL;
    HRESULT         hRes = S_OK;

    // Get IProfAdmin interface
    hRes = MAPIAdminProfiles(0, &lpProfAdmin);
    if (FAILED(hRes))
    {
        logdebug(TEXT("failed to get profile admin interface\n"));
        return -2;
    }

    hRes = lpProfAdmin->DeleteProfile((LPTSTR)mProfileName, NULL);

    if (hRes != S_OK && hRes != MAPI_E_NOT_FOUND)
        rc = -3;

    return rc;
}

int CreateProfile(LPTSTR szPSTPath)
{
    int rc = 0;
    LPPROFADMIN   lpProfAdmin = NULL;
    LPSERVICEADMIN lpServiceAdmin = NULL;
    HRESULT         hRes = S_OK;

    if (szPSTPath == NULL)
        return -1;

    DeleteProfile();

    // Get IProfAdmin interface
    hRes = MAPIAdminProfiles(0, &lpProfAdmin);
    if (FAILED(hRes))
    {
        logdebug(TEXT("failed to get profile admin interface\n"));
        return -2;
    }

    hRes = lpProfAdmin->DeleteProfile((LPTSTR)mProfileName, NULL);
    if (hRes != S_OK && hRes != MAPI_E_NOT_FOUND)
    {
        logdebug(TEXT("failed to delete profile %s 0x%08X\n"), mProfileName, hRes);
        rc = -7;
    }

    if (rc == 0)
    {
        // Create profile
        hRes = lpProfAdmin->CreateProfile((LPTSTR)mProfileName,     // Name of new profile.
                            NULL,          // Password for profile.
                            NULL,          // Handle to parent window.
                            0);         // Flags.
        if (FAILED(hRes))
        {
            logdebug(TEXT("failed to create profile %s 0x%08X\n"), mProfileName, hRes);
            rc = -2;
        }
    }

    if (rc == 0)
    {
        // Get an IMsgServiceAdmin interface
        hRes = lpProfAdmin->AdminServices((LPTSTR)mProfileName,     // Profile that we want to modify.
                                           (LPWSTR)"",          // Password for that profile.
                                           NULL,          // Handle to parent window.
                                           0,             // Flags.
                                           &lpServiceAdmin); // Pointer to new IMsgServiceAdmin.
        if (FAILED(hRes))
        {
            logdebug(TEXT("failed to get IMsgServiceAdmin interface\n"));
            rc = -3;
        }
    }

    if (rc == 0)
    {
        LPSERVICEADMIN2 lpServiceAdmin2 = NULL;
        hRes = lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpServiceAdmin2);

        if (SUCCEEDED(hRes))
        {
            // Call CreateMsgServiceEx to create service
            MAPIUID uidService = { 0 };
            LPMAPIUID lpuidService = &uidService;
            hRes = lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)"MSUPST MS",
                                                       (LPTSTR)"MSUPST MS",
                                                       0,
                                                       0,
                                                       &uidService);
            if (SUCCEEDED(hRes))
            {
                logdebug(TEXT("Created profile %S!\n"), mProfileName);
                SPropValue PropVal[1];

                PropVal[0].ulPropTag = CHANGE_PROP_TYPE(PR_PST_PATH, PT_TSTRING);
                PropVal[0].Value.LPSZ = szPSTPath;

                hRes = lpServiceAdmin2->ConfigureMsgService( lpuidService,
                                                            NULL,
                                                            0,
                                                            1,
                                                            PropVal);
                if (FAILED(hRes))
                {
                    rc = -4;
                    logdebug(TEXT("failed to set pst path:  0x%08X\n"), hRes);
                }
            }
            else
            {
                rc = -5;
                logdebug(TEXT("failed to create unicode pst service:  0x%08X\n"), hRes);
            }
        }
        else
        {
            rc = -6;
            logdebug(TEXT("failed to acquire IID_IMsgServiceAdmin2 interface\n"));
        }

        if (lpServiceAdmin2)
            lpServiceAdmin->Release();
    }

    if (lpProfAdmin)
        lpProfAdmin->Release();

    return rc;
}

HRESULT OpenOutlookPST(const wchar_t *pstName, LPMDB* lppMDB, LPMAPISESSION* sppSession)
{
    MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
    HRESULT hRes = S_OK;

    ForceOutlookMAPI(true);

    MAPIInitialize(&MAPIINIT);

    CreateProfile((LPTSTR)pstName);

    //Logon session
    *sppSession = NULL;
    hRes = MAPILogonEx(NULL, (LPTSTR)mProfileName, NULL, 0, sppSession);

    if (SUCCEEDED(hRes))
    {
        logdebug(TEXT("MAPILogonEx successful\n"));
    }
    else
        logdebug(TEXT("Error 0x%08X\n"), hRes);

    // Open default MDB
    logdebug(TEXT("open default mdb\n"));

    *lppMDB = NULL;

    hRes = OpenDefaultMsgStore(*sppSession, lppMDB);
    return hRes;
}

void CloseOutlookPST(LPMDB lpMDB, LPMAPISESSION spSession)
{
    // release MDB
    if (lpMDB)
    {
        lpMDB->Release();
        lpMDB = NULL;
    }

    DeleteProfile();
    // log off session
    if (spSession)
    {
        spSession->Logoff(NULL, 0, NULL);
        spSession->Release();
        spSession = NULL;
    }

    MAPIUninitialize();
}

HRESULT OpenPSTRootFolder(LPMDB lpMDB, LPMAPIFOLDER* folderP)
{
    return OpenDefaultMAPIFolder(lpMDB, PR_IPM_SUBTREE_ENTRYID, folderP);
}

HRESULT OpenPSTFolder(LPMAPIFOLDER parentFolderP, wchar_t* folderName, LPMAPIFOLDER* folderP)
{
    HRESULT hRes = S_OK;

    hRes = parentFolderP->CreateFolder(
                            FOLDER_GENERIC,
                            folderName,
                            NULL,           // comment
                            NULL,           // interface
                            MAPI_UNICODE | OPEN_IF_EXISTS,
                            folderP);

    // commit changes
    if (SUCCEEDED(hRes))
        (*folderP)->SaveChanges(KEEP_OPEN_READWRITE);

    return hRes;
}

HRESULT ProcessFile_MSG(LPMDB lpMDB, LPMAPISESSION spSession, folder_s *folderP, wchar_t *fileName)
{
    HRESULT hRes = S_OK;
    LPMAPIFOLDER lpFolder = folderP->folderP;
    LPMESSAGE   lpMessage = NULL;
    LPMESSAGE   pIMsg = NULL;
    LPSTORAGE   pStorage = NULL;
    LPMALLOC    lpMalloc = NULL;
    wchar_t     msgFileName[MAX_PATH] = { 0 };

    if (folderP == NULL || folderP->folderP == NULL || folderP->folder_name == NULL || fileName == NULL)
        return MAPI_E_INVALID_PARAMETER;

    wsprintf(msgFileName, TEXT("%s\\%s"), folderP->folder_name, fileName);
    logdebug(TEXT("LoadMSGToMessage %s\n"), msgFileName);

    if (SUCCEEDED(hRes))
    {
        logdebug(TEXT("create message\n"));
        hRes = lpFolder->CreateMessage(NULL, NULL, &lpMessage);
    }

    // get memory allocation function
    lpMalloc = MAPIGetDefaultMalloc();

    if (SUCCEEDED(hRes) && lpMalloc)
    {
        // Open the compound file
        hRes = ::StgOpenStorage(
            msgFileName,
            NULL,
            STGM_TRANSACTED,
            NULL,
            0,
            &pStorage);

        if (SUCCEEDED(hRes) && pStorage)
        {
            // Open an IMessage interface on an IStorage object
            hRes = OpenIMsgOnIStg(NULL,
                MAPIAllocateBuffer,
                MAPIAllocateMore,
                MAPIFreeBuffer,
                lpMalloc,
                NULL,
                pStorage,
                NULL,
                0,
                0,
                &pIMsg);
        }
    }

    if (pStorage)
        pStorage->Release();

    if (SUCCEEDED(hRes))
    {
        logdebug(TEXT("copy message to pst file\n"));
        hRes = pIMsg->CopyTo(
            0,
            NULL,
            NULL,
            NULL,
            NULL,
            &IID_IMessage,
            lpMessage,
            0,
            NULL); // ignore problem array for now

        if (!FAILED(hRes))
        {
            logdebug(TEXT("save changes!\n"));
            hRes = lpMessage->SaveChanges(KEEP_OPEN_READWRITE);
        }
    }

    if (pIMsg)
        pIMsg->Release();

    // Release message
    if (lpMessage)
    {
        lpMessage->Release();
        lpMessage = NULL;
    }
    return hRes;
}

HRESULT ProcessFile_EML(LPMDB lpMDB, LPMAPISESSION spSession, folder_s *folderP, wchar_t *fileName)
{
    HRESULT hRes = S_OK;
    LPMAPIFOLDER lpFolder = folderP->folderP;
    LPMESSAGE   lpMessage = NULL;
    IConverterSession* lpConverter = NULL;
    wchar_t     msgFileName[MAX_PATH] = { 0 };

    if (folderP == NULL || folderP->folderP == NULL || folderP->folder_name == NULL || fileName == NULL)
        return MAPI_E_INVALID_PARAMETER;

    wsprintf(msgFileName, TEXT("%s\\%s"), folderP->folder_name, fileName);
    logdebug(TEXT("LoadEMLToMessage %s\n"), msgFileName);

    if (SUCCEEDED(hRes))
    {
        logdebug(TEXT("create message\n"));
        hRes = lpFolder->CreateMessage(NULL, NULL, &lpMessage);
    }

    if (SUCCEEDED(hRes) && lpMessage)
    {
        lpConverter = CreateConverter();

        if (lpConverter)
        {
            LPSTREAM lpEMLStm = NULL;

            hRes = OpenStreamOnFileW(MAPIAllocateBuffer,
                MAPIFreeBuffer,
                STGM_READ,
                msgFileName,
                NULL,
                &lpEMLStm);
            if (SUCCEEDED(hRes) && lpEMLStm)
            {
                hRes = lpConverter->MIMEToMAPI(lpEMLStm,
                        lpMessage,
                        NULL,       // Must be NULL
                        0x0022);    // CCSF_SMTP | CCSF_INCLUDE_BCC
                if (SUCCEEDED(hRes))
                {
                    hRes = lpMessage->SaveChanges(NULL);
                }
            }

            if (lpEMLStm) lpEMLStm->Release();
        }
        if (lpMessage)
            lpMessage->Release();
    }

    return hRes;
}

HRESULT ProcessFile(LPMDB lpMDB, LPMAPISESSION spSession, folder_s *folderP, wchar_t *fileName)
{
    HRESULT hRes = S_FALSE;

    if (folderP == NULL || folderP->folderP == NULL || folderP->folder_name == NULL || fileName == NULL)
        return MAPI_E_INVALID_PARAMETER;

    wchar_t *ext = wcsrchr(fileName, '.');
    if (ext != NULL)
    {
        if (wcscmp(ext, L".msg") == 0)
        {
            hRes = ProcessFile_MSG(lpMDB, spSession, folderP, fileName);
        }
        else if (wcscmp(ext, L".eml") == 0)
        {
            hRes = ProcessFile_EML(lpMDB, spSession, folderP, fileName);
        }
        else
            logdebug(TEXT("Not supported file %s"), fileName);
    }
    else
        logdebug(TEXT("Not supported file %s"), fileName);

    return hRes;
}

HRESULT ProcessDirectory(LPMDB lpMDB, LPMAPISESSION spSession, queue<folder_s*> *folderQ)
{
    DWORD dwError = 0;
    HANDLE hFind = INVALID_HANDLE_VALUE;
    WCHAR szDir[MAX_PATH];
    WIN32_FIND_DATA ffd;
    HRESULT hRes = S_OK;
    folder_s* cur_folder = NULL;

    if (folderQ == NULL || folderQ->empty())
        return dwError;

    cur_folder = folderQ->front();
    if (cur_folder == NULL || cur_folder->folder_name == NULL || cur_folder->folderP == NULL)
    {
        logdebug(TEXT("Internal error\n"));
        return -1;
    }
    folderQ->pop();

    logdebug(TEXT("Process directory <DIR>  %s\n"), cur_folder->folder_name);

    wcsncpy_s(szDir, cur_folder->folder_name, MAX_PATH);
    wcsncat_s(szDir, TEXT("\\*"), MAX_PATH);

    hFind = FindFirstFile(szDir, &ffd);

    if (INVALID_HANDLE_VALUE == hFind)
    {
        logdebug(TEXT("FindFirstFile error! dw = %d\n"), GetLastError());
        return dwError;
    }

    do
    {
        if (ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            if (wcsncmp(ffd.cFileName, L".", 2) != 0 && wcsncmp(ffd.cFileName, L"..", 4) != 0)
            {
                wchar_t* newDir = (wchar_t *)malloc(MAX_PATH * sizeof(wchar_t *));
                wsprintf(newDir, TEXT("%s\\%s"), cur_folder->folder_name, ffd.cFileName);

                LPMAPIFOLDER lpSubFolder = NULL;
                hRes = OpenPSTFolder(cur_folder->folderP, ffd.cFileName, &lpSubFolder);
                if (SUCCEEDED(hRes))
                {
                    folder_s* subFolder = new folder_s();
                    subFolder->folder_name = newDir;
                    subFolder->folderP = lpSubFolder;
                    folderQ->push(subFolder);
                }
            }
        }
        else if (!(ffd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) &&
                 !(ffd.dwFileAttributes & FILE_ATTRIBUTE_DEVICE) &&
                 !(ffd.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN))
        {
            hRes = ProcessFile(lpMDB, spSession, cur_folder, ffd.cFileName);
        }
    } while (FindNextFile(hFind, &ffd) != 0);

    dwError = GetLastError();
    if (dwError != ERROR_NO_MORE_FILES)
    {
        logdebug(TEXT("FindFirstFile error! dw = %d\n"), GetLastError());
    }

    FindClose(hFind);

    // free
    if (cur_folder->folder_name)
        delete cur_folder->folder_name;

    if (cur_folder->folderP)
    {
        cur_folder->folderP->Release();
        cur_folder->folderP = NULL;
    }
    delete cur_folder;

    return dwError;
}

HRESULT ProcessRequest(LPMDB lpMDB, LPMAPISESSION spSession, wchar_t* dirName)
{
    HRESULT hRes = S_OK;
    LPMAPIFOLDER lpRootFolder = NULL;

    queue<folder_s*> *folderQ = new queue<folder_s*>();
    folder_s* root_folder = new folder_s();
    root_folder->folder_name = _wcsdup(dirName);
    hRes = OpenPSTRootFolder(lpMDB, &lpRootFolder);
    root_folder->folderP = lpRootFolder;

    folderQ->push(root_folder);

    while (!folderQ->empty())
    {
        ProcessDirectory(lpMDB, spSession, folderQ);
    }

    delete folderQ;
    return hRes;
}

int usage()
{
    printf("PSTART.exe <msg_file_dir> <pst_file_name>\n");
    return 0;
}

int wmain(int argc, wchar_t *argv[])
{
    LPMDB lpMDB = NULL;
    LPMAPISESSION spSession = NULL;
    wchar_t *dir = NULL;
    wchar_t *pst_filename = NULL;

    if (argc < 3)
        return usage();

    dir = argv[1];
    pst_filename = argv[2];

    logdebug(L"begin\n");

    HRESULT hRes = OpenOutlookPST(pst_filename, &lpMDB, &spSession);

    if (SUCCEEDED(hRes))
    {
        ProcessRequest(lpMDB, spSession, dir);
    }

    DeleteConverter();
    CloseOutlookPST(lpMDB, spSession);

    printf("Press key to continue...\n");
    getchar();
    return 0;
}

