#pragma once

// missing definitions for Conversion interface

#include "mimeole.h"

interface IConverterSession : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE SetAdrBook(LPADRBOOK pab);

    virtual HRESULT STDMETHODCALLTYPE SetEncoding(ENCODINGTYPE et);

    virtual HRESULT PlaceHolder1();

    virtual HRESULT STDMETHODCALLTYPE MIMEToMAPI(LPSTREAM pstm,
        LPMESSAGE pmsg,
        LPCSTR pszSrcSrv,
        ULONG ulFlags);

    virtual HRESULT STDMETHODCALLTYPE MAPIToMIMEStm(LPMESSAGE pmsg,
        LPSTREAM pstm,
        ULONG ulFlags);

    virtual HRESULT PlaceHolder2();
    virtual HRESULT PlaceHolder3();
    virtual HRESULT PlaceHolder4();

    virtual HRESULT STDMETHODCALLTYPE SetTextWrapping(bool  fWrapText,
        ULONG ulWrapWidth);

    virtual HRESULT STDMETHODCALLTYPE SetSaveFormat(MIMESAVETYPE mstSaveFormat);

    virtual HRESULT PlaceHolder5();

    virtual HRESULT STDMETHODCALLTYPE SetCharset(
        bool fApply,
        HCHARSET hcharset,
        CSETAPPLYTYPE csetapplytype);
};

STDAPI OpenStreamOnFileW(_In_ LPALLOCATEBUFFER lpAllocateBuffer,
    _In_ LPFREEBUFFER lpFreeBuffer,
    ULONG ulFlags,
    _In_z_ LPCWSTR lpszFileName,
    _In_opt_z_ LPCWSTR lpszPrefix,
    _Out_ LPSTREAM FAR * lppStream);

