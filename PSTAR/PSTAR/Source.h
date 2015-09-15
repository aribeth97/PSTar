#ifndef __SOURCE_H__
#define __SOURCE_H__

// Exported from StubUtils.cpp
HMODULE GetMAPIHandle();
void UnLoadPrivateMAPI();
void ForceOutlookMAPI(bool fForce);
void ForceSystemMAPI(bool fForce);
void SetMAPIHandle(HMODULE hinstMAPI);
HMODULE GetPrivateMAPI();
bool GetComponentPath(LPCSTR szComponent, LPSTR szQualifier, LPSTR szDllPath, DWORD cchBufferSize, bool fInstall);


#endif
