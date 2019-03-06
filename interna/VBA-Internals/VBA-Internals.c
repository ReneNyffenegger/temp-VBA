//
//  cl /LD VBA-Internals.c    user32.lib OleAut32.lib /FeVBA-Internals.dll /link /def:VBA-Internals.def
//
//  gcc -c VBA-Internals.c
//  gcc -shared             *.o -lOleAut32 -o VBA-Internals.dll -Wl,--add-stdcall-alias
//  gcc -shared VBA-Internals.o -lOleAut32 -o VBA-Internals.dll -Wl,--add-stdcall-alias
//
//  TODO
//    Detouring: https://reverseengineering.stackexchange.com/a/13694

#include <windows.h>
#include <Shlobj.h>
#include <imagehlp.h>
#include <stdlib.h>

#include "WinAPI-typedefs.h"  

#include "VCOMInitializerStruct.h"

VCOMInitializerStruct *m_loader;


#define USE_SEARCH

#ifdef USE_SEARCH
#else
#define NOF_BREAKPOINTS 2
#endif


LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr);

#define msg(txt)
#define msg_2(txt) MessageBox(0, txt, 0, 0);

// #include "c:\about\libc\search\t\search.c"
#include <search.h>

#define TQ84_DEBUG_ENABLED
#define TQ84_DEBUG_TO_FILE
#define TQ84_DEBUG_FILENAME_WIDTH "20"
#define TQ84_DEBUG_FUNCNAME_WIDTH "50"
#include "c:\lib\tq84-c-debug\tq84_debug.c"

// #include "C:\temp\mhook\mhook-lib\mhook.h"
#include "mhook\mhook.h"

// --------------------------------------------------------------------

typedef struct {
  IDispatch_vTable * vtbl;
}
ERROBJ;

typedef HRESULT (STDMETHODCALLTYPE *fn_guts_m_36)(int *p1, int *p2, int *p3);

typedef struct {
    
  long *m_00;
  long *m_01;
  long *m_02;
  long ***m_03;
  long  m_04;
  long ***m_05;
  long *m_06;
  long *m_07;
  ERROBJ      *errObj;
  long *m_09;
  long *m_10;
  long *m_11;
  long  m_12;
  long *m_13;
  long *m_14;
  long *m_15;
  long  m_16;
  long *m_17;
  long *m_18;
  long *m_19;
  long *m_20;
  long *m_21;
  long *m_22;
  long *m_23;
  long *m_24;
  long *m_25;
  long *m_26;
  long *m_27;
  long *m_28;
  long *m_29;
  long *m_30;
  long *m_31;
  long *m_32;
  long *m_33;
  long *m_34;
  long *m_35;
//  long *m_36;
  fn_guts_m_36          fn_m_36;
//GUTS_OBJ_36 *guts_obj_36;
  long *m_37;
  long *m_38;
  long *m_39;
  long *m_40;
  long *m_41;
  long *m_42;
  long *m_43;
  long *m_44;
  long *m_45;
  long *m_46;
  long *m_47;
  long *m_48;
  long *m_49;

} VBA_guts;

typedef HRESULT (STDMETHODCALLTYPE *fn_hook_params_1)(ERROBJ *p_errObj); // , int p2, int p3, int p4, int p5, int p6, int p7);
typedef HRESULT (STDMETHODCALLTYPE *fn_hook)(ERROBJ *p_errObj, int *p2); // , int p2, int p3, int p4, int p5, int p6, int p7);
typedef HRESULT (STDMETHODCALLTYPE *fn_hook_params_3)(ERROBJ *p_errObj, int *p2, int *p3); // , int p2, int p3, int p4, int p5, int p6, int p7);
typedef HRESULT (STDMETHODCALLTYPE *fn_hook_params_4)(ERROBJ *p_errObj, int *p2, int *p3, int *p4);
typedef HRESULT (STDMETHODCALLTYPE *fn_hook_params_5)(ERROBJ *p_errObj, int *p2, int *p3, int *p4, int *p5);

typedef HRESULT   (STDMETHODCALLTYPE *fn_hook_07)(ERROBJ *p_errObj, int *p2);
typedef HRESULT   (STDMETHODCALLTYPE *fn_hook_11)(ERROBJ *p_errObj, int *p2);
typedef VBA_guts* (STDMETHODCALLTYPE *fn_hook_18)(ERROBJ *p_errObj);
typedef HRESULT   (STDMETHODCALLTYPE *fn_hook_19)(ERROBJ *p_errObj, int *p2);
//typedef HRESULT (*fn_hook)(int p1, int p2, int p3, int p4, int p5, int p6, int p7);


// typedef struct {
//   IDispatch_vTable * vtbl;
// }
// GUTS_OBJ_36;



VBA_guts* g_VBA_guts = NULL;

ERROBJ  *errObj = NULL;
//minimalVBAObject* errObj = NULL;
// IUnknown_vTable * errObj = NULL;


void printVal(const char* txt, long val) {
  TQ84_DEBUG_INDENT_T(txt);
  TQ84_DEBUG("val = %d", val);
  long addr = val;
  while (ReadProcessMemory(GetCurrentProcess(), (LPCVOID) addr, &val, 4, NULL)) {
    TQ84_DEBUG("* -> %d / %c%c%c%c%c%c", val, *((char*) addr), *(((char*) addr) + 2), *(((char*) addr) + 4), *(((char*) addr) + 6), *(((char*) addr) + 8), *(((char*) addr) + 10));
    addr = val;
  }

}


fn_guts_m_36 orig_guts_m_36 = NULL;
HRESULT STDMETHODCALLTYPE hook_guts_m_36(int *p1, int *p2, int *p3) { 
  TQ84_DEBUG_INDENT_T("hook_guts_m_36: p1 = %d, p2 = %d, p3 = %d", p1, p2, p3); 
//TQ84_DEBUG("*p2 = %d", *p2); 
  HRESULT ret = orig_guts_m_36(p1, p2, p3); 
  TQ84_DEBUG("ret = %d");
  return ret; 
}

void printGuts(VBA_guts* guts) {

  const char* t = "T.h.i.s. .i.s. .a";
  printVal("t", (long)t);

  TQ84_DEBUG("guts         = %d,                           ", guts          );
  TQ84_DEBUG("guts->m_00   = %d,                           ", guts->m_00    );
  TQ84_DEBUG("guts->m_01   = %d,                           ", guts->m_01    );
  TQ84_DEBUG("guts->m_02   = %d,                           ", guts->m_02    );
  printVal("guts->m_03", (long) guts->m_03);
  TQ84_DEBUG("guts->m_04   = %d,                           ", guts->m_04    );
  TQ84_DEBUG("guts->m_05   = %d, *guts->m_05 = %d , ** = %d, *** = %d", guts->m_05    , *guts->m_05, **guts->m_05, ***guts->m_05);
  TQ84_DEBUG("guts->m_06   = %d, *guts->m_06 = %d          ", guts->m_06    , *guts->m_06);
  TQ84_DEBUG("guts->m_07   = %d, *guts->m_07 = %d          ", guts->m_07    , *guts->m_07);
  TQ84_DEBUG("guts->errObj = %d, *guts->errObj = %d        ", guts->errObj  , *guts->errObj);
  TQ84_DEBUG("guts->m_09   = %d,                           ", guts->m_09    );
  TQ84_DEBUG("guts->m_10   = %d,                           ", guts->m_10    );
  TQ84_DEBUG("guts->m_11   = %d,                           ", guts->m_11    );
  TQ84_DEBUG("guts->m_12   = %d,                           ", guts->m_12    );
  TQ84_DEBUG("guts->m_13   = %d,                           ", guts->m_13    );
  TQ84_DEBUG("guts->m_14   = %d,                           ", guts->m_14    );
  TQ84_DEBUG("guts->m_15   = %d,                           ", guts->m_15    );
  TQ84_DEBUG("guts->m_16   = %d,                           ", guts->m_16    );
  TQ84_DEBUG("guts->m_17   = %d,                           ", guts->m_17    );
//  TQ84_DEBUG("guts->m_18   = %d, *guts->m_18 = %u          ", guts->m_18    , *guts->m_18);
  printVal("guts->m_18", (long) guts->m_18);
  TQ84_DEBUG("guts->m_19   = %d,                           ", guts->m_19    );
  TQ84_DEBUG("guts->m_20   = %d,                           ", guts->m_20    );
  TQ84_DEBUG("guts->m_21   = %d,                           ", guts->m_21    );
  TQ84_DEBUG("guts->m_22   = %d,                           ", guts->m_22    );
  TQ84_DEBUG("guts->m_23   = %d,                           ", guts->m_23    );
  TQ84_DEBUG("guts->m_24   = %d,                           ", guts->m_24    );
  TQ84_DEBUG("guts->m_25   = %d,                           ", guts->m_25    );
  TQ84_DEBUG("guts->m_26   = %d,                           ", guts->m_26    );
  TQ84_DEBUG("guts->m_27   = %d,                           ", guts->m_27    );
//TQ84_DEBUG("guts->m_28   = %d,                           ", guts->m_28    );
  printVal("guts->m_28", (long) guts->m_28);
  TQ84_DEBUG("guts->m_30   = %d,                           ", guts->m_30    );
  TQ84_DEBUG("guts->m_31   = %d,                           ", guts->m_31    );
  TQ84_DEBUG("guts->m_32   = %d,                           ", guts->m_32    );
  TQ84_DEBUG("guts->m_33   = %d,                           ", guts->m_33    );
  TQ84_DEBUG("guts->m_34   = %d,                           ", guts->m_34    );
  TQ84_DEBUG("guts->m_35   = %d,                           ", guts->m_35    );
  TQ84_DEBUG("guts->fn_m_36 = %d                         ", guts->fn_m_36    );

  if (guts->fn_m_36 && !orig_guts_m_36) {

    orig_guts_m_36 = guts->fn_m_36;

    if (! Mhook_SetHook((PVOID*) &orig_guts_m_36, (PVOID) hook_guts_m_36)) { \
         MessageBox(0, "Sorry, could not hook thing_18_0", 0, 0); \
    } 

//   TQ84_DEBUG("guts->guts_obj_36->vtbl->QueryInterface   = %d            ", guts->guts_obj_36->vtbl->QueryInterface);
//   TQ84_DEBUG("guts->guts_obj_36->vtbl->AddRef           = %d            ", guts->guts_obj_36->vtbl->AddRef);
//   TQ84_DEBUG("guts->guts_obj_36->vtbl->Release          = %d            ", guts->guts_obj_36->vtbl->Release);
//   TQ84_DEBUG("guts->guts_obj_36->vtbl->GetTypeInfoCount = %d            ", guts->guts_obj_36->vtbl->GetTypeInfoCount);
//   TQ84_DEBUG("guts->guts_obj_36->vtbl->GetTypeInfo      = %d            ", guts->guts_obj_36->vtbl->GetTypeInfo);
//   TQ84_DEBUG("guts->guts_obj_36->vtbl->GetIDsOfNames    = %d            ", guts->guts_obj_36->vtbl->GetIDsOfNames);
//   TQ84_DEBUG("guts->guts_obj_36->vtbl->Invoke           = %d            ", guts->guts_obj_36->vtbl->Invoke);
  }

  TQ84_DEBUG("guts->m_37   = %d,                           ", guts->m_37    );
  TQ84_DEBUG("guts->m_38   = %d,                           ", guts->m_38    );
  TQ84_DEBUG("guts->m_39   = %d,                           ", guts->m_39    );
  TQ84_DEBUG("guts->m_40   = %d,                           ", guts->m_40    );
  TQ84_DEBUG("guts->m_41   = %d,                           ", guts->m_41    );
  TQ84_DEBUG("guts->m_42   = %d,                           ", guts->m_42    );
  TQ84_DEBUG("guts->m_43   = %d,                           ", guts->m_43    );
  TQ84_DEBUG("guts->m_44   = %d,                           ", guts->m_44    );
  TQ84_DEBUG("guts->m_45   = %d,                           ", guts->m_45    );
  TQ84_DEBUG("guts->m_46   = %d,                           ", guts->m_46    );
  TQ84_DEBUG("guts->m_47   = %d,                           ", guts->m_47    );
  TQ84_DEBUG("guts->m_48   = %d,                           ", guts->m_48    );
  TQ84_DEBUG("guts->m_49   = %d,                           ", guts->m_49    );

}

typedef void* (CALLBACK *fn_rtcErrObj          )(); fn_rtcErrObj orig_rtcErrObj          ;


// Hook functions // {

BOOL WINAPI hook_ChooseColorA(LPCHOOSECOLOR lpcc) { // {
  TQ84_DEBUG("ChooseColor");
  return orig_ChooseColorA(lpcc);
} // }

BOOL WINAPI hook_GetFileVersionInfoA(LPCSTR lptstrFilename, DWORD dwHandle, DWORD dwLen, LPVOID lpData) { // {
  TQ84_DEBUG("GetFileVersionInfo, lptstrFilename = %s, dwHandle = %d, dwLen = %d", lptstrFilename, dwHandle, dwLen);
  return orig_GetFileVersionInfoA(lptstrFilename, dwHandle, dwLen, lpData);
} // }

DWORD WINAPI hook_GetFileVersionInfoSizeA(LPCSTR lptstrFilename, LPDWORD lpdwHandle) { // {
  TQ84_DEBUG_INDENT_T("GetFileVersionInfoSizeA, lptstrFilename = %s", lptstrFilename);
  DWORD ret = orig_GetFileVersionInfoSizeA(lptstrFilename, lpdwHandle);
  TQ84_DEBUG("ret = %d", ret);
  return ret;
} // }

HMODULE WINAPI hook_GetModuleHandleA(LPCSTR lpModuleName) { // {
  TQ84_DEBUG_INDENT_T("GetModuleHandleA, lpModuleName = %s", lpModuleName);

  HMODULE ret = orig_GetModuleHandleA(lpModuleName);

  TQ84_DEBUG("ret = %d", ret);

  return ret;
} // }

FARPROC WINAPI hook_GetProcAddress(HMODULE hModule, LPCSTR lpProcName) { // {
  TQ84_DEBUG_INDENT_T("GetProcAddress, hModule = %d, lpProcName = %s", hModule, lpProcName);

//if (! strcmp(lpProcName, "GetProcAddress")) {
//  TQ84_DEBUG("lpProcName = GetProcAddress, returning hook_GetProcAddress");
//  return hook_GetProcAddress;
//}

  FARPROC ret = orig_GetProcAddress(hModule, lpProcName);

  TQ84_DEBUG("ret = %d", ret);

  return ret;
} // }

HMODULE WINAPI hook_LoadLibraryA(LPCSTR lpLibFileName) { // {
  TQ84_DEBUG_INDENT_T("LoadLibraryA, lpLibFileName = %s", lpLibFileName);
  HMODULE ret = orig_LoadLibraryA(lpLibFileName);
  TQ84_DEBUG("ret = %d", ret);
  return ret;
  
} // }

BOOL WINAPI hook_MapAndLoad(PCSTR ImageName, PCSTR DllPath, PLOADED_IMAGE LoadedImage, WINBOOL DotDll, WINBOOL ReadOnly) { // {
  TQ84_DEBUG_INDENT_T("MapAndLoad, ImageName = %s, DllPath = %s", ImageName, DllPath);
  int ret = orig_MapAndLoad(ImageName, DllPath, LoadedImage, DotDll, ReadOnly);
  TQ84_DEBUG("ret = %d", ret);

  return ret;
} // }

LSTATUS WINAPI hook_RegCloseKey      ( HKEY hKey){ // {
  TQ84_DEBUG("RegCloseKey, hKey = %d", hKey);
  return orig_RegCloseKey(hKey);
} // }

VOID WINAPI hook_RtlUnwind(PVOID TargetFrame, PVOID TargetIp, PEXCEPTION_RECORD ExceptionRecord, PVOID ReturnValue) { // {
  TQ84_DEBUG_INDENT_T("RtlUnwind, TargetFrame = %d, TargetIp = %d", TargetFrame, TargetIp);

  orig_RtlUnwind(TargetIp, TargetIp, ExceptionRecord, ReturnValue);

} // }

void printSpecialhKey(HKEY hKey) { // {

  if (hKey == HKEY_CLASSES_ROOT) {
     TQ84_DEBUG("HKEY_CLASSES_ROOT");
  }
  else if (hKey == HKEY_CURRENT_CONFIG ) {
     TQ84_DEBUG("HKEY_CURRENT_CONFIG");
  }
  else if (hKey == HKEY_CURRENT_USER  ) {
     TQ84_DEBUG("HKEY_CURRENT_USER");
  }
  else if (hKey == HKEY_LOCAL_MACHINE   ) {
     TQ84_DEBUG("HKEY_LOCAL_MACHINE");
  }
  else if (hKey == HKEY_USERS) {
     TQ84_DEBUG("HKEY_USERS");
  }
} // }

LSTATUS WINAPI hook_RegOpenKeyExW    ( HKEY hKey, LPCWSTR lpSubKey, DWORD ulOptions, REGSAM samDesired, PHKEY phkResult)               { // {

  char x[500];
  wcstombs(x, lpSubKey, 499);

  TQ84_DEBUG_INDENT_T("RegOpenKeyExW, hKey = %d, lpSubKey = %s, ulOptions = %d, samDesired = %d", hKey, x, ulOptions, samDesired);

  printSpecialhKey(hKey);

//TQ84_DEBUG_INDENT_T("RegOpenKeyExW, lpSubKey = %S", lpSubKey);
  LSTATUS ret = orig_RegOpenKeyExW(hKey, lpSubKey, ulOptions, samDesired, phkResult);
  TQ84_DEBUG("hkeyResult = %d, ret = %d", *phkResult, ret);
  return ret;
} // }

LSTATUS WINAPI hook_RegOpenKeyExA    ( HKEY hKey, LPCSTR lpSubKey, DWORD  ulOptions, REGSAM samDesired, PHKEY  phkResult)                   { // {
  TQ84_DEBUG_INDENT_T("RegOpenKeyExA, hKey = %d, lpSubKey = %s, ulOptions = %d, samDesired = %d", hKey, lpSubKey, ulOptions, samDesired);

  printSpecialhKey(hKey);

//TQ84_DEBUG_INDENT_T("RegOpenKeyExA");
//TQ84_DEBUG("lpSubKey = %s", lpSubKey);
  LSTATUS ret = orig_RegOpenKeyExA(hKey, lpSubKey, ulOptions, samDesired, phkResult);

  TQ84_DEBUG("hkeyResult = %d, ret = %d", *phkResult, ret);
  return ret;
} // }

LSTATUS WINAPI hook_RegQueryValueExA ( HKEY hKey, LPCSTR  lpValueName, LPDWORD lpReserved, LPDWORD lpType, LPBYTE  lpData, LPDWORD lpcbData){ // {
  TQ84_DEBUG_INDENT_T("RegQueryValueExA, hKey = %d, lpValueName = %s", hKey, lpValueName);
  LSTATUS ret = orig_RegQueryValueExA(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData);
  TQ84_DEBUG("ret = %d", ret);
  return ret;
} // }

LSTATUS WINAPI hook_RegQueryValueExW ( HKEY hKey, LPCWSTR lpValueName, LPDWORD lpReserved, LPDWORD lpType, LPBYTE  lpData, LPDWORD lpcbData){ // {
  char x[500];
  wcstombs(x, lpValueName, 499);
  TQ84_DEBUG_INDENT_T("RegQueryValueExW hKey = %d, lpValueName = %s", hKey, x);
  return orig_RegQueryValueExW(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData);
} // }

LSTATUS WINAPI hook_RegSetValueExA   ( HKEY hKey, LPCSTR lpValueName, DWORD Reserved, DWORD dwType, const BYTE *lpData, DWORD cbData       ){ // {
  TQ84_DEBUG_INDENT_T("RegSetValueExA, hKey = %d, lpValueName = %s", hKey, lpValueName);
  return orig_RegSetValueExA(hKey, lpValueName, Reserved, dwType, lpData, cbData);
} // }

HINSTANCE WINAPI hook_ShellExecuteA( HWND hwnd, LPCSTR lpOperation, LPCSTR lpFile, LPCSTR lpParameters, LPCSTR lpDirectory, INT nShowCmd ) { // {
  TQ84_DEBUG_INDENT_T("ShellExecuteA, lpOperation = %s, lpFile = %s, lpParameters = %s, lpDirectory = %s, nShowCmd = %d", lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd);  
  return orig_ShellExecuteA(hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd);
} // }

HRESULT  hook_SHGetFolderPathW(HWND hwnd, int csidl, HANDLE hToken, DWORD dwFlags, LPWSTR pszPath) { // {

//char x[500];

//TQ84_DEBUG_INDENT_T("SHGetFolderPathW, hwnd = %d, csidl = %d, hToken = %d, dwFlags = %d", hwnd, csidl, hToken, dwFlags);

  HRESULT ret = orig_SHGetFolderPathW(hwnd, csidl, hToken, dwFlags, pszPath);

//int nof = wcstombs(x, pszPath, 499);

//TQ84_DEBUG("pszPath = %s, nof = %d", x, nof);

  return ret;
} // }

void hook_SysFreeString(BSTR bstrString) { // {

  TQ84_DEBUG_INDENT_T("SysFreeString, bstrString = %d", bstrString);
  

  if (bstrString) {

    char x[500];
    wcstombs(x, bstrString, 499);

    TQ84_DEBUG("bstrString = %s", x);
  }

  orig_SysFreeString(bstrString);
} // }

BOOL WINAPI hook_ShellExecuteExW(SHELLEXECUTEINFOW *pExecInfo) { // {
  TQ84_DEBUG_INDENT_T("ShellExecuteExW");
  return orig_ShellExecuteExW(pExecInfo);
} // }

BOOL WINAPI hook_UnMapAndLoad(PLOADED_IMAGE LoadedImage) { // {
  TQ84_DEBUG("UnMapAndLoad");
  return orig_UnMapAndLoad(LoadedImage);
} // }

BOOL WINAPI hook_VerQueryValueA(LPCVOID pBlock, LPCSTR lpSubBlock, LPVOID *lplpBuffer, PUINT puLen) { // {
  TQ84_DEBUG("VerQueryValue, lpSubBlock = %s, puLen = %d", lpSubBlock, puLen);
  return orig_VerQueryValueA(pBlock, lpSubBlock, lplpBuffer, puLen);
} // }

LPVOID WINAPI hook_VirtualAlloc(LPVOID lpAddress, SIZE_T dwSize, DWORD  flAllocationType, DWORD flProtect) { // {

  TQ84_DEBUG_INDENT_T("VirtualAlloc, lpAddress = %d, dwSize = %d, flAllocationType = %d, flProtect = %d", lpAddress, dwSize, flAllocationType, flProtect);
  LPVOID ret = orig_VirtualAlloc(lpAddress, dwSize, flAllocationType, flProtect);
  TQ84_DEBUG("ret = %d\n", ret);
  return ret;

} // }

int hook_WideCharToMultiByte(UINT CodePage, DWORD dwFlags, LPCWCH lpWideCharStr, int cchWideChar, LPSTR lpMultiByteStr, int cbMultiByte, LPCCH lpDefaultChar, LPBOOL lpUsedDefaultChar) { // {
  TQ84_DEBUG_INDENT_T("WideCharToMultiByte, codePage = %d, dwFlags = %d, cchWideChar = %d, cbMultiByte = %d", CodePage, dwFlags, cchWideChar, cbMultiByte);

//   char nullTerminated[500];
  int ret = orig_WideCharToMultiByte(CodePage, dwFlags, lpWideCharStr, cchWideChar, lpMultiByteStr, cbMultiByte, lpDefaultChar, lpUsedDefaultChar);

//  if (cchWideChar > 0) {
//    memcpy(nullTerminated, lpMultiByteStr,  cchWideChar);
//    nullTerminated[cchWideChar] = 0;
//  }
//  else {
//    strcpy(nullTerminated, lpMultiByteStr);
//  }


//TQ84_DEBUG("ret = %d, lpMultiByteStr = %s", ret, nullTerminated);
  TQ84_DEBUG("ret = %d", ret);

  return ret;
} // }

BOOL WINAPI hook_WriteProcessMemory( HANDLE hProcess, LPVOID lpBaseAddress, LPCVOID lpBuffer, SIZE_T nSize, SIZE_T *lpNumberOfBytesWritten) { // {
  TQ84_DEBUG_INDENT_T("WriteProcessMemory, hProcess = %d, lpBaseAddress = %d, lpBuffer = %d, nSize = %d", hProcess, lpBaseAddress, lpBuffer, nSize, lpNumberOfBytesWritten);

  BOOL ret = orig_WriteProcessMemory(hProcess, lpBaseAddress, lpBaseAddress, nSize, lpNumberOfBytesWritten);
  return ret;
} // }

funcPtr_IUnknown_AddRef         orig_errObj_AddRef;
funcPtr_IUnknown_QueryInterface orig_errObj_QueryInterface;
funcPtr_IDispatch_GetIDsOfNames orig_IDispatch_GetIDsOfNames;
funcPtr_IDispatch_Invoke        orig_IDispatch_Invoke;

HRESULT STDMETHODCALLTYPE hook_IUnknown_AddRef        (void *self) { // {
  TQ84_DEBUG_INDENT_T("hook_IUnknown_AddRef");
  ULONG ret = orig_errObj_AddRef(self);
  TQ84_DEBUG("ret = %d", ret);
  return ret;

} // }

HRESULT STDMETHODCALLTYPE hook_errObj_QueryInterface(void *self, REFIID riid, void **ppvObj) { // {
  TQ84_DEBUG_INDENT_T("hook_errObj_QueryInterface");
  TQ84_DEBUG("self = %d, errObj = %d", self, errObj);

  HRESULT ret = orig_errObj_QueryInterface(self, riid, ppvObj);
//HRESULT ret = errObj->vtbl->QueryInterface(self, riid, ppvObj);
//HRESULT ret = errObj->QueryInterface(self, riid, ppvObj);

  TQ84_DEBUG("ret = %d, *ppvObj = %d", ret, *ppvObj);

  return ret;
} // }

HRESULT STDMETHODCALLTYPE hook_IDispatch_GetIDsOfNames(void *self, REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) { // {
  TQ84_DEBUG_INDENT_T("hook_IDispatch_GetIDsOfNames");
  HRESULT ret = orig_IDispatch_GetIDsOfNames(self, riid, rgszNames, cNames, lcid, rgDispId);
  return ret;
} // }

HRESULT STDMETHODCALLTYPE hook_IDispatch_Invoke(void *self, DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pvarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) { // {
  TQ84_DEBUG_INDENT_T("hook_IDispatch_Invoke");
  HRESULT ret = orig_IDispatch_Invoke(self, dispidMember, riid, lcid, wFlags, pDispParams, pvarResult, pExcepInfo, puArgErr);
  return ret;
} // }




// HRESULT STDMETHODCALLTYPE hook_##n(ERROBJ *p_errObj, int *p2) { \

//  TQ84_DEBUG_INDENT_T("hook_" #n ": p1: %d, p2: %d, p3: %d, p4: %d, p5: %d, p6: %d, p7: %d", p1, p2, p3, p4, p5, p6, p7); \

//  HRESULT ret = orig_##n(p1, p2, p3, p4, p5, p6, p7); \


HRESULT (STDMETHODCALLTYPE *orig_thing_18_0)(int *p1, int *p2);
HRESULT STDMETHODCALLTYPE hook_thing_18_0(int *p1, int *p2) {

  TQ84_DEBUG_INDENT_T("hook_thing_18_0, p1 = %d, p2 = %d", p1, p2);
  HRESULT ret = orig_thing_18_0(p1, p2);
  return ret;

}


fn_hook_07 orig_07;
fn_hook_11 orig_11;
fn_hook_18 orig_18;
fn_hook_19 orig_19;

HRESULT STDMETHODCALLTYPE hook_07(ERROBJ* p_errObj, int *p2) {
  TQ84_DEBUG_INDENT_T("hook_07 (called after boom and _19): p_errObj = %d, p2 = %d, *p2 = %d", p_errObj, p2, *p2);
  TQ84_DEBUG("*p2 = %d", *p2);
  HRESULT ret = orig_07(p_errObj, p2);
  printGuts(g_VBA_guts);
  TQ84_DEBUG("ret = %d, *p2 = %d", ret, *p2);
  return ret;
}

HRESULT STDMETHODCALLTYPE hook_11(ERROBJ* p_errObj, int *p2) {
  TQ84_DEBUG_INDENT_T("hook_11: p_errObj = %d, p2 = %d, *p2 = %d", p_errObj, p2, *p2);
  TQ84_DEBUG("*p2 = %d", *p2);
  HRESULT ret = orig_07(p_errObj, p2);
  printGuts(g_VBA_guts);
  TQ84_DEBUG("ret = %d, *p2 = %d", ret, *p2);
  return ret;
}

VBA_guts* STDMETHODCALLTYPE hook_18(ERROBJ* p_errObj) {
  TQ84_DEBUG_INDENT_T("hook_18 (Used to set error proc name?): p_errObj = %d", p_errObj);
//  VBA_guts *ret = orig_18(p_errObj);
  g_VBA_guts = orig_18(p_errObj);




  if (g_VBA_guts->errObj != p_errObj) {
    MessageBox(0, "guts->errObj != p_errObj", 0, 0);
  }
  if (g_VBA_guts->errObj != errObj) {
    MessageBox(0, "guts->errObj != errObj", 0, 0);
  }

  printGuts(g_VBA_guts);


  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *g_VBA_guts);
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+1));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+2));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+3));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+4));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+5));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+6));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+7));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+8));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+9));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+10));
  TQ84_DEBUG("g_VBA_guts = %d, *g_VBA_guts = %d", g_VBA_guts, *( (int*) g_VBA_guts+11));

//  if (! Mhook_SetHook((PVOID*) & ret, (PVOID) hook_thing_18_0)) { \
//       MessageBox(0, "Sorry, could not hook thing_18_0", 0, 0); \
//  } 

  return g_VBA_guts;
}

HRESULT STDMETHODCALLTYPE hook_19(ERROBJ* p_errObj, int *p2) {
  TQ84_DEBUG_INDENT_T("hook_19 (called after boom and before _07): p_errObj = %d, p2 = %d, *p2 = %d", p_errObj, p2, *p2);
  TQ84_DEBUG("*p2 = %d", *p2);
  HRESULT ret = orig_19(p_errObj, p2);
  TQ84_DEBUG("ret = %d, *p2 = %d", ret, *p2);
  printGuts(g_VBA_guts);
  return ret;
}


#define def_hook_func(n)                   \
  fn_hook orig_##n;                        \
HRESULT STDMETHODCALLTYPE hook_##n(ERROBJ *p_errObj, int *p2) { \
  TQ84_DEBUG_INDENT_T("hook_" #n ": p_errObj = %d, p2 = %d", p_errObj, p2); \
  TQ84_DEBUG("*p2 = %d", *p2); \
  HRESULT ret = orig_##n(p_errObj, p2); \
  TQ84_DEBUG("ret = %d, *p2 = %d", ret, *p2); \
  return ret; \
}

#define def_hook_func_params_1(n)                   \
  fn_hook_params_1 orig_##n;                        \
HRESULT STDMETHODCALLTYPE hook_##n(ERROBJ *p_errObj) { \
  TQ84_DEBUG_INDENT_T("hook_" #n ": p_errObj = %d"); \
  HRESULT ret = orig_##n(p_errObj); \
  TQ84_DEBUG("ret = %d", ret); \
  return ret; \
}

#define def_hook_func_params_3(n)                   \
  fn_hook_params_3 orig_##n;                        \
HRESULT STDMETHODCALLTYPE hook_##n(ERROBJ *p_errObj, int *p2, int *p3) { \
  TQ84_DEBUG_INDENT_T("hook_" #n ": p_errObj = %d, p2 = %d, p3 = %d", p_errObj, p2, p3); \
  TQ84_DEBUG("*p2 = %d", *p2); \
  TQ84_DEBUG("*p3 = %d", *p3); \
  HRESULT ret = orig_##n(p_errObj, p2, p3); \
  TQ84_DEBUG("ret = %d, *p2 = %d, *p3 = %d", ret, *p2, *p3); \
  return ret; \
}

#define def_hook_func_params_4(n)                   \
  fn_hook_params_4 orig_##n;                        \
HRESULT STDMETHODCALLTYPE hook_##n(ERROBJ *p_errObj, int *p2, int *p3, int *p4) { \
  TQ84_DEBUG_INDENT_T("hook_" #n ": p_errObj = %d, p2 = %d, p3 = %d, p4 = %d", p_errObj, p2, p3, p4); \
  TQ84_DEBUG("*p2 = %d", *p2); \
  TQ84_DEBUG("*p3 = %d", *p3); \
/*  TQ84_DEBUG("*p4 = %d", *p4); */ \
  HRESULT ret = orig_##n(p_errObj, p2, p3, p4); \
/*  TQ84_DEBUG("ret = %d, *p2 = %d, *p3 = %d, *p4 = %d", ret, *p2, *p3, *p4); */ \
  return ret; \
}

#define def_hook_func_params_5(n)                   \
  fn_hook_params_5 orig_##n;                        \
HRESULT STDMETHODCALLTYPE hook_##n(ERROBJ *p_errObj, int *p2, int *p3, int *p4, int *p5) { \
  TQ84_DEBUG_INDENT_T("hook_" #n ": p_errObj = %d, p2 = %d, p3 = %d, p4 = %d, p5 = %d", p_errObj, p2, p3, p4, p5); \
  TQ84_DEBUG("*p2 = %d", *p2); \
  TQ84_DEBUG("*p3 = %d", *p3); \
/*  TQ84_DEBUG("*p4 = %d", *p4); */ \
  HRESULT ret = orig_##n(p_errObj, p2, p3, p4, p5); \
/*  TQ84_DEBUG("ret = %d, *p2 = %d, *p3 = %d, *p4 = %d", ret, *p2, *p3, *p4); */ \
  return ret; \
}


#define hook_func(n)                         \
    orig_##n = *(((int*) errObj->vtbl) + n); \
     if (! Mhook_SetHook((PVOID*) & orig_##n, (PVOID) hook_##n)) { \
       MessageBox(0, "Sorry, could not hook " #n, 0, 0); \
     } 


//def_hook_func(4);
// def_hook_func( 7);
def_hook_func( 8);
def_hook_func( 9);
def_hook_func(10);
// def_hook_func(11);
def_hook_func(12);
def_hook_func(13);
def_hook_func(14);
def_hook_func(15);
def_hook_func(16);
def_hook_func(17);
//def_hook_func_params_1(18);
// def_hook_func(19);
def_hook_func(20);
def_hook_func(21);
def_hook_func(22);
def_hook_func(26);
def_hook_func(27);


def_hook_func(31);
def_hook_func(32);
def_hook_func(33);
def_hook_func(34);
def_hook_func(35);
def_hook_func(36);
def_hook_func(37);
def_hook_func(38);
def_hook_func(39);
def_hook_func(40);

def_hook_func(42);

def_hook_func(48);
def_hook_func(49);


ERROBJ* CALLBACK hook_rtcErrObj() { // {
//minimalVBAObject* CALLBACK hook_rtcErrObj() { // 
// IUnknown_vTable * CALLBACK hook_rtcErrObj() { // 
  TQ84_DEBUG_INDENT_T("hook_rtcErrObj, hook_rtcErrObj = %d", hook_rtcErrObj);

  TQ84_DEBUG("orig_rtcErrObj = %d", orig_rtcErrObj);

  ERROBJ *ret = orig_rtcErrObj();
//  minimalVBAObject* ret = orig_rtcErrObj();
//IUnknown_vTable* ret = orig_rtcErrObj();
  if (errObj) {
    if (errObj != ret) {
       MessageBox(0, "errObj != ret", "Error", 0);
    }
    else {
      TQ84_DEBUG("errObj == ret");
    }

  }
  else {
     TQ84_DEBUG("setting errObj = ret");

     errObj = ret;

     TQ84_DEBUG("errObj->QueryInterface   = %d", errObj->vtbl->QueryInterface  ); // {
     TQ84_DEBUG("errObj->AddRef           = %d", errObj->vtbl->AddRef          );
     TQ84_DEBUG("errObj->Release          = %d", errObj->vtbl->Release         );
     TQ84_DEBUG("errObj->GetTypeInfoCount = %d", errObj->vtbl->GetTypeInfoCount);
     TQ84_DEBUG("errObj->GetTypeInfo      = %d", errObj->vtbl->GetTypeInfo     );
     TQ84_DEBUG("errObj->GetIDsOfNames    = %d", errObj->vtbl->GetIDsOfNames   );
     TQ84_DEBUG("errObj->Invoke           = %d", errObj->vtbl->Invoke          );

//   hook_func( 9);
//   hook_func(10);
//   TQ84_DEBUG(orig_j->4                 = %d", 
//     TQ84_DEBUG("errObj->4                = %d", *(((int*) errObj->vtbl) + 4)     );
//     TQ84_DEBUG("errObj->5                = %d", *(((int*) errObj->vtbl) + 5)     );
       TQ84_DEBUG("errObj->6                = %d", *(((int*) errObj->vtbl) + 6)     );
//     TQ84_DEBUG("errObj->7                = %d", *(((int*) errObj->vtbl) + 7)     );
//     TQ84_DEBUG("errObj->8                = %d", *(((int*) errObj->vtbl) + 8)     );
//     TQ84_DEBUG("errObj->9                = %d", *(((int*) errObj->vtbl) + 9)     );
//     TQ84_DEBUG("errObj->10               = %d", *(((int*) errObj->vtbl) + 10)     );
//     TQ84_DEBUG("errObj->11               = %d", *(((int*) errObj->vtbl) + 11)     );
       TQ84_DEBUG("errObj->12               = %d", *(((int*) errObj->vtbl) + 12)     );
       TQ84_DEBUG("errObj->13               = %d", *(((int*) errObj->vtbl) + 13)     );
       TQ84_DEBUG("errObj->14               = %d", *(((int*) errObj->vtbl) + 14)     );
       TQ84_DEBUG("errObj->15               = %d", *(((int*) errObj->vtbl) + 15)     );
       TQ84_DEBUG("errObj->16               = %d", *(((int*) errObj->vtbl) + 16)     );
       TQ84_DEBUG("errObj->17               = %d", *(((int*) errObj->vtbl) + 17)     );
       TQ84_DEBUG("errObj->18               = %d", *(((int*) errObj->vtbl) + 18)     );
       TQ84_DEBUG("errObj->19               = %d", *(((int*) errObj->vtbl) + 19)     );
       TQ84_DEBUG("errObj->20               = %d", *(((int*) errObj->vtbl) + 20)     );
       TQ84_DEBUG("errObj->21               = %d", *(((int*) errObj->vtbl) + 21)     );
       TQ84_DEBUG("errObj->22               = %d", *(((int*) errObj->vtbl) + 22)     );
       TQ84_DEBUG("errObj->23               = %d", *(((int*) errObj->vtbl) + 23)     );
       TQ84_DEBUG("errObj->24               = %d", *(((int*) errObj->vtbl) + 24)     );
       TQ84_DEBUG("errObj->25               = %d", *(((int*) errObj->vtbl) + 25)     );
       TQ84_DEBUG("errObj->26               = %d", *(((int*) errObj->vtbl) + 26)     );
       TQ84_DEBUG("errObj->27               = %d", *(((int*) errObj->vtbl) + 27)     );
       TQ84_DEBUG("errObj->28               = %d", *(((int*) errObj->vtbl) + 28)     );
       TQ84_DEBUG("errObj->29               = %d", *(((int*) errObj->vtbl) + 29)     );
       TQ84_DEBUG("errObj->30               = %d", *(((int*) errObj->vtbl) + 30)     );
       TQ84_DEBUG("errObj->31               = %d", *(((int*) errObj->vtbl) + 31)     );
       printVal("errObj->vtbl + 31", (long) (((int*) errObj->vtbl) + 31));
       TQ84_DEBUG("errObj->32               = %d", *(((int*) errObj->vtbl) + 32)     );
       TQ84_DEBUG("errObj->33               = %d", *(((int*) errObj->vtbl) + 33)     );
       TQ84_DEBUG("errObj->34               = %d", *(((int*) errObj->vtbl) + 34)     );
       TQ84_DEBUG("errObj->35               = %d", *(((int*) errObj->vtbl) + 35)     );
       TQ84_DEBUG("errObj->36               = %d", *(((int*) errObj->vtbl) + 36)     );
       TQ84_DEBUG("errObj->37               = %d", *(((int*) errObj->vtbl) + 37)     );
       TQ84_DEBUG("errObj->38               = %d", *(((int*) errObj->vtbl) + 38)     );
       TQ84_DEBUG("errObj->39               = %d", *(((int*) errObj->vtbl) + 39)     );
       TQ84_DEBUG("errObj->40               = %d", *(((int*) errObj->vtbl) + 40)     );
       TQ84_DEBUG("errObj->41               = %d", *(((int*) errObj->vtbl) + 41)     );
       TQ84_DEBUG("errObj->42               = %d", *(((int*) errObj->vtbl) + 42)     );
       TQ84_DEBUG("errObj->43               = %d", *(((int*) errObj->vtbl) + 43)     );
       TQ84_DEBUG("errObj->44               = %d", *(((int*) errObj->vtbl) + 44)     );
       TQ84_DEBUG("errObj->45               = %d", *(((int*) errObj->vtbl) + 45)     );
       TQ84_DEBUG("errObj->46               = %d", *(((int*) errObj->vtbl) + 46)     );
       TQ84_DEBUG("errObj->47               = %d", *(((int*) errObj->vtbl) + 47)     );
       TQ84_DEBUG("errObj->48               = %d", *(((int*) errObj->vtbl) + 48)     );
       TQ84_DEBUG("errObj->49               = %d", *(((int*) errObj->vtbl) + 49)     ); 
       
       
       
       
       
       
       
       // }

     hook_func(07);
     hook_func( 8);
     hook_func( 9);
     hook_func(10);
     hook_func(11);
     hook_func(12);
     hook_func(13);
     hook_func(14);
     hook_func(15);
     hook_func(16);
     hook_func(17);
     hook_func(18);
     hook_func(19);
     hook_func(20);
     hook_func(21);
     hook_func(22);

     hook_func(26);
     hook_func(27);

//   hook_func(31);  // Why can't this one be hooked?
     hook_func(32);
     hook_func(33);
     hook_func(34);
     hook_func(35);
     hook_func(36);
     hook_func(37);
     hook_func(38);
     hook_func(39);
     hook_func(40);

//   hook_func(42);  // Why can't this one be hooked?

     hook_func(48);
//   hook_func(49);  // Why can't this one be hooked?

//   orig_errObj_QueryInterface = errObj->QueryInterface;
     orig_errObj_AddRef         = errObj->vtbl->AddRef;
     orig_errObj_QueryInterface = errObj->vtbl->QueryInterface;
     orig_IDispatch_GetIDsOfNames = errObj->vtbl->GetIDsOfNames;
     orig_IDispatch_Invoke        = errObj->vtbl->Invoke;


     if (! Mhook_SetHook((PVOID*) & orig_errObj_AddRef, (PVOID) hook_IUnknown_AddRef)) {
       MessageBox(0, "Sorry, could not hook orig_errObj_AddRef", 0, 0);
     }

     if (! Mhook_SetHook((PVOID*) & orig_errObj_QueryInterface, (PVOID) hook_errObj_QueryInterface)) {
       MessageBox(0, "Sorry, could not hook orig_errObj_QueryInterface", 0, 0);
     }

     if (! Mhook_SetHook((PVOID*) & orig_IDispatch_GetIDsOfNames, (PVOID) hook_IDispatch_GetIDsOfNames)) {
       MessageBox(0, "Sorry, could not hook orig_IDispatch_GetIDsOfNames", 0, 0);
     }

     if (! Mhook_SetHook((PVOID*) & orig_IDispatch_Invoke, (PVOID) hook_IDispatch_Invoke)) {
       MessageBox(0, "Sorry, could not hook orig_IDispatch_Invoke", 0, 0);
     }

     TQ84_DEBUG("errObj->QueryInterface = %d", errObj->vtbl->QueryInterface);
     TQ84_DEBUG("errObj->AddRef         = %d", errObj->vtbl->AddRef        );
     TQ84_DEBUG("errObj->Release        = %d", errObj->vtbl->Release       );
  }

  TQ84_DEBUG("returning ret (errObj)= %d", ret);

  return ret;
} // }

 // }

// --------------------------------------------------------------------

typedef unsigned char   instruction;
typedef LPVOID address;


HANDLE curProcess;

static int initialized = 0;

instruction replaceInstruction(address addr, instruction instr) { // {
    TQ84_DEBUG_INDENT_T("replaceInstruction, addr=%d, byte = %u", addr, instr);

    char txt[400];
    SIZE_T nofBytes;

    instruction oldInstr;
    DWORD       oldProtection;

//  wsprintf(txt, "replace instruction at %d with %x", addr, instr);
//  msg_2(txt);


//  TQ84_DEBUG("VirtualProtect");
    VirtualProtect(addr, 1, PAGE_EXECUTE_READWRITE, &oldProtection);

//  oldInstr =  *((instruction*) addr);
    if (! ReadProcessMemory(curProcess, addr, &oldInstr, 1, &nofBytes)) {
      MessageBox(0, "Could not ReadProcessMemory", 0, 0);
    }
//  TQ84_DEBUG("nofBytes read: %d, byte = %u, now going to change to %u", nofBytes, oldInstr, instr);

// *((instruction*) addr) = instr;
    if (! WriteProcessMemory(curProcess, addr, &instr, 1, &nofBytes)) {
      MessageBox(0, "Could not ReadProcessMemory", 0, 0);
    }
//  TQ84_DEBUG("nofBytes written: %d", nofBytes);

//  VirtualProtect(addr, 1, oldProtection, &oldProtection);

//  if (! FlushInstructionCache(curProcess, addr, 10)) {
//    MessageBox(0, "Could not FlushInstructionCache", 0, 0);
//  }

//  TQ84_DEBUG("Returning old byte: %u", oldInstr); 

//  wsprintf(txt, "replace instruction %x with %x at %d", oldInstr, instr, addr);
//  msg_2("replaced");

    return oldInstr;
} // }

typedef struct { // {
    address       addr;
    char         *name;
    instruction   origInstruction;
    char         *format;
}
breakpoint; // }

int compareBreakpoints(const void *const f1, const void *const f2) { // {

// int compareBreakpoints(const breakpoint *const f1, const breakpoint *const f2) 

  const breakpoint *bp1 = (breakpoint*) f1;
  const breakpoint *bp2 = (breakpoint*) f2;
  char txt[200];
  wsprintf(txt, "Comparing f1 (%d, addr: %d, %s) with f2 (%d, addr: %d, %s)", bp1, (long) bp1->addr, bp1->name, bp2, (long) bp2->addr, bp2->name);
  msg_2(txt);
  return ((long)(bp1->addr)) - ((long)(bp2->addr));
 

// was: breakpoint *breakpointTreeRoot = 0 ...
              void *breakpointTreeRoot = 0;

//  TQ84_DEBUG_INDENT_T("compareBreakpoints, addr 1 = %d, addr 2 = %d", ((breakpoint*)f1)->addr, ((breakpoint*)f2)->addr);

//  char txt[400];
//wsprintf(txt, "Comparing f1 (%d, addr: %d, %s) with f2 (%d, addr: %d, %s)", f1, ((breakpoint*)f1)->addr, ((breakpoint*)f1)->name, f2, ((breakpoint*)f2)->addr, ((breakpoint*)f2)->name);
// MM  wsprintf(txt, "Comparing f1 (%d, addr: %d    ) with f2 (%d, addr: %d    )", f1, ((breakpoint*)f1)->addr,                          f2, ((breakpoint*)f2)->addr                         );
// MM  msg(txt);
//  return ((long)( ((breakpoint*)f1)->addr)) - ((long)(((breakpoint*)f2)->addr));
} // }

// breakpoint *breakpointTreeRoot = 0;
   void       *breakpointTreeRoot = 0;

__declspec(dllexport) void __stdcall addBreakpoint(address addr, char* name, char* format) { // {

    TQ84_DEBUG_INDENT_T("addBreakpoint, addr = %d, name = %s", addr, name);
    
    char txt[400]; 
    breakpoint *f;



// MM    wsprintf(txt, "addBreakpoint %s at %d", name, addr);
// MM    msg(txt);


    f = malloc(sizeof(breakpoint));
// MM    wsprintf(txt, "malloc, f= %d, sizeof(breakpoint) = %d",f, sizeof(breakpoint));
// MM    msg(txt);

    f->addr            = addr;
    f->name            = strdup(name  );
    f->format          = strdup(format);
//  msg("calling replaceInstruction");

// MM    wsprintf(txt, "addBreakpoint, after f->name = strdup(name): %s, addr(name) = %d, f = %d", f->name, f->name, f);
// MM    msg(txt);

//  TQ84_DEBUG("going to call replaceInstruction");
    f->origInstruction = replaceInstruction(addr, 0xcc); // 0xcc = INT3
//  msg("returned from replaceInstruction");

// MM    wsprintf(txt, "addBreakpoint f->name = %s, addr f->name= %d, addr = %d, f = %d, f->origInstruction: %x", f->name, f->name, f->addr, f, f->origInstruction);
// MM    msg_2(txt);
//  TQ84_DEBUG("going to call tsearch");
    tsearch(f, &breakpointTreeRoot, compareBreakpoints);
//  tsearch( (void*) addr, &breakpointTreeRoot, compareBreakpoints);
//  msg("leaving addBreakpoint");
} // }

__declspec(dllexport) void __stdcall addDllFunctionBreakpoint(char* module, char* funcName, char* format) { // {

  return;   /// HERE HERE HERE ////

     address addr;
     HMODULE mod;
     char    breakpointName[200];
     char    txt           [400];

     TQ84_DEBUG_INDENT_T("addDllFunctionBreakpoint, module = %s, funcName = s, format = %s", module, funcName, format);

// MM     wsprintf(txt, "addDllFunctionBreakpoint %s / %s", module, funcName);
// MM     msg(txt);

     mod = GetModuleHandle(module);
     if (! mod) {
       TQ84_DEBUG("! mod");
       MessageBox(0, module, "GetModuleHandle", 0);
       return;
     }

     addr = GetProcAddress(mod, funcName);
     if (!addr) {
       TQ84_DEBUG("! addr");
          MessageBox(0, funcName, "GetProcAddress", 0);
          return;
     }

     wsprintf(breakpointName, "%s.%s", module, funcName);
// MM     msg(breakpointName);

     addBreakpoint(addr, breakpointName, format);
//   msg("leaving addDllFunctionBreakpoint");

} // }


#ifdef USE_SEARCH // {
address      addrLastBreakpointHit;
 // }
#else // {
address      func_addrs[NOF_BREAKPOINTS];
instruction  old_instr [NOF_BREAKPOINTS];
int   nrLastFuncBreakpointHit;
#endif // }



// typedef void (WINAPI   *addrCallBack_t)(BSTR);
   typedef void (CALLBACK *addrCallBack_t)(BSTR);
addrCallBack_t addrCallBack;

void WinAPI_hook_GetModuleHandleA(LPCSTR lpModuleName) {

}

void callCallback(char* txt) { // {
     TQ84_DEBUG_INDENT_T("callCallback, txt = %s", txt);
     int wslen;
     BSTR bstr;
 
     wslen = MultiByteToWideChar(CP_ACP, 0, txt, strlen(txt),    0,     0); bstr  = SysAllocStringLen(0, wslen);
             MultiByteToWideChar(CP_ACP, 0, txt, strlen(txt), bstr, wslen);
 
    (*addrCallBack)(bstr);
 
     SysFreeString(bstr);
} // }

LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr) { // {

    int i;
    char txt [400];
    char txt_[400];

    if      (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT ) { // {

       TQ84_DEBUG("EXCEPTION_BREAKPOINT, addr = %d", exPtr->ContextRecord->Eip);

#ifdef USE_SEARCH // {



       breakpoint **breakpointFound;
       breakpoint  *breakpointHit;
       address addressHitByBreakpoint = (address) exPtr->ContextRecord->Eip;

//     TQ84_DEBUG("addressHitByBreakpoint = %d", exPtr->ContextRecord->Eip);
//     addressHitByBreakpoint = (address) (( (DWORD) addressHitByBreakpoint) - 1);

//MM     wsprintf(txt, "addressHitByBreakpoint = %d", addressHitByBreakpoint);
//MM     msg(txt);


       breakpoint findMe;
       findMe.addr = addressHitByBreakpoint;

       breakpointFound = tfind(&findMe, &breakpointTreeRoot, compareBreakpoints);


       if (!breakpointFound) { // {
         TQ84_DEBUG("Did not tfind breakpointHit...");
         MessageBox(0, "Did not tfind breakpointHit...", 0, 0);
       } // }
       else { // {

          breakpointHit = *breakpointFound;

//        wsprintf(txt, "&findMe = %d, breakpointHit = %d", &findMe, breakpointHit);
//        msg(txt);
//        wsprintf(txt, "breakpointHit = %d", breakpointHit);

//        wsprintf(txt ,"Hit Breakpoint, breakpointHit = %d, addr = %d (name = %s, addr(name) = %d), orig instr1stByte: %d", breakpointHit, breakpointHit->addr, breakpointHit->name, breakpointHit->name, breakpointHit->origInstruction);
          wsprintf(txt_,">>%s<<: >>%s<<", breakpointHit->name, breakpointHit->format);
//        TQ84_DEBUG("txt_: %s", txt_);

          unsigned long* esp = (unsigned long*) exPtr->ContextRecord->Esp;
          wsprintf(txt , txt_, *(esp+1), *(esp+2), *(esp+3), *(esp+4),*(esp+5), *(esp+6), *(esp+7), *(esp+8));
          TQ84_DEBUG("txt = %s", txt);

//        wsprintf(txt, "Hit Breakpoint %s", breakpointHit->name);
//        msg_2(txt);
          callCallback(txt);
   
          addrLastBreakpointHit = addressHitByBreakpoint;
   
          replaceInstruction(addressHitByBreakpoint, breakpointHit->origInstruction);

       // Single Step:
          exPtr->ContextRecord->EFlags |= 0x100;

       // TQ84_DEBUG("Going to call SetThreadContext");

          TQ84_DEBUG_DEDENT();
          SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
       } // }
 // }
#else // {


       for (i=0; i<NOF_BREAKPOINTS; i++) {
   
         // The EIP register contains the address of the break point
         // instruction that triggered the exception.
         // This allows to determine the number of the breakpoint:
           if (exPtr->ContextRecord->Eip == (int) func_addrs[i]) {
 
               wsprintf(txt, "Breakpoint %d was hit", i);
               callCallback(txt);
 
            //
            // Store the number of the breakpoint so that we can
            // set the breakpoint again after single stepping
            // the »current« instruction:
            //
               nrLastFuncBreakpointHit = i;
   
            //
            // In order to proceed with the execution of the program, we
            // restore the old value of the byte that was replaced by
            // the int-3 instruction:
            //
               replaceInstruction(func_addrs[i], old_instr[i]);
   
            //
            // Set TF bit in order to single step the next
            // instruction. This allows to set the break point
            // again after the single step instruction.
            //
   
               exPtr->ContextRecord->EFlags |= 0x100;
   
            //
            // Resume execution:
            //
               SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
           }
        }
#endif // }

       TQ84_DEBUG  ("EXCEPTION_BREAKPOINT: no matching address found");
       callCallback("EXCEPTION_BREAKPOINT: no matching address found");
       return EXCEPTION_CONTINUE_SEARCH;

    } // }
    else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP) { // {
       TQ84_DEBUG("EXCEPTION_SINGLE_STEP: Address: %d", exPtr->ContextRecord->Eip);


#ifdef USE_SEARCH  // {

//     TQ84_DEBUG("EXCEPTION_SINGLE_STEP: going to replaceInstruction at lastAddrBreakpoint hit: %d", addrLastBreakpointHit);
       replaceInstruction(addrLastBreakpointHit, 0xcc); // INT3              

 // }
#else // {

// QQ      replaceInstruction(func_addrs[nrLastFuncBreakpointHit], 0xcc);


     // callCallback("EXCEPTION_SINGLE_STEP");
 
     //
     // The processor is one instruction behind the last breakpoint that was
     // hit. We can now set the breakpoint again.
 
 
 //     TODO: Uncomment me:
 //     *func_addrs[nrLastFuncBreakpointHit] = 0xcc;
 // QQ  replaceInstruction(func_addrs[nrLastFuncBreakpointHit], 0xcc);
 
 //     msg_2("Single Step: Replacing instruction in single step");
 //     msg_2("Single Step: instruction was replaced, going to call SetThreadContext");
 
     //
     // Apparently, the TF flag is reset after the single execution, it
     // needs not be reset.
     //
     // Thus, we can resume execution again:
     //
 
     // callCallback("EXCEPTION_SINGLE_STEP -> SetThreadContext");

#endif // }

//     TQ84_DEBUG("Removing 0x100");
//     exPtr->ContextRecord->EFlags &= ~(0x100);
//     TQ84_DEBUG("Going to call SetThreadContext");
//     SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
       return EXCEPTION_CONTINUE_EXECUTION;



    } // }
    else { // {

       TQ84_DEBUG("EXCEPTION else with code %x at address %d", exPtr->ExceptionRecord->ExceptionCode, exPtr->ContextRecord->Eip);
    //
    // Should never be reached.
    //
    
           if      (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ACCESS_VIOLATION           ) { wsprintf(txt, "Unexpected exception EXCEPTION_ACCESS_VIOLATION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_DATATYPE_MISALIGNMENT      ) { wsprintf(txt, "Unexpected exception EXCEPTION_DATATYPE_MISALIGNMENT    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_BREAKPOINT                 ) { wsprintf(txt, "Unexpected exception EXCEPTION_BREAKPOINT               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_SINGLE_STEP                ) { wsprintf(txt, "Unexpected exception EXCEPTION_SINGLE_STEP              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ARRAY_BOUNDS_EXCEEDED      ) { wsprintf(txt, "Unexpected exception EXCEPTION_ARRAY_BOUNDS_EXCEEDED    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_DENORMAL_OPERAND       ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_DENORMAL_OPERAND     , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_DIVIDE_BY_ZERO         ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_DIVIDE_BY_ZERO       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_INEXACT_RESULT         ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_INEXACT_RESULT       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_INVALID_OPERATION      ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_INVALID_OPERATION    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_OVERFLOW               ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_OVERFLOW             , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_STACK_CHECK            ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_STACK_CHECK          , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_FLT_UNDERFLOW              ) { wsprintf(txt, "Unexpected exception EXCEPTION_FLT_UNDERFLOW            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INT_DIVIDE_BY_ZERO         ) { wsprintf(txt, "Unexpected exception EXCEPTION_INT_DIVIDE_BY_ZERO       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INT_OVERFLOW               ) { wsprintf(txt, "Unexpected exception EXCEPTION_INT_OVERFLOW             , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_PRIV_INSTRUCTION           ) { wsprintf(txt, "Unexpected exception EXCEPTION_PRIV_INSTRUCTION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_IN_PAGE_ERROR              ) { wsprintf(txt, "Unexpected exception EXCEPTION_IN_PAGE_ERROR            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_ILLEGAL_INSTRUCTION        ) { wsprintf(txt, "Unexpected exception EXCEPTION_ILLEGAL_INSTRUCTION      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_NONCONTINUABLE_EXCEPTION   ) { wsprintf(txt, "Unexpected exception EXCEPTION_NONCONTINUABLE_EXCEPTION , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_STACK_OVERFLOW             ) { wsprintf(txt, "Unexpected exception EXCEPTION_STACK_OVERFLOW           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INVALID_DISPOSITION        ) { wsprintf(txt, "Unexpected exception EXCEPTION_INVALID_DISPOSITION      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_GUARD_PAGE                 ) { wsprintf(txt, "Unexpected exception EXCEPTION_GUARD_PAGE               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }
           else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_INVALID_HANDLE             ) { wsprintf(txt, "Unexpected exception EXCEPTION_INVALID_HANDLE           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }


           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_WAIT_0                        ) { wsprintf(txt, "Unexpected exception STATUS_WAIT_0                      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ABANDONED_WAIT_0              ) { wsprintf(txt, "Unexpected exception STATUS_ABANDONED_WAIT_0            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_USER_APC                      ) { wsprintf(txt, "Unexpected exception STATUS_USER_APC                    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_TIMEOUT                       ) { wsprintf(txt, "Unexpected exception STATUS_TIMEOUT                     , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_PENDING                       ) { wsprintf(txt, "Unexpected exception STATUS_PENDING                     , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_EXCEPTION_HANDLED                ) { wsprintf(txt, "Unexpected exception DBG_EXCEPTION_HANDLED              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTINUE                         ) { wsprintf(txt, "Unexpected exception DBG_CONTINUE                       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SEGMENT_NOTIFICATION          ) { wsprintf(txt, "Unexpected exception STATUS_SEGMENT_NOTIFICATION        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FATAL_APP_EXIT                ) { wsprintf(txt, "Unexpected exception STATUS_FATAL_APP_EXIT              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_TERMINATE_THREAD                 ) { wsprintf(txt, "Unexpected exception DBG_TERMINATE_THREAD               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_TERMINATE_PROCESS                ) { wsprintf(txt, "Unexpected exception DBG_TERMINATE_PROCESS              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTROL_C                        ) { wsprintf(txt, "Unexpected exception DBG_CONTROL_C                      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_PRINTEXCEPTION_C                 ) { wsprintf(txt, "Unexpected exception DBG_PRINTEXCEPTION_C               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_RIPEXCEPTION                     ) { wsprintf(txt, "Unexpected exception DBG_RIPEXCEPTION                   , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                       
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_CONTROL_BREAK                    ) { wsprintf(txt, "Unexpected exception DBG_CONTROL_BREAK                  , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_COMMAND_EXCEPTION                ) { wsprintf(txt, "Unexpected exception DBG_COMMAND_EXCEPTION              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_GUARD_PAGE_VIOLATION          ) { wsprintf(txt, "Unexpected exception STATUS_GUARD_PAGE_VIOLATION        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DATATYPE_MISALIGNMENT         ) { wsprintf(txt, "Unexpected exception STATUS_DATATYPE_MISALIGNMENT       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_BREAKPOINT                    ) { wsprintf(txt, "Unexpected exception STATUS_BREAKPOINT                  , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SINGLE_STEP                   ) { wsprintf(txt, "Unexpected exception STATUS_SINGLE_STEP                 , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                         
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_LONGJUMP                      ) { wsprintf(txt, "Unexpected exception STATUS_LONGJUMP                    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_UNWIND_CONSOLIDATE            ) { wsprintf(txt, "Unexpected exception STATUS_UNWIND_CONSOLIDATE          , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                
           else if (exPtr->ExceptionRecord->ExceptionCode == DBG_EXCEPTION_NOT_HANDLED            ) { wsprintf(txt, "Unexpected exception DBG_EXCEPTION_NOT_HANDLED          , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ACCESS_VIOLATION              ) { wsprintf(txt, "Unexpected exception STATUS_ACCESS_VIOLATION            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_IN_PAGE_ERROR                 ) { wsprintf(txt, "Unexpected exception STATUS_IN_PAGE_ERROR               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_HANDLE                ) { wsprintf(txt, "Unexpected exception STATUS_INVALID_HANDLE              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_PARAMETER             ) { wsprintf(txt, "Unexpected exception STATUS_INVALID_PARAMETER           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_NO_MEMORY                     ) { wsprintf(txt, "Unexpected exception STATUS_NO_MEMORY                   , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                       
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ILLEGAL_INSTRUCTION           ) { wsprintf(txt, "Unexpected exception STATUS_ILLEGAL_INSTRUCTION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_NONCONTINUABLE_EXCEPTION      ) { wsprintf(txt, "Unexpected exception STATUS_NONCONTINUABLE_EXCEPTION    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                      
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_DISPOSITION           ) { wsprintf(txt, "Unexpected exception STATUS_INVALID_DISPOSITION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ARRAY_BOUNDS_EXCEEDED         ) { wsprintf(txt, "Unexpected exception STATUS_ARRAY_BOUNDS_EXCEEDED       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_DENORMAL_OPERAND        ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_DENORMAL_OPERAND      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_DIVIDE_BY_ZERO          ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_DIVIDE_BY_ZERO        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_INEXACT_RESULT          ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_INEXACT_RESULT        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_INVALID_OPERATION       ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_INVALID_OPERATION     , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                     
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_OVERFLOW                ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_OVERFLOW              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_STACK_CHECK             ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_STACK_CHECK           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_UNDERFLOW               ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_UNDERFLOW             , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                             
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INTEGER_DIVIDE_BY_ZERO        ) { wsprintf(txt, "Unexpected exception STATUS_INTEGER_DIVIDE_BY_ZERO      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INTEGER_OVERFLOW              ) { wsprintf(txt, "Unexpected exception STATUS_INTEGER_OVERFLOW            , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                              
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_PRIVILEGED_INSTRUCTION        ) { wsprintf(txt, "Unexpected exception STATUS_PRIVILEGED_INSTRUCTION      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_STACK_OVERFLOW                ) { wsprintf(txt, "Unexpected exception STATUS_STACK_OVERFLOW              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DLL_NOT_FOUND                 ) { wsprintf(txt, "Unexpected exception STATUS_DLL_NOT_FOUND               , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                           
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ORDINAL_NOT_FOUND             ) { wsprintf(txt, "Unexpected exception STATUS_ORDINAL_NOT_FOUND           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ENTRYPOINT_NOT_FOUND          ) { wsprintf(txt, "Unexpected exception STATUS_ENTRYPOINT_NOT_FOUND        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_CONTROL_C_EXIT                ) { wsprintf(txt, "Unexpected exception STATUS_CONTROL_C_EXIT              , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                            
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_DLL_INIT_FAILED               ) { wsprintf(txt, "Unexpected exception STATUS_DLL_INIT_FAILED             , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                             
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_MULTIPLE_FAULTS         ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_MULTIPLE_FAULTS       , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                   
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_FLOAT_MULTIPLE_TRAPS          ) { wsprintf(txt, "Unexpected exception STATUS_FLOAT_MULTIPLE_TRAPS        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_REG_NAT_CONSUMPTION           ) { wsprintf(txt, "Unexpected exception STATUS_REG_NAT_CONSUMPTION         , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                 
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_STACK_BUFFER_OVERRUN          ) { wsprintf(txt, "Unexpected exception STATUS_STACK_BUFFER_OVERRUN        , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                  
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_INVALID_CRUNTIME_PARAMETER    ) { wsprintf(txt, "Unexpected exception STATUS_INVALID_CRUNTIME_PARAMETER  , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                        
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_ASSERTION_FAILURE             ) { wsprintf(txt, "Unexpected exception STATUS_ASSERTION_FAILURE           , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                               
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SXS_EARLY_DEACTIVATION        ) { wsprintf(txt, "Unexpected exception STATUS_SXS_EARLY_DEACTIVATION      , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                    
           else if (exPtr->ExceptionRecord->ExceptionCode == STATUS_SXS_INVALID_DEACTIVATION      ) { wsprintf(txt, "Unexpected exception STATUS_SXS_INVALID_DEACTIVATION    , addr = %d", exPtr->ContextRecord->Eip); msg_2(txt); }                                                      

//         else if (exPtr->ExceptionRecord->ExceptionCode == EXCEPTION_POSSIBLE_DEADLOCK          ) MessageBox(0, "Unexpected exception EXCEPTION_POSSIBLE_DEADLOCK        ",  0, 0);
           else {

               TQ84_DEBUG("Unexpected exception with code %d at address %d", exPtr->ExceptionRecord->ExceptionCode, exPtr->ContextRecord->Eip);
//             wsprintf(txt, "Unexpected exception with code %d at address %d", exPtr->ExceptionRecord->ExceptionCode, exPtr->ContextRecord->Eip);
//             TQ84_DEBUG(txt);
//             msg_2(txt);
//             callCallback(txt);

           }

     //    SetThreadContext(GetCurrentThread(), exPtr->ContextRecord);
           return EXCEPTION_CONTINUE_SEARCH;
    } // }


    return EXCEPTION_CONTINUE_SEARCH;
} // }

#define TQ84_HOOK_FUNCTION(API_name, module_name )                                                     \
     orig_ ## API_name = (fn_ ## API_name) GetProcAddress(GetModuleHandle( #module_name ), #API_name); \
     if (! orig_ ## API_name) {                                                                            \
       MessageBox(0, "GetProcAddress: " #API_name ", " #module_name, 0, 0);                                \
     }                                                                                                     \
     else {                                                                                                \
       if (! Mhook_SetHook((PVOID*) &orig_ ## API_name, (PVOID) hook_ ## API_name )) {                     \
         MessageBox(0, "Sorry, could not hook " #API_name, 0, 0);                                          \
       }                                                                                                   \
     }

FARPROC WINAPI hook_GetProcAddress(HMODULE hModule, LPCSTR lpProcName);

__declspec(dllexport) void __stdcall VBAInternalsInit(addrCallBack_t addrCallBack_) { // {

    if (!initialized) {
       TQ84_DEBUG_OPEN("c:\\temp\\vba-dbg.out", "w");
       initialized ++;
       curProcess = GetCurrentProcess();

    AddVectoredExceptionHandler(1, VectoredHandler);

    TQ84_DEBUG("Address of VBAInternalsInit            = %d", VBAInternalsInit           );
    TQ84_DEBUG("Address of VectoredHandler             = %d", VectoredHandler            );
    TQ84_DEBUG("Address of compareBreakpoints          = %d", compareBreakpoints         );
    TQ84_DEBUG("Address of callCallback                = %d", callCallback               );
    TQ84_DEBUG("Address of hook_GetProcAddress         = %d", hook_GetProcAddress        );

     { TQ84_DEBUG_INDENT_T("Initializing hooks");

// TODO: uncoment me           TQ84_HOOK_FUNCTION(rtcErrObj                , VBE7.dll    )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(ChooseColorA             , Comdlg32.dll)
// TODO: uncoment me           TQ84_HOOK_FUNCTION(GetFileVersionInfoA      , version.dll )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(GetFileVersionInfoSizeA  , version.dll )
// TODO: uncoment me// !!!!    TQ84_HOOK_FUNCTION(GetProcAddress           , Kernel32.dll)
// TODO: uncoment me           TQ84_HOOK_FUNCTION(GetModuleHandleA         , Kernel32.dll)
                               TQ84_HOOK_FUNCTION(LoadLibraryA             , Kernel32.dll)
                               TQ84_HOOK_FUNCTION(MapAndLoad               , Imagehlp.dll)
// TODO: uncoment me   
// TODO: uncoment me           TQ84_HOOK_FUNCTION(RegCloseKey              , Advapi32.dll )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(RegOpenKeyExW            , Advapi32.dll )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(RegOpenKeyExA            , Advapi32.dll )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(RegQueryValueExA         , Advapi32.dll )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(RegQueryValueExW         , Advapi32.dll )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(RegSetValueExA           , Advapi32.dll )
                               TQ84_HOOK_FUNCTION(RtlUnwind                , Kernel32.dll )
// TODO: uncoment me   
// TODO: uncoment me   //      TQ84_HOOK_FUNCTION(SHGetFolderPathW         , Shell32.dll  )                                 // TODO !!!
// TODO: uncoment me   
// TODO: uncoment me           TQ84_HOOK_FUNCTION(ShellExecuteA            , Shell32.dll  )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(ShellExecuteExW          , Shell32.dll  )
// TODO: uncoment me   
                               TQ84_HOOK_FUNCTION(UnMapAndLoad             , Imagehlp.dll)
// TODO: uncoment me   //      TQ84_HOOK_FUNCTION(VerQueryValue            , Api-ms-win-core-version-l1-1-0.dll)
// TODO: uncoment me           TQ84_HOOK_FUNCTION(VerQueryValueA           , version.dll )
// TODO: uncoment me           TQ84_HOOK_FUNCTION(VirtualAlloc             , Kernel32.dll)
// TODO: uncoment me   //      TQ84_HOOK_FUNCTION(WideCharToMultiByte      , Kernel32.dll)
                               TQ84_HOOK_FUNCTION(WriteProcessMemory       , Kernel32.dll)
   


          orig_rtcErrObj = (fn_rtcErrObj) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcErrObj");
          TQ84_DEBUG("orig_rtcErrObj = %d", orig_rtcErrObj);
   
          if (! Mhook_SetHook((PVOID*) &orig_rtcErrObj, (PVOID) hook_rtcErrObj)) {
            TQ84_DEBUG("Sorry, could not hook rtcErrObj");
          }
          TQ84_DEBUG("orig_rtcErrObj = %d", orig_rtcErrObj);
   


     }

    }

return; /* HERE HERE HERE */

    TQ84_DEBUG_INDENT_T("VBAInternalsInit");

    int i, res;

    char txt[400];

//  wsprintf(txt, "VBAInternalsInit, sizeof(address) = %d", sizeof(address));
//  msg(txt);

    addrCallBack = addrCallBack_;
    callCallback("addrCallBack assigned");

//  AddVectoredExceptionHandler(1, VectoredHandler);

//  msg("VBAInternalsInit leaving");

//  func_addrs[0] = (char*) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcMsgBox");
//  func_addrs[1] = (char*) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcRound" );

 // func_addrs[2] = (char*) func_2;
 // func_addrs[3] = (char*) func_3;

#ifdef USE_SEARCH
#else
    for (i=0; i<NOF_BREAKPOINTS; i++) {
       callCallback("Breakpoint");
       TQ84_DEBUG_INDENT_T("i = %d", i);
     //
     // Setting the breakpoints at the functions addresses.
     //
     // First, we need to make the code segment writable to be able to
     // insert the breakpoint instruction. Otherwise, the modification of
     // the (read-only) code segment would cause an exception.
     //
        DWORD oldProtection;
        VirtualProtect(func_addrs[i], 1, PAGE_EXECUTE_READWRITE, &oldProtection);

     //
     // Inject the int-3 instruction (0xcc):
     // 
     // We also want to store the value of the byte before we set the int-3
     // instruction:
     //
        old_instr[i] = replaceInstruction(func_addrs[i], 0xcc);

    }
#endif

} // }

__declspec(dllexport) void __stdcall dbg(char *txt) { // {
  TQ84_DEBUG("dbg = %s", txt);
} // }


funcPtr_IUnknown_QueryInterface m_loader_queryInterface;
// funcPtr_IUnknown_AddRef         m_loader_AddRef;
// funcPtr_IUnknown_Release        m_loader_Release;

// fn_GetProcAddress               REAL_orig_GetProcAddress;

// fn_WideCharToMultiByte          orig_WideCharToMultiByte;

// HRESULT STDMETHODCALLTYPE hook_AddRef(void *self) {
// 
//   TQ84_DEBUG_INDENT_T("hook_AddRef, self = %d", self);
//   HRESULT ret = m_loader_AddRef(self);
//   TQ84_DEBUG("ret = %d", ret);
//   return ret;
// 
// }
// 
// HRESULT STDMETHODCALLTYPE hook_Release(void *self) {
// 
//   TQ84_DEBUG_INDENT_T("hook_Release, self = %d", self);
//   HRESULT ret = m_loader_Release(self);
//   TQ84_DEBUG("ret = %d", ret);
//   return ret;
// 
// }

funcPtr_IUnknown_QueryInterface orig_rootObject_QueryInterface;
HRESULT STDMETHODCALLTYPE hook_rootObject_QueryInterface(void *self, REFIID riid, void **pObj) { // {
  TQ84_DEBUG_INDENT_T("hook_rootObject_QueryInterface, self = %d", self);
  LPOLESTR iid_string;
  StringFromIID(riid, &iid_string);

  char x[500];
  wcstombs(x, iid_string, 499);
  TQ84_DEBUG("iid = %s", x);

  CoTaskMemFree(iid_string);

  HRESULT ret = orig_rootObject_QueryInterface(self, riid, pObj);
  TQ84_DEBUG("ret = %d", ret);
  TQ84_DEBUG("pObj = %d", *pObj);

  return ret;

} // }

funcPtr_IUnknown_QueryInterface orig_classFactory_QueryInterface;
HRESULT STDMETHODCALLTYPE hook_classFactory_QueryInterface(void *self, REFIID riid, void **pObj) { // {
  TQ84_DEBUG_INDENT_T("hook_classFactory_QueryInterface, self = %d", self);
  LPOLESTR iid_string;
  StringFromIID(riid, &iid_string);

  char x[500];
  wcstombs(x, iid_string, 499);
  TQ84_DEBUG("iid = %s", x);

  CoTaskMemFree(iid_string);

  HRESULT ret = orig_classFactory_QueryInterface(self, riid, pObj);
  TQ84_DEBUG("ret = %d", ret);
  TQ84_DEBUG("pObj = %d", *pObj);

  return ret;

} // }

funcPtr_IDispatch_GetIDsOfNames orig_classFactory_GetIDsOfNames;
HRESULT STDMETHODCALLTYPE hook_classFactory_GetIDsOfNames(void *self, REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) { // {
  int i;
    char x[500];
  TQ84_DEBUG_INDENT_T("hook_classFactory_GetIDsOfNames, cNames = %d, lcid = %d", cNames, lcid);


  HRESULT ret = orig_classFactory_GetIDsOfNames(self, riid, rgszNames, cNames, lcid, rgDispId);

  for (i=0; i<cNames; i++) {
    wcstombs(x, rgszNames[i], 499);
    TQ84_DEBUG("%d -> %s, dispid: %d", i, x, rgDispId[i]);
  }
  return ret;

} // }

funcPtr_IDispatch_Invoke orig_classFactory_Invoke;
HRESULT STDMETHODCALLTYPE hook_classFactory_Invoke(void *self, DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pvarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) { // {
  TQ84_DEBUG_INDENT_T("hook_classFactory_Invoke, dispidMember = %d, pvarResult = %d", dispidMember, pvarResult);

  HRESULT ret = orig_classFactory_Invoke(self, dispidMember, riid, lcid, wFlags, pDispParams, pvarResult, pExcepInfo, puArgErr);
  TQ84_DEBUG("pvarResult = %d", pvarResult);

  return ret;
} // }


funcPtr_IDispatch_GetIDsOfNames orig_m_VCOMObject_GetIDsOfNames;
HRESULT STDMETHODCALLTYPE hook_m_VCOMObject_GetIDsOfNames(void *self, REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) { // {
  int i;
    char x[500];
  TQ84_DEBUG_INDENT_T("hook_m_VCOMObject_GetIDsOfNames, cNames = %d, lcid = %d", cNames, lcid);


  HRESULT ret = orig_m_VCOMObject_GetIDsOfNames(self, riid, rgszNames, cNames, lcid, rgDispId);

  for (i=0; i<cNames; i++) {
    wcstombs(x, rgszNames[i], 499);
    TQ84_DEBUG("%d -> %s, dispid: %d", i, x, rgDispId[i]);
  }
  return ret;

} // }

funcPtr_IDispatch_Invoke orig_m_VCOMObject_Invoke;
HRESULT STDMETHODCALLTYPE hook_m_VCOMObject_Invoke(void *self, DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pvarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) { // {
  TQ84_DEBUG_INDENT_T("hook_m_VCOMObject_Invoke, dispidMember = %d, pvarResult = %d", dispidMember, pvarResult);

  HRESULT ret = orig_m_VCOMObject_Invoke(self, dispidMember, riid, lcid, wFlags, pDispParams, pvarResult, pExcepInfo, puArgErr);
  TQ84_DEBUG("pvarResult = %d", pvarResult);

  return ret;
} // }

VBAObject_withDisp *m_VCOMObject;

__declspec(dllexport) void __stdcall set_m_VCOMObject(VBAObject_withDisp **obj) { // {
  TQ84_DEBUG_INDENT_T("set_m_VCOMObject, obj = %d");
  TQ84_DEBUG("*obj = %d", *obj);

  m_VCOMObject = *obj;

  TQ84_DEBUG("m_VCOMObject->vtbl = %d", m_VCOMObject->vtbl);

  TQ84_DEBUG("m_VCOMObject->vtbl->QueryInterface   = %d",  m_VCOMObject->vtbl->QueryInterface  );
  TQ84_DEBUG("m_VCOMObject->vtbl->AddRef           = %d",  m_VCOMObject->vtbl->AddRef          );
  TQ84_DEBUG("m_VCOMObject->vtbl->Release          = %d",  m_VCOMObject->vtbl->Release         );
  TQ84_DEBUG("m_VCOMObject->vtbl->GetTypeInfoCount = %d",  m_VCOMObject->vtbl->GetTypeInfoCount);
  TQ84_DEBUG("m_VCOMObject->vtbl->GetTypeInfo      = %d",  m_VCOMObject->vtbl->GetTypeInfo     );
  TQ84_DEBUG("m_VCOMObject->vtbl->GetIDsOfNames    = %d",  m_VCOMObject->vtbl->GetIDsOfNames   );
  TQ84_DEBUG("m_VCOMObject->vtbl->Invoke           = %d",  m_VCOMObject->vtbl->Invoke          );


  orig_m_VCOMObject_GetIDsOfNames = m_VCOMObject->vtbl->GetIDsOfNames;
  if (! Mhook_SetHook((PVOID*) &orig_m_VCOMObject_GetIDsOfNames, hook_m_VCOMObject_GetIDsOfNames)) {
       MessageBox(0, "Sorry, could not hook hook_m_VCOMObject_GetIDsOfNames", 0, 0);
  }

  orig_m_VCOMObject_Invoke = m_VCOMObject->vtbl->Invoke;
  if (! Mhook_SetHook((PVOID*) &orig_m_VCOMObject_Invoke, hook_m_VCOMObject_Invoke)) {
       MessageBox(0, "Sorry, could not hook hook_m_VCOMObject_Invoke", 0, 0);
  }

} // }

HRESULT STDMETHODCALLTYPE hook_QueryInterface(void *self, REFIID riid, void **pObj) { // {
  TQ84_DEBUG_INDENT_T("QueryInterface, self = %d", self);

  TQ84_DEBUG("m_loader_queryInterface = %d", m_loader_queryInterface);

  LPOLESTR iid_string;
  StringFromIID(riid, &iid_string);

  char x[500];
  wcstombs(x, iid_string, 499);
  TQ84_DEBUG("iid = %s", x);

  CoTaskMemFree(iid_string);


  TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->QueryInterface   = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->QueryInterface  );
  HRESULT ret = m_loader_queryInterface(self, riid, pObj);

  TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->QueryInterface   = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->QueryInterface  );
  TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->AddRef           = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->AddRef          );
  TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->Release          = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->Release         );
  TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->GetTypeInfoCount = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->GetTypeInfoCount);
  TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->GetTypeInfo      = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->GetTypeInfo     );
  TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->GetIDsOfNames    = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->GetIDsOfNames   );
  TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->Invoke           = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->Invoke          );

  orig_rootObject_QueryInterface = m_loader->rootObject->vtbl->QueryInterface;
  if (! Mhook_SetHook((PVOID*) &orig_rootObject_QueryInterface, hook_rootObject_QueryInterface)) {
       MessageBox(0, "Sorry, could not hook hook_rootObject_QueryInterface", 0, 0);
  }

  TQ84_DEBUG("ret = %d", ret);
  TQ84_DEBUG("pObj = %d", *pObj);

  return ret;
} // }

__declspec(dllexport) void __stdcall addrOf_m_Loader(void *addr) { // {
    TQ84_DEBUG_INDENT_T("addrOf_m_Loader = %d", addr);

 // Sanity check

    m_loader = (VCOMInitializerStruct*) addr;

    HANDLE h = GetModuleHandle("kernel32.dll");


    TQ84_DEBUG("m_loader.kernel32Handle = %d, GetModuleHandle() = %d", m_loader->kernel32Handle, h);


    TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->QueryInterface = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->QueryInterface);
    TQ84_DEBUG("Trying to hook");

    m_loader_queryInterface = m_loader->rootObject->vtbl->QueryInterface;
//  m_loader_AddRef         = m_loader->rootObject->vtbl->AddRef;
//  m_loader_Release        = m_loader->rootObject->vtbl->Release;

    TQ84_DEBUG("Before hooking hook_QueryInterface");
    if (! Mhook_SetHook((PVOID*) &m_loader_queryInterface, hook_QueryInterface)) {
         MessageBox(0, "Sorry, could not hook", 0, 0);
    }
//  TQ84_DEBUG("Before hooking hook_AddRef");
//  if (! Mhook_SetHook((PVOID*) &m_loader_AddRef, hook_AddRef)) {
//       MessageBox(0, "Sorry, could not hook", 0, 0);
//  }
//  TQ84_DEBUG("Before hooking hook_Release");
//  if (! Mhook_SetHook((PVOID*) &m_loader_Release, hook_Release)) {
//       MessageBox(0, "Sorry, could not hook", 0, 0);
//  }

    orig_GetProcAddress      = m_loader->GetProcAddress;
//  REAL_orig_GetProcAddress = orig_GetProcAddress;

    TQ84_DEBUG("before hooking GetProcAddress: orig_GetProcAddress = %d", orig_GetProcAddress);
    if (! Mhook_SetHook((PVOID*) &orig_GetProcAddress, hook_GetProcAddress)) {
//  if (! Mhook_SetHook((PVOID*) &(m_loader->GetProcAddress), hook_GetProcAddress)) {
         MessageBox(0, "Sorry, could not hook GetProcAddress", 0, 0);
    }
    TQ84_DEBUG("after hooking GetProcAddress: orig_GetProcAddress = %d", orig_GetProcAddress);


//  orig_SysFreeString = m_loader->SysFreeString;
//  if (! Mhook_SetHook((PVOID*) &orig_SysFreeString, hook_SysFreeString)) {
//       MessageBox(0, "Sorry, could not hook SysFreeString", 0, 0);
//  }
    

//  m_loader_queryInterface                    = m_loader->rootObject->vtbl->QueryInterface;
//  m_loader->rootObject->vtbl->QueryInterface = QueryInterface;


    if (h != m_loader->kernel32Handle) {
      MessageBox(0, "m_loader.kernel32Handle!", 0, 0);
    }

    if (m_loader->vTablePtr != m_loader) {
      MessageBox(0, "m_loader.vTablePtr!", 0, 0);
    }

    if (m_loader->GetProcAddress != GetProcAddress) {
      MessageBox(0, "GetProcAddress (1)!", 0, 0);
    }

//  if (m_loader->GetProcAddress != GetProcAddress(h, "GetProcAddress")) {
//    MessageBox(0, "GetProcAddress (2)!", 0, 0);
//  }

    if (m_loader->SysFreeString != GetProcAddress(GetModuleHandle("oleaut32.dll"), "SysFreeString")) {
      MessageBox(0, "SysFreeString!", 0, 0);
    }


    TQ84_DEBUG("m_loader->iDispatch.QueryInterface       = %d", m_loader->iDispatch.QueryInterface);
    TQ84_DEBUG("m_loader->iDispatch.AddRef               = %d", m_loader->iDispatch.AddRef        );
    TQ84_DEBUG("m_loader->iDispatch.Release              = %d", m_loader->iDispatch.Release       );
    TQ84_DEBUG("m_loader->iDispatch.GetTypeInfoCount     = %d", m_loader->iDispatch.GetTypeInfoCount );
    TQ84_DEBUG("m_loader->iDispatch.GetTypeInfo          = %d", m_loader->iDispatch.GetTypeInfo      );
    TQ84_DEBUG("m_loader->iDispatch.GetIDsOfNames        = %d", m_loader->iDispatch.GetIDsOfNames );
    TQ84_DEBUG("m_loader->iDispatch.Invoke               = %d", m_loader->iDispatch.Invoke        );

    TQ84_DEBUG("m_loader->SysFreeString                  = %d", m_loader->SysFreeString);
    TQ84_DEBUG("m_loader->WideCharToMultiByte            = %d", m_loader->WideCharToMultiByte);
    TQ84_DEBUG("m_loader->GetProcAddress                 = %d", m_loader->GetProcAddress);

//  TQ84_DEBUG("1st element (* ) in m_loader: %d", *  ((int* ) addr));
//  TQ84_DEBUG("1st element (**) in m_loader: %d", ** ((int**) addr));

    TQ84_DEBUG("nativeCode (at %d)", m_loader->nativeCode);
    TQ84_DEBUG("loaderMem  = %d", m_loader->loaderMem);
    TQ84_DEBUG("vTablePtr  = %d", m_loader->vTablePtr  );
    TQ84_DEBUG("rootObject = %d", m_loader->rootObject  );
    TQ84_DEBUG("classFactory = %d", m_loader->classFactory  );


    TQ84_DEBUG("m_loader->rootObject(%d)->vtbl(%d)->QueryInterface = %d", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->QueryInterface);

//  TQ84_DEBUG("& rootObject->QueryInterface = %d", & (m_loader->rootObject->QueryInterface));
 
    if (m_loader->rootObject->vtbl != m_loader->vTablePtr) {
       MessageBox(0, "m_loader->rootObject->vtbl", 0, 0);
    }

    if (& (m_loader->rootObject->vtbl) != m_loader->rootObject) {
       MessageBox(0, "& (m_loader->rootObject->QueryInterface)", 0, 0);

    }

    TQ84_DEBUG("Returning");

} // }

__declspec(dllexport) void __stdcall beforeSettingRootObjectToNothing() { // {

  TQ84_DEBUG_INDENT_T("beforeSettingRootObjectToNothing");


  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->QueryInterface   = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->QueryInterface  );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->AddRef           = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->AddRef          );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->Release          = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->Release         );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->GetTypeInfoCount = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->GetTypeInfoCount);
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->GetTypeInfo      = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->GetTypeInfo     );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->GetIDsOfNames    = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->GetIDsOfNames   );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->Invoke           = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->Invoke          );

  if (m_loader->rootObject->vtbl != m_loader->classFactory->vtbl) {
    MessageBox(0, "m_loader->rootObject->vtbl != m_loader->classFactory->vtbl", "Assertion failed", 0);
    TQ84_DEBUG("m_loader->rootObject(%d)->vtbl (%d) != m_loader->classFactory(%d)->vtbl (%d)", m_loader->rootObject, m_loader->rootObject->vtbl, m_loader->classFactory, m_loader->classFactory->vtbl);
    TQ84_DEBUG("m_loader->rootObject->vtbl(%d)->QueryInterface = %d, m_loader->classFactory->vtbl->QueryInterface = %d", m_loader->rootObject->vtbl, m_loader->rootObject->vtbl->QueryInterface, m_loader->classFactory->vtbl->QueryInterface);
  }
  else {
     TQ84_DEBUG("m_loader->rootObject->vtbl == m_loader->classFactory->vtbl ");
  }

  orig_classFactory_QueryInterface = m_loader->classFactory->vtbl->QueryInterface;
  if (! Mhook_SetHook((PVOID*) &orig_classFactory_QueryInterface, hook_classFactory_QueryInterface)) {
       MessageBox(0, "Sorry, could not hook hook_classFactory_QueryInterface", 0, 0);
  }
  orig_classFactory_GetIDsOfNames = m_loader->classFactory->vtbl->GetIDsOfNames;
  if (! Mhook_SetHook((PVOID*) &orig_classFactory_GetIDsOfNames, hook_classFactory_GetIDsOfNames)) {
       MessageBox(0, "Sorry, could not hook hook_classFactory_GetIDsOfNames", 0, 0);
  }
  orig_classFactory_Invoke = m_loader->classFactory->vtbl->Invoke;
  if (! Mhook_SetHook((PVOID*) &orig_classFactory_Invoke, hook_classFactory_Invoke)) {
       MessageBox(0, "Sorry, could not hook hook_classFactory_Invoke", 0, 0);
  }

} // }

__declspec(dllexport) void __stdcall beforeCallingClassFactoryInit() { // {

  TQ84_DEBUG_INDENT_T("beforeCallingClassFactoryInit");


  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->QueryInterface   = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->QueryInterface  );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->AddRef           = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->AddRef          );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->Release          = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->Release         );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->GetTypeInfoCount = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->GetTypeInfoCount);
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->GetTypeInfo      = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->GetTypeInfo     );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->GetIDsOfNames    = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->GetIDsOfNames   );
  TQ84_DEBUG("m_loader->classFactory(%d)->vtbl(%d)->Invoke           = %d", m_loader->classFactory, m_loader->classFactory->vtbl, m_loader->classFactory->vtbl->Invoke          );

} // }

BOOL WINAPI DllMain( // {
  _In_ HINSTANCE hinstDLL,
  _In_ DWORD     fdwReason,
  _In_ LPVOID    lpvReserved

) {

   int i;

// if (!initialized) {
//    TQ84_DEBUG_OPEN("c:\\temp\\dbg.vba", "w");
//    initialized ++;
// }

// TQ84_DEBUG_INDENT_T("DllMain");

// TQ84_DEBUG("EXCEPTION_BREAKPOINT = %x, EXCEPTION_SINGLE_STEP = %x", EXCEPTION_BREAKPOINT, EXCEPTION_SINGLE_STEP);

// if      (fdwReason == DLL_PROCESS_ATTACH) {    MessageBox(0, "DllMain DLL_PROCESS_ATTACH", 0, 0)    ;}
// else if (fdwReason == DLL_PROCESS_DETACH) {    MessageBox(0, "DllMain DLL_PROCESS_DETACH", 0, 0)    ;}
// else if (fdwReason == DLL_THREAD_ATTACH ) { /* MessageBox(0, "DllMain DLL_THREAD_ATTACH ", 0, 0) */ ;}
// else if (fdwReason == DLL_THREAD_DETACH ) { /* MessageBox(0, "DllMain DLL_THREAD_DETACH ", 0, 0) */ ;}
// else                                      {    MessageBox(0, "DllMain hmmmm???"          , 0, 0)    ;}

   if (fdwReason == DLL_PROCESS_ATTACH) { // {
 //   TQ84_DEBUG("fdwReason == DLL_PROCESS_ATTACH");
 //   curProcess = GetCurrentProcess();
#ifdef USE_SEARCH
#else
      func_addrs[0] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcMsgBox");
      func_addrs[1] = (address) GetProcAddress(GetModuleHandle("VBE7.dll"), "rtcRound" );
#endif

   } // }

   if (fdwReason == DLL_PROCESS_DETACH) { // {
     msg_2("DLL_PROCESS_DETACH")
//    TQ84_DEBUG("DLL_PROCESS_DETACH");
#ifdef USE_SEARCH
#else
      for (i=0; i<NOF_BREAKPOINTS; i++) {
//      TQ84_DEBUG("going to call replaceInstruction");
        replaceInstruction(func_addrs[i], old_instr[i]);
      }
#endif
   } // }

   if (fdwReason == DLL_THREAD_DETACH) { // {
//    TQ84_DEBUG("DLL_THREAD_DETACH");
   } // }

   if (fdwReason == DLL_THREAD_ATTACH) { // {
//    TQ84_DEBUG("DLL_THREAD_ATTACH");
   } // }

   return TRUE;

} // }
