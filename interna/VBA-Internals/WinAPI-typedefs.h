
typedef BOOL      (WINAPI *fn_ChooseColorA           )(LPCHOOSECOLOR lpcc                                                                                   ); fn_ChooseColorA            orig_ChooseColorA       ;
typedef BOOL      (WINAPI *fn_GetFileVersionInfoA    )(LPCSTR lptstrFilename, DWORD dwHandle, DWORD  dwLen, LPVOID lpData                                   ); fn_GetFileVersionInfoA     orig_GetFileVersionInfoA;
typedef DWORD     (WINAPI *fn_GetFileVersionInfoSizeA)(LPCSTR lptstrFilename, LPDWORD lpdwHandle                                                            ); fn_GetFileVersionInfoSizeA orig_GetFileVersionInfoSizeA;
typedef FARPROC   (WINAPI *fn_GetProcAddress         )(HMODULE hModule, LPCSTR lpProcName                                                                   ); fn_GetProcAddress          orig_GetProcAddress         ;
typedef HMODULE   (WINAPI *fn_GetModuleHandleA       )(LPCSTR lpModuleName                                                                                  ); fn_GetModuleHandleA        orig_GetModuleHandleA       ;

typedef HMODULE   (WINAPI *fn_LoadLibraryA           )(LPCSTR lpLibName                                                                                     ); fn_LoadLibraryA            orig_LoadLibraryA           ;

typedef int       (WINAPI *fn_MessageBoxA            )(HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType                                               ); fn_MessageBoxA             orig_MessageBoxA            ;

// typedef BOOL      (WINAPI *fn_MapAndLoad             )(PCSTR ImageName,PCSTR DllPath,PLOADED_IMAGE LoadedImage,WINBOOL DotDll,WINBOOL ReadOnly              ); fn_MapAndLoad              orig_MapAndLoad             ;
   typedef BOOL      (WINAPI *fn_MapAndLoad             )(PCSTR ImageName,PCSTR DllPath,PLOADED_IMAGE LoadedImage,   BOOL DotDll,   BOOL ReadOnly              ); fn_MapAndLoad              orig_MapAndLoad             ;

typedef LSTATUS   (WINAPI *fn_RegCloseKey            )( HKEY hKey                                                                                           ); fn_RegCloseKey             orig_RegCloseKey            ;
typedef LSTATUS   (WINAPI *fn_RegOpenKeyExW          )( HKEY hKey, LPCWSTR lpSubKey, DWORD   ulOptions, REGSAM  samDesired, PHKEY   phkResult)               ; fn_RegOpenKeyExW           orig_RegOpenKeyExW          ;
typedef LSTATUS   (WINAPI *fn_RegOpenKeyExA          )( HKEY hKey, LPCSTR lpSubKey, DWORD  ulOptions, REGSAM samDesired, PHKEY  phkResult)                   ; fn_RegOpenKeyExA           orig_RegOpenKeyExA          ;
typedef LSTATUS   (WINAPI *fn_RegQueryValueExA       )( HKEY hKey, LPCSTR  lpValueName, LPDWORD lpReserved, LPDWORD lpType, LPBYTE  lpData, LPDWORD lpcbData                                                      ); fn_RegQueryValueExA        orig_RegQueryValueExA       ;
typedef LSTATUS   (WINAPI *fn_RegQueryValueExW       )( HKEY hKey, LPCWSTR lpValueName, LPDWORD lpReserved, LPDWORD lpType, LPBYTE  lpData, LPDWORD lpcbData                                                      ); fn_RegQueryValueExW        orig_RegQueryValueExW       ;
typedef LSTATUS   (WINAPI *fn_RegSetValueExA         )( HKEY hKey, LPCSTR lpValueName, DWORD Reserved, DWORD dwType, const BYTE *lpData, DWORD cbData                                                             ); fn_RegSetValueExA          orig_RegSetValueExA         ;

/* NTSYSAPI ? */
typedef VOID      (WINAPI *fn_RtlUnwind            )(PVOID TargetFrame, PVOID TargetIp, PEXCEPTION_RECORD ExceptionRecord, PVOID ReturnValue                                                                    ); fn_RtlUnwind               orig_RtlUnwind             ;

typedef HINSTANCE (WINAPI *fn_ShellExecuteA          )( HWND hwnd, LPCSTR lpOperation, LPCSTR lpFile, LPCSTR lpParameters, LPCSTR lpDirectory, INT nShowCmd                                                       ); fn_ShellExecuteA           orig_ShellExecuteA          ;
typedef BOOL      (WINAPI *fn_ShellExecuteExW        )( SHELLEXECUTEINFOW *pExecInfo                                                                                                                              ); fn_ShellExecuteExW         orig_ShellExecuteExW        ;

typedef HRESULT   (WINAPI *fn_SHGetFolderPathW       )( HWND hwnd, int csidl, HANDLE hToken, DWORD dwFlags, LPWSTR pszPath                                                                                        ); fn_SHGetFolderPathW        orig_SHGetFolderPathW       ;

typedef void      (WINAPI *fn_SysFreeString          )(BSTR bstrString                                                                                                                                            ); fn_SysFreeString           orig_SysFreeString;

typedef BOOL      (WINAPI *fn_UnMapAndLoad           )(PLOADED_IMAGE LoadedImage                                                                                                                                  ); fn_UnMapAndLoad            orig_UnMapAndLoad           ;
typedef int       (WINAPI *fn_WideCharToMultiByte    )( UINT CodePage, DWORD dwFlags, LPCWCH lpWideCharStr, int cchWideChar, LPSTR lpMultiByteStr, int cbMultiByte, LPCCH lpDefaultChar, LPBOOL lpUsedDefaultChar ); fn_WideCharToMultiByte     orig_WideCharToMultiByte    ;
typedef BOOL      (WINAPI *fn_VerQueryValueA         )(LPCVOID pBlock, LPCSTR lpSubBlock, LPVOID *lplpBuffer, PUINT puLen                                                                                         ); fn_VerQueryValueA          orig_VerQueryValueA         ;
typedef LPVOID    (WINAPI *fn_VirtualAlloc           )(LPVOID lpAddress, SIZE_T dwSize, DWORD  flAllocationType, DWORD flProtect                                                                                  ); fn_VirtualAlloc            orig_VirtualAlloc           ;

typedef BOOL (WINAPI *fn_WriteProcessMemory )(HANDLE hProcess, LPVOID lpBaseAddress, LPCVOID lpBuffer, SIZE_T nSize, SIZE_T *lpNumberOfBytesWritten                                                               ); fn_WriteProcessMemory      orig_WriteProcessMemory     ;
