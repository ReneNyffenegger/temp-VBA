#include <windows.h>
#include <dbghelp.h>

#include <stdio.h>

#include "mhook.h"

#include "../WinAPI-typedefs.h"

//
//    https://stackoverflow.com/a/45061320/180275
//


HMODULE WINAPI hook_GetModuleHandleA(LPCSTR lpModuleName) {

  HMODULE ret = orig_GetModuleHandleA(lpModuleName);

  printf("GetModuleHandleA, lpModuleName = %s, ret = %d\n", lpModuleName, ret);

  return ret;
}

FARPROC WINAPI hook_GetProcAddress(HMODULE hModule, LPCSTR lpProcName) {

  FARPROC ret = orig_GetProcAddress(hModule, lpProcName);

  printf("GetProcAddress, hModule = %d, lpProcName = %s, ret = %d\n", hModule, lpProcName, ret);

  return ret;
}

int WINAPI hook_MessageBoxA(HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType) { 
    printf("MessageBox lpText = %s, lpCaption = %s\n", lpText, lpCaption);

    return orig_MessageBoxA(hWnd, lpText, lpCaption, uType);
}



int main() {

    orig_GetModuleHandleA = (fn_GetModuleHandleA) GetProcAddress(GetModuleHandleA("Kernel32.dll"), "GetModuleHandleA"); 

    if (!orig_GetModuleHandleA) { MessageBoxA(0, "orig_GetModuleHandleA", 0, 0); }

    if (! Mhook_SetHook( (PVOID*) &orig_GetModuleHandleA, hook_GetModuleHandleA)) {
       MessageBox(0, "Could not hook GetModuleHandelA", 0, 0);
    }

 // ------------------------------------------------------------------------------------------------------------

    orig_GetProcAddress = (fn_GetProcAddress) GetProcAddress(GetModuleHandleA("Kernel32.dll"), "GetProcAddress");

    printf("orig_GetProcAddress before hooking: %d\n", orig_GetProcAddress);

    if (!orig_GetProcAddress) { MessageBoxA(0, "orig_GetProcAddress", 0, 0); }

    if (! Mhook_SetHook( (PVOID*) &orig_GetProcAddress, hook_GetProcAddress)) {
       MessageBox(0, "Could not hook GetProcAddress", 0, 0);
    }

    printf("orig_GetProcAddress after hooking: %d\n", orig_GetProcAddress);

 // ------------------------------------------------------------------------------------------------------------

    orig_MessageBoxA = (fn_MessageBoxA) orig_GetProcAddress(GetModuleHandleA("User32.dll"), "MessageBoxA");
    if (!orig_MessageBoxA) { MessageBoxA(0, "orig_MessageBoxA", 0, 0); }

    orig_MessageBoxA(0, "Unhooked message box", 0, 0);

    if (! Mhook_SetHook( (PVOID*) &orig_MessageBoxA, hook_MessageBoxA)) {
       MessageBox(0, "Could not hook MessageBoxA", 0, 0);
    }


    MessageBoxA(0, "Hooked Message Box", "Info", 0);


 // ------------------------------------------------------------------------------------------------------------
 
    fn_GetProcAddress again_GetProcAddress = GetProcAddress(GetModuleHandleA("Kernel32.dll"), "GetProcAddress");
    printf("again_GetProcAddress = %d\n", again_GetProcAddress);

    fn_MessageBoxA    again_MessageBoxA    = GetProcAddress(GetModuleHandleA("User32.dll"  ), "MessageBoxA"   );
    printf("again_MessageBoxA = %d\n", again_MessageBoxA);
    MessageBoxA(0, "again_MessageBoxA", "again", 0);


    Mhook_Unhook(&orig_MessageBoxA     );
    Mhook_Unhook(&orig_GetModuleHandleA);
    Mhook_Unhook(&orig_GetProcAddress  );

}
