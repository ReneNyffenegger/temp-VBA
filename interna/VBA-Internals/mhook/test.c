#include <windows.h>
#include <dbghelp.h>

#include <stdio.h>

#include "mhook.h"

#include "../WinAPI-typedefs.h"

//
//    https://stackoverflow.com/a/45061320/180275
//


HMODULE WINAPI hook_GetModuleHandleA(LPCSTR lpModuleName) {

  printf("GetModuleHandeA, lpModuleName = %s\n", lpModuleName);
  return orig_GetModuleHandleA(lpModuleName);

}

int WINAPI hook_MessageBoxA(HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType) { 
    printf("MessageBox lpText = %s, lpCaption = %s\n", lpText, lpCaption);

    return orig_MessageBoxA(hWnd, lpText, lpCaption, uType);
}



int main() {

    orig_GetModuleHandleA = (fn_GetModuleHandleA) GetProcAddress(GetModuleHandleA("Kernel32.dll"), "GetModuleHandleA"); 
    if (!orig_GetModuleHandleA) { MessageBoxA(0, "orig_GetModuleHandleA", 0, 0); }

    if (! Mhook_SetHook( (PVOID*) &orig_GetModuleHandleA, hook_GetModuleHandleA)) {
       MessageBox(0, "Could not hook GetModuleHandeA", 0, 0);
    }


    orig_GetProcAddress = (fn_GetProcAddress) GetProcAddress(GetModuleHandleA("Kernel32.dll"), "GetProcAddress");

    if (!orig_GetProcAddress) { MessageBoxA(0, "orig_GetProcAddress", 0, 0); }

    orig_MessageBoxA = (fn_MessageBoxA) orig_GetProcAddress(GetModuleHandleA("User32.dll"), "MessageBoxA");
    if (!orig_MessageBoxA) { MessageBoxA(0, "orig_MessageBoxA", 0, 0); }

    orig_MessageBoxA(0, "Unhooked message box", 0, 0);

    if (! Mhook_SetHook( (PVOID*) &orig_MessageBoxA, hook_MessageBoxA)) {
       MessageBox(0, "Could not hook MessageBoxA", 0, 0);
    }

    MessageBoxA(0, "Hooked Message Box", "Info", 0);

    Mhook_Unhook(&orig_MessageBoxA     );
    Mhook_Unhook(&orig_GetModuleHandleA);

}
