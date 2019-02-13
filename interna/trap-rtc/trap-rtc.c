#include "msg.h"
#include "txt.h"
#include "funcsInDll.h"
#include "int3.h"

HANDLE msgFile;

void callbackFuncInDll(char *funcName, DWORD address) {
    writeToFile(txt("callbackFuncInDll received %s @ %d ?", funcName, address));
 // writeToFile(txt("callbackFuncInDll received %s @                       ));


    HANDLE   h = GetModuleHandle("VBE7.dll");
    FARPROC  a = GetProcAddress(h, funcName);
    writeToFile(txt("proc addr: %d\n", a) );

//writeToFile("\n");

}

LONG WINAPI VectoredHandler(PEXCEPTION_POINTERS exPtr) { // {

} // }


__declspec(dllexport) void __stdcall Init () { // {

    createNewOutputFile(TEXT("c:\\temp\\trap-rtc.msg"));

    writeToFile("Initialized\n");

    iterateOverFuncsInDll("VBE7.dll", "C:\\Program Files\\Common Files\\microsoft shared\\VBA\\VBA7", callbackFuncInDll);

} // }

BOOL DllMain(HINSTANCE hInst, DWORD reason, LPVOID reserved) { // {
    return TRUE;
} // }
